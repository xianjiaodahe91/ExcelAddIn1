using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelAddIn1.UI;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private sealed class PaneEntry
        {
            public Excel.Window Window;
            public CustomTaskPane Pane;
            public NavPaneControl Ui;
            public PaneHost Host;
            public string Signature;
            public int Key;
        }

        private readonly Dictionary<int, PaneEntry> _panes = new Dictionary<int, PaneEntry>();
        private bool _busy;
        private bool _suppressStateSave;
        private bool _pendingRestore;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WindowActivate += OnWindowActivate;
            Application.SheetActivate += OnSheetActivate;
            Application.WorkbookNewSheet += OnWorkbookNewSheet;
            Application.WorkbookBeforeClose += OnWorkbookBeforeClose;

            // 利用 Excel 启动的间隙，在后台把 WPF 程序集先映射进进程，
            // 等真正建面板时就不用再走磁盘加载（这是冷启动里最贵的一段）
            PreloadWpfAssemblies();

            // 上次是开着的就恢复。外壳是纯 WinForms，创建几乎零成本，
            // 所以可以直接在启动阶段建，让面板和 Excel 同时出现（WPF 内容稍后再挂）
            if (PaneOptions.LastPaneVisible)
            {
                try
                {
                    var window = SafeActiveWindow();
                    if (window != null) EnsurePane(window);
                    else _pendingRestore = true;
                }
                catch
                {
                    _pendingRestore = true;
                }
            }
        }

        /// <summary>
        /// 后台预热 WPF：只加载程序集（不动任何 UI 对象），线程安全，
        /// 能把面板创建时那 200~400ms 的加载开销提前摊掉。
        /// </summary>
        private static void PreloadWpfAssemblies()
        {
            try
            {
                System.Threading.ThreadPool.QueueUserWorkItem(delegate
                {
                    var names = new[]
                    {
                        "WindowsBase",
                        "PresentationCore",
                        "PresentationFramework",
                        "System.Xaml",
                        "WindowsFormsIntegration"
                    };
                    for (var i = 0; i < names.Length; i++)
                    {
                        try
                        {
                            System.Reflection.Assembly.Load(names[i]);
                        }
                        catch
                        {
                        }
                    }
                });
            }
            catch
            {
            }
        }

        private Excel.Window SafeActiveWindow()
        {
            try
            {
                return Application.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 退出前按面板的实际可见状态落盘（拆解过程中还会触发"已隐藏"，所以随后要屏蔽）
            try
            {
                foreach (var pair in _panes)
                {
                    if (pair.Value.Pane == null) continue;
                    PaneOptions.LastPaneVisible = pair.Value.Pane.Visible;
                    break;
                }
            }
            catch
            {
            }
            _suppressStateSave = true;

            try
            {
                Application.WindowActivate -= OnWindowActivate;
                Application.SheetActivate -= OnSheetActivate;
                Application.WorkbookNewSheet -= OnWorkbookNewSheet;
                Application.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            }
            catch
            {
            }
        }

        // ==================== Excel 事件 ====================

        private void OnWindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            try
            {
                // 启动时没有可用窗口、但需要恢复面板的情况
                if (_pendingRestore)
                {
                    _pendingRestore = false;
                    EnsurePane(wn);
                    return;
                }

                // 只刷新已经打开过面板的窗口，不主动创建
                PaneEntry entry;
                if (TryGetPane(wn, out entry)) Refresh(entry, false);
            }
            catch
            {
            }
        }

        private void OnSheetActivate(object sh)
        {
            try
            {
                PaneEntry entry;
                if (TryGetPane(Application.ActiveWindow, out entry)) Refresh(entry, false);
            }
            catch
            {
            }
        }

        private void OnWorkbookNewSheet(Excel.Workbook wb, object sh)
        {
            try
            {
                PaneEntry entry;
                if (TryGetPane(Application.ActiveWindow, out entry)) Refresh(entry, true);
            }
            catch
            {
            }
        }

        private void OnWorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            // 关闭工作簿时把这个工作簿对应窗口的面板一起清掉，避免残留与泄漏
            var deadKeys = new List<int>();
            foreach (var pair in _panes)
            {
                try
                {
                    var window = pair.Value.Window;
                    if (window == null) continue;
                    var owner = window.Parent as Excel.Workbook;
                    if (owner != null && SameWorkbook(owner, wb)) deadKeys.Add(pair.Key);
                }
                catch
                {
                }
            }
            for (var i = 0; i < deadKeys.Count; i++) RemovePane(deadKeys[i]);
        }

        // ==================== 面板管理 ====================

        /// <summary>只查找已存在的面板，不创建。</summary>
        private bool TryGetPane(Excel.Window window, out PaneEntry entry)
        {
            entry = null;
            if (window == null) return false;
            try
            {
                return _panes.TryGetValue(window.Hwnd, out entry);
            }
            catch
            {
                return false;
            }
        }

        private PaneEntry EnsurePane(Excel.Window window)
        {
            if (window == null) return null;

            int key;
            try
            {
                key = window.Hwnd;
            }
            catch
            {
                return null;
            }

            PaneEntry existing;
            if (_panes.TryGetValue(key, out existing))
            {
                existing.Window = window;
                return existing;
            }

            try
            {
                // 只建外壳（纯 WinForms），WPF 内容留给 EnsureContent 稍后挂上
                var host = new PaneHost();
                var pane = CustomTaskPanes.Add(host, "工作表导航", window);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                pane.Visible = true;
                ApplyWidth(pane);

                var entry = new PaneEntry
                {
                    Window = window,
                    Pane = pane,
                    Host = host,
                    Key = key
                };
                _panes[key] = entry;
                WirePane(entry);

                // 创建时就是可见的，这种"从无到有"不会触发 VisibleChanged，必须手动记一次
                if (!_suppressStateSave) PaneOptions.LastPaneVisible = true;

                ScheduleContent(entry);
                return entry;
            }
            catch (Exception ex)
            {
                ShowError("无法创建“工作表导航”侧边栏。", ex);
                return null;
            }
        }

        /// <summary>外壳已经出现，WPF 内容晚一拍再挂，把初始化挪出 Excel 启动路径。</summary>
        private void ScheduleContent(PaneEntry entry)
        {
            var timer = new System.Windows.Forms.Timer { Interval = 120 };
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                timer.Dispose();
                Guarded(() => EnsureContent(entry));
            };
            timer.Start();
        }

        private void EnsureContent(PaneEntry entry)
        {
            if (entry == null || entry.Ui != null || entry.Host == null) return;

            var ui = new NavPaneControl();
            entry.Host.Attach(ui);
            entry.Ui = ui;
            WireUi(entry);
            ui.BeginLoading();

            ui.Dispatcher.BeginInvoke(new Action(() =>
            {
                PaneEntry alive;
                if (!_panes.TryGetValue(entry.Key, out alive) || !ReferenceEquals(alive, entry)) return;
                Guarded(() => Refresh(entry, true));
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>面板级的订阅：显隐、宽度、状态记忆。</summary>
        private void WirePane(PaneEntry entry)
        {
            entry.Pane.VisibleChanged += (s, e) => Guarded(() =>
            {
                // 记住状态，供下次启动恢复（关闭工作簿导致的隐藏不记）
                if (!_suppressStateSave) PaneOptions.LastPaneVisible = entry.Pane.Visible;
                SyncRibbon(entry.Pane.Visible);
                if (!entry.Pane.Visible) return;

                // 面板被重新显示时 Office 可能把宽度改回默认值，这里补一次
                ApplyWidth(entry.Pane);
                if (entry.Ui == null) EnsureContent(entry);
                else Refresh(entry, false);
            });
        }

        /// <summary>WPF 内容级的订阅。</summary>
        private void WireUi(PaneEntry entry)
        {
            entry.Ui.RefreshRequested += (s, e) => Refresh(entry, true);
            entry.Ui.NewSheetRequested += (s, e) => Guarded(() => CreateSheet(entry));
            entry.Ui.ActivateRequested += (s, e) => Guarded(() => ActivateSheet(entry, e.Sheet.Name));
            entry.Ui.VisibilityRequested += (s, e) => Guarded(() => ToggleVisibility(entry, e.Sheet));
            entry.Ui.RenameRequested += (s, e) => Guarded(() => RenameSheet(entry, e.Sheet));
            entry.Ui.DeleteRequested += (s, e) => Guarded(() => DeleteSheet(entry, e.Sheet));
            entry.Ui.WidthChangeRequested += width => Guarded(() =>
            {
                if (entry.Pane != null) entry.Pane.Width = (int)width;
            });
        }

        private static void ApplyWidth(CustomTaskPane pane)
        {
            if (pane == null) return;
            try
            {
                pane.Width = (int)PaneOptions.PaneWidth;
            }
            catch
            {
            }
        }

        private void RemovePane(int key)
        {
            PaneEntry entry;
            if (!_panes.TryGetValue(key, out entry)) return;
            _panes.Remove(key);

            _suppressStateSave = true;
            try
            {
                if (entry.Host != null)
                {
                    entry.Host.Detach();
                    entry.Host.Dispose();
                }
            }
            catch
            {
            }
            try
            {
                if (entry.Pane != null) CustomTaskPanes.Remove(entry.Pane);
            }
            catch
            {
            }
            finally
            {
                _suppressStateSave = false;
            }
        }

        // 暴露给 Ribbon：显示/隐藏
        public void TogglePane()
        {
            var window = SafeActiveWindow();
            if (window == null) return;

            PaneEntry entry;
            if (TryGetPane(window, out entry) && entry.Pane != null)
            {
                // 已经建过：切换显隐
                entry.Pane.Visible = !entry.Pane.Visible;
            }
            else
            {
                // 第一次点：直接创建并显示（EnsurePane 内部已经 Visible=true）
                entry = EnsurePane(window);
            }

            if (entry != null && entry.Pane != null) SyncRibbon(entry.Pane.Visible);
        }

        public bool IsPaneVisible
        {
            get
            {
                try
                {
                    if (Application.ActiveWindow == null) return false;
                    var key = Application.ActiveWindow.Hwnd;
                    PaneEntry entry;
                    return _panes.TryGetValue(key, out entry) && entry.Pane != null && entry.Pane.Visible;
                }
                catch
                {
                    return false;
                }
            }
        }

        // ==================== 数据刷新 ====================

        private void Refresh(PaneEntry entry, bool force)
        {
            if (entry == null || entry.Ui == null || _busy) return;

            _busy = true;
            try
            {
                Excel.Workbook wb = null;
                try
                {
                    wb = entry.Window.Parent as Excel.Workbook;
                }
                catch
                {
                }

                if (wb == null)
                {
                    entry.Ui.SetWorkbook(new List<SheetInfo>(), null);
                    return;
                }

                var activeName = GetActiveSheetName(wb);
                var sheets = ReadSheets(wb);

                var signature = BuildSignature(sheets, activeName);
                if (!force && signature == entry.Signature)
                {
                    entry.Ui.SetActiveSheet(activeName);
                    return;
                }
                entry.Signature = signature;

                for (var i = 0; i < sheets.Count; i++)
                {
                    sheets[i].IsActive = string.Equals(sheets[i].Name, activeName, StringComparison.OrdinalIgnoreCase);
                }

                entry.Ui.SetWorkbook(sheets, activeName);
            }
            catch (Exception ex)
            {
                ShowError("刷新工作表列表时出错。", ex);
            }
            finally
            {
                _busy = false;
            }
        }

        private Excel.Workbook CurrentWorkbook(PaneEntry entry)
        {
            if (entry == null || entry.Window == null) return null;
            try
            {
                return entry.Window.Parent as Excel.Workbook;
            }
            catch
            {
                return null;
            }
        }

        private string GetActiveSheetName(Excel.Workbook wb)
        {
            object active = null;
            try
            {
                active = Application.ActiveSheet;
                if (active == null) return null;

                Excel.Workbook owner = null;
                try
                {
                    owner = GetSheetParent(active);
                }
                catch
                {
                }
                if (owner == null || !SameWorkbook(owner, wb)) return null;
                return GetSheetName(active);
            }
            catch
            {
                return null;
            }
            finally
            {
                Release(active);
            }
        }

        private List<SheetInfo> ReadSheets(Excel.Workbook wb)
        {
            var result = new List<SheetInfo>();
            Excel.Sheets collection = null;
            try
            {
                collection = wb.Sheets;
                var count = collection.Count;
                for (var i = 1; i <= count; i++)
                {
                    object item = null;
                    try
                    {
                        item = collection[i];
                        var info = ToSheetInfo(item, i);
                        if (info != null) result.Add(info);
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                Release(collection);
            }
            return result;
        }

        private static SheetInfo ToSheetInfo(object item, int index)
        {
            var worksheet = item as Excel.Worksheet;
            if (worksheet != null)
            {
                string name;
                try
                {
                    name = worksheet.Name;
                }
                catch
                {
                    return null;
                }

                var visible = true;
                try
                {
                    visible = worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                }
                catch
                {
                }

                var info = new SheetInfo(name, index, SheetKind.Worksheet, ReadTabColor(worksheet.Tab));
                info.Visible = visible;
                return info;
            }

            var chart = item as Excel._Chart;
            if (chart != null)
            {
                string name;
                try
                {
                    name = chart.Name;
                }
                catch
                {
                    return null;
                }

                var visible = true;
                try
                {
                    visible = chart.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                }
                catch
                {
                }

                var info = new SheetInfo(name, index, SheetKind.Chart, ReadTabColor(chart.Tab));
                info.Visible = visible;
                return info;
            }

            return null;
        }

        private static System.Windows.Media.Color? ReadTabColor(Excel.Tab tab)
        {
            if (tab == null) return null;
            try
            {
                var colorIndex = Convert.ToInt32(tab.ColorIndex);
                // -4142 = xlColorIndexNone, -4105 = xlColorIndexAutomatic
                if (colorIndex == -4142 || colorIndex == -4105 || colorIndex == 0) return null;

                var bgr = Convert.ToInt32(tab.Color);
                var r = (byte)(bgr & 0xFF);
                var g = (byte)((bgr >> 8) & 0xFF);
                var b = (byte)((bgr >> 16) & 0xFF);
                return System.Windows.Media.Color.FromRgb(r, g, b);
            }
            catch
            {
                return null;
            }
            finally
            {
                Release(tab);
            }
        }

        private static string BuildSignature(List<SheetInfo> sheets, string activeName)
        {
            var builder = new System.Text.StringBuilder();
            builder.Append(activeName).Append('#');
            for (var i = 0; i < sheets.Count; i++)
            {
                builder.Append(sheets[i].Name).Append(':')
                       .Append(sheets[i].Visible ? 'v' : 'h')
                       .Append(':')
                       .Append(sheets[i].TabColor.HasValue ? sheets[i].TabColor.Value.ToString() : "-")
                       .Append('|');
            }
            return builder.ToString();
        }

        // ==================== 工作表动作 ====================

        private void ActivateSheet(PaneEntry entry, string name)
        {
            var wb = CurrentWorkbook(entry);
            if (wb == null || string.IsNullOrEmpty(name)) return;

            object sheet = null;
            try
            {
                sheet = wb.Sheets[name];
                if (sheet == null) return;
                try
                {
                    wb.Activate();
                }
                catch
                {
                }

                var worksheet = sheet as Excel.Worksheet;
                if (worksheet != null)
                {
                    worksheet.Activate();
                }
                else
                {
                    var chart = sheet as Excel._Chart;
                    if (chart != null) chart.Activate();
                }
            }
            finally
            {
                Release(sheet);
            }

            Refresh(entry, false);
        }

        private void ToggleVisibility(PaneEntry entry, SheetInfo sheet)
        {
            var wb = CurrentWorkbook(entry);
            if (wb == null || sheet == null) return;

            object target = null;
            try
            {
                target = wb.Sheets[sheet.Name];
                if (target == null) return;

                var wantVisible = !sheet.Visible;
                if (!wantVisible)
                {
                    // Excel 不允许隐藏活动工作表，先切到另一张可见表
                    var activeName = GetActiveSheetName(wb);
                    if (string.Equals(activeName, sheet.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        var fallback = FindOtherVisible(wb, sheet.Name);
                        if (fallback == null)
                        {
                            ShowMessage("至少需要保留一个可见的工作表。");
                            return;
                        }
                        try
                        {
                            wb.Sheets[fallback].Activate();
                        }
                        catch
                        {
                        }
                    }
                }

                var targetVisibility = wantVisible
                    ? Excel.XlSheetVisibility.xlSheetVisible
                    : Excel.XlSheetVisibility.xlSheetHidden;

                var worksheet = target as Excel.Worksheet;
                if (worksheet != null)
                {
                    worksheet.Visible = targetVisibility;
                }
                else
                {
                    var chart = target as Excel._Chart;
                    if (chart != null) chart.Visible = targetVisibility;
                }
            }
            catch (Exception ex)
            {
                ShowError("切换工作表显示状态失败。", ex);
            }
            finally
            {
                Release(target);
            }

            Refresh(entry, true);
        }

        private string FindOtherVisible(Excel.Workbook wb, string excludeName)
        {
            Excel.Sheets collection = null;
            try
            {
                collection = wb.Sheets;
                var count = collection.Count;
                for (var i = 1; i <= count; i++)
                {
                    object item = null;
                    try
                    {
                        item = collection[i];
                        var worksheet = item as Excel.Worksheet;
                        if (worksheet != null)
                        {
                            if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible &&
                                !string.Equals(worksheet.Name, excludeName, StringComparison.OrdinalIgnoreCase))
                            {
                                return worksheet.Name;
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                Release(collection);
            }
            return null;
        }

        private void RenameSheet(PaneEntry entry, SheetInfo sheet)
        {
            var wb = CurrentWorkbook(entry);
            if (wb == null || sheet == null) return;

            var newName = UiPrompt.Ask(GetExcelHandle(), "重命名工作表", "新的工作表名称：", sheet.Name);
            if (newName == null) return;
            newName = newName.Trim();
            if (newName.Length == 0 || newName == sheet.Name) return;

            object target = null;
            try
            {
                target = wb.Sheets[sheet.Name];
                if (target == null) return;

                var worksheet = target as Excel.Worksheet;
                if (worksheet != null)
                {
                    worksheet.Name = newName;
                }
                else
                {
                    var chart = target as Excel._Chart;
                    if (chart != null) chart.Name = newName;
                }
            }
            catch (Exception ex)
            {
                ShowError("重命名失败（名称不能为空、不能重名，且不能包含 : \\ / ? * [ ]）。", ex);
            }
            finally
            {
                Release(target);
            }

            Refresh(entry, true);
        }

        private void DeleteSheet(PaneEntry entry, SheetInfo sheet)
        {
            var wb = CurrentWorkbook(entry);
            if (wb == null || sheet == null) return;

            var answer = System.Windows.Forms.MessageBox.Show(
                "确定要删除工作表“" + sheet.Name + "”吗？此操作无法撤销。",
                "工作表导航",
                System.Windows.Forms.MessageBoxButtons.OKCancel,
                System.Windows.Forms.MessageBoxIcon.Warning,
                System.Windows.Forms.MessageBoxDefaultButton.Button2);
            if (answer != System.Windows.Forms.DialogResult.OK) return;

            object target = null;
            try
            {
                target = wb.Sheets[sheet.Name];
                if (target == null) return;
                var worksheet = target as Excel.Worksheet;
                if (worksheet != null)
                {
                    worksheet.Delete();
                }
                else
                {
                    var chart = target as Excel._Chart;
                    if (chart != null) chart.Delete();
                }
            }
            catch (Exception ex)
            {
                ShowError("删除工作表失败（工作簿可能处于保护状态）。", ex);
            }
            finally
            {
                Release(target);
            }

            Refresh(entry, true);
        }

        private void CreateSheet(PaneEntry entry)
        {
            var wb = CurrentWorkbook(entry);
            if (wb == null) return;

            Excel.Sheets worksheets = null;
            object last = null;
            object added = null;
            try
            {
                worksheets = wb.Worksheets;
                last = wb.Sheets[wb.Sheets.Count];
                // Worksheets.Add 返回 object，需要显式转换
                added = worksheets.Add(Type.Missing, last, Type.Missing, Type.Missing);
                var created = added as Excel.Worksheet;
                if (created != null) created.Activate();
            }
            catch (Exception ex)
            {
                ShowError("新建工作表失败。", ex);
            }
            finally
            {
                Release(added);
                Release(last);
                Release(worksheets);
            }

            Refresh(entry, true);
        }

        // ==================== 工具方法 ====================

        private static Excel.Workbook GetSheetParent(object sheet)
        {
            var worksheet = sheet as Excel.Worksheet;
            if (worksheet != null) return worksheet.Parent as Excel.Workbook;
            var chart = sheet as Excel._Chart;
            if (chart != null) return chart.Parent as Excel.Workbook;
            return null;
        }

        private static string GetSheetName(object sheet)
        {
            var worksheet = sheet as Excel.Worksheet;
            if (worksheet != null) return worksheet.Name;
            var chart = sheet as Excel._Chart;
            if (chart != null) return chart.Name;
            return null;
        }

        private static bool SameWorkbook(Excel.Workbook left, Excel.Workbook right)
        {
            if (left == null || right == null) return false;
            try
            {
                return string.Equals(left.Name, right.Name, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private IntPtr GetExcelHandle()
        {
            try
            {
                return new IntPtr(Application.Hwnd);
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        private static void Release(object comObject)
        {
            if (comObject == null) return;
            try
            {
                if (Marshal.IsComObject(comObject)) Marshal.ReleaseComObject(comObject);
            }
            catch
            {
            }
        }

        private void Guarded(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                ShowError("操作未完成。", ex);
            }
        }

        private void SyncRibbon(bool visible)
        {
            try
            {
                var ribbon = Globals.Ribbons.RibbonNavigation;
                if (ribbon != null && ribbon.toggleButton1 != null)
                {
                    ribbon.toggleButton1.Checked = visible;
                }
            }
            catch
            {
                // Ribbon 还没加载，忽略
            }
        }

        private static void ShowMessage(string text)
        {
            System.Windows.Forms.MessageBox.Show(text, "工作表导航",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private static void ShowError(string text, Exception ex)
        {
            var message = text;
            if (ex != null) message += Environment.NewLine + Environment.NewLine + ex.Message;
            System.Windows.Forms.MessageBox.Show(message, "工作表导航",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
