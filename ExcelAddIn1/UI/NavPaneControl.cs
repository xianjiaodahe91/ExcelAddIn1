using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelAddIn1.UI
{
    /// <summary>
    /// 面板根控件：左侧图标栏 + 右侧页面（工作表 / 设置）。
    /// 完全不引用 Excel 类型，所有动作通过事件交给加载项处理。
    /// </summary>
    public sealed class NavPaneControl : UserControl
    {
        private readonly SheetListPage _sheetPage;
        private readonly SettingsPage _settingsPage;
        private readonly List<RadioButton> _railButtons = new List<RadioButton>();

        public event EventHandler RefreshRequested;
        public event EventHandler NewSheetRequested;
        public event EventHandler<SheetActionEventArgs> ActivateRequested;
        public event EventHandler<SheetActionEventArgs> RenameRequested;
        public event EventHandler<SheetActionEventArgs> DeleteRequested;
        public event EventHandler<SheetActionEventArgs> VisibilityRequested;
        public event Action<double> WidthChangeRequested;

        public NavPaneControl()
        {
            FontFamily = UiTheme.Font;
            FontSize = UiTheme.FontBody;
            UseLayoutRounding = true;
            SnapsToDevicePixels = true;
            Background = UiTheme.PaneBg;

            var root = new Grid();
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(UiTheme.RailWidth) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // ---------- 左栏 ----------
            var rail = new Border
            {
                Background = UiTheme.RailBg,
                BorderBrush = UiTheme.Divider,
                BorderThickness = new Thickness(0, 0, 1, 0)
            };
            var railStack = new Grid();
            railStack.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            railStack.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            railStack.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var topStack = new StackPanel { Margin = new Thickness(0, 6, 0, 0) };
            topStack.Children.Add(MakeRailButton(UiTheme.Icons.Sheets, "工作表", 0));
            railStack.Children.Add(topStack);

            var bottomStack = new StackPanel { Margin = new Thickness(0, 0, 0, 6) };
            bottomStack.Children.Add(MakeRailButton(UiTheme.Icons.Settings, "设置", 1));
            Grid.SetRow(bottomStack, 2);
            railStack.Children.Add(bottomStack);

            rail.Child = railStack;
            root.Children.Add(rail);

            // ---------- 内容区 ----------
            var host = new Grid { Background = UiTheme.PaneBg };
            _sheetPage = new SheetListPage { ShowHiddenSheets = PaneOptions.ShowHiddenSheets };
            _settingsPage = new SettingsPage();
            host.Children.Add(_sheetPage);
            host.Children.Add(_settingsPage);
            Grid.SetColumn(host, 1);
            root.Children.Add(host);

            Content = root;

            // ---------- 事件转发 ----------
            _sheetPage.RefreshRequested += (s, e) => Raise(RefreshRequested);
            _sheetPage.NewSheetRequested += (s, e) => Raise(NewSheetRequested);
            _sheetPage.ActivateRequested += (s, e) => Forward(ActivateRequested, e);
            _sheetPage.RenameRequested += (s, e) => Forward(RenameRequested, e);
            _sheetPage.DeleteRequested += (s, e) => Forward(DeleteRequested, e);
            _sheetPage.VisibilityRequested += (s, e) => Forward(VisibilityRequested, e);
            _settingsPage.ShowHiddenSheetsChanged += (s, e) =>
            {
                _sheetPage.ShowHiddenSheets = PaneOptions.ShowHiddenSheets;
                _sheetPage.RefreshList();
            };
            _settingsPage.WidthChangeRequested += width =>
            {
                var handler = WidthChangeRequested;
                if (handler != null) handler(width);
            };

            ShowPage(0);
        }

        // ---------- 对外 API ----------

        public void SetWorkbook(IList<SheetInfo> sheets, string activeSheet)
        {
            _sheetPage.SetSheets(sheets);
            if (!string.IsNullOrEmpty(activeSheet)) _sheetPage.SetActive(activeSheet);
        }

        /// <summary>显示“载入中”占位，等数据到了再填。</summary>
        public void BeginLoading()
        {
            _sheetPage.BeginLoading();
        }

        public void SetActiveSheet(string name)
        {
            _sheetPage.SetActive(name);
        }

        public void ShowPage(int index)
        {
            for (var i = 0; i < _railButtons.Count; i++)
            {
                _railButtons[i].IsChecked = i == index;
            }
            _sheetPage.Visibility = index == 0 ? Visibility.Visible : Visibility.Collapsed;
            _settingsPage.Visibility = index == 1 ? Visibility.Visible : Visibility.Collapsed;
        }

        public void FocusSheetPage()
        {
            ShowPage(0);
        }

        // ---------- 内部 ----------

        private RadioButton MakeRailButton(string iconData, string tooltip, int pageIndex)
        {
            var button = new RadioButton
            {
                GroupName = "navRail",
                ToolTip = tooltip,
                Style = (Style)UiTheme.Styles["RailButton"],
                Content = UiTheme.Icon(iconData, UiTheme.IconRail, UiTheme.IconNormal, 1.3)
            };
            button.Checked += (s, e) => ShowPage(pageIndex);
            _railButtons.Add(button);
            return button;
        }

        private void Raise(EventHandler handler)
        {
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void Forward(EventHandler<SheetActionEventArgs> handler, SheetActionEventArgs args)
        {
            if (handler != null) handler(this, args);
        }
    }
}
