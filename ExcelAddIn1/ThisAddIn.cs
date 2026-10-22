using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;


namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        // 用窗口句柄做 key，每个窗口维护一个 TaskPane
        private Dictionary<Excel.Window, CustomTaskPane> _taskPanes
            = new Dictionary<Excel.Window, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 为启动时已打开的工作簿创建
            CreateTaskPaneForWindow(this.Application.ActiveWindow);

            // 监听新建/打开工作簿事件
            this.Application.WorkbookActivate += App_WorkbookActivate;
            this.Application.WorkbookBeforeClose += App_WorkbookBeforeClose;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        // 切换到任意工作簿时触发
        private void App_WorkbookActivate(Excel.Workbook wb)
        {
            CreateTaskPaneForWindow(this.Application.ActiveWindow);
        }

        // 工作簿关闭时清理，避免内存泄漏
        private void App_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            var window = this.Application.ActiveWindow;
            if (_taskPanes.TryGetValue(window, out var pane))
            {
                this.CustomTaskPanes.Remove(pane);
                _taskPanes.Remove(window);
            }
        }

        // 暴露给 Ribbon 按钮控制显示/隐藏
        public void TogglePane()
        {
            var window = this.Application.ActiveWindow;
            if (_taskPanes.TryGetValue(window, out var pane))
                pane.Visible = !pane.Visible;
        }

        private void CreateTaskPaneForWindow(Excel.Window window)
        {
            if (window == null) return;

            // 已存在则跳过
            if (_taskPanes.ContainsKey(window)) return;

            var navControl = new NavigationPane();
            var pane = this.CustomTaskPanes.Add(navControl, "工作表导航", window);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Width = 200;
            pane.Visible = true;

            _taskPanes[window] = pane;
        }
        public bool IsPaneVisible
        {
            get
            {
                var window = this.Application.ActiveWindow;
                return _taskPanes.TryGetValue(window, out var pane) && pane.Visible;
            }
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
