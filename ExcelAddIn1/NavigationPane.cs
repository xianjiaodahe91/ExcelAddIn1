using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class NavigationPane : UserControl
    {
        public NavigationPane()
        {
            InitializeComponent();
            InitNavItems();
        }

        private void InitNavItems()
        {
            // 添加导航按钮或 TreeView
            var btn1 = new Button { Text = "📊 数据总览", Dock = DockStyle.Top, Height = 40 };
            var btn2 = new Button { Text = "📈 图表分析", Dock = DockStyle.Top, Height = 40 };
            var btn3 = new Button { Text = "⚙️ 设置", Dock = DockStyle.Top, Height = 40 };

            btn1.Click += (s, e) => NavigateTo("overview");
            btn2.Click += (s, e) => NavigateTo("chart");
            btn3.Click += (s, e) => NavigateTo("settings");

            this.Controls.AddRange(new Control[] { btn3, btn2, btn1 });
        }


        private void NavigateTo(string page)
        {
            // 操作 Excel，例如跳转到指定 Sheet
            var app = Globals.ThisAddIn.Application;
            switch (page)
            {
                case "overview":
                    app.Worksheets["总览"].Activate(); break;
                case "chart":
                    app.Worksheets["图表"].Activate(); break;
            }
        }

    }
}
