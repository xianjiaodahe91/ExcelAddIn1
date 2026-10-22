using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelAddIn1.UI;

namespace ExcelAddIn1
{
    /// <summary>
    /// VSTO 的 CustomTaskPane 只接受 WinForms 的 UserControl。
    /// 这里把它做成"两层"：
    ///   1) 外壳是纯 WinForms（白底 + “载入中…”），创建几乎零成本，
    ///      可以跟着 Excel 一起出现，不会有"左栏晚半拍"的观感；
    ///   2) WPF 内容稍后再挂进来（Attach），把 WPF 那 300ms 初始化挪出启动路径。
    /// </summary>
    public sealed class PaneHost : UserControl
    {
        private ElementHost _host;
        private Label _placeholder;

        public PaneHost()
        {
            BackColor = Color.White;

            _placeholder = new Label
            {
                Dock = DockStyle.Fill,
                Text = "载入中…",
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(0x8A, 0x8A, 0x86),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.White
            };
            Controls.Add(_placeholder);
        }

        /// <summary>内容是否已经挂上。</summary>
        public bool HasContent
        {
            get { return _host != null; }
        }

        /// <summary>把 WPF 面板挂进来，替换掉占位层。</summary>
        public void Attach(NavPaneControl content)
        {
            if (content == null || _host != null) return;

            _host = new ElementHost
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Child = content
            };
            Controls.Add(_host);
            _host.BringToFront();

            if (_placeholder != null)
            {
                Controls.Remove(_placeholder);
                _placeholder.Dispose();
                _placeholder = null;
            }
        }

        /// <summary>释放前先断开 WPF 元素，避免 ElementHost 泄漏。</summary>
        public void Detach()
        {
            try
            {
                if (_host != null) _host.Child = null;
            }
            catch
            {
            }
        }
    }
}
