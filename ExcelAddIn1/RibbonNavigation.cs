using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace ExcelAddIn1
{
    public partial class RibbonNavigation
    {
        private void RibbonNavigation_Load(object sender, RibbonUIEventArgs e)
        {
            // 按钮：图标在上、文字在下（ToggleButton 本身就是大按钮样式）
            try
            {
                var image = CreateSheetIcon(32);
                if (image != null) toggleButton1.Image = image;
                toggleButton1.ShowImage = true;
                toggleButton1.ShowLabel = true;
                toggleButton1.ScreenTip = "显示 / 隐藏工作表导航面板";
                toggleButton1.Label = "\u5BFC\u2060\u822A"; // 导 + 词连接符 + 航，避免被逐字换行
            }
            catch
            {
                // 图标生成失败不影响功能
            }

            // 让 Ribbon 上的开关与面板实际状态保持一致
            try
            {
                toggleButton1.Checked = Globals.ThisAddIn.IsPaneVisible;
            }
            catch
            {
            }
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TogglePane();
        }

        /// <summary>
        /// 画一个“工作表”图标（32×32）：白页 + 灰边框 + 橙色表头 + 网格。
        /// 用 GDI+ 直接画，不走 WPF——Ribbon 加载发生在 Excel 启动阶段，
        /// 这时候初始化 WPF 会拖慢首帧并触发 LoaderLock。
        /// </summary>
        private static Image CreateSheetIcon(int size)
        {
            try
            {
                var k = size / 32f;
                var bitmap = new Bitmap(size, size);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.SmoothingMode = SmoothingMode.None;
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.Clear(Color.Transparent);

                    using (var border = new Pen(Color.FromArgb(0x8C, 0x8C, 0x8C), Math.Max(1f, k)))
                    using (var grid = new Pen(Color.FromArgb(0xC8, 0xC8, 0xC8), Math.Max(1f, k * 0.8f)))
                    using (var header = new SolidBrush(Color.FromArgb(0xE8, 0x5D, 0x2A)))
                    using (var side = new SolidBrush(Color.FromArgb(0xEF, 0xEF, 0xEF)))
                    {
                        // 页面 3..29（26×26）
                        var page = new RectangleF(3f * k, 3f * k, 26f * k, 26f * k);
                        g.FillRectangle(Brushes.White, page);
                        g.DrawRectangle(border, page.X, page.Y, page.Width, page.Height);

                        // 左侧行号列 4..10
                        g.FillRectangle(side, 4f * k, 4f * k, 6f * k, 24f * k);
                        // 橙色表头 10..29 × 4..10
                        g.FillRectangle(header, 10f * k, 4f * k, 19f * k, 6f * k);

                        // 竖线
                        for (var i = 0; i < 3; i++)
                        {
                            var x = (15f + i * 5f) * k;
                            g.DrawLine(grid, x, 3f * k, x, 29f * k);
                        }
                        // 横线
                        g.DrawLine(grid, 3f * k, 10f * k, 29f * k, 10f * k);
                        g.DrawLine(grid, 3f * k, 16f * k, 29f * k, 16f * k);
                        g.DrawLine(grid, 3f * k, 22f * k, 29f * k, 22f * k);
                        // 行号列与网格的分隔线
                        g.DrawLine(grid, 10f * k, 3f * k, 10f * k, 29f * k);
                    }
                }
                return bitmap;
            }
            catch
            {
                return null;
            }
        }
    }
}
