using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace ExcelAddIn1.UI
{
    /// <summary>轻量输入框，用于重命名工作表（避免弹 WinForms 的丑陋对话框）。</summary>
    public static class UiPrompt
    {
        public static string Ask(IntPtr owner, string title, string label, string initialValue)
        {
            var window = new Window
            {
                Title = title,
                Width = 330,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ShowInTaskbar = false,
                Background = UiTheme.PaneBg,
                FontFamily = UiTheme.Font,
                FontSize = UiTheme.FontBody
            };
            window.Resources.MergedDictionaries.Add(UiTheme.Styles);

            if (owner != IntPtr.Zero)
            {
                try
                {
                    new WindowInteropHelper(window).Owner = owner;
                }
                catch
                {
                    // 设置 Owner 失败不影响弹窗
                }
            }

            var root = new StackPanel { Margin = new Thickness(14, 12, 14, 12) };
            root.Children.Add(new TextBlock
            {
                Text = label,
                Foreground = UiTheme.TextSecondary,
                FontSize = UiTheme.FontMid,
                Margin = new Thickness(0, 0, 0, 6)
            });

            var box = new TextBox
            {
                Text = initialValue ?? string.Empty,
                Height = 28,
                FontSize = UiTheme.FontTitle,
                Padding = new Thickness(6, 0, 6, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderBrush = UiTheme.InputBorder,
                BorderThickness = new Thickness(1)
            };
            root.Children.Add(box);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };

            string result = null;
            var ok = new Button
            {
                Content = "确定",
                Width = 70,
                Style = (Style)UiTheme.Styles["TextButton"]
            };
            ok.Click += (s, e) => { result = box.Text; window.DialogResult = true; };

            var cancel = new Button
            {
                Content = "取消",
                Width = 70,
                Margin = new Thickness(8, 0, 0, 0),
                Style = (Style)UiTheme.Styles["TextButton"]
            };
            cancel.Click += (s, e) => { window.DialogResult = false; };

            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            root.Children.Add(buttons);

            window.Content = root;
            window.Loaded += (s, e) =>
            {
                box.Focus();
                box.SelectAll();
            };
            window.KeyDown += (s, e) =>
            {
                if (e.Key == System.Windows.Input.Key.Enter)
                {
                    result = box.Text;
                    window.DialogResult = true;
                }
                else if (e.Key == System.Windows.Input.Key.Escape)
                {
                    window.DialogResult = false;
                }
            };

            var shown = window.ShowDialog();
            return shown == true ? result : null;
        }
    }
}
