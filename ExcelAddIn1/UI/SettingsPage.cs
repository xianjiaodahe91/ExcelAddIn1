using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelAddIn1.UI
{
    /// <summary>“设置”页。</summary>
    public sealed class SettingsPage : Grid
    {
        public event EventHandler ShowHiddenSheetsChanged;
        public event Action<double> WidthChangeRequested;

        private readonly CheckBox _showHidden;
        private readonly TextBox _widthBox;

        public SettingsPage()
        {
            RowDefinitions.Add(new RowDefinition { Height = new GridLength(UiTheme.HeaderHeight) });
            RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            Background = UiTheme.PaneBg;

            var header = new Grid { Background = UiTheme.HeaderBg };
            var title = new TextBlock
            {
                Text = "设置",
                FontSize = UiTheme.FontTitle,
                FontWeight = FontWeights.SemiBold,
                Foreground = UiTheme.TextPrimary,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(10, 0, 0, 0)
            };
            header.Children.Add(title);

            header.Children.Add(new Border
            {
                Height = 1,
                Background = UiTheme.DividerSoft,
                VerticalAlignment = VerticalAlignment.Bottom
            });

            Children.Add(header);

            var host = new StackPanel { Margin = new Thickness(10, 10, 10, 10) };

            host.Children.Add(new TextBlock
            {
                Text = "列表",
                FontSize = UiTheme.FontSmall,
                FontWeight = FontWeights.SemiBold,
                Foreground = UiTheme.Accent,
                Margin = new Thickness(0, 0, 0, 8)
            });

            _showHidden = new CheckBox
            {
                Content = "显示被隐藏的工作表",
                IsChecked = PaneOptions.ShowHiddenSheets,
                Style = (Style)UiTheme.Styles["PaneCheckBox"],
                Margin = new Thickness(0, 0, 0, 4)
            };
            _showHidden.Checked += OnShowHiddenChanged;
            _showHidden.Unchecked += OnShowHiddenChanged;
            host.Children.Add(_showHidden);

            host.Children.Add(new TextBlock
            {
                Text = "关闭后，隐藏的工作表不会出现在列表里。",
                FontSize = UiTheme.FontSmall,
                Foreground = UiTheme.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(22, 0, 0, 0)
            });

            host.Children.Add(new TextBlock
            {
                Text = "外观",
                FontSize = UiTheme.FontSmall,
                FontWeight = FontWeights.SemiBold,
                Foreground = UiTheme.Accent,
                Margin = new Thickness(0, 16, 0, 8)
            });

            var widthRow = new Grid();
            widthRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            widthRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            widthRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            widthRow.Children.Add(new TextBlock
            {
                Text = "面板宽度",
                FontSize = UiTheme.FontMid,
                Foreground = UiTheme.TextSecondary,
                VerticalAlignment = VerticalAlignment.Center
            });

            _widthBox = new TextBox
            {
                Text = ((int)PaneOptions.PaneWidth).ToString(CultureInfo.InvariantCulture),
                Width = 54,
                Height = 26,
                FontSize = UiTheme.FontBody,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                BorderBrush = UiTheme.InputBorder,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(0)
            };
            Grid.SetColumn(_widthBox, 1);
            widthRow.Children.Add(_widthBox);

            var apply = new Button
            {
                Content = "应用",
                Style = (Style)UiTheme.Styles["TextButton"],
                Margin = new Thickness(6, 0, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            apply.Click += (s, e) => ApplyWidth();
            Grid.SetColumn(apply, 2);
            widthRow.Children.Add(apply);
            host.Children.Add(widthRow);

            host.Children.Add(new TextBlock
            {
                Text = "范围 220 – 700 像素，也可以直接拖动面板边缘。",
                FontSize = UiTheme.FontSmall,
                Foreground = UiTheme.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(80, 6, 0, 0)
            });

            host.Children.Add(new Border
            {
                Height = 1,
                Background = UiTheme.DividerSoft,
                Margin = new Thickness(0, 18, 0, 10)
            });

            host.Children.Add(new TextBlock
            {
                Text = "工作表导航 · VSTO + WPF\n点击列表行跳转，眼睛图标切换隐藏，右键更多操作。",
                FontSize = UiTheme.FontSmall,
                Foreground = UiTheme.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 18
            });

            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = host
            };
            Grid.SetRow(scroll, 1);
            Children.Add(scroll);
        }

        private void OnShowHiddenChanged(object sender, RoutedEventArgs e)
        {
            PaneOptions.ShowHiddenSheets = _showHidden.IsChecked == true;
            var handler = ShowHiddenSheetsChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void ApplyWidth()
        {
            double value;
            if (!double.TryParse(_widthBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return;
            PaneOptions.PaneWidth = value;
            _widthBox.Text = ((int)PaneOptions.PaneWidth).ToString(CultureInfo.InvariantCulture);
            var handler = WidthChangeRequested;
            if (handler != null) handler(PaneOptions.PaneWidth);
        }
    }
}
