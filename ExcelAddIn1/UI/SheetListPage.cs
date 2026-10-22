using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;

namespace ExcelAddIn1.UI
{
    /// <summary>
    /// “工作表”页：新建/重命名/删除/刷新 + 搜索 + 工作表列表（点击跳转、眼睛图标切换隐藏）。
    /// 只依赖 SheetInfo，不引用任何 Excel 类型。
    /// </summary>
    public sealed class SheetListPage : Grid
    {
        public event EventHandler RefreshRequested;
        public event EventHandler NewSheetRequested;
        public event EventHandler<SheetActionEventArgs> ActivateRequested;
        public event EventHandler<SheetActionEventArgs> RenameRequested;
        public event EventHandler<SheetActionEventArgs> DeleteRequested;
        public event EventHandler<SheetActionEventArgs> VisibilityRequested;

        private readonly Run _countRun;
        private readonly TextBox _search;
        private readonly TextBlock _placeholder;
        private readonly ItemsControl _list;
        private readonly TextBlock _empty;
        private readonly Button _renameButton;
        private readonly Button _deleteButton;

        private IList<SheetInfo> _all = new List<SheetInfo>();
        private SheetInfo _active;

        /// <summary>是否把隐藏的工作表也列出来（默认列出，否则无法从面板里取消隐藏）。</summary>
        public bool ShowHiddenSheets { get; set; }

        public SheetListPage()
        {
            ShowHiddenSheets = true;
            Resources.MergedDictionaries.Add(UiTheme.Styles);

            RowDefinitions.Add(new RowDefinition { Height = new GridLength(UiTheme.HeaderHeight) });
            RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            Background = UiTheme.PaneBg;

            // ---------- 标题行 ----------
            // 标题列用 Star + 省略号，宽度不够时先压缩标题，保证右侧按钮永远完整显示；
            // 计数用 Run 挂在标题里，避免被 Star 列推远
            var header = new Grid { Background = UiTheme.HeaderBg };
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titleBlock = new TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(10, 0, 6, 0),
                TextTrimming = TextTrimming.CharacterEllipsis,
                TextWrapping = TextWrapping.NoWrap
            };
            titleBlock.Inlines.Add(new Run("工作表")
            {
                FontSize = UiTheme.FontTitle,
                FontWeight = FontWeights.SemiBold,
                Foreground = UiTheme.TextPrimary
            });
            _countRun = new Run(string.Empty)
            {
                FontSize = UiTheme.FontSmall,
                Foreground = UiTheme.TextSecondary
            };
            titleBlock.Inlines.Add(_countRun);
            header.Children.Add(titleBlock);

            var actions = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            };
            var newButton = MakeIconButton(UiTheme.Icons.Plus, "新建工作表");
            newButton.Click += (s, e) => Raise(NewSheetRequested);
            _renameButton = MakeIconButton(UiTheme.Icons.Pencil, "重命名当前工作表");
            _renameButton.Click += (s, e) => RaiseSheet(RenameRequested, _active);
            _deleteButton = MakeIconButton(UiTheme.Icons.Trash, "删除当前工作表");
            _deleteButton.Click += (s, e) => RaiseSheet(DeleteRequested, _active);
            var refreshButton = MakeIconButton(UiTheme.Icons.Refresh, "刷新列表");
            refreshButton.Click += (s, e) => Raise(RefreshRequested);

            actions.Children.Add(newButton);
            actions.Children.Add(_renameButton);
            actions.Children.Add(_deleteButton);
            var sep = new Border
            {
                Width = 1,
                Height = 14,
                Background = UiTheme.Divider,
                Margin = new Thickness(4, 0, 4, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            actions.Children.Add(sep);
            actions.Children.Add(refreshButton);
            Grid.SetColumn(actions, 1);
            header.Children.Add(actions);

            var headerLine = new Border
            {
                Height = 1,
                Background = UiTheme.DividerSoft,
                VerticalAlignment = VerticalAlignment.Bottom
            };
            Grid.SetColumnSpan(headerLine, 2);
            header.Children.Add(headerLine);

            Children.Add(header);

            // ---------- 搜索行 ----------
            var searchBox = new Border
            {
                Height = 28,
                Margin = new Thickness(8, 3, 8, 6),
                CornerRadius = new CornerRadius(3),
                BorderBrush = UiTheme.InputBorder,
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            var searchGrid = new Grid();
            searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(28) });
            searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var searchIcon = UiTheme.Icon(UiTheme.Icons.Search, UiTheme.IconSmall, UiTheme.IconMuted, 1.3);
            searchIcon.Margin = new Thickness(7, 0, 0, 0);
            Grid.SetColumn(searchIcon, 0);
            searchGrid.Children.Add(searchIcon);

            var textHost = new Grid();
            _placeholder = new TextBlock
            {
                Text = "搜索工作表",
                FontSize = UiTheme.FontBody,
                Foreground = UiTheme.TextDisabled,
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false,
                Margin = new Thickness(1, 0, 0, 0)
            };
            _search = new TextBox
            {
                Style = (Style)UiTheme.Styles["SearchBox"],
                Margin = new Thickness(0, 0, 6, 0)
            };
            _search.TextChanged += (s, e) =>
            {
                _placeholder.Visibility = _search.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
                ApplyFilter();
            };
            textHost.Children.Add(_search);
            textHost.Children.Add(_placeholder);
            Grid.SetColumn(textHost, 1);
            searchGrid.Children.Add(textHost);

            searchBox.Child = searchGrid;
            var searchRow = new Border { Child = searchBox };
            Grid.SetRow(searchRow, 1);
            Children.Add(searchRow);

            // ---------- 列表 ----------
            _list = new ItemsControl();
            _list.ItemTemplate = BuildRowTemplate();
            _list.MouseLeftButtonUp += OnListClick;
            _list.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(OnListButtonClick));
            _list.PreviewMouseRightButtonUp += OnListRightClick;
            _list.Margin = new Thickness(0, 0, 0, 4);

            var scroller = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = _list,
                Padding = new Thickness(0, 0, 2, 0)
            };
            Grid.SetRow(scroller, 2);
            Children.Add(scroller);

            _empty = new TextBlock
            {
                Text = "没有匹配的工作表",
                FontSize = UiTheme.FontBody,
                Foreground = UiTheme.TextSecondary,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 20, 0, 0),
                Visibility = Visibility.Collapsed
            };
            Grid.SetRow(_empty, 2);
            Children.Add(_empty);

            UpdateActionState();
        }

        // ---------- 对外 API ----------

        /// <summary>面板刚创建、数据还没读时的占位状态。</summary>
        public void BeginLoading()
        {
            _countRun.Text = string.Empty;
            _empty.Text = "载入中…";
            _empty.Visibility = Visibility.Visible;
        }

        public void SetSheets(IList<SheetInfo> sheets)
        {
            _empty.Text = "没有匹配的工作表";
            _all = sheets ?? new List<SheetInfo>();
            _active = null;
            for (var i = 0; i < _all.Count; i++)
            {
                if (_all[i].IsActive) _active = _all[i];
            }
            ApplyFilter();
            UpdateActionState();
        }

        public void SetActive(string sheetName)
        {
            SheetInfo found = null;
            for (var i = 0; i < _all.Count; i++)
            {
                var item = _all[i];
                var isActive = string.Equals(item.Name, sheetName, StringComparison.OrdinalIgnoreCase);
                item.IsActive = isActive;
                if (isActive) found = item;
            }
            _active = found;
            UpdateActionState();
            ScrollToActive();
        }

        public void FocusSearch()
        {
            _search.Focus();
        }

        /// <summary>选项变化后重新过滤列表。</summary>
        public void RefreshList()
        {
            ApplyFilter();
        }

        // ---------- 内部 ----------

        private void UpdateActionState()
        {
            var has = _active != null;
            _renameButton.IsEnabled = has;
            _deleteButton.IsEnabled = has;
            _deleteButton.ToolTip = has ? "删除当前工作表" : "删除当前工作表";
        }

        private void ApplyFilter()
        {
            var keyword = (_search.Text ?? string.Empty).Trim();
            var view = new List<SheetInfo>();
            for (var i = 0; i < _all.Count; i++)
            {
                var item = _all[i];
                if (!ShowHiddenSheets && !item.Visible) continue;
                if (keyword.Length > 0 &&
                    item.Name.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) < 0) continue;
                view.Add(item);
            }

            _list.ItemsSource = view;
            _countRun.Text = string.Format("  · {0}", view.Count);
            _empty.Visibility = view.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ScrollToActive()
        {
            if (_active == null) return;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var container = _list.ItemContainerGenerator.ContainerFromItem(_active) as FrameworkElement;
                if (container != null) container.BringIntoView();
            }), DispatcherPriority.Background);
        }

        private void OnListClick(object sender, MouseButtonEventArgs e)
        {
            var hit = e.OriginalSource as DependencyObject;
            if (hit == null) return;
            if (UiHelpers.FindButton(hit) != null) return;
            var info = HitSheet(hit);
            if (info == null) return;
            RaiseSheet(ActivateRequested, info);
        }

        private void OnListButtonClick(object sender, RoutedEventArgs e)
        {
            var button = e.OriginalSource as Button;
            if (button == null) return;
            var info = button.DataContext as SheetInfo;
            if (info == null) return;
            if (string.Equals(button.Name, "Eye", StringComparison.Ordinal))
            {
                RaiseSheet(VisibilityRequested, info);
            }
        }

        private void OnListRightClick(object sender, MouseButtonEventArgs e)
        {
            var hit = e.OriginalSource as DependencyObject;
            var info = HitSheet(hit);
            if (info == null) return;

            var menu = new ContextMenu();
            menu.Style = (Style)UiTheme.Styles[typeof(ContextMenu)];
            menu.Items.Add(MakeMenuItem("跳转到此工作表", () => RaiseSheet(ActivateRequested, info)));
            menu.Items.Add(new Separator());
            menu.Items.Add(MakeMenuItem("重命名…", () => RaiseSheet(RenameRequested, info)));
            menu.Items.Add(MakeMenuItem(info.Visible ? "隐藏" : "取消隐藏", () => RaiseSheet(VisibilityRequested, info)));
            menu.Items.Add(new Separator());
            var deleteItem = MakeMenuItem("删除工作表…", () => RaiseSheet(DeleteRequested, info));
            deleteItem.Foreground = new SolidColorBrush(Color.FromRgb(0xC0, 0x39, 0x2B));
            menu.Items.Add(deleteItem);

            menu.PlacementTarget = _list;
            menu.Placement = PlacementMode.MousePoint;
            menu.IsOpen = true;
            e.Handled = true;
        }

        private SheetInfo HitSheet(DependencyObject hit)
        {
            if (hit == null) return null;
            var container = ItemsControl.ContainerFromElement(_list, hit) as FrameworkElement;
            if (container == null) return null;
            return container.DataContext as SheetInfo;
        }

        private static Button MakeIconButton(string iconData, string tooltip)
        {
            var button = new Button();
            button.Style = (Style)UiTheme.Styles["IconButton"];
            button.ToolTip = tooltip;
            button.Content = UiTheme.Icon(iconData, UiTheme.IconTool, UiTheme.IconNormal, 1.3);
            return button;
        }

        private static MenuItem MakeMenuItem(string header, Action action)
        {
            var item = new MenuItem { Header = header, FontSize = UiTheme.FontBody };
            item.Style = (Style)UiTheme.Styles[typeof(MenuItem)];
            item.Click += (s, e) => action();
            return item;
        }

        private void Raise(EventHandler handler)
        {
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void RaiseSheet(EventHandler<SheetActionEventArgs> handler, SheetInfo sheet)
        {
            if (handler != null && sheet != null) handler(this, new SheetActionEventArgs(sheet));
        }

        private static DataTemplate BuildRowTemplate()
        {
            var xaml = RowTemplateXaml
                .Replace("__BODY__", UiTheme.Num(UiTheme.FontBody))
                .Replace("__ROW__", UiTheme.Num(UiTheme.RowHeight))
                .Replace("__ROWICON__", UiTheme.Num(UiTheme.IconRow));
            return (DataTemplate)XamlReader.Parse(xaml);
        }

        private const string RowTemplateXaml = @"
<DataTemplate xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
              xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">
  <Border x:Name=""row"" Height=""__ROW__"" Cursor=""Hand"" SnapsToDevicePixels=""True"">
    <Border.Style>
      <Style TargetType=""Border"">
        <Setter Property=""Background"" Value=""Transparent""/>
        <Style.Triggers>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter Property=""Background"" Value=""#F2F2F0""/>
          </Trigger>
          <DataTrigger Binding=""{Binding IsActive}"" Value=""True"">
            <Setter Property=""Background"" Value=""#C6E0B4""/>
          </DataTrigger>
          <MultiDataTrigger>
            <MultiDataTrigger.Conditions>
              <Condition Binding=""{Binding IsActive}"" Value=""True""/>
              <Condition Binding=""{Binding IsMouseOver, RelativeSource={RelativeSource Self}}"" Value=""True""/>
            </MultiDataTrigger.Conditions>
            <Setter Property=""Background"" Value=""#BBD8A7""/>
          </MultiDataTrigger>
        </Style.Triggers>
      </Style>
    </Border.Style>
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width=""3""/>
        <ColumnDefinition Width=""22""/>
        <ColumnDefinition Width=""*""/>
        <ColumnDefinition Width=""26""/>
      </Grid.ColumnDefinitions>
      <Rectangle Grid.Column=""0"" Fill=""{Binding TabBrush}"" Margin=""0,4,0,4""/>
      <Path Grid.Column=""1"" Data=""{Binding KindIconData}"" Stroke=""#6A6A66"" StrokeThickness=""1.25""
            Width=""__ROWICON__"" Height=""__ROWICON__"" Stretch=""Uniform"" StrokeLineJoin=""Round""
            StrokeStartLineCap=""Round"" StrokeEndLineCap=""Round""
            HorizontalAlignment=""Center"" VerticalAlignment=""Center""/>
      <TextBlock Grid.Column=""2"" Text=""{Binding Name}"" VerticalAlignment=""Center""
                 FontSize=""__BODY__"" Foreground=""#323232"" TextTrimming=""CharacterEllipsis""
                 Opacity=""{Binding NameOpacity}"" Margin=""0,0,4,0"">
        <TextBlock.Style>
          <Style TargetType=""TextBlock"">
            <Style.Triggers>
              <DataTrigger Binding=""{Binding IsActive}"" Value=""True"">
                <Setter Property=""FontWeight"" Value=""SemiBold""/>
              </DataTrigger>
            </Style.Triggers>
          </Style>
        </TextBlock.Style>
      </TextBlock>
      <Button x:Name=""Eye"" Grid.Column=""3"" Width=""24"" Height=""24"" Cursor=""Hand""
              ToolTip=""{Binding EyeTooltip}"" Background=""Transparent"" BorderThickness=""0""
              Focusable=""False"">
        <Button.Template>
          <ControlTemplate TargetType=""Button"">
            <Border x:Name=""bd"" Background=""Transparent"" CornerRadius=""3"">
              <ContentPresenter HorizontalAlignment=""Center"" VerticalAlignment=""Center""/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property=""IsMouseOver"" Value=""True"">
                <Setter TargetName=""bd"" Property=""Background"" Value=""#40FFFFFF""/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Button.Template>
        <Path Data=""{Binding EyeIconData}"" Stroke=""#6A6A66"" StrokeThickness=""1.2""
              Width=""__ROWICON__"" Height=""__ROWICON__"" Stretch=""Uniform"" StrokeLineJoin=""Round""
              StrokeStartLineCap=""Round"" StrokeEndLineCap=""Round""/>
      </Button>
    </Grid>
  </Border>
</DataTemplate>";
    }
}
