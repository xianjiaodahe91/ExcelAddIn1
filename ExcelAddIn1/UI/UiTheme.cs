using System.Windows;
using System.Windows.Markup;
using System.Windows.Media;

namespace ExcelAddIn1.UI
{
    /// <summary>
    /// 统一的视觉规范：颜色、字体、图标几何、控件样式。
    /// 样式用 XamlReader 在运行时解析，因此不需要项目支持 XAML 编译（VSTO 旧式 csproj 无需改造）。
    /// </summary>
    internal static class UiTheme
    {
        // ---------- 颜色 ----------
        public static readonly Brush PaneBg = Frozen(0xFF, 0xFF, 0xFF);
        public static readonly Brush RailBg = Frozen(0xFA, 0xFA, 0xF9);
        public static readonly Brush HeaderBg = Frozen(0xFF, 0xFF, 0xFF);
        public static readonly Brush Divider = Frozen(0xE6, 0xE6, 0xE4);
        public static readonly Brush DividerSoft = Frozen(0xF0, 0xF0, 0xEE);
        public static readonly Brush TextPrimary = Frozen(0x32, 0x32, 0x32);
        public static readonly Brush TextSecondary = Frozen(0x77, 0x77, 0x77);
        public static readonly Brush TextDisabled = Frozen(0xAE, 0xAE, 0xAE);
        public static readonly Brush Accent = Frozen(0x21, 0x73, 0x46);
        public static readonly Brush AccentSoft = Frozen(0xE9, 0xF3, 0xED);
        public static readonly Brush RowHover = Frozen(0xF2, 0xF2, 0xF0);
        public static readonly Brush RowActiveBg = Frozen(0xC6, 0xE0, 0xB4);
        public static readonly Brush RowActiveHoverBg = Frozen(0xBB, 0xD8, 0xA7);
        public static readonly Brush IconNormal = Frozen(0x4A, 0x4A, 0x4A);
        public static readonly Brush IconMuted = Frozen(0x8A, 0x8A, 0x86);
        public static readonly Brush InputBorder = Frozen(0xD5, 0xD5, 0xD2);

        public static readonly FontFamily Font = new FontFamily("Microsoft YaHei UI, Microsoft YaHei, Segoe UI, SimSun");

        // 字号：想整体调大/调小，只改这四个值即可
        public const double FontTitle = 12.0;   // 页面标题
        public const double FontBody = 11.5;    // 正文 / 列表项 / 按钮
        public const double FontMid = 11.0;     // 信息页的标签与值
        public const double FontSmall = 10.5;   // 计数、说明文字

        // 图标尺寸：同样集中在这里
        public const double IconRail = 15.0;    // 左侧图栏
        public const double IconTool = 16.0;    // 页面右上工具按钮
        public const double IconRow = 15.0;     // 列表行内图标 / 眼睛
        public const double IconSmall = 14.0;   // 次要图标（搜索框等）

        /// <summary>把字号转成 XAML 里可用的文本（不变式小数）。</summary>
        internal static string Num(double value)
        {
            return value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }

        public const double RowHeight = 26.0;
        public const double RailWidth = 36.0;
        public const double HeaderHeight = 32.0;

        private static Brush Frozen(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        // ---------- 图标（24x24 视口，描边式） ----------
        public static class Icons
        {
            public const string Sheets = "M4,5 H20 V19 H4 Z M4,9.5 H20 M4,14 H20 M9.5,5 V19 M15,5 V19";
            public const string Info = "M12,3.6 A8.4,8.4 0 1,1 11.99,3.6 Z M12,11 V16.4 M12,7.9 V8.3";
            public const string Settings = "M19.4,15 A1.65,1.65 0 0,0 19.73,16.82 L19.79,16.88 A2,2 0 0,1 19.79,19.71 A2,2 0 0,1 16.96,19.71 L16.9,19.65 A1.65,1.65 0 0,0 15.08,19.32 A1.65,1.65 0 0,0 14.08,20.83 L14.08,21 A2,2 0 0,1 12.08,23 A2,2 0 0,1 10.08,21 L10.08,20.91 A1.65,1.65 0 0,0 9,19.4 A1.65,1.65 0 0,0 7.18,19.73 L7.12,19.79 A2,2 0 0,1 4.29,19.79 A2,2 0 0,1 4.29,16.96 L4.35,16.9 A1.65,1.65 0 0,0 4.68,15.08 A1.65,1.65 0 0,0 3.17,14.08 L3,14.08 A2,2 0 0,1 1,12.08 A2,2 0 0,1 3,10.08 L3.09,10.08 A1.65,1.65 0 0,0 4.6,9 A1.65,1.65 0 0,0 4.27,7.18 L4.21,7.12 A2,2 0 0,1 4.21,4.29 A2,2 0 0,1 7.04,4.29 L7.1,4.35 A1.65,1.65 0 0,0 8.92,4.68 L9,4.68 A1.65,1.65 0 0,0 10,3.17 L10,3 A2,2 0 0,1 12,1 A2,2 0 0,1 14,3 L14,3.09 A1.65,1.65 0 0,0 15,4.6 A1.65,1.65 0 0,0 16.82,4.27 L16.88,4.21 A2,2 0 0,1 19.71,4.21 A2,2 0 0,1 19.71,7.04 L19.65,7.1 A1.65,1.65 0 0,0 19.32,8.92 L19.32,9 A1.65,1.65 0 0,0 20.83,10 L21,10 A2,2 0 0,1 23,12 A2,2 0 0,1 21,14 L20.91,14 A1.65,1.65 0 0,0 19.4,15 Z M12,8.5 A3.5,3.5 0 1,1 11.99,8.5 Z";
            public const string Search = "M10.8,4.2 A6.1,6.1 0 1,1 10.79,4.2 Z M15.3,15.3 L20,20";
            public const string Plus = "M12,5.2 V18.8 M5.2,12 H18.8";
            public const string Pencil = "M4,20 L4.6,16.2 L15.8,5 L19,8.2 L7.8,19.4 Z M14.2,6.6 L17.4,9.8";
            public const string Trash = "M5.4,7 H18.6 M9.4,7 V4.8 H14.6 V7 M6.6,7 L7.5,19.5 H16.5 L17.4,7 M10.2,10.4 V16.2 M13.8,10.4 V16.2";
            public const string Eye = "M2.6,12 C5,7.7 8.8,5.7 12,5.7 C15.2,5.7 19,7.7 21.4,12 C19,16.3 15.2,18.3 12,18.3 C8.8,18.3 5,16.3 2.6,12 Z M12,9.1 A2.9,2.9 0 1,1 11.99,9.1 Z";
            public const string EyeOff = "M3.4,3.4 L20.6,20.6 M7.2,7.4 C5.5,8.8 4,10.4 2.6,12 C5,16.3 8.8,18.3 12,18.3 C13.5,18.3 15.1,17.9 16.6,17.1 M10.4,6 C10.9,5.8 11.4,5.7 12,5.7 C15.2,5.7 19,7.7 21.4,12 C20.6,13.4 19.7,14.5 18.8,15.4";
            public const string Refresh = "M20,11.6 A8,8 0 1,1 17.6,6.1 M17.6,3.2 V6.6 H14.2";
            public const string Chart = "M4.5,19.5 V10.5 H8.2 V19.5 Z M10.1,19.5 V4.5 H13.8 V19.5 Z M15.7,19.5 V13 H19.4 V19.5 Z";
            public const string ChevronRight = "M9.5,6 L15.5,12 L9.5,18";
            public const string Grid = "M4,4 H10 V10 H4 Z M14,4 H20 V10 H14 Z M4,14 H10 V20 H4 Z M14,14 H20 V20 H14 Z";
        }

        /// <summary>创建一个描边图标。</summary>
        public static System.Windows.Shapes.Path Icon(string data, double size, Brush stroke, double thickness)
        {
            var path = new System.Windows.Shapes.Path();
            path.Data = Geometry.Parse(data);
            path.Stroke = stroke;
            path.StrokeThickness = thickness;
            path.StrokeStartLineCap = PenLineCap.Round;
            path.StrokeEndLineCap = PenLineCap.Round;
            path.StrokeLineJoin = PenLineJoin.Round;
            path.Fill = null;
            path.Stretch = Stretch.Uniform;
            path.Width = size;
            path.Height = size;
            path.HorizontalAlignment = HorizontalAlignment.Center;
            path.VerticalAlignment = VerticalAlignment.Center;
            path.SnapsToDevicePixels = true;
            return path;
        }

        // ---------- 运行时解析的样式表 ----------
        private static ResourceDictionary _styles;

        public static ResourceDictionary Styles
        {
            get
            {
                if (_styles == null)
                {
                    var xaml = StylesXaml
                        .Replace("__BODY__", Num(FontBody))
                        .Replace("__RAIL__", Num(RailWidth))
                        .Replace("__RAILBTN__", Num(RailWidth - 2))
                        .Replace("__ROWICON__", Num(IconRow));
                    _styles = (ResourceDictionary)XamlReader.Parse(xaml);
                }
                return _styles;
            }
        }

        private const string StylesXaml = @"
<ResourceDictionary xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
                    xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">

  <!-- 普通图标按钮 -->
  <Style x:Key=""IconButton"" TargetType=""Button"">
    <Setter Property=""Width"" Value=""26""/>
    <Setter Property=""Height"" Value=""26""/>
    <Setter Property=""Background"" Value=""Transparent""/>
    <Setter Property=""BorderThickness"" Value=""0""/>
    <Setter Property=""Cursor"" Value=""Hand""/>
    <Setter Property=""Focusable"" Value=""False""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""Button"">
          <Border x:Name=""bd"" Background=""{TemplateBinding Background}"" CornerRadius=""3"" SnapsToDevicePixels=""True"">
            <ContentPresenter HorizontalAlignment=""Center"" VerticalAlignment=""Center""/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property=""IsMouseOver"" Value=""True"">
              <Setter TargetName=""bd"" Property=""Background"" Value=""#EFEFED""/>
            </Trigger>
            <Trigger Property=""IsPressed"" Value=""True"">
              <Setter TargetName=""bd"" Property=""Background"" Value=""#E0E0DC""/>
            </Trigger>
            <Trigger Property=""IsEnabled"" Value=""False"">
              <Setter Property=""Opacity"" Value=""0.35""/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- 左侧栏图标按钮（RadioButton，选中时左侧色条 + 淡绿底） -->
  <Style x:Key=""RailButton"" TargetType=""RadioButton"">
    <Setter Property=""Width"" Value=""__RAIL__""/>
    <Setter Property=""Height"" Value=""__RAILBTN__""/>
    <Setter Property=""Cursor"" Value=""Hand""/>
    <Setter Property=""Focusable"" Value=""False""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""RadioButton"">
          <Grid Background=""Transparent"">
            <Border x:Name=""bg"" Background=""Transparent""/>
            <Rectangle x:Name=""bar"" Width=""3"" HorizontalAlignment=""Left"" Fill=""#217346"" Visibility=""Collapsed""/>
            <ContentPresenter x:Name=""cp"" HorizontalAlignment=""Center"" VerticalAlignment=""Center""/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property=""IsMouseOver"" Value=""True"">
              <Setter TargetName=""bg"" Property=""Background"" Value=""#F1F1EF""/>
            </Trigger>
            <Trigger Property=""IsChecked"" Value=""True"">
              <Setter TargetName=""bg"" Property=""Background"" Value=""#E9F3ED""/>
              <Setter TargetName=""bar"" Property=""Visibility"" Value=""Visible""/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- 普通按钮（文字） -->
  <Style x:Key=""TextButton"" TargetType=""Button"">
    <Setter Property=""Height"" Value=""26""/>
    <Setter Property=""Padding"" Value=""10,0""/>
    <Setter Property=""Background"" Value=""#FFFFFF""/>
    <Setter Property=""BorderBrush"" Value=""#D5D5D2""/>
    <Setter Property=""BorderThickness"" Value=""1""/>
    <Setter Property=""Foreground"" Value=""#323232""/>
    <Setter Property=""FontSize"" Value=""__BODY__""/>
    <Setter Property=""Cursor"" Value=""Hand""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""Button"">
          <Border x:Name=""bd"" Background=""{TemplateBinding Background}"" BorderBrush=""{TemplateBinding BorderBrush}""
                  BorderThickness=""{TemplateBinding BorderThickness}"" CornerRadius=""3"" SnapsToDevicePixels=""True"">
            <ContentPresenter HorizontalAlignment=""Center"" VerticalAlignment=""Center"" Margin=""{TemplateBinding Padding}""/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property=""IsMouseOver"" Value=""True"">
              <Setter TargetName=""bd"" Property=""Background"" Value=""#F5F7F4""/>
              <Setter TargetName=""bd"" Property=""BorderBrush"" Value=""#217346""/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- 搜索框 -->
  <Style x:Key=""SearchBox"" TargetType=""TextBox"">
    <Setter Property=""BorderThickness"" Value=""0""/>
    <Setter Property=""Background"" Value=""Transparent""/>
    <Setter Property=""Foreground"" Value=""#323232""/>
    <Setter Property=""CaretBrush"" Value=""#217346""/>
    <Setter Property=""VerticalContentAlignment"" Value=""Center""/>
    <Setter Property=""FontSize"" Value=""__BODY__""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""TextBox"">
          <ScrollViewer x:Name=""PART_ContentHost"" Margin=""0,1,0,0"" VerticalAlignment=""Center""/>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- 复选框 -->
  <Style x:Key=""PaneCheckBox"" TargetType=""CheckBox"">
    <Setter Property=""Foreground"" Value=""#323232""/>
    <Setter Property=""FontSize"" Value=""__BODY__""/>
    <Setter Property=""Cursor"" Value=""Hand""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""CheckBox"">
          <StackPanel Orientation=""Horizontal"" Background=""Transparent"">
            <Border x:Name=""box"" Width=""14"" Height=""14"" CornerRadius=""2""
                    BorderBrush=""#B8B8B4"" BorderThickness=""1"" Background=""#FFFFFF""
                    VerticalAlignment=""Center"">
              <Path x:Name=""check"" Data=""M3,7.2 L5.9,10.2 L11.2,4"" Stroke=""#FFFFFF""
                    StrokeThickness=""1.6"" Stretch=""None"" Visibility=""Collapsed""
                    StrokeStartLineCap=""Round"" StrokeEndLineCap=""Round"" StrokeLineJoin=""Round""/>
            </Border>
            <ContentPresenter Margin=""7,0,0,0"" VerticalAlignment=""Center""/>
          </StackPanel>
          <ControlTemplate.Triggers>
            <Trigger Property=""IsChecked"" Value=""True"">
              <Setter TargetName=""box"" Property=""Background"" Value=""#217346""/>
              <Setter TargetName=""box"" Property=""BorderBrush"" Value=""#217346""/>
              <Setter TargetName=""check"" Property=""Visibility"" Value=""Visible""/>
            </Trigger>
            <Trigger Property=""IsMouseOver"" Value=""True"">
              <Setter TargetName=""box"" Property=""BorderBrush"" Value=""#217346""/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- 细滚动条 -->
  <Style TargetType=""ScrollBar"">
    <Setter Property=""Width"" Value=""9""/>
    <Setter Property=""Background"" Value=""Transparent""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""ScrollBar"">
          <Grid Background=""Transparent"">
            <Track x:Name=""PART_Track"" IsDirectionReversed=""True"" Focusable=""False"">
              <Track.Thumb>
                <Thumb MinHeight=""24"">
                  <Thumb.Template>
                    <ControlTemplate TargetType=""Thumb"">
                      <Border x:Name=""tb"" Background=""#CDCDC8"" CornerRadius=""2.5"" Margin=""2,0,2,0""/>
                      <ControlTemplate.Triggers>
                        <Trigger Property=""IsMouseOver"" Value=""True"">
                          <Setter TargetName=""tb"" Property=""Background"" Value=""#A8A8A2""/>
                        </Trigger>
                      </ControlTemplate.Triggers>
                    </ControlTemplate>
                  </Thumb.Template>
                </Thumb>
              </Track.Thumb>
            </Track>
          </Grid>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- 右键菜单 -->
  <Style TargetType=""ContextMenu"">
    <Setter Property=""Background"" Value=""#FFFFFF""/>
    <Setter Property=""BorderBrush"" Value=""#DCDCD8""/>
    <Setter Property=""BorderThickness"" Value=""1""/>
    <Setter Property=""Padding"" Value=""2""/>
    <Setter Property=""HasDropShadow"" Value=""True""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""ContextMenu"">
          <Border Background=""{TemplateBinding Background}"" BorderBrush=""{TemplateBinding BorderBrush}""
                  BorderThickness=""{TemplateBinding BorderThickness}"" CornerRadius=""3"" Padding=""{TemplateBinding Padding}"">
            <ItemsPresenter/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType=""MenuItem"">
    <Setter Property=""Foreground"" Value=""#323232""/>
    <Setter Property=""FontSize"" Value=""__BODY__""/>
    <Setter Property=""Padding"" Value=""10,5""/>
    <Setter Property=""Template"">
      <Setter.Value>
        <ControlTemplate TargetType=""MenuItem"">
          <Border x:Name=""bd"" Background=""Transparent"" CornerRadius=""3"" Padding=""{TemplateBinding Padding}"">
            <ContentPresenter ContentSource=""Header"" RecognizesAccessKey=""True"" VerticalAlignment=""Center""/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property=""IsHighlighted"" Value=""True"">
              <Setter TargetName=""bd"" Property=""Background"" Value=""#E9F3ED""/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

</ResourceDictionary>";
    }
}
