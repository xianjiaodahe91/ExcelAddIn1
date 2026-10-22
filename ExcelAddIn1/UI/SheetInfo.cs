using System;
using System.ComponentModel;
using System.Windows.Media;

namespace ExcelAddIn1.UI
{
    /// <summary>工作表类型，仅用于显示图标。</summary>
    public enum SheetKind
    {
        Worksheet,
        Chart,
        Macro,
        Dialog,
        Other
    }

    /// <summary>
    /// 工作表的只读视图模型。UI 层不引用任何 Excel 类型，
    /// 由加载项把 Excel.Worksheet 映射成这个对象。
    /// </summary>
    public sealed class SheetInfo : INotifyPropertyChanged
    {
        private bool _visible = true;
        private bool _isActive;

        public SheetInfo(string name, int index, SheetKind kind, Color? tabColor)
        {
            Name = name;
            Index = index;
            Kind = kind;
            TabColor = tabColor;
        }

        public string Name { get; private set; }

        public int Index { get; private set; }

        public SheetKind Kind { get; private set; }

        public Color? TabColor { get; private set; }

        /// <summary>Excel 没有给标签上色时用透明，避免出现黑块。</summary>
        public Brush TabBrush
        {
            get
            {
                if (TabColor.HasValue && TabColor.Value.A > 0)
                {
                    var brush = new SolidColorBrush(TabColor.Value);
                    brush.Freeze();
                    return brush;
                }
                return Brushes.Transparent;
            }
        }

        public string KindIconData
        {
            get
            {
                switch (Kind)
                {
                    case SheetKind.Chart:
                        return UiTheme.Icons.Chart;
                    case SheetKind.Macro:
                    case SheetKind.Dialog:
                    case SheetKind.Other:
                        return UiTheme.Icons.Grid;
                    default:
                        return UiTheme.Icons.Sheets;
                }
            }
        }

        public bool Visible
        {
            get { return _visible; }
            set
            {
                if (_visible == value) return;
                _visible = value;
                Raise("Visible");
                Raise("EyeIconData");
                Raise("EyeTooltip");
                Raise("NameOpacity");
            }
        }

        public bool IsActive
        {
            get { return _isActive; }
            set
            {
                if (_isActive == value) return;
                _isActive = value;
                Raise("IsActive");
            }
        }

        public string EyeIconData
        {
            get { return _visible ? UiTheme.Icons.Eye : UiTheme.Icons.EyeOff; }
        }

        public string EyeTooltip
        {
            get { return _visible ? "隐藏工作表" : "取消隐藏"; }
        }

        /// <summary>被隐藏的工作表名字淡一点，和可见表区分。</summary>
        public double NameOpacity
        {
            get { return _visible ? 1.0 : 0.45; }
        }

        public string DisplayName
        {
            get { return Name; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void Raise(string property)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(property));
        }
    }

    /// <summary>工作表相关动作的事件参数。</summary>
    public sealed class SheetActionEventArgs : EventArgs
    {
        public SheetActionEventArgs(SheetInfo sheet)
        {
            Sheet = sheet;
        }

        public SheetInfo Sheet { get; private set; }
    }
}
