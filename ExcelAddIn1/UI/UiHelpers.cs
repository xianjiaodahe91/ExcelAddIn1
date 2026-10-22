using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace ExcelAddIn1.UI
{
    internal static class UiHelpers
    {
        public static T FindAncestor<T>(DependencyObject node) where T : DependencyObject
        {
            var current = node;
            while (current != null)
            {
                var typed = current as T;
                if (typed != null) return typed;
                current = GetParent(current);
            }
            return null;
        }

        private static DependencyObject GetParent(DependencyObject node)
        {
            var visual = node as Visual;
            if (visual != null)
            {
                DependencyObject parent = VisualTreeHelper.GetParent(visual);
                if (parent != null) return parent;
            }
            var element = node as FrameworkElement;
            if (element != null && element.Parent != null) return element.Parent;
            var content = node as FrameworkContentElement;
            if (content != null) return content.Parent;
            return null;
        }

        /// <summary>判断点击是否落在某个按钮上（用于区分“点行跳转”和“点眼睛/图钉”）。</summary>
        public static ButtonBase FindButton(DependencyObject node)
        {
            return FindAncestor<ButtonBase>(node);
        }
    }
}
