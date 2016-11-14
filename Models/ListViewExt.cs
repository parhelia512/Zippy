using System;
using System.Windows.Controls;
using System.Windows.Input;

namespace Zippy.Models
{
    public static class ListViewExt
    {
        public static void PerformDoubleClick(this ListView tExt)
        {
            MouseButtonEventArgs e = new MouseButtonEventArgs(Mouse.PrimaryDevice,
         0,
         MouseButton.Left)
            { RoutedEvent = Control.MouseDoubleClickEvent };

            tExt.RaiseEvent(e);
        }
    }
}
