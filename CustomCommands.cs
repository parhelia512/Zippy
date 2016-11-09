using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Zippy
{
    public static class CustomCommands
    {
        public static RoutedCommand NewGZip = new RoutedCommand();
        public static RoutedCommand ExtractTo = new RoutedCommand();
    }

}
