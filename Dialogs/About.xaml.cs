using SharpCompress.Archives;
using SharpCompress.Writers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using SharpCompress.Archives.Zip;
using System.Diagnostics;

namespace Zippy.Dialogs
{
    /// <summary>
    /// Interaction logic for About.xaml, Includes the basic compression functions.
    /// </summary>
    public partial class About
    {
        public About()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("EULA.rtf");
        }
    }
}
