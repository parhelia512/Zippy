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

namespace Zippy.Dialogs
{
    /// <summary>
    /// Interaction logic for ProgressDialog.xaml, Includes the basic compression functions.
    /// </summary>
    public partial class ProgressDialog
    {
        public ProgressDialog()
        {
            InitializeComponent();
        }
        public void StartExtract(List<IArchiveEntry> entries, string dir)
        {
            Show();
            progress.Maximum = entries.Count;
            var th = new Thread(delegate ()
            {
                entries.ToList().ForEach(delegate (IArchiveEntry entry)
                {
                    var path = Path.Combine(dir, entry.Key.Replace("/", "\\"));
                    entry.WriteToFile(path);
                    progress.Dispatcher.Invoke(delegate ()
                    {
                        progress.Value += 1;
                    });
                });
                Dispatcher.Invoke(Close);
            });
            th.Start();
        }

        public void StartCompress(ZipArchive archive, string fileName, WriterOptions writerOptions)
        {
            Message.Text = "Compressing...";
            this.Show();
            progress.IsIndeterminate = true;
            var th = new Thread(delegate ()
            {
                archive.SaveTo(fileName, writerOptions);
                Dispatcher.Invoke(Close);
            });
            th.Start();
        }

    }
}
