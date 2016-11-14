#region usages

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Archives.SevenZip;
using SharpCompress.Archives.Tar;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using SharpCompress.Readers;
using SharpCompress.Writers;
using Zippy.Dialogs;
using Zippy.Models;
using Zippy.Utils;
using Image = System.Windows.Controls.Image;
using Point = System.Windows.Point;

#endregion

namespace Zippy
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private const string DummyNode = "none";
        private readonly string _temporaryDirectory;
        private string _archivePath;
        private Point _start;
        private ArchiveTypes _type;

        public MainWindow()
        {
            InitializeComponent();
            var directoryInfo = new DirectoryInfo(Path.GetTempPath()).Parent;
            if (directoryInfo != null)
                _temporaryDirectory = directoryInfo.FullName + "\\Zippy\\";
        }

        public bool ArchiveMode { get; set; }

        public ImageSource BmtoImgSource(Bitmap source)
        {
            return Imaging.CreateBitmapSourceFromHBitmap(
                source.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(_temporaryDirectory))
            {
                Directory.CreateDirectory(_temporaryDirectory);
            }
            //Clear the temporary directory
            foreach (var file in Directory.GetFiles(_temporaryDirectory))
            {
                File.Delete(file);
            }
            //Add the treeview items for every drive
            foreach (var s in Directory.GetLogicalDrives())
            {
                var item = new TreeViewItem
                {
                    Header = GetStackpanel(s),
                    Tag = s,
                    FontWeight = FontWeights.Normal
                };
                item.Items.Add(DummyNode);
                item.Expanded += folder_Expanded;
                FoldersItem.Items.Add(item);
            }
            //Adds a treeview item for the home folder:
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var item2 = new TreeViewItem
            {
                Header = GetStackpanel(home),
                Tag = home,
                FontWeight = FontWeights.Normal
            };
            item2.Items.Add(DummyNode);
            item2.Expanded += folder_Expanded;
            item2.IsSelected = true;
            FoldersItem.Items.Add(item2);
            //Checks if Zippy has command line args, and if so, it uses them:
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                var path = args[1];
                if (File.Exists(path))
                {
                    _type = IsArchive(path);
                    if (_type == ArchiveTypes.None)
                    {
                        var parent = new FileInfo(path).Directory?.FullName;
                        SelectFolder(parent);
                    }
                    else
                    {
                        OpenArchive(path);
                    }
                }
                else if (Directory.Exists(path))
                {
                    var info = new DirectoryInfo(path);
                    Zip(ArchiveTypes.Zip, null,
                        new List<FileItem> { new FileItem(path) { IsDir = true, Path = path, Name = info.Name } }, true);
                }
                else
                    switch (path)
                    {
                        case "-e":
                        case "-h":
                            path = args[2];
                            var info = new FileInfo(path);

                            if (IsArchive(path) != ArchiveTypes.None)
                                if (args[1] == "-e")
                                {
                                    var parent = new FileInfo(path).Directory?.FullName;
                                    var dialog = new CommonOpenFileDialog
                                    {
                                        IsFolderPicker = true,
                                        InitialDirectory = parent
                                    };

                                    var result = dialog.ShowDialog();
                                    if (result == CommonFileDialogResult.Ok)
                                        if (Directory.Exists(dialog.FileName))
                                            WriteToDirectory(ArchiveFactory.Open(path), dialog.FileName);
                                }
                                else
                                {
                                    var toDir = info.Directory?.FullName;
                                    WriteToDirectory(ArchiveFactory.Open(path), toDir);
                                    if (toDir != null) Process.Start(toDir);
                                }
                            break;
                        case "-a":
                            var saveAs = new CommonSaveFileDialog();
                            saveAs.Filters.Add(new CommonFileDialogFilter("Zip Archive", ".zip"));
                            saveAs.DefaultExtension = ".zip";
                            var files = args.Skip(2).ToArray();
                            saveAs.InitialDirectory = new FileInfo(files[0]).Directory?.FullName;
                            if (saveAs.ShowDialog() == CommonFileDialogResult.Ok)
                            {
                                var archive = ZipArchive.Create();
                                var tmp = _temporaryDirectory + "\\" + DateTime.Now.Ticks;
                                Directory.CreateDirectory(tmp);
                                foreach (var file in files)
                                    if (File.Exists(file))
                                        File.Copy(file, Path.Combine(tmp, new FileInfo(file).Name));
                                    else if (Directory.Exists(file))
                                        DirectoryCopy(file, Path.Combine(tmp, new DirectoryInfo(file).Name));
                                archive.AddAllFromDirectory(tmp);
                                if (File.Exists(saveAs.FileName))
                                    File.Delete(saveAs.FileName);

                                SaveTo(archive, saveAs.FileName, new WriterOptions(CompressionType.LZMA));
                                try
                                {
                                    Directory.Delete(tmp, true);
                                }
                                catch
                                {
                                    // TODO: error handling
                                }
                            }
                            break;
                        default:
                            break;
                    }
            }
        }


        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs = true)
        {
            var dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);

            var dirs = dir.GetDirectories();
            if (!Directory.Exists(destDirName))
                Directory.CreateDirectory(destDirName);

            var files = dir.GetFiles();
            foreach (var file in files)
            {
                var temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }
            if (!copySubDirs) return;
            {
                foreach (var subdir in dirs)
                {
                    var temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath);
                }
            }
        }

        private WrapPanel GetStackpanel(string home)
        {
            var sp = new WrapPanel { IsItemsHost = false };
            sp.Children.Add(new Image
            {
                Source = BmtoImgSource(ImageUtilities.GetRegisteredIcon(home).ToBitmap()),
                Width = 16
            });
            sp.Children.Add(new TextBlock { Text = new DirectoryInfo(home).Name });
            return sp;
        }

        private void folder_Expanded(object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)sender;
            if (item.Items.Count != 1) return;
            item.Items.Clear();
            try
            {
                foreach (var s in Directory.GetDirectories(item.Tag.ToString()).Where(d => !new DirectoryInfo(d).Attributes.HasFlag(FileAttributes.Hidden)))
                {
                    var subitem = new TreeViewItem
                    {
                        Header = GetStackpanel(s),
                        Tag = s,
                        FontWeight = FontWeights.Normal
                    };
                    subitem.Items.Add(DummyNode);
                    subitem.Expanded += folder_Expanded;
                    item.Items.Add(subitem);
                }
            }
            catch
            {
                // TODO: error handling
            }
        }

        public ArchiveTypes IsArchive(string path)
        {
            var type = ArchiveTypes.None;

            if (RarArchive.IsRarFile(path))
                type = ArchiveTypes.RAR;
            if (ZipArchive.IsZipFile(path))
                type = ArchiveTypes.Zip;
            if (SevenZipArchive.IsSevenZipFile(path))
                type = ArchiveTypes.SevenZip;
            if (TarArchive.IsTarFile(path))
                type = ArchiveTypes.Deflate;

            return type;
        }

        private void foldersItem_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            ArchiveMode = false;
            Refresh();
        }

        private void listView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (FileItem)((ListView)sender).SelectedItem;
            if (ArchiveMode)
            {
                if (!item.IsDir)
                {
                    var i = (IArchiveEntry)item.Entry;

                    var path = _temporaryDirectory + DateTime.Now.ToFileTime();
                    Directory.CreateDirectory(path);
                    i.WriteToDirectory(path, new ExtractionOptions { ExtractFullPath = true });
                    Process.Start(path + "\\" + i.Key);
                }
                else
                {
                    var i = (IArchiveEntry)item.Entry;
                    var archive = i.Archive;
                    var path = _temporaryDirectory + DateTime.Now.ToFileTime();
                    Directory.CreateDirectory(path);
                    WriteToDirectory(archive, path);
                    ArchiveMode = false;
                    Process.Start(path);
                }
            }
            else
            {
                var path = item.Path;
                _type = IsArchive(path);
                OpenArchive(path);
            }
        }

        private void OpenArchive(string path)
        {
            if (Directory.Exists(path))
            {
                SelectFolder(path);
            }
            else if (_type != ArchiveTypes.None)
            {
                listView.Items.Clear();
                var archive = ArchiveFactory.Open(path);
                _archivePath = path;
                archive.Entries.ToList().ForEach(delegate (IArchiveEntry values)
                {
                    var extension = "." + values.Key.Split('.').Last();
                    var isDir = values.IsDirectory;
                    var exePath = IconManager.FindIconForFilename(extension, false);
                    listView.Items.Add(new FileItem(exePath) { Name = values.Key, Entry = values, IsDir = isDir });
                });
                ArchiveMode = true;
            }
            else
            {
                Process.Start(path);
            }
        }

        public static string Extension_GetExePath(string strExtension)
        {
            var strExePath = "C:\\Windows";
            try
            {
                //We need a leading dot, so add it if it's missing.
                if (!strExtension.StartsWith(".")) strExtension = "." + strExtension;

                //Get the class-name associated with the passed extension
                var rkClassName
                    = Registry.ClassesRoot.OpenSubKey(strExtension);
                //Exit, if not found
                if (rkClassName == null) return string.Empty;
                var strClassName = rkClassName.GetValue("").ToString();

                //Get the shell-command for the retrieved executable 
                //This key is found at HKCR\[ClassName]\shell\open\command\(Default)
                //One or more of the paths may be missing, so each of them is being tested
                //separately.
                var rkShellCommandRoot
                    = Registry.ClassesRoot.OpenSubKey(strClassName);

                var rkShell
                    = rkShellCommandRoot?.OpenSubKey("shell");
                if (rkShell == null) return string.Empty;

                var rkOpen
                    = rkShell.OpenSubKey("open");

                var rkCommand
                    = rkOpen?.OpenSubKey("command");
                if (rkCommand == null) return string.Empty;

                var strShellCommand = rkCommand.GetValue("").ToString();

                //The shell-command may contain additional parameters and may be wrapped in double quotes,
                //so parse out the exe-path.
                strExePath = strShellCommand.StartsWith(@"")
                    ? strShellCommand.Split('"')[1]
                    : strShellCommand.Split(' ')[0];
            }
            catch
            {
                // TODO: error handling
            }
            return strExePath;
        }


        public void ZipSelected(ArchiveTypes type, WriterOptions options1 = null)
        {
            var items = listView.SelectedItems;
            Zip(type, options1, items);
        }

        private void Zip(ArchiveTypes type, WriterOptions options1, IList items, bool close = false)
        {
            if (items.Count <= 0) return;
            switch (type)
            {
                case ArchiveTypes.Zip:
                    {
                        if (close)
                            IsEnabled = false;
                        var ar = ZipArchive.Create();
                        var p = "";
                        foreach (var x in items)
                        {
                            var item = (FileItem)x;
                            if (Directory.Exists(item.Path))
                                ar.AddAllFromDirectory(item.Path);
                            else
                                ar.AddEntry(item.Name, File.Open(item.Path, FileMode.Open, FileAccess.Read));
                            p = item.Path;
                        }
                        var fpath = new FileInfo(p).Directory + "\\" + ((FileItem)items[0]).Name.Split('.')[0];
                        if (File.Exists(fpath + ".zip"))
                            fpath += "_";

                        if (options1 == null) options1 = new WriterOptions(CompressionType.LZMA);
                        SaveTo(ar, fpath + ".zip", options1);
                    }
                    break;
                case ArchiveTypes.Deflate:
                    {
                        var ar = ZipArchive.Create();
                        var p = "";
                        foreach (var x in items)
                        {
                            var item = (FileItem)x;
                            if (Directory.Exists(item.Path))
                                ar.AddAllFromDirectory(item.Path);
                            else
                                ar.AddEntry(item.Name, File.OpenRead(item.Path));
                            p = item.Path;
                        }
                        Cursor = Cursors.Wait;
                        loading.IsIndeterminate = true;
                        var thread = new Thread(delegate ()
                        {
                            var v = new FileInfo(p).Directory + "\\" + ((FileItem)items[0]).Name.Split('.')[0] + ".zip";
                            using (var _open = File.Open(v, FileMode.OpenOrCreate))
                            {
                                ar.SaveTo(_open, new WriterOptions(CompressionType.Deflate));
                                _open.Close();
                            }
                            Dispatcher.Invoke(delegate
                            {
                                Cursor = Cursors.Arrow;
                                loading.IsIndeterminate = false;
                            });
                        });
                        thread.Start();
                    }
                    break;
                case ArchiveTypes.None:
                    break;
                case ArchiveTypes.RAR:
                    break;
                case ArchiveTypes.SevenZip:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }
            Refresh();
        }

        private void goto_Click(object sender, RoutedEventArgs e)
        {
            var path = pathBox.Text;
            SelectFolder(path);
        }

        public void SelectFolder(string path)
        {
            if (!Directory.Exists(path)) return;
            listView.Visibility = Visibility.Hidden;
            new Thread(delegate ()
            {
                var nodes = path.Split('\\');
                var items = FoldersItem.Items;
                for (var i = 0; i < nodes.Length; i++)
                {
                    var node = nodes[i];
                    if (i == 0)
                        node += "\\";

                    for (int[] ix = { 0 }; ix[0] < items.Count; ix[0]++)
                    {
                        Dispatcher.Invoke(delegate
                        {
                            var obj = (TreeViewItem)items[ix[0]];

                            var header = ((WrapPanel)obj.Header).Children[1] as TextBlock;
                            if (header?.Text != node) return;
                            obj.IsExpanded = true;
                            items = obj.Items;
                            ix[0] = 0;
                            obj.IsSelected = true;
                        });
                    }
                }
                listView.Dispatcher.Invoke(delegate { listView.Visibility = Visibility.Visible; });
            }).Start();
        }

        private void listView_MouseMove(object sender, MouseEventArgs e)
        {
            var mpos = e.GetPosition(null);
            var diff = _start - mpos;

            if (e.LeftButton != MouseButtonState.Pressed ||
                !(Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance) ||
                !(Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)) return;
            if (listView.SelectedItems.Count == 0)
                return;

            var files = new string[listView.SelectedItems.Count];
            var ix = 0;
            foreach (var nextSel in listView.SelectedItems)
            {
                if (!ArchiveMode)
                {
                    files[ix] = ((FileItem)nextSel).Path;
                }
                else
                {
                    var fileItem = ((FileItem)nextSel);
                    var path = Path.Combine(_temporaryDirectory, ((IArchiveEntry)fileItem.Entry).Key);
                    ExtractFileItem(_temporaryDirectory, fileItem);
                    files[ix] = path;
                }
                ix++;
            }
            var dataFormat = DataFormats.FileDrop;
            var dataObject = new DataObject(dataFormat, files);
            DragDrop.DoDragDrop(listView, dataObject, DragDropEffects.Copy);
        }

        private void listView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _start = e.GetPosition(null);
        }


        public void Refresh()
        {
            try
            {
                var path = ((TreeViewItem)FoldersItem.SelectedItem).Tag.ToString();
                pathBox.Text = path;
                listView.Items.Clear();
                Directory.GetDirectories(path)
                    .Where(d => !new DirectoryInfo(d).Attributes.HasFlag(FileAttributes.Hidden))
                    .ToList()
                    .ForEach(
                        delegate (string s)
                        {
                            listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
                        });
                Directory.GetFiles(path)
                    .Where(f => !new FileInfo(f).Attributes.HasFlag(FileAttributes.Hidden))
                    .ToList()
                    .ForEach(
                        delegate (string s)
                        {
                            listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
                        });
            }
            catch
            {
                // TODO: error handling
            }
        }

        private void folder_up_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ArchiveMode)
                {
                    ArchiveMode = false;
                    SelectFolder(pathBox.Text);
                }
                else
                {
                    var path = pathBox.Text;
                    var directoryInfo = new DirectoryInfo(path).Parent;
                    if (directoryInfo == null) return;
                    var parentPath = directoryInfo.FullName;
                    SelectFolder(parentPath);
                }
            }
            catch
            {
                // TODO: error handling
            }
        }

        private void newZip_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ZipSelected(ArchiveTypes.Zip);
        }

        private void newZipG_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ZipSelected(ArchiveTypes.Deflate);
        }

        private void delete_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var items = listView.SelectedItems;
            if (items.Count <= 0) return;
            if (ArchiveMode)
            {
                foreach (FileItem item in items)
                    if (_type == ArchiveTypes.Zip)
                    {
                        var x = (ZipArchiveEntry)item.Entry;
                        var archive = (ZipArchive)x.Archive;
                        archive.RemoveEntry(x);
                        archive.SaveTo(_archivePath.Replace(".zip", "_.zip"), CompressionType.Deflate);
                        ArchiveMode = false;
                        //back.IsEnabled = false;
                    }
            }
            else
            {
                foreach (FileItem file in items)
                {
                    var path = file.Path;
                    if (Directory.Exists(path))
                        Directory.Delete(path, true);
                    else
                        File.Delete(path);
                }
            }

            Refresh();
        }

        private void open_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (listView.SelectedItem != null)
            {
                string args = $"/e, /select, \"{((FileItem)listView.SelectedItem).Path}\"";

                var info = new ProcessStartInfo
                {
                    FileName = "explorer",
                    Arguments = args
                };
                Process.Start(info);
            }
            else
            {
                Process.Start(pathBox.Text);
            }
        }

        private void OutArchive_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (!ArchiveMode && listView.SelectedIndex != -1)
                e.CanExecute = true;
        }

        private void extractTo_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (ArchiveMode && listView.SelectedIndex != -1)
                e.CanExecute = true;
        }

        private void extractTo_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog { IsFolderPicker = true };

            dialog.ShowDialog();
            if (!Directory.Exists(dialog.FileName)) return;
            foreach (FileItem item in listView.SelectedItems)
                ExtractFileItem(dialog.FileName, item);
            SelectFolder(dialog.FileName);
        }

        private static void ExtractFileItem(string filename, FileItem item)
        {
            var i = (IArchiveEntry)item.Entry;
            i.WriteToDirectory(filename);
        }

        private void WriteToDirectory(IArchive archive, string dir)
        {
            IsEnabled = false;

            var controller = new ProgressDialog { Title = "Busy" };
            controller.Closed += delegate
            {
                IsEnabled = true;
                SelectFolder(dir);
            };
            controller.StartExtract(archive.Entries.ToList(), dir);
        }

        private void SaveTo(ZipArchive archive, string fileName, WriterOptions writerOptions)
        {
            IsEnabled = false;

            var controller = new ProgressDialog { Title = "Busy" };
            controller.Closed += delegate { IsEnabled = true; };
            controller.StartCompress(archive, fileName, writerOptions);
        }

        private void about_Click(object sender, RoutedEventArgs e)
        {
            new About().ShowDialog();
        }

        public class FileItem
        {
            public FileItem(string path)
            {
                Icon = Imaging.CreateBitmapSourceFromHBitmap(
                    ImageUtilities.GetRegisteredIcon(path).ToBitmap().GetHbitmap(), IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            }

            public FileItem(ImageSource src)
            {
                Icon = src;
            }

            public ImageSource Icon { get; set; }
            public string Name { get; set; }
            public string Path { get; set; }
            public object Entry { get; set; }
            public bool IsDir { get; set; }
        }

        private void listView_DragLeave(object sender, DragEventArgs e)
        {

        }

        private void MainWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Back:
                    folder_up.PerformClick();
                    break;
                case Key.Enter:
                    if (listView.SelectedIndex > -1)
                    {
                        listView.PerformDoubleClick();
                    }
                    e.Handled = true;
                    break;
                case Key.Down:
                    if (listView.SelectedIndex == -1)
                    {
                        listView.SelectedIndex = 0;
                    }
                    else
                    {
                        var i = listView.SelectedIndex;
                        listView.SelectedIndex = i != listView.Items.Count - 1? ++i : 0;
                    }
                    break;
                case Key.Up:
                    if (listView.SelectedIndex == -1)
                    {
                        listView.SelectedIndex = 0;
                    }
                    var ix = listView.SelectedIndex;
                    listView.SelectedIndex = ix != 0 ? --ix : 0;
                    break;
                default:
                    break;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {

        }
    }
}