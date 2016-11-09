﻿using SharpCompress.Archives.Rar;
using SharpCompress.Archives.SevenZip;
using SharpCompress.Archives.Zip;
using SharpCompress.Archives;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using SharpCompress.Writers;
using SharpCompress.Archives.Tar;
using SharpCompress.Common;
using System.Threading;
using MahApps.Metro.Controls;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Zippy
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        string dummyNode = "none";
        public MainWindow()
        {
            InitializeComponent();
            tmp = new DirectoryInfo(System.IO.Path.GetTempPath()).Parent.FullName + "\\Zippy\\";
        }
        public ImageSource BmtoImgSource(System.Drawing.Bitmap source)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
      source.GetHbitmap(),
      IntPtr.Zero,
      Int32Rect.Empty,
      BitmapSizeOptions.FromEmptyOptions());
        }
        private void Window_Loaded(Object sender, RoutedEventArgs e)
        {
            foreach (string s in Directory.GetLogicalDrives())
            {
                TreeViewItem item = new TreeViewItem();
                item.Header = GetStackpanel(s);
                item.Tag = s;
                item.FontWeight = FontWeights.Normal;
                item.Items.Add(dummyNode);
                item.Expanded += new RoutedEventHandler(folder_Expanded);
                foldersItem.Items.Add(item);
            }
            string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            TreeViewItem item2 = new TreeViewItem();
            item2.Header = GetStackpanel(home);
            item2.Tag = home;
            item2.FontWeight = FontWeights.Normal;
            item2.Items.Add(dummyNode);
            item2.Expanded += new RoutedEventHandler(folder_Expanded);
            item2.IsSelected = true;
            foldersItem.Items.Add(item2);
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                string path = args[1];
                if (File.Exists(path))
                {
                    MessageBox.Show(path);
                    type = IsArchive(path);
                    OpenArchive(path);
                }
            }
        }

        private StackPanel GetStackpanel(string home)
        {
            var sp = new StackPanel();
            sp.Orientation = Orientation.Horizontal;
            sp.IsItemsHost = false;
            sp.Children.Add(new Image() { Source = BmtoImgSource(ImageUtilities.GetRegisteredIcon(home).ToBitmap()), Width = 16 });
            sp.Children.Add(new TextBlock() { Text = new DirectoryInfo(home).Name });
            return sp;
        }

        void folder_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = (TreeViewItem)sender;
            if (item.Items.Count == 1)
            {
                item.Items.Clear();
                try
                {
                    foreach (string s in Directory.GetDirectories(item.Tag.ToString()))
                    {
                        TreeViewItem subitem = new TreeViewItem();
                        subitem.Header = GetStackpanel(s);
                        subitem.Tag = s;
                        subitem.FontWeight = FontWeights.Normal;
                        subitem.Items.Add(dummyNode);
                        subitem.Expanded += new RoutedEventHandler(folder_Expanded);
                        item.Items.Add(subitem);
                    }
                }
                catch (Exception) { }
            }
        }
        public ArchiveTypes IsArchive(string path)
        {
            ArchiveTypes type = ArchiveTypes.None;

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
        private void foldersItem_SelectedItemChanged(Object sender, RoutedPropertyChangedEventArgs<Object> e)
        {
            ArchiveMode = false;
            Refresh();
        }

        public static ImageSource GetIcon(string strPath, bool bSmall)
        {
            Interop.SHFILEINFO info = new Interop.SHFILEINFO(true);
            int cbFileInfo = Marshal.SizeOf(info);
            Interop.SHGFI flags;
            if (bSmall)
                flags = Interop.SHGFI.Icon | Interop.SHGFI.SmallIcon | Interop.SHGFI.UseFileAttributes;
            else
                flags = Interop.SHGFI.Icon | Interop.SHGFI.LargeIcon | Interop.SHGFI.UseFileAttributes;

            Interop.SHGetFileInfo(strPath, 256, out info, (uint)cbFileInfo, flags);

            IntPtr iconHandle = info.hIcon;
            //if (IntPtr.Zero == iconHandle) // not needed, always return icon (blank)
            //  return DefaultImgSrc;
            ImageSource img = Imaging.CreateBitmapSourceFromHIcon(
                        iconHandle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
            Interop.DestroyIcon(iconHandle);
            return img;
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
            public Boolean isDir { get; set; }
        }
        private bool _ArchiveMode = false;
        public Boolean ArchiveMode {
            get { return _ArchiveMode; }
            set {
                _ArchiveMode = value;
                back.IsEnabled = value;
            }
        }
        private ArchiveTypes type;
        private System.Windows.Point start;
        private String tmp;
        private String archivePath;

        private void listView_MouseDoubleClick(Object sender, MouseButtonEventArgs e)
        {
            var item = ((FileItem)((ListView)sender).SelectedItem);
            if (ArchiveMode)
            {
                if (!item.isDir)
                {
                    var i = (IArchiveEntry)item.Entry;

                    string path = tmp + DateTime.Now.ToFileTime().ToString();
                    Directory.CreateDirectory(path);
                    i.WriteToDirectory(path, new SharpCompress.Readers.ExtractionOptions() { ExtractFullPath = true });
                    Process.Start(path + "\\" + i.Key);
                }
                else
                {
                    var i = (IArchiveEntry)item.Entry;
                    var archive = i.Archive;
                    string path = tmp + DateTime.Now.ToFileTime().ToString();
                    Directory.CreateDirectory(path);

                    foreach (var entry in archive.Entries)
                    {
                        if (!entry.IsDirectory)
                        {
                            Console.WriteLine(entry.Key);
                            entry.WriteToDirectory(path, new SharpCompress.Readers.ExtractionOptions() { ExtractFullPath = true });
                        }
                    }
                    ArchiveMode = false;
                    Process.Start(path);
                }
            }
            else
            {
                string path = item.Path.ToString();
                type = IsArchive(path);
                OpenArchive(path);
            }
        }

        private void OpenArchive(string path)
        {
            if (Directory.Exists(path))
            {
                SelectFolder(path);
            }
            else if (type != ArchiveTypes.None)
            {
                listView.Items.Clear();
                var archive = ArchiveFactory.Open(path);
                archivePath = path;
                archive.Entries.ToList().ForEach(new Action<IArchiveEntry>(delegate (IArchiveEntry values)
                {
                    string extension = "." + values.Key.Split('.').Last();
                    bool isDir = false;
                    if (values.IsDirectory)
                        isDir = true;
                    var exePath = IconManager.FindIconForFilename(extension, false);
                    listView.Items.Add(new FileItem(exePath) { Name = values.Key, Entry = values, isDir = isDir });
                }));
                ArchiveMode = true;
            }
            else
            {
                Process.Start(path);
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            zipSelected(ArchiveTypes.Zip);
        }

        private void ExtractionBegin(Object sender, ReaderExtractionEventArgs<IEntry> e)
        {

        }

        public static string Extension_GetExePath(string strExtension)
        {
            string strExePath = "C:\\Windows";
            try
            {
                //We need a leading dot, so add it if it's missing.
                if (!strExtension.StartsWith(".")) strExtension = "." + strExtension;

                //Get the class-name associated with the passed extension
                Microsoft.Win32.RegistryKey rkClassName
                   = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(strExtension);
                //Exit, if not found
                if (rkClassName == null) return string.Empty;
                string strClassName = rkClassName.GetValue("").ToString();

                //Get the shell-command for the retrieved executable 
                //This key is found at HKCR\[ClassName]\shell\open\command\(Default)
                //One or more of the paths may be missing, so each of them is being tested
                //separately.
                Microsoft.Win32.RegistryKey rkShellCommandRoot
                   = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(strClassName);
                if (rkShellCommandRoot == null) return string.Empty;

                Microsoft.Win32.RegistryKey rkShell
                   = rkShellCommandRoot.OpenSubKey("shell");
                if (rkShell == null) return string.Empty;

                Microsoft.Win32.RegistryKey rkOpen
                   = rkShell.OpenSubKey("open");
                if (rkOpen == null) return string.Empty;

                Microsoft.Win32.RegistryKey rkCommand
                   = rkOpen.OpenSubKey("command");
                if (rkCommand == null) return string.Empty;

                string strShellCommand = rkCommand.GetValue("").ToString();

                //The shell-command may contain additional parameters and may be wrapped in double quotes,
                //so parse out the exe-path.
                if (strShellCommand.StartsWith(@""))
                    //Extract path (wrapped in double-quotes)
                    strExePath = strShellCommand.Split('"')[1];
                else
                    //Extract first word (until first space char)
                    strExePath = strShellCommand.Split(' ')[0];

            }
            catch { }
            return strExePath;
        }



        public void zipSelected(ArchiveTypes type, WriterOptions options1 = null)
        {
            var items = ((listView).SelectedItems);
            if (items.Count > 0)
            {
                if (type == ArchiveTypes.Zip)
                {
                    var ar = ZipArchive.Create();
                    string p = "";
                    foreach (var x in items)
                    {
                        FileItem item = (FileItem)x;
                        if (Directory.Exists(item.Path))
                        {
                            ar.AddAllFromDirectory(item.Path);
                        }
                        else
                        {
                            ar.AddEntry(item.Name, File.Open(item.Path,FileMode.Open,FileAccess.Read));
                        }
                        p = item.Path;
                    }
                    this.Cursor = Cursors.Wait;
                    loading.IsIndeterminate = true;
                    Thread thread = new Thread(delegate ()
                     {
                         var fpath = new FileInfo(p).Directory + "\\" + ((FileItem)items[0]).Name.Split('.')[0];
                         if (File.Exists(fpath + ".zip"))
                             fpath += "_";
                         var open = File.Open(fpath + ".zip",FileMode.OpenOrCreate);
                         
                         if (options1 == null) options1 = new WriterOptions(CompressionType.LZMA);
                         ar.SaveTo(open, options1);
                         open.Dispose();
                         this.Dispatcher.Invoke(delegate ()
                         {
                             this.Cursor = Cursors.Arrow;
                             loading.IsIndeterminate = false;
                         });
                     });
                    thread.Start();

                }
                else if (type == ArchiveTypes.Deflate)
                {
                    var ar = ZipArchive.Create();
                    string p = "";
                    foreach (var x in items)
                    {
                        FileItem item = (FileItem)x;
                        if (Directory.Exists(item.Path))
                        {
                            ar.AddAllFromDirectory(item.Path);
                        }
                        else
                        {   
                            ar.AddEntry(item.Name, File.OpenRead(item.Path));
                        }
                        p = item.Path;
                    }
                    this.Cursor = Cursors.Wait;
                    loading.IsIndeterminate = true;
                    Thread thread = new Thread(delegate ()
                    {
                        var v = new FileInfo(p).Directory + "\\" + ((FileItem)items[0]).Name.Split('.')[0] + ".zip";
                        var open = File.Open(v, FileMode.OpenOrCreate);
                        ar.SaveTo(open, new WriterOptions(CompressionType.Deflate));
                        open.Close();
                        this.Dispatcher.Invoke(delegate ()
                        {
                            this.Cursor = Cursors.Arrow;
                            loading.IsIndeterminate = false;
                        });
                    });
                    thread.Start();

                }
                Refresh();
            }
        }

        private void back_Click(Object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)foldersItem.SelectedItem;
            item.IsSelected = false;

            item.IsSelected = true;
            ArchiveMode = false;
        }

        private void tar_Click(Object sender, RoutedEventArgs e)
        {
            zipSelected(ArchiveTypes.Deflate);
        }

        private void goto_Click(Object sender, RoutedEventArgs e)
        {
            string path = pathBox.Text;
            SelectFolder(path);
        }
        public void SelectFolder(string path)
        {
            if (Directory.Exists(path))
            {
                string[] nodes = path.Split('\\');
                var items = foldersItem.Items;

                for (int i = 0; i < nodes.Length; i++)
                {
                    string node = nodes[i];
                    if (i == 0)
                    {
                        node += "\\";
                    }

                    for (int ix = 0; ix < items.Count; ix++)
                    {
                        TreeViewItem obj = (TreeViewItem)items[ix];
                        var header = ((StackPanel)obj.Header).Children[1] as TextBlock;
                        if ((header.Text.ToString() == node))
                        {
                            obj.IsExpanded = true;
                            items = obj.Items;
                            ix = 0;
                            obj.IsSelected = true;
                        }
                    }
                }
            }
        }

        private void listView_MouseMove(Object sender, MouseEventArgs e)
        {
            System.Windows.Point mpos = e.GetPosition(null);
            Vector diff = this.start - mpos;

            if (e.LeftButton == MouseButtonState.Pressed &&
                Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                if (this.listView.SelectedItems.Count == 0)
                    return;

                // right about here you get the file urls of the selected items.  
                // should be quite easy, if not, ask.  
                string[] files = new String[listView.SelectedItems.Count];
                int ix = 0;
                foreach (object nextSel in listView.SelectedItems)
                {
                    files[ix] = ((FileItem)nextSel).Path;
                    ++ix;
                }
                string dataFormat = DataFormats.FileDrop;
                DataObject dataObject = new DataObject(dataFormat, files);
                DragDrop.DoDragDrop(this.listView, dataObject, DragDropEffects.Copy);
            }
        }

        private void listView_PreviewMouseLeftButtonDown(Object sender, MouseButtonEventArgs e)
        {
            this.start = e.GetPosition(null);
        }


        public void Refresh()
        {
            try
            {
                string path = ((TreeViewItem)foldersItem.SelectedItem).Tag.ToString();
                pathBox.Text = path;
                listView.Items.Clear();
                Directory.GetDirectories(path).Where(d => !new DirectoryInfo(d).Attributes.HasFlag(FileAttributes.Hidden)).ToList().ForEach(new Action<string>(delegate (string s)
                {
                    listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
                }));
                Directory.GetFiles(path).Where(f => ! new FileInfo(f).Attributes.HasFlag(FileAttributes.Hidden)).ToList().ForEach(new Action<string>(delegate (string s)
                {
                    listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
                }));
            }
            catch { }
        }

        private void folder_up_Click(Object sender, RoutedEventArgs e)
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
                    string path = pathBox.Text;
                    string parentPath = new DirectoryInfo(path).Parent.FullName;
                    SelectFolder(parentPath);
                }
            }
            catch { }
        }

        private void newZip_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            zipSelected(ArchiveTypes.Zip);

        }

        private void newZipG_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            zipSelected(ArchiveTypes.Deflate);

        }

        private void delete_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var items = ((listView).SelectedItems);
            if (items.Count > 0)
            {
                if (ArchiveMode)
                {
                    foreach (FileItem item in items)
                    {
                        if (type == ArchiveTypes.Zip)
                        {
                            var x = (ZipArchiveEntry)item.Entry;
                            var Archive = (ZipArchive)x.Archive;
                            Archive.RemoveEntry(x);
                            Archive.SaveTo(archivePath.Replace(".zip", "_.zip"), CompressionType.Deflate);
                            ArchiveMode = false;
                            back.IsEnabled = false;
                        }
                    }
                }
                else
                {
                    foreach (FileItem item in items)
                    {
                        string path = item.Path;
                        if (Directory.Exists(path))
                        {
                            Directory.Delete(path, true);
                        }
                        else
                        {
                            File.Delete(path);
                        }
                    }
                }
                Refresh();
            }
        }

        private void open_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (listView.SelectedItem != null)
            {
                string args = string.Format("/e, /select, \"{0}\"", ((FileItem)(listView.SelectedItem)).Path);

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "explorer";
                info.Arguments = args;
                Process.Start(info);
            }
            else
            {
                Process.Start(pathBox.Text);
            }
        }

        private void OutArchive_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if(!ArchiveMode && listView.SelectedIndex != -1)
            {
                e.CanExecute = true;
            }
        }

        private void advancedArchiving_Executed(object sender, ExecutedRoutedEventArgs e)
        {

        }

        private void extractTo_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (ArchiveMode && listView.SelectedIndex != -1)
            {
                e.CanExecute = true;
            }
        }

        private void extractTo_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;

            var result = dialog.ShowDialog();
            if (Directory.Exists(dialog.FileName))
            {
                foreach (FileItem item in listView.SelectedItems)
                {
                    var i = (IArchiveEntry)item.Entry;
                    i.WriteToDirectory(dialog.FileName);
                }
                SelectFolder(dialog.FileName);
            }
        }
    }
    public static class CustomCommands
    {
        public static RoutedCommand NewGZip = new RoutedCommand();
        public static RoutedCommand ExtractTo = new RoutedCommand();
    }
    public static class ImageUtilities
    {
        public static System.Drawing.Icon GetRegisteredIcon(string filePath)
        {
            var shinfo = new SHfileInfo();
            Win32.SHGetFileInfo(filePath, 0, ref shinfo, (uint)Marshal.SizeOf(shinfo), Win32.SHGFI_ICON | Win32.SHGFI_SMALLICON);
            return System.Drawing.Icon.FromHandle(shinfo.hIcon);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SHfileInfo
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }


    internal sealed class Win32
    {
        public const uint SHGFI_ICON = 0x100;
        public const uint SHGFI_LARGEICON = 0x0; // large
        public const uint SHGFI_SMALLICON = 0x1; // small

        [System.Runtime.InteropServices.DllImport("shell32.dll")]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHfileInfo psfi, uint cbSizeFileInfo, uint uFlags);
    }
}
