using SharpCompress.Archives.Rar;
using SharpCompress.Archives.SevenZip;
using SharpCompress.Archives.Zip;
using SharpCompress.Archives;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using SharpCompress.Writers;
using SharpCompress.Archives.Tar;

namespace Zippy
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        string dummyNode = "test";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(Object sender, RoutedEventArgs e)
        {
            foreach (string s in Directory.GetLogicalDrives())
            {
                TreeViewItem item = new TreeViewItem();
                item.Header = s;
                item.Tag = s;
                item.FontWeight = FontWeights.Normal;
                item.Items.Add(dummyNode);
                item.Expanded += new RoutedEventHandler(folder_Expanded);
                foldersItem.Items.Add(item);
            }
            TreeViewItem item2 = new TreeViewItem();
            item2.Header = "Home";
            item2.Tag = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            item2.FontWeight = FontWeights.Normal;
            item2.Items.Add(dummyNode);
            item2.Expanded += new RoutedEventHandler(folder_Expanded);
            foldersItem.Items.Add(item2);
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
                        subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
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
                type = ArchiveTypes.Gzip;

            return type;
        }
        private void foldersItem_SelectedItemChanged(Object sender, RoutedPropertyChangedEventArgs<Object> e)
        {
            try
            {
                string path = ((TreeViewItem)e.NewValue).Tag.ToString();
                pathBox.Text = path;
                listView.Items.Clear();
                Directory.GetDirectories(path).ToList().ForEach(new Action<string>(delegate (string s)
                {
                    listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
                }));
                Directory.GetFiles(path).ToList().ForEach(new Action<string>(delegate (string s)
                {
                    listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
                }));
            }
            catch { }
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
          BitmapSizeOptions.FromEmptyOptions()); ;

            }
            public FileItem()
            {
                Icon = GetIcon("", false);
            }
            public ImageSource Icon { get; set; }
            public string Name { get; set; }
            public string Path { get; set; }
            public object Entry { get; set; }
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
        private Point start;

        private void listView_MouseDoubleClick(Object sender, MouseButtonEventArgs e)
        {
            var item = ((FileItem)((ListView)sender).SelectedItem);
            if (ArchiveMode)
            {
                if (type == ArchiveTypes.Zip)
                {
                    var Entry = (ZipArchiveEntry)item.Entry;
                    string toPath = System.IO.Path.GetTempPath() + "\\" + Entry.Key;
                    Entry.WriteToFile(toPath);
                    Process.Start(toPath);
                }
                if (type == ArchiveTypes.RAR)
                {
                    var Entry = (RarArchiveEntry)item.Entry;
                    string toPath = System.IO.Path.GetTempPath() + "\\" + Entry.Key;
                    Entry.WriteToFile(toPath);
                    Process.Start(toPath);
                }
                if (type == ArchiveTypes.SevenZip)
                {
                    var Entry = (SevenZipArchiveEntry)item.Entry;
                    string toPath = System.IO.Path.GetTempPath() + "\\" + Entry.Key;
                    Entry.WriteToFile(toPath);
                    Process.Start(toPath);
                }
                if (type == ArchiveTypes.Gzip)
                {
                    var Entry = (TarArchiveEntry)item.Entry;
                    string toPath = System.IO.Path.GetTempPath() + "\\" + Entry.Key;
                    Entry.WriteToFile(toPath);
                    Process.Start(toPath);
                }
            }
            else
            {
                string path = item.Path.ToString();
                type = IsArchive(path);
                if (Directory.Exists(path))
                {
                    SelectFolder(path);
                }
                else if (type != ArchiveTypes.None)
                {
                    listView.Items.Clear();
                    if (type == ArchiveTypes.Zip)
                    {
                        var archive = ZipArchive.Open(path);
                        archive.Entries.Select(i => new KeyValuePair<string, object>(i.Key, i)).ToList().ForEach(new Action<KeyValuePair<string, object>>(delegate (KeyValuePair<string, object> values)
                        {
                            listView.Items.Add(new FileItem() { Name = values.Key, Entry = values.Value });
                        }));
                    }
                    else if (type == ArchiveTypes.RAR)
                    {
                        var archive = RarArchive.Open(path);
                        archive.Entries.Select(i => new KeyValuePair<string, object>(i.Key, i)).ToList().ForEach(new Action<KeyValuePair<string, object>>(delegate (KeyValuePair<string, object> values)
                        {
                            listView.Items.Add(new FileItem() { Name = values.Key, Entry = values.Value });
                        }));
                    }
                    else if (type == ArchiveTypes.SevenZip)
                    {
                        var archive = SevenZipArchive.Open(path);
                        archive.Entries.Select(i => new KeyValuePair<string, object>(i.Key, i)).ToList().ForEach(new Action<KeyValuePair<string, object>>(delegate (KeyValuePair<string, object> values)
                        {
                            listView.Items.Add(new FileItem() { Name = values.Key, Entry = values.Value });
                        }));
                    }
                    else if (type == ArchiveTypes.Gzip)
                    {
                        var archive = TarArchive.Open(path);
                        archive.Entries.Select(i => new KeyValuePair<string, object>(i.Key, i)).ToList().ForEach(new Action<KeyValuePair<string, object>>(delegate (KeyValuePair<string, object> values)
                        {
                            listView.Items.Add(new FileItem() { Name = values.Key, Entry = values.Value });
                        }));
                    }
                    ArchiveMode = true;
                }
                else
                {
                    Process.Start(path);
                }
            }
        }

        private void MenuItem_Click(Object sender, RoutedEventArgs e)
        {
            zipSelected(ArchiveTypes.Zip);
        }

        public void zipSelected(ArchiveTypes type)
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
                            ar.AddEntry(item.Name, File.OpenRead(item.Path));
                        }
                        p = item.Path;
                    }
                    this.Cursor = Cursors.Wait;
                    loading.IsIndeterminate = true;
                    var open = File.OpenWrite(new FileInfo(p).Directory + "\\" + ((FileItem)items[0]).Name.Split('.')[0] + ".zip");
                    ar.SaveTo(open);
                    open.Close();
                    this.Cursor = Cursors.Arrow;
                    loading.IsIndeterminate = false;
                }
                else if (type == ArchiveTypes.Gzip)
                {
                    var ar = SharpCompress.Archives.GZip.GZipArchive.Create();
                    string p = "";
                    foreach (var x in items)
                    {
                        FileItem item = (FileItem)x;
                        ar.AddEntry(item.Name, File.OpenRead(item.Path));
                        p = item.Path;
                    }
                    loading.IsIndeterminate = true;
                    this.Cursor = Cursors.Wait;
                    var open = File.OpenWrite(new FileInfo(p).Directory + "\\" + ((FileItem)items[0]).Name.Split('.')[0] + ".gz");

                    ar.SaveTo(open, new WriterOptions(SharpCompress.Common.CompressionType.GZip));
                    open.Close();
                    this.Cursor = Cursors.Arrow;
                    loading.IsIndeterminate = false;

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
            zipSelected(ArchiveTypes.Gzip);
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
                        if ((obj.Header.ToString() == node))
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
            Point mpos = e.GetPosition(null);
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

        private void open_Click(Object sender, RoutedEventArgs e)
        {
            if (listView.SelectedItem != null)
            {
                string args = string.Format("/e, /select, \"{0}\"", ((FileItem)(listView.SelectedItem)).Path);

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "explorer";
                info.Arguments = args;
                Process.Start(info);
            }
        }

        private void delete_Click(Object sender, RoutedEventArgs e)
        {
            var items = ((listView).SelectedItems);
            if (items.Count > 0)
            {
                foreach (FileItem item in items)
                {
                    string path = item.Path;
                    if (Directory.Exists(path))
                    {
                        Directory.Delete(path);
                    }
                    else
                    {
                        File.Delete(path);
                    }
                }
                Refresh();
            }
        }
        public void Refresh()
        {
            string path = ((TreeViewItem)foldersItem.SelectedItem).Tag.ToString();
            pathBox.Text = path;
            listView.Items.Clear();
            Directory.GetDirectories(path).ToList().ForEach(new Action<string>(delegate (string s)
            {
                listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
            }));
            Directory.GetFiles(path).ToList().ForEach(new Action<string>(delegate (string s)
            {
                listView.Items.Add(new FileItem(s) { Name = new FileInfo(s).Name, Path = s });
            }));
        }

        private void folder_up_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                string path = pathBox.Text;
                string parentPath = new DirectoryInfo(path).Parent.FullName;
                SelectFolder(parentPath);

            }
            catch { }
        }
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
