using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Archives.Zip;
using SharpCompress.Readers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Zippy
{
    public class Decompressor
    {
        public string Archive { get; set; }
        public string Output { get; set; }
        public ArchiveTypes ArchiveType { get; set; }
        public ExtractionOptions Options { get; set; }
        public Decompressor(string _archive, string _output, ArchiveTypes type)
        {
            Archive = _archive;
            ArchiveType = type;
            Output = _output;
            Options = new ExtractionOptions()
            {
                ExtractFullPath = true,
                Overwrite = true
            };
        }
        public Decompressor(string _archive, string _output, ArchiveTypes type, ExtractionOptions _options)
        {
            Archive = _archive;
            ArchiveType = type;
            Output = _output;
            Options = _options;
        }
        public void Extract()
        {
            if(ArchiveType == ArchiveTypes.RAR)
            {
                using (var archive = RarArchive.Open(Archive))
                {
                    foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                    {
                        entry.WriteToDirectory(Output, Options);
                    }
                }
            }
            else if (ArchiveType == ArchiveTypes.Zip)
            {
                using (var archive = ZipArchive.Open(Archive))
                {
                    foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                    {
                        entry.WriteToDirectory(Output, Options);
                    }
                }
            }
            else if (ArchiveType == ArchiveTypes.SevenZip)
            {
                using (var archive = SharpCompress.Archives.SevenZip.SevenZipArchive.Open(Archive))
                {
                    foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                    {
                        entry.WriteToDirectory(Output, Options);
                    }
                }
            }
        }
    }
    public enum ArchiveTypes { RAR, Zip, SevenZip,
        None,
        Deflate
    }
}
