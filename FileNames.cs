using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;

namespace Document_Converter
{
    public class FileNames
    {
        //fields
        string _fileNames;

        //Properties
        public string filesNames {
            get { return _fileNames; }
            set { _fileNames = value; }
        }

        //Constructors
        public FileNames()
        {}

        public FileNames(string filename)
        {
            _fileNames = filename;
        }
   }

    public class FileNameCollection : Collection<FileNames>
    {
        public FileNames this[int ctr]
        {
            get { return this.Items[ctr]; }
            set { this.Items[ctr] = value; }
        }

        new public FileNames Add(FileNames newFilename)
        {
            this.Items.Add(newFilename);
            return(FileNames)this.Items[this.Items.Count-1];
        }
    }

    public sealed class FileNameDataList
    {
        static FileNameCollection _newFileNameDataList = new FileNameCollection();
        public static FileNameCollection FileNameList
        {
            get { return _newFileNameDataList; }
        }
    }
}
