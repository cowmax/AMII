using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ImportExcelToDB
{
    public class TemplateManager
    {
        TemplateCollection _templates;
        string _folderPath;
        public int Count;

        public TemplateManager(string path)
        {
            this.Count = loadTemplates(path);
        }

        private int loadTemplates(string path)
        {
            _folderPath = path;

            string[] filePaths = Directory.GetFiles(path);
            _templates = new TemplateCollection(filePaths);

            return _templates.Count;
        }

        public string Path
        {
            get
            {
                return _folderPath;
            }
        }

        public TemplateCollection Templates
        {
            get
            {
                return _templates;
            }
        }

        /// <returns>TemplateInfo</returns>
        public TemplateInfo GetTemplateInfo(string tmplId)
        {
            return _templates.GetTemplate(tmplId);
        }

        public void AddNewTemplate()
        {
            throw new System.NotImplementedException();
        }
    }

    public class TemplateCollection
    {
        private TemplateInfo[] _items;

        public TemplateCollection(string[] tmplFiles)
        {
            int len = tmplFiles.Length;
            _items = new TemplateInfo[len];

            TemplateInfo ti = null;
            for (int i = 0; i < len; i++ )
            {
                ti = new TemplateInfo(tmplFiles[i]);
                _items[i] = ti;
            }
        }

        public TemplateInfo[] Items
        {
            get
            {
                return _items;
            }
        }

        public int AddNew(TemplateInfo tmplInfo)
        {
            throw new System.NotImplementedException();
        }

        public bool Delete(string tmplId)
        {
            throw new System.NotImplementedException();
        }

        public TemplateInfo GetTemplate(string tmplId)
        {
            throw new System.NotImplementedException();
        }

        public TemplateInfo[] GetOtherVersion(string tmplId)
        {
            throw new System.NotImplementedException();
        }

        public int Count {
            get { return _items.Length; }
            }
    }



    public class TemplateInfo
    {
        // 模板文件名称规范 : <name>.<id>.<version>.xls
        // 例: LocalSaleRecords.141225223055123.FFFF.xls
        // LocalSaleRecords :　模板的可读名称，相式相同的模板的可读名称应完全相同
        // 20141225223055123000 : 系统范围内模板文件的唯一标识，20位随机数字串（使用系统时间创建）
        // FFFF ： 系统范围内同类模板的序号，这里称为版本号

        static Regex rgxTmplFile = new Regex(@"(?<name>(\w+))\.(?<id>(\d+))\.(?<ver>([\d\w]+))\.", RegexOptions.IgnoreCase);
 
        string _filePath;
        string _id;
        string _name;
        int _version;
        FileInfo _fileInfo;

        #region BEGIN : Private Method
        private void parseTemplateInfo()
        {
            string fnm = Path.GetFileName(_filePath);

            Match mch = rgxTmplFile.Match(_filePath);
            if (mch.Length > 1)
            {
                _name = mch.Groups["name"].Value;
                _id = mch.Groups["id"].Value;

                if (!int.TryParse(mch.Groups["ver"].Value, NumberStyles.HexNumber,
                    new CultureInfo("zh-CN"), out _version))
                {
                    _version = 1; // Set to default version
                }
            }
        }
        #endregion END : Private Method

        public TemplateInfo(string filePath)
        {
            _filePath = filePath;

            // 1. query file-info
            _fileInfo = new FileInfo(filePath);

            parseTemplateInfo();
        }
    
        public FileInfo fileInfo
        {
            get
            {
                return _fileInfo;
            }
        }

        public string id
        {
            get
            {
                return _id;
            }
        }

        public string name
        {
            get
            {
                return _name;
            }
        }

        public int version
        {
            get
            {
                return _version;
            }
        }
    }
}
