using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using TFlex.Model;

namespace AppLinkReplace.Class
{
    public class CTflexFileVariableInfo
    {
        #region constructor

        /// <summary>
        /// пустой конструктор
        /// </summary>
        public CTflexFileVariableInfo() { }

        /// <summary>
        /// конструктор принимает документ TFlex CADa
        /// </summary>
        /// <param name="doc">документ T-Flex CAD</param>
        public CTflexFileVariableInfo(Document doc)
        {
            if (doc == null) throw new ArgumentNullException("doc", "Переданно null");
            InitVariable(doc);
        }

        #endregion constructor

        #region private method

        /// <summary>
        /// метод осуществляет поик нужных переменных и заполняет соответствующие поля атрибутов
        /// </summary>
        /// <param name="doc">документ T-Flex CAD</param>
        private void InitVariable(Document doc)
        {
            var varTemp = doc.FindVariable("$format");
            if (varTemp != null)
            {
                _format = varTemp.TextValue;
            }
            varTemp = doc.FindVariable("$name");
            if (varTemp != null)
            {
                _name = varTemp.TextValue;
            }
            varTemp = doc.FindVariable("type_doc");
            if (varTemp != null)
            {
                _docType = varTemp.RealValue.ToString(CultureInfo.InvariantCulture);
            }
            varTemp = doc.FindVariable("$se4_mkmd");
            if (varTemp != null)
            {
                _sechMkmd = varTemp.TextValue.Replace("\\n", " ");
            }
            varTemp = doc.FindVariable("m");
            if (varTemp != null)
            {
                _massa = varTemp.RealValue;
            }
            else
            {
                varTemp = doc.FindVariable("mm");
                if (varTemp != null)
                {
                    _massa = varTemp.RealValue;
                }
            }
            varTemp = doc.FindVariable("area");
            if (varTemp != null)
            {
                _area = varTemp.RealValue;
            }
            varTemp = doc.FindVariable("$mat");
            if (varTemp != null)
            {
                _material = varTemp.TextValue.Replace("\\n", " ");
            }
            varTemp = doc.FindVariable("$pr");
            if (varTemp != null)
            {
                _comment = varTemp.TextValue.Replace("\\n", " ");
            }
            varTemp = doc.FindVariable("id");
            if (varTemp != null)
            {
                _docsID = varTemp.RealValue;
            }
            varTemp = doc.FindVariable("$list");
            if (varTemp != null)
            {
                _list = varTemp.TextValue;
            }
            varTemp = doc.FindVariable("$poz");
            if (varTemp != null)
            {
                _poz = varTemp.TextValue;
            }
        }

        #endregion private method

        #region public property

        public String Name
        {
            get { return _name; }
        }

        /// <summary>
        /// переменная формат "$format"
        /// </summary>
        public String Format
        {
            get { return _format; }
        }

        /// <summary>
        /// переменная тип документа "type_doc"
        /// </summary>
        public String DocType
        {
            get { return _docType; }
        }

        /// <summary>
        /// переменная масса "m"
        /// </summary>
        public Double Massa
        {
            get { return _massa; }
        }

        /// <summary>
        /// переменная площадь "area"
        /// </summary>
        public Double Area
        {
            get { return _area; }
        }

        /// <summary>
        /// переменная материал "$mat"
        /// </summary>
        public String Material
        {
            get { return _material; }
        }

        /// <summary>
        /// переменная сечения для MKMD "$se4_mkmd"
        /// </summary>
        public String SechMKMD
        {
            get { return _sechMkmd; }
        }

        /// <summary>
        /// переменная примечания "$pr
        /// </summary>
        public String Comment
        {
            get { return _comment; }
        }

        /// <summary>
        /// переменная идентификатор DOCsID "$pr
        /// </summary>
        public Double DOCsID
        {
            get { return _docsID; }
        }

        public String List
        {
            get { return _list; }
        }

        public String Poz
        {
            get { return _poz; }
        }

        #endregion public property

        #region variable

        /// <summary>
        /// переменная формат "$format"
        /// </summary>
        private String _format;

        /// <summary>
        /// переменная тип документа "type_doc"
        /// </summary>
        private String _docType;

        /// <summary>
        /// переменная масса "m"
        /// </summary>
        private Double _massa;

        /// <summary>
        /// переменная площадь "area"
        /// </summary>
        private Double _area;

        /// <summary>
        /// переменная материал "$mat"
        /// </summary>
        private String _material;

        /// <summary>
        /// переменная сечения для MKMD "$se4_mkmd"
        /// </summary>
        private String _sechMkmd;

        /// <summary>
        /// переменная примечания "$pr"
        /// </summary>
        private String _comment;

        /// <summary>
        /// идентификатор DOCsID
        /// </summary>
        private Double _docsID;

        /// <summary>
        /// переменная лист "$list"
        /// </summary>
        private String _list;

        /// <summary>
        /// переменная позиция "$poz"
        /// </summary>
        private String _poz;

        private String _name;

        #endregion variable
    }

    public class CTflexFile : INotifyPropertyChanged
    {
        public CTflexFile(FileInfo tflexFileInfo, CTflexFileVariableInfo tflexVariableInfo)
        {
            _isCheck = false;
            _fileName = tflexFileInfo.Name;
            _fullFileName = tflexFileInfo.FullName;
            _fileSize = Math.Ceiling((Double)tflexFileInfo.Length / 1024);
            _isReadOnly = tflexFileInfo.IsReadOnly;
            _tflexVariableInfo = tflexVariableInfo;
        }



        #region var

        private String _fileName;
        private String _fullFileName;
        private Boolean _isCheck;
        private Boolean _isReadOnly;
        private Double _fileSize;
        private CTflexFileVariableInfo _tflexVariableInfo;
        // private String _name;

        #endregion var

        #region property

        public Boolean Check
        {
            get { return _isCheck; }
            set
            {
                _isCheck = value;
                NotifyPropertyChanged("Check");
            }
        }

        public String FileName
        {
            get { return _fileName; }
            set
            {
                _fileName = value;
                NotifyPropertyChanged("FileName");
            }
        }

        public CTflexFileVariableInfo TflexVariableInfo
        {
            get { return _tflexVariableInfo; }
        }

        public String FullFileName
        {
            get { return _fullFileName; }

            set
            {
                _fullFileName = value;
                NotifyPropertyChanged("FullFileName");
            }
        }

        public String Name
        {
            get { return _fullFileName; }

            set
            {
                _fullFileName = value;
                NotifyPropertyChanged("FullFileName");
            }
        }

        public Double FileSize
        {
            get { return _fileSize; }
            set
            {
                _fileSize = value;
                NotifyPropertyChanged("FileSize");
            }
        }

        public Boolean ReadOnly
        {
            get { return _isReadOnly; }

            set
            {
                _isReadOnly = value;
                NotifyPropertyChanged("ReadOnly");
            }
        }

        // ReSharper disable UnusedAutoPropertyAccessor.Local
        private long MFileSize { get; set; }

        // ReSharper restore UnusedAutoPropertyAccessor.Local

        #endregion property

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion INotifyPropertyChanged Members

        #region Private Helpers

        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion Private Helpers
    }
}