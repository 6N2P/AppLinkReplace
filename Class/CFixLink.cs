using System;
using System.Collections.Generic;
using System.IO;
using TFlex.Model;

using TFlex.Model.Model2D;

namespace MKMDPlugin.CFixLink
{
    /*
        public class CDiagnosticMessage
        {
            public CDiagnosticMessage(String fragmentName, String bedLink, String fixLink, String message)
            {
                _mFragmentName = fragmentName;
                _mBedLink = bedLink;
                _mFixLink = fixLink;
                _mMessage = message;
            }

            #region variable

            private String _mFragmentName;
            private String _mBedLink;
            private String _mFixLink;
            private String _mMessage;

            #endregion variable
        }
    */

    /// <summary>
    /// Клас для работы со ссылками фрагментов позволят найти и исправить не верные ссылки на фрагменты документов
    /// </summary>
    public class CFixLink
    {
        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="flexActiveDoc">Передается документ T-Flex CAD для обработки</param>
        public CFixLink(Document flexActiveDoc)
        {
            _mTfDoc = flexActiveDoc;
        }

        /// <summary>
        /// Метод осуществляет поиск и замену битых ссылок на фрагменты
        /// </summary>
        public void FixLink()
        {
            if (_mTfDoc != null)
            {
                foreach (var fragment in _mTfDoc.GetFragments())
                {
                    //fragment.OpenPart( Fragment.OpenPartOptions
                    var fl = fragment.FileLink;
                    //var originalLink = (String)fl.FullFilePath.Clone();
                    if (fl.FilePath != "" && !fl.FilePath.Contains("$"))
                    {
                        var fileInfo = new FileInfo(fl.FullFilePath);
                        if (!fileInfo.Exists || fileInfo.Exists)
                        {
                            _mLinkDocument = _mTfDoc.FileName;
                            _mBedLink = fl.FullFilePath;
                            _mBedLinkPath = fl.FilePath;

                            var result = BuildFileLink(fragment);
                            if (result == 1)
                            {
                                _mTfDoc.Diagnostics.Add(new DiagnosticsMessage(DiagnosticsMessageType.Information, "Ссылка исправлена", fragment));
                            }
                            else
                            {
                                if (result == 3)
                                { continue; }
                                var res = BuildFileLinkSubDir(fragment);
                                if (res == 0)
                                {
                                    _mTfDoc.Diagnostics.Add(new DiagnosticsMessage(DiagnosticsMessageType.FileError, "Ссылку не удалось найти", fragment));
                                    //         System.Windows.Forms.MessageBox.Show("Возможно файл расположен в нижней вложенности \n Плохая ссылка " + this.m_bed_link);
                                }
                                else
                                {
                                    if (res == 3)
                                    { continue; }

                                    _mTfDoc.Diagnostics.Add(new DiagnosticsMessage(DiagnosticsMessageType.Information, "Ссылка исправлена", fragment));
                                }
                            }
                        }
                        else
                        {
                            if (!fl.FilePath.Contains(".."))
                            {
                                _mLinkDocument = _mTfDoc.FileName;
                                _mBedLink = fl.FullFilePath;
                                _mBedLinkPath = fl.FilePath;

                                Int32 result = BuildFileLink(fragment);
                                if (result == 1)
                                {
                                    _mTfDoc.Diagnostics.Add(new DiagnosticsMessage(DiagnosticsMessageType.Information, "Ссылка исправлена", fragment));
                                }
                                else
                                {
                                    if (result == 3)
                                    { continue; }

                                    Int32 res = BuildFileLinkSubDir(fragment);
                                    if (res == 0)
                                    {
                                        _mTfDoc.Diagnostics.Add(new DiagnosticsMessage(DiagnosticsMessageType.FileError, "Ссылку не уадалось найти", fragment));
                                        //            System.Windows.Forms.MessageBox.Show("Возможно файл расположен в нижней вложенности \n Плохая ссылка " + this.m_bed_link);
                                    }
                                    else
                                    {
                                        if (res == 3)
                                        { continue; }

                                        _mTfDoc.Diagnostics.Add(new DiagnosticsMessage(DiagnosticsMessageType.Information, "Ссылка исправлена", fragment));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Метод исправления битых ссылок поиск вверх до первого нахождения
        /// </summary>
        /// <param name="frag">фрагмент</param>
        /// <returns>код отработки метода 0 - неудалось исправить, 1 - удачное исправление, 3 - линки совпадают</returns>
        private Int32 BuildFileLink(Fragment frag)
        {
            Int32 result = 0;
            if (frag != null && _mLinkDocument != null)
            {
                var directoryName = new FileInfo(_mLinkDocument).DirectoryName;
                if (directoryName != null)
                    Directory.SetCurrentDirectory(directoryName);
                var fileBed = _mBedLink.Split('\\');

                var fiLocal = new FileInfo(fileBed[fileBed.Length - 1]);
                if (fiLocal.Exists)
                {
                    if (_mBedLinkPath == fileBed[fileBed.Length - 1])
                    {
                        return 3;
                    }
                    _mTfDoc.BeginChanges("Ремонтируем ссылки");
                    frag.FilePath = fileBed[fileBed.Length - 1];
                    _mTfDoc.EndChanges();
                    return 1;
                }

                String bedName = fileBed[fileBed.Length - 2] + @"\" + fileBed[fileBed.Length - 1];
                Boolean isFind = false;
                const string sep = @"..\";
                String tempDir = bedName;
                Int32 i = 0;
                while (!isFind)
                {
                    i++;
                    var fi = new FileInfo(tempDir);
                    if (fi.Exists)
                    {
                        if (_mBedLinkPath == tempDir)
                        {
                            result = 3;
                            break;
                        }
                        try
                        {
                            _mTfDoc.BeginChanges("Ремонтируем ссылки");
                            frag.FilePath = tempDir;
                            // isFind = true;
                            result = 1;
                            break;
                        }
                        catch
                        {
                            //isFind = false;
                            result = 0;
                            break;
                        }
                        finally
                        {
                            _mTfDoc.EndChanges();
                        }
                    }
                    tempDir = sep + tempDir;

                    if (i == 20)
                    { isFind = true; result = 0; }
                }
            }
            return result;
        }

        /// <summary>
        /// Метод для поиска директории в массиве
        /// </summary>
        /// <param name="bk">искуеммый элемент</param>
        /// <returns></returns>
        private bool FindDir(String bk)
        {
            if (bk == _mFindName)
            {
                return true;
            }
            {
                return false;
            }
        }

        /// <summary>
        /// Переменная для сравнимания
        /// </summary>
        private String _mFindName;

        /// <summary>
        /// Осуществляет поис битых ссылок вниз по директории до первого нахождения
        /// </summary>
        /// <param name="frag">Проверяемый фрагмент</param>
        /// <returns>код отработки метода 0 - неудалось исправить, 1 - удачное исправление, 3 - линки совпадают</returns>
        private Int32 BuildFileLinkSubDir(Fragment frag)
        {
            Int32 result = 0;
            if (frag != null && _mLinkDocument != null)
            {
                var directoryName = new FileInfo(_mLinkDocument).DirectoryName;
                if (directoryName != null)
                    Directory.SetCurrentDirectory(directoryName);
                String[] fileParent = _mLinkDocument.Split('\\');
                String[] fileBed = _mBedLink.Split('\\');

                String test2 = FindFullPath(fileBed[fileBed.Length - 1], _mLinkDocument);
                if (test2 != null)
                {
                    fileBed = test2.Split('\\');
                }

                var listFileBed = new List<string>();
                listFileBed.AddRange(fileBed);
                Int32 res = 0;
                String newDir = @"..";
                for (Int32 i = listFileBed.Count; i >= 0; i--)
                {
                    _mFindName = fileParent[i - 2];
                    res = listFileBed.FindIndex(FindDir);
                    if (res > 0)
                    {
                        break;
                    }
                }

                if (res > 0)
                {
                    for (Int32 i = res + 1; i < listFileBed.Count; i++)
                    {
                        newDir += "\\" + fileBed[i];
                    }
                }
                if (_mBedLinkPath == newDir)
                {
                    result = 3;
                }
                else
                {
                    try
                    {
                        _mTfDoc.BeginChanges("Ремонтируем ссылки");
                        frag.FilePath = newDir;
                        result = 1;
                    }
                    catch
                    {
                        result = 0;
                    }
                    finally
                    {
                        _mTfDoc.EndChanges();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Ищет полный путь к файлу
        /// </summary>
        /// <param name="searchFile">Имя искомого файла</param>
        /// <param name="filePathParent">Путь расположения родительского файла</param>
        /// <returns>Найденный путь</returns>
        private static String FindFullPath(String searchFile, String filePathParent)
        {
            if (searchFile != null && filePathParent != null)
            {
                var fi = new FileInfo(filePathParent);
                var tempDir = fi.DirectoryName;

                if (tempDir != null)
                {
                    var pathParent = tempDir.Split('\\');

                    var res = Directory.GetFiles(tempDir, searchFile, SearchOption.AllDirectories);
                    if (res.Length != 0)
                    {
                        return res[0];
                    }
                    for (Int32 i = pathParent.Length; i > 0; i--)
                    {
                        String removeName = "\\" + pathParent[i - 1];
                        tempDir = tempDir.Replace(removeName, "");
                        //BinarySearch(file_parent[i-2]);
                        try
                        {
                            res = Directory.GetFiles(tempDir, searchFile, SearchOption.AllDirectories);
                        }
                        catch { return null; }
                        if (res.Length != 0)
                        { return res[0]; }
                    }
                }
            }
            return null;
        }

        #region variable

        /// <summary>
        /// Ссылка на документ T-Flex CAD
        /// </summary>
        private readonly Document _mTfDoc;

        /// <summary>
        /// Содержит битый ссылочный путь
        /// </summary>
        private String _mBedLink;

        /// <summary>
        /// Содержит ссылку на файл документа
        /// </summary>
        private String _mLinkDocument;

        /// <summary>
        /// Содержит битую ссылочную директорию
        /// </summary>
        private String _mBedLinkPath;

        #endregion variable
    }
}