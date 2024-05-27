using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace AppLinkReplace.Class
{
    public class CAPILoader : IAPILoader
    {
        #region winapi

        //[DllImport(@"C:\Windows\SysWOW64\kernel32.dll")]


        [DllImport("kernel32.dll")]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll")]
        public static extern bool SetProcessWorkingSetSize(IntPtr handle, int minimumWorkingSetSize, int maximumWorkingSetSize);

        //[return: MarshalAs(UnmanagedType.Bool)]
        //static extern bool FreeLibrary(IntPtr hModule);

        #endregion winapi

        public CAPILoader()
        { }

        //прописываем путь к CADу и DOCSу
        public void Initialize()
        {
            if (_folders != null)
                return;

            _folders = new List<string>();

            string path = GetTopSystemsTFlexCadPath();
            if (string.IsNullOrEmpty(path))
                throw new System.IO.FileNotFoundException("T-FLEX CAD not installed");

            _folders.Add(path);

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolve);
            //  System.Windows.Forms.MessageBox.Show("Подписался на событие");
        }

        public bool InitializeTFlexCADAPI(String loadBomdll)
        {
            if (_folders == null)
                throw new InvalidOperationException("Call Initialize first");

            //Перед работой с API T-FLEX CAD его необходимо инициализировать
            //В зависимости от параметров инициализации, будут или не будут
            //доступны функции изменения документов и сохранение документов в файл.
            //За это отвечает параметр setup.ReadOnly.
            //Если setup.ReadOnly = false, то для работы программы требуется
            //лицензия на сам T-FLEX CAD
            try
            {
                TFlex.ApplicationSessionSetup setup = new TFlex.ApplicationSessionSetup();

                setup.ReadOnly = false;
                //TFlex.Application.EnableNotRespondingDialog(false);
                TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.AutoRefresh;
                m_is_init = TFlex.Application.InitSession(setup);
                if (loadBomdll != null)
                {
                    if (File.Exists(loadBomdll))
                    {
                        TFlex.Application.BOMSectionsDatabase = loadBomdll;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(Properties.Resources.INFORMATION_BOM_FILE_PATH);
                        return false;
                    }

                    IntPtr _mLibrary = LoadLibrary("tfbom");
                    if (_mLibrary == IntPtr.Zero)
                    {
                        System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_NOT_LOAD_TFBOMDLL);
                        return false;
                    }
                }
                return m_is_init;
            }
            catch (Exception e) { throw new Exception("Ошибка инициализации T-Flex CAD", e.InnerException); }
        }

        public void Terminate()
        {
            try
            {
                if (_folders == null)
                    return;

                TFlex.Application.ExitSession();
            }
            catch (Exception e) { throw new Exception("Ошибка выхода из сессии T-Flex CAD", e.InnerException); }
            finally
            {
                AppDomain.CurrentDomain.AssemblyResolve -= new ResolveEventHandler(AssemblyResolve);

                _folders = null;
                m_is_init = false;
            }
        }

        public bool IsInit { get { return m_is_init; } }

        protected bool m_is_init = false;

        private List<string> _folders;

        private String GetTopSystemsTFlexCadPath()
        {
            RegistryKey key = null;
            try
            {
                #region поиск TFLEX CAD 17

                if (IntPtr.Size == 8)
                {
                    key =
                        Registry.LocalMachine.OpenSubKey(
                            string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 17\Rus"),
                            RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);

                    _tflexVersionApi = "17";
                }
                //else
                //{
                //    key =
                //        Registry.LocalMachine.OpenSubKey(
                //            string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 17\Rus"),
                //            RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                //    _tflexVersionApi = "17";
                //}
                #endregion поиск TFLEX CAD 17

                #region поиск TFLEX CAD 15
                if (key == null)
                {

                    if (IntPtr.Size == 8)
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 15 x64\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                        if (key == null)
                        {
                            key =
                                Registry.LocalMachine.OpenSubKey(
                                    string.Format(@"SOFTWARE\Wow6432Node\Top Systems\{0}\", @"T-FLEX CAD 3D 15\Rus"),
                                    RegistryKeyPermissionCheck.ReadSubTree,
                                    System.Security.AccessControl.RegistryRights.ReadKey);
                        }
                        _tflexVersionApi = "15";
                    }
                    else
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 15\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                        _tflexVersionApi = "15";
                    }
                }
                #endregion поиск TFLEX CAD 15

                #region поиск TFLEX CAD 14

                if (key == null)
                {
                    if (IntPtr.Size == 8)
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 14 x64\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                        if (key == null)
                        {
                            key =
                                Registry.LocalMachine.OpenSubKey(
                                    string.Format(@"SOFTWARE\Wow6432Node\Top Systems\{0}\", @"T-FLEX CAD 3D 14\Rus"),
                                    RegistryKeyPermissionCheck.ReadSubTree,
                                    System.Security.AccessControl.RegistryRights.ReadKey);
                        }
                        _tflexVersionApi = "14";
                    }
                    else
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 14\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                        _tflexVersionApi = "14";
                    }
                }

                #endregion поиск TFLEX CAD 14

                #region поиск TFLEX CAD 12

                if (key == null)
                {
                    if (IntPtr.Size == 8)
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 12 x64\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                        if (key == null)
                        {
                            key =
                                Registry.LocalMachine.OpenSubKey(
                                    string.Format(@"SOFTWARE\Wow6432Node\Top Systems\{0}\", @"T-FLEX CAD 3D 12\Rus"),
                                    RegistryKeyPermissionCheck.ReadSubTree,
                                    System.Security.AccessControl.RegistryRights.ReadKey);
                        }
                        _tflexVersionApi = "12";
                    }
                    else
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 12\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                        _tflexVersionApi = "12";
                    }
                }

                #endregion поиск TFLEX CAD 12

                #region поиск TFLEX CAD 11

                if (key == null)
                {
                    if (IntPtr.Size == 8)
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 11.0 x64\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree,
                                System.Security.AccessControl.RegistryRights.ReadKey);
                        if (key == null)
                        {
                            key =
                                Registry.LocalMachine.OpenSubKey(
                                    string.Format(@"SOFTWARE\Wow6432Node\Top Systems\{0}\", @"T-FLEX CAD 3D 11.0\Rus"),
                                    RegistryKeyPermissionCheck.ReadSubTree,
                                    System.Security.AccessControl.RegistryRights.ReadKey);
                        }
                        _tflexVersionApi = "11";
                    }
                    else
                    {
                        key =
                            Registry.LocalMachine.OpenSubKey(
                                string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 11.0\Rus"),
                                RegistryKeyPermissionCheck.ReadSubTree,
                                System.Security.AccessControl.RegistryRights.ReadKey);
                        _tflexVersionApi = "11";
                    }
                }

                #endregion поиск TFLEX CAD 11
            }
            catch (Exception e) { throw new Exception("Ошибка при поиске ключа рееста T-Flex CAD", e.InnerException); }

            //var path = GetCurrentPath(key);
            //if (key != null)
            //{ key.Close(); }

            #region Код для 17 тифлекса
            var path = (string)key.GetValue("ProgramFolder", string.Empty);
            if (string.IsNullOrEmpty(path))
                path = (string)key.GetValue("SetupHelpPath", string.Empty);
            key.Close();
            if (path.Length > 0 && path[path.Length - 1] != '\\')
                path += @"\";
            #endregion Код для 17 тифлекса

            return path;
        }

        public String TflexVersionApi
        {
            get { return _tflexVersionApi; }
        }

        private String _tflexVersionApi = null;

        private String GetCurrentPath(RegistryKey Key)
        {
            String path = null;
            try
            {
                if (Key != null)
                {
                    path = (string)Key.GetValue("ProgramFolder", string.Empty);
                    if (string.IsNullOrEmpty(path))
                        path = (string)Key.GetValue("SetupHelpPath", string.Empty);

                    if (path.Length > 0 && path[path.Length - 1] != '\\')
                        path += @"\";
                }
            }
            catch (Exception e) { throw new Exception("Ошибка получение пути из ключа реестра", e.InnerException); }
            return path;
        }

        protected System.Reflection.Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            //System.Windows.Forms.MessageBox.Show("В методе AssemblyResolve");
            if (_folders == null || _folders.Count == 0)
            { /*System.Windows.Forms.MessageBox.Show("В методе AssemblyResolve пустой массив");*/ return null; }

            //System.Windows.Forms.MessageBox.Show("В методе AssemblyResolve " + _folders.Count.ToString());

            try
            {
                string name = args.Name;

                int index = name.IndexOf(",");
                if (index > 0)
                    name = name.Substring(0, index);

                foreach (string path in _folders)
                {
                    if (!System.IO.Directory.Exists(path))
                        continue;

                    string fileName = string.Format("{0}{1}.dll", path, name);

                    if (!System.IO.File.Exists(fileName))
                    {
                        //  System.Windows.Forms.MessageBox.Show("Не нашол бибилиотеку " + fileName);
                        continue;
                    }
                    System.IO.Directory.SetCurrentDirectory(path);
                    //System.Windows.Forms.MessageBox.Show("Загрузил сборку " + fileName);
                    return System.Reflection.Assembly.LoadFile(fileName);
                }
            }
            catch (Exception ex)
            {
                //System.Windows.Forms.MessageBox.Show(string.Format("Ошибка загрузки сборки {0}.\n\nОписание:\n{1}", args.Name, ex.Message),
                //    "Ошибка", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                throw new Exception(string.Format("Ошибка загрузки сборки {0}.\n\nОписание:\n{1}", args.Name, ex.Message, ex.InnerException));

                
            }
            return null;
        }
    }
}