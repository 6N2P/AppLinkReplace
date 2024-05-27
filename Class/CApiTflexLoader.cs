using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace AppLinkReplace.Class
{
    public class CApiTflexLoader
    {
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

            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolve;
            //  System.Windows.Forms.MessageBox.Show("Подписался на событие");
        }
 
        public bool InitializeTFlexCadapi()
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
                var setup = new TFlex.ApplicationSessionSetup
                                {
                                    ReadOnly = false,
                                    //DOCsAPIVersion = TFlex.ApplicationSessionSetup.DOCsVersion.Version12
                                };

                TFlex.Application.EnableNotRespondingDialog(false);
                TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.AutoRefresh;
                return MIsInit = TFlex.Application.InitSession(setup);
            }
            catch (Exception e) { throw new Exception("Ошибка инициализации T-Flex CAD",e.InnerException); }
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
                AppDomain.CurrentDomain.AssemblyResolve -= AssemblyResolve;

                _folders = null;
                MIsInit = false;
            }
        }

        public bool IsInit { get { return MIsInit; } }
 
        protected bool MIsInit = false;

        private List<string> _folders;

        private String GetTopSystemsTFlexCadPath()
        {
            RegistryKey key;
            try
            {
                if (IntPtr.Size == 8)
                {
                    key = Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 12 x64\Rus"), RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey) ??
                          Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Wow6432Node\Top Systems\{0}\", @"T-FLEX CAD 3D 12\Rus"), RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                }
                else
                {
                    key = Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Top Systems\{0}\", @"T-FLEX CAD 3D 12\Rus"), RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
                }
            }
            catch (Exception e) { throw new Exception("Ошибка при поиске ключа рееста T-Flex CAD",e.InnerException); }
            var path = GetCurrentPath(key);
            if (key != null)
            {key.Close();}
            return path;
        }
        private String GetCurrentPath(RegistryKey key)
        {
            String path =null;
            try
            {
                if (key != null)
                {
                    path = (string)key.GetValue("ProgramFolder", string.Empty);
                    if (string.IsNullOrEmpty(path))
                        path = (string)key.GetValue("SetupHelpPath", string.Empty);


                    if (path.Length > 0 && path[path.Length - 1] != '\\')
                        path += @"\";

                }
            }
            catch (Exception e) { throw new Exception("Ошибка получение пути из ключа реестра",e.InnerException); }
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
                var name = args.Name;

                if (name != null)
                {
                    var index = name.IndexOf(",", StringComparison.Ordinal);
                    if (index > 0)
                        name = name.Substring(0, index);
                }

                foreach (var path in _folders)
                {
                    var fileName = string.Format("{0}{1}.dll", path, name);

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
                throw new Exception(string.Format("Ошибка загрузки сборки {0}.\n\nОписание:\n{1}", args.Name, ex.Message));
            }
            return null;
        }
    }
}
