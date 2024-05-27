using AppLinkReplace.Class;
using CTFlex.ExportDXF;
using Microsoft.Win32;
using MKMDPlugin.CFixLink;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;


//using System.Runtime.InteropServices;

//using PdfSharp.Drawing;
//using PdfSharp.Pdf;

namespace AppLinkReplace
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        #region winapi

        //[DllImport(@"C:\Windows\SysWOW64\kernel32.dll")]
        /*
                [DllImport("kernel32.dll")]
                private static extern IntPtr LoadLibrary(string lpFileName);
        */
        /*
                [DllImport("kernel32.dll")]
                public static extern bool SetProcessWorkingSetSize(IntPtr handle, int minimumWorkingSetSize, int maximumWorkingSetSize);
        */
        //[return: MarshalAs(UnmanagedType.Bool)]
        //static extern bool FreeLibrary(IntPtr hModule);

        /*
                private IntPtr _mLibrary;
        */

        #endregion winapi

        private static string _initDir;

        //private CApiTflexLoader MApiTflexLoader = null;
        private readonly IAPILoader _api;
       
        // private TFlex.ApplicationSessionSetup _setup = null;
        public MainWindow()
        {
            InitializeComponent();
            SetItemsCB();
            _api = new CAPILoader();               
        
             _api.Initialize();          

             _api.InitializeTFlexCADAPI(Properties.Settings.Default.BOMSectionsDatabase);

            RegistryKey test = Registry.CurrentUser.OpenSubKey(@"Software\TSM Utility", RegistryKeyPermissionCheck.ReadWriteSubTree);

            _initDir = "";

            if (test != null)
            {
                _initDir = (string)test.GetValue("Initial Directory");
                test.Close();
            }
            Title += String.Format(" используется TFLEX API v{0}", _api.TflexVersionApi);

            //string bomFile = Properties.Settings.Default.BOMSectionsDatabase;
            //if (File.Exists(bomFile))
            //{
            //    TFlex.Application.BOMSectionsDatabase = bomFile;
            //}
            //else
            //{
            //    System.Windows.Forms.MessageBox.Show(Properties.Resources.INFORMATION_BOM_FILE_PATH);
            //    return;
            //}

            //_mLibrary = LoadLibrary("tfbom");
            //text_status.Text = _mLibrary.ToString();
            //if (_mLibrary == IntPtr.Zero)
            //{
            //    System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_NOT_LOAD_TFBOMDLL);
            //}

            //////MApiTflexLoader = new CApiTflexLoader();
            //////MApiTflexLoader.Initialize();
            //////if (MApiTflexLoader.InitializeTFlexCadapi())
            //////{
            //////    string bomFile = Properties.Settings.Default.BOMSectionsDatabase;
            //////    if (File.Exists(bomFile))
            //////    {
            //////        TFlex.Application.BOMSectionsDatabase = bomFile;
            //////    }
            //////    else
            //////    {
            //////        System.Windows.Forms.MessageBox.Show(Properties.Resources.INFORMATION_BOM_FILE_PATH);
            //////        return; /*Close();*/
            //////    }

            //////    _mLibrary = LoadLibrary("tfbom");
            //////    text_status.Text = _mLibrary.ToString();
            //////    if (_mLibrary == IntPtr.Zero)
            //////    {
            //////        System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_NOT_LOAD_TFBOMDLL);
            //////    }
            //////}

            //_setup = new TFlex.ApplicationSessionSetup();
            //_setup.ReadOnly = false;
            //TFlex.Application.EnableNotRespondingDialog(false);
            //TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.AutoRefresh;
            ////================================================================================================================================
            //RegistryKey test = Registry.CurrentUser.OpenSubKey(@"Software\TSM Utility", RegistryKeyPermissionCheck.ReadWriteSubTree);

            //InitDir = "";

            //if (test != null)
            //{
            //    InitDir = (string)test.GetValue("Initial Directory");
            //    test.Close();
            //}
            ////test.Close();
            ////================================================================================================================================
            //if (TFlex.Application.SystemPath == @"\")
            //{
            //    try
            //    {
            //        TFlex.Application.InitSession(_setup);
            //        string bomFile = Properties.Settings.Default.BOMSectionsDatabase;
            //        if (File.Exists(bomFile))
            //        {
            //            TFlex.Application.BOMSectionsDatabase = bomFile;
            //        }
            //        else { System.Windows.Forms.MessageBox.Show(Properties.Resources.INFORMATION_BOM_FILE_PATH); return; /*Close();*/ }
            //    }
            //    catch /*(Exception e)*/ { System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_NOT_LOAD_TFLEX); return; }
            //    _mLibrary = LoadLibrary("tfbom");
            //    text_status.Text = _mLibrary.ToString();
            //    if (_mLibrary == IntPtr.Zero)
            //    {
            //        System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_NOT_LOAD_TFBOMDLL);
            //    }
            //}
            //cb_version.SelectedIndex = 0;
        }

        private void SetItemsCB()
        {
            cb_layer.ItemsSource = new List<string>() { "FREI", "ENGRAVING"};
            cb_layer.SelectedIndex = 0;
        }

        private void BOpenDirClick(object sender, RoutedEventArgs e)
        {
            // OFD = new OpenFileDialog();
            var fbd = new System.Windows.Forms.FolderBrowserDialog();

            FileInfo[] fileInfo = null;
            fbd.SelectedPath = _initDir;
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                 fileInfo = new DirectoryInfo(fbd.SelectedPath).GetFiles("*.grb", SearchOption.AllDirectories);
                _initDir = fbd.SelectedPath;
            }
            FillGrid(fileInfo);
        }

        private void FillGrid(IEnumerable<FileInfo> flexFileInfo)
        {
            _mIsThStop = false;

            // var swatch = new Stopwatch(); // создаем объект
            // swatch.Start(); // старт
            BlockUnBlockControls();
            if (flexFileInfo != null)
            {
                int count = flexFileInfo.Count();

                _mCountSelected = count;
                var waitTh1 = new System.Threading.Thread(StartProgress);
                waitTh1.SetApartmentState(System.Threading.ApartmentState.STA);
                waitTh1.Start();
                while (_windowProgress == null)
                {
                    System.Threading.Thread.Sleep(100);
                }

                int i = 0;
                foreach (FileInfo fi in flexFileInfo)
                {
                    if (_mIsThStop)
                    {
                        StopProgress();
                        BlockUnBlockControls();
                        _tflexList.Clear();
                        return;
                    }
                    i++;
                    String text = String.Format("Предварительная обработка. Обработанно {1} из {0} имя файла {2}", count,
                                                i, fi.Name);
                    //String text = "Предварительная обработка. Обработанно "+ count.ToString() +" из "+ i.ToString() +" имя файла "+fi.Name;

                    SetProgress(text, i);

                   
                    Document doc = TFlex.Application.OpenFragmentDocument(fi.FullName, false, false);
                    TFlex.Application.IdleSession();

                    _tflexList.Add(new CTflexFile(fi, new CTflexFileVariableInfo(doc)));
                    doc.Close();
                }
                cTflexFileDataGrid.DataContext = null;
                cTflexFileDataGrid.DataContext = _tflexList;
                StopProgress();
            }

            // swatch.Stop();

            // System.Windows.MessageBox.Show(swatch.Elapsed.ToString()); // результат
            _mCountSelected = 0;
            BlockUnBlockControls();
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            var cTflexFileViewSource = ((System.Windows.Data.CollectionViewSource)(FindResource("cTflexFileViewSource")));
            if (cTflexFileViewSource == null) throw new ArgumentNullException("sender");
        }

        private void BCheckSelectedClick(object sender, RoutedEventArgs e)
        {
            if (cTflexFileDataGrid.Items == null) return;
            foreach (CTflexFile tff in cTflexFileDataGrid.SelectedItems)
            {
                //if (tff.ReadOnly)
                //{
                //    continue;
                //}
                tff.Check = !tff.Check;
            }
        }

        private void BCheckAllClick(object sender, RoutedEventArgs e)
        {
            if (cTflexFileDataGrid.Items == null) return;
            foreach (CTflexFile tff in cTflexFileDataGrid.Items)
            {
                //if (tff.ReadOnly)
                //{
                //    continue;
                //}
                tff.Check = !tff.Check;
            }
        }

        private void BExitClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BStartClick(object sender, RoutedEventArgs e)
        {
            //_th = new System.Threading.Thread(StartAction) {Name = "ThreadWork"};
            _mCountSelected = GetCountSelectedElements();
            //m_is_scep_update = cb_recalc_bom.IsChecked.Value;
            //m_is_export = cb_export_png.IsChecked.Value;
            //m_is_link_replase = cb_link_replase.IsChecked.Value;
            //_mProp = (Int32)sliderCoef.Value;
            if (_mIsExport)
            {
                if (_mDirectorySave == null)
                {
                    System.Windows.MessageBox.Show("Не указана директория для Экспорта...");
                    return;
                }
            }
            _mInvNum = tb_inv_num.Text;
            _mIsThStop = false;
            //m_export_autocad = cb_autocad.IsChecked.Value;
            //type_dxf = (AutocadExportFileVersionType)cb_version.SelectedItem;

            //_th.Start();
            //var st = new System.Diagnostics.Stopwatch();
            //st.Start();
            if (_mCountSelected > 0)
            {
                StartAction();
            }
            //st.Stop();
            //TimeSpan ts = st.Elapsed;
            //string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
            //ts.Hours, ts.Minutes, ts.Seconds,
            //ts.Milliseconds / 10);
            //System.Windows.MessageBox.Show(elapsedTime);
        }

        #region методы отвечающие за отображение хода выполнения

        private WindowProgress _windowProgress;

        private void StartProgress()
        {
            _windowProgress = new WindowProgress { ProgressMaximum = _mCountSelected };
            _windowProgress.OnStopEvent += new WindowProgress.OnStopVoidHandler(_windowProgress_OnStopEvent);
            _windowProgress.ShowDialog();
        }

        private void _windowProgress_OnStopEvent(WindowProgress sender)
        {
            if (sender != null)
            {
                _mIsThStop = true;
            }
        }

        private void StopProgress()
        {
            if (_windowProgress != null)
            {
                Dispatcher targetDisp = _windowProgress.Dispatcher;
                if (targetDisp.CheckAccess())
                {
                    _windowProgress.Close();
                }
                else
                {
                    targetDisp.Invoke(DispatcherPriority.Background, new
                                    Action(() => _windowProgress.Close()));
                }
                _windowProgress.OnStopEvent -= new WindowProgress.OnStopVoidHandler(_windowProgress_OnStopEvent);

                _windowProgress = null;
            }
        }

        private void SetProgress(string text, double progress)
        {
            if (_windowProgress != null)
            {
                Dispatcher targetDisp = _windowProgress.Dispatcher;
                if (targetDisp.CheckAccess())
                {
                    _windowProgress.SetProgress(text, progress);
                }
                else
                {
                    //targetDisp.Invoke(
                    //    DispatcherPriority.Normal,
                    //    (System.Windows.Forms.MethodInvoker)delegate() { _windowProgress.SetProgress(text, progress); });

                    targetDisp.Invoke(DispatcherPriority.Background, new
                                    Action(() => _windowProgress.SetProgress(text, progress)));
                }
            }
        }

        #endregion методы отвечающие за отображение хода выполнения

        private void StartAction()
        {
            BlockUnBlockControls();

            #region работа с PDF

            PdfDocument pdfDoc = null;
            if (tbMonohrom.IsChecked != null && (tbMonohrom.IsChecked != null && tbMonohrom.IsChecked.Value))
            {
                pdfDoc = new PdfDocument();
            }
            pdfDoc = new PdfDocument();

            #endregion работа с PDF

            var waitTh = new System.Threading.Thread(StartProgress);
            waitTh.SetApartmentState(System.Threading.ApartmentState.STA);
            waitTh.Start();
            //"Выбранно {0} из {1} Обработка {2} действие {3}"
            Int32 i = 0;
            if (cTflexFileDataGrid.Items != null)
            {
                foreach (CTflexFile tff in cTflexFileDataGrid.Items)
                {
                    if (_mIsThStop)
                    {
                        StopProgress();
                        BlockUnBlockControls();
                        return;
                    }
                    //if (tff.ReadOnly || !tff.Check /*&& (!m_is_link_replase || !m_is_scep_update) */)
                    //{
                    //    if (!m_is_export && !tff.Check)
                    //    {
                    //        continue;
                    //    }
                    //}
                    if (!tff.Check)
                    { continue; }

                    i++;
                    SetProgress(String.Format(MStatus, _mCountSelected.ToString(CultureInfo.InvariantCulture), _mCountAll.ToString(CultureInfo.InvariantCulture), tff.FileName, "Открытие", i.ToString(CultureInfo.InvariantCulture)), i);
                    Document tdoc;
                    try
                    {
                        tdoc = TFlex.Application.OpenDocument(tff.FullFileName, false);
                        TFlex.Application.IdleSession();
                        //tdoc.BeginChanges("перерисовка");
                        //var regOpt = new RegenerateOptions { Full = false, UpdateAllLinks = false, Projections = true };
                        //tdoc.Regenerate(regOpt);
                        //tdoc.Redraw(true);
                        //tdoc.EndChanges();
                    }
                    catch /*(Exception e)*/ { System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_TFLEX_FILE_OPEN); return; }
                    //if (tdoc.LastSavedVersion > TFlex.Application.Version)
                    //{
                    //    System.Windows.MessageBox.Show("Документа сохранен в TFlex CAD старшей версии");
                    //    continue;
                    //}
                    if (tdoc != null)
                    {
                        #region fix link

                        if (_mIsLinkReplase)
                        {
                            SetProgress(String.Format(MStatus, _mCountSelected.ToString(CultureInfo.InvariantCulture), _mCountAll.ToString(CultureInfo.InvariantCulture), tff.FileName, "Замена сылок", i.ToString(CultureInfo.InvariantCulture)), i);

                            var fixLink = new CFixLink(tdoc);
                            fixLink.FixLink();
                        }

                        #endregion fix link

                        #region specification update

                        if (_mIsScepUpdate)
                        {
                            SetProgress(String.Format(MStatus, _mCountSelected.ToString(CultureInfo.InvariantCulture), _mCountAll.ToString(CultureInfo.InvariantCulture), tff.FileName, "Обновление спецификации", i.ToString(CultureInfo.InvariantCulture)), i);

                            foreach (Text text in tdoc.GetTexts())
                            {
                                if (text.SubType != TextType.Undefined)
                                { continue; }
                                var bomTflex = text as BOMObject;

                                if (bomTflex == null) continue;
                                bomTflex.Document.BeginChanges("Обновление спецификации.");
                                //bom_tflex.BeginEdit();
                                //bom_tflex.UpdateRecord();
                                try
                                {
                                    bomTflex.Refresh(true);
                                }
                                catch { System.Windows.Forms.MessageBox.Show(Properties.Resources.ERROR_UPDATE_SPECIFICATION + bomTflex.FriendlyName); }
                                //bom_tflex.EndEdit();
                                bomTflex.Document.EndChanges();
                            }
                        }

                        #endregion specification update

                        #region экспорт файлов в PNG

                        if (_mIsExport)
                        {
                            SetProgress(String.Format(MStatus, _mCountSelected.ToString(CultureInfo.InvariantCulture), _mCountAll.ToString(CultureInfo.InvariantCulture), tff.FileName, "Экспорт документа в *.PNG", i.ToString(CultureInfo.InvariantCulture)), i);

                            #region делаем черно-белую картинку

                            //tdoc.BeginChanges("Изменения цвета слоя");
                            //foreach (Layer frDocLater in tdoc.Layers)
                            //{
                            //    if (frDocLater == null) continue;
                            //    frDocLater.Color = 0;
                            //}
                            //tdoc.EndChanges();

                            #endregion делаем черно-белую картинку

                            var export = tdoc.ExportToBitmap;
                            //поиск нужной страницы

                            foreach (Fragment fr in tdoc.GetFragments())
                            {
                                if (fr.FilePath == "{$formatka}" || fr.FilePath.Contains("форматка"))
                                {
                                    export.Page = fr.Page;
                                    //Double height = ((fr.Page.Rectangle.Height / 10) / 2.54 * _mProp);
                                    //Double width = ((fr.Page.Rectangle.Width / 10) / 2.54 * _mProp);
                                    Double height = ((fr.Page.Rectangle.Height * 300) / 25.4);
                                    Double width = ((fr.Page.Rectangle.Width * 300) / 25.4);

                                    export.Height = (int)height;
                                    export.Width = (int)width;
                                    String[] fileName = tdoc.FileName.Split('\\');

                                    #region если нада вставить подписи

                                    FragmentVariableValue varCryptLocal = fr.GetVariables().FirstOrDefault(varCrypt => varCrypt.Name == "podpis");
                                    //                                  Variable varDisp = tdoc.Variables.Cast<Variable>().FirstOrDefault(varDispD => varDispD.Name == "disp");

                                    #region изменяем свойство переменной подписи

                                    Document frDoc = fr.OpenPart();
                                    if (frDoc == null) continue;
                                    if (varCryptLocal != null)
                                    {
                                        try
                                        {
                                            if (_mIsCrypt)
                                            {
                                                frDoc.BeginChanges("Edit variable");
                                                tdoc.BeginChanges("doc");
                                                varCryptLocal.RealValue = 1;
                                                //                                                varDisp.Expression = "1";
                                                tdoc.EndChanges();
                                                frDoc.EndChanges();
                                            }
                                            else
                                            {
                                                frDoc.BeginChanges("Edit variable");

                                                tdoc.BeginChanges("doc");
                                                varCryptLocal.RealValue = 0;
                                                //varDisp.RealValue = 0;
                                                tdoc.EndChanges();
                                                frDoc.EndChanges();
                                            }
                                            RegenerateOptions regopt = new RegenerateOptions();
                                            //regopt.Full = true;

                                            //tdoc.Regenerate(regopt);
                                        }
                                        catch (Exception e)
                                        {
                                            throw new Exception("Фрагмент доступен только для чтения (" + fr.FilePath + ")", e.InnerException);
                                        }
                                    }

                                    #endregion изменяем свойство переменной подписи

                                    #endregion если нада вставить подписи

                                    #region экспорт в PDF

                                    tdoc.BeginChanges("");
                                    foreach (var layer in tdoc.GetLayers())
                                    {
                                        layer.ColorVariable = null;
                                        layer.Color = 0;
                                        layer.Monochrome = true;
                                    }
                                    tdoc.ApplyChanges();

                                    #region Делаем проверку на тип документа для правильного формирования имени пдф файлов

                                    var fnPdf = fileName[fileName.Length - 1];
                                    var varTypeDoc = tdoc.FindVariable("type_doc");
                                    if (varTypeDoc != null)
                                    {
                                        switch (varTypeDoc.RealValue.ToString())
                                        {
                                            case "1":
                                            case "2":
                                                {
                                                    var varMarka = tdoc.FindVariable("$marka");
                                                    var varName = tdoc.FindVariable("$name");
                                                    if (varMarka != null && varName != null)
                                                    {
                                                        fnPdf = varMarka.TextValue.Replace('/', '_') + "_" + varName.TextValue.Replace('/', '_');
                                                    }
                                                }
                                                break;

                                            case "0":
                                            case "4":
                                                {
                                                    var varPoz = tdoc.FindVariable("$poz");
                                                    var varName = tdoc.FindVariable("$name");
                                                    if (varPoz != null && varName != null)
                                                    {
                                                        fnPdf = "Поз." + varPoz.TextValue + "_" + varName.TextValue.Replace('/', '_');
                                                    }
                                                }
                                                break;
                                        }
                                    }

                                    #endregion Делаем проверку на тип документа для правильного формирования имени пдф файлов

                                    var pdfExp = new ExportToPDF(tdoc);

                                    const string pageName = "чертеж";
                                    var wPage = tdoc.GetPages().FirstOrDefault(page => page.Name == pageName);

                                    // добавленно в 12.63 API CAD
                                    pdfExp.ExportPages.Add(wPage);

                                    pdfExp.Export(_mDirectorySave + @"\" + fnPdf + ".pdf");
                                    tdoc.CancelChanges();

                                    var inputDoc = PdfReader.Open(_mDirectorySave + @"\" + fnPdf + ".pdf", PdfDocumentOpenMode.Import);
                                    int count = inputDoc.PageCount;
                                    for (int idx = 0; idx < count; idx++)
                                    {
                                        pdfDoc.AddPage(inputDoc.Pages[idx]);
                                    }
                                    //  inputDoc.Close();

                                    #endregion экспорт в PDF

                                    if (export.Export(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + ".png", ImageExport.ScreenLayers, ImageExportFormat.Png))
                                    {
                                        //                          export.Export(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + "Constructions.png",
                                        //                                        ImageExport.Constructions, ImageExportFormat.Png);

                                        //                          export.Export(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + "Default.png",
                                        //ImageExport.Default, ImageExportFormat.Png);

                                        //                      //    export.Export(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + "None.png",
                                        //                      //                  ImageExport.None, ImageExportFormat.Png);

                                        //                          export.Export(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + "ScreenLayers.png",
                                        //                                        ImageExport.ScreenLayers, ImageExportFormat.Png);

                                        //                          export.Export(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + "ScreenLayers_None.png",
                                        //                          ImageExport.ScreenLayers | ImageExport.None, ImageExportFormat.Png);

                                        if (tbMonohrom.IsChecked != null && (tbMonohrom.IsChecked != null && tbMonohrom.IsChecked.Value) && pdfDoc != null)
                                        {
                                            //var colourBitMaps =
                                            //    new Bitmap(_mDirectorySave + @"\" + fileName[fileName.Length - 1] +
                                            //               ".png");
                                            //colourBitMaps = MakeBW(colourBitMaps);
                                            //colourBitMaps.SetResolution(300, 300);
                                            //colourBitMaps.Save(_mDirectorySave + @"\" + fileName[fileName.Length - 1] +
                                            //                   "_BW.png");
                                            //var pageP = new PdfPage {Height = height, Width = width};

                                            //pdfDoc.Pages.Add(pageP);
                                            //  var inputDoc = PdfReader.Open(_mDirectorySave + @"\" + fileName[fileName.Length - 1] + ".pdf");
                                            //  pdfDoc.AddPage(inputDoc.Pages[0]);
                                            //var xgr = XGraphics.FromPdfPage(pdfDoc.Pages[i - 1]);
                                            //var img =
                                            //    XImage.FromBitmapSource(CreateBitmapSourceFromBitmap(colourBitMaps));
                                            //xgr.DrawImage(img, 0, 0, width, height);
                                        }
                                    }

                                    if (frDoc != null) frDoc.Close();

                                    #region экспорт в dwg

                                    if (tb_export_dwg.IsChecked.Value == true)
                                    {
                                        var exportdwg = tdoc.ExportToDWG;
                                        exportdwg.AutocadExportFileVersion = AutocadExportFileVersionType.efACAD2000;
                                        exportdwg.Page = tdoc.GetPages().FirstOrDefault(page => page.Name == pageName);
                                        exportdwg.Export(_mDirectorySave + @"\" + fnPdf + ".dwg");
                                    }

                                    #endregion экспорт в dwg

                                    #region экспорт в step

                                    if (tb_export_step.IsChecked.Value == true)
                                    {
                                        var exportParasolid = tdoc.ExportToParasolid;
                                        exportParasolid.Export(_mDirectorySave + @"\" + fnPdf + ".x_t");
                                    }

                                    #endregion экспорт в step
                                }
                            }
                        }

                        #endregion экспорт файлов в PNG

                        #region экспорт в DXF

                        if (_mIsExportDxf)
                        {
                            SetProgress(String.Format(MStatus, _mCountSelected.ToString(CultureInfo.InvariantCulture), _mCountAll.ToString(CultureInfo.InvariantCulture), tff.FileName, "Экспорт документа в *.DXF", i.ToString(CultureInfo.InvariantCulture)), i);

                            ParseDocument(tdoc);
                            //ExportDXFNoTFlex(tdoc);
                            //ExportDXF(tdoc);
                        }

                        #endregion экспорт в DXF

                        #region Временная изменение переменных

                        if (_is_insert_value)
                        {
                            tdoc.BeginChanges("Замена выражения переменной");
                            var variable = tdoc.FindVariable("m");
                            if (variable != null)
                            {
                                variable.Expression = "round(abs(s)*8.5e-6*pl,0.1)";
                            }
                            variable = tdoc.FindVariable("$mat");
                            if (variable != null)
                            {
                                variable.Expression = "\"" + "15ХСНД-3\\nГОСТ 6713-91" + "\"";
                            }
                            variable = tdoc.FindVariable("$prover");
                            if (variable != null)
                            {
                                variable.Expression = "\"" + "Тришин" + "\"";
                            }
                            tdoc.EndChanges();
                            tdoc.Save();
                        }

                        #endregion Временная изменение переменных

                        #region сохраняем документ если нада

                        if ((_mIsLinkReplase || _mIsScepUpdate) && !tff.ReadOnly)
                        {
                            SetProgress(String.Format(MStatus, _mCountSelected.ToString(CultureInfo.InvariantCulture), _mCountAll.ToString(CultureInfo.InvariantCulture), tff.FileName, "Сохранение", i.ToString(CultureInfo.InvariantCulture)), i);

                            tdoc.Save();
                        }
                        tdoc.Close();

                        #endregion сохраняем документ если нада
                    }
                }
            }
            if (tbMonohrom.IsChecked != null && (tbMonohrom.IsChecked != null && tbMonohrom.IsChecked.Value) && pdfDoc != null)
            {
                System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
                saveDialog.Title = "Сохранить PDF книгу как ?";
                saveDialog.Filter = "PDF Файлы | *.pdf";
                saveDialog.DefaultExt = "pdf";
                saveDialog.FileName = "Книга чертежей";
                saveDialog.InitialDirectory = _mDirectorySave;
                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    pdfDoc.Save(/*_mDirectorySave + @"\Книга чертежей на инвентарный №.pdf"*/saveDialog.FileName);
                }
                pdfDoc.Close();
            }
            StopProgress();
            BlockUnBlockControls();
        }

        private BitmapSource CreateBitmapSourceFromBitmap(Bitmap bitmap)
        {
            if (bitmap == null)
                throw new ArgumentNullException("bitmap");

            using (var memoryStream = new MemoryStream())
            {
                try
                {
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    var bitmapDecoder = BitmapDecoder.Create(
                        memoryStream,
                        BitmapCreateOptions.PreservePixelFormat,
                        BitmapCacheOption.OnLoad);

                    var writable =
            new WriteableBitmap(bitmapDecoder.Frames.Single());
                    writable.Freeze();

                    return writable;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        private Bitmap MakeBW(Bitmap source)
        {
            ////использование промежуточных переменных ускоряет
            ////код в несколько раз
            //var width = source.Width;
            //var height = source.Height;

            //var sourceData = source.LockBits(new Rectangle(new System.Drawing.Point(0, 0), source.Size),
            //                 ImageLockMode.ReadWrite,
            //                 source.PixelFormat);

            //var sourceStride = sourceData.Stride;

            //var sourceScan0 = sourceData.Scan0;

            //var resultPixelSize = sourceStride / width;

            //unsafe
            //{
            //    for (var y = 0; y < height; y++)
            //    {
            //        var sourceRow = (byte*)sourceScan0 + (y * sourceStride);
            //        for (var x = 0; x < width; x++)
            //        {
            //            var v = (byte)((sourceRow[x * resultPixelSize + 2] + sourceRow[x * resultPixelSize + 1] + sourceRow[x * resultPixelSize]) / 3);

            //            v = (byte)(v != 255 ? 0 : 255);
            //            sourceRow[x * resultPixelSize] = v;
            //            sourceRow[x * resultPixelSize + 1] = v;
            //            sourceRow[x * resultPixelSize + 2] = v;
            //        }
            //    }

            //}

            //source.UnlockBits(sourceData);
            //return source;
            return null;
        }

        /*
		 *  private Bitmap MakeBW(Bitmap source)
		{
			//использование промежуточных переменных ускоряет
			//код в несколько раз
			var width = source.Width;
			var height = source.Height;

			var sourceData = source.LockBits(new Rectangle(new System.Drawing.Point(0, 0), source.Size),
							 ImageLockMode.ReadOnly,
							 source.PixelFormat);

			var result = new Bitmap(width, height, source.PixelFormat);
			var resultData = result.LockBits(new Rectangle(new System.Drawing.Point(0, 0), result.Size),
							 ImageLockMode.ReadWrite,
							 source.PixelFormat);

			var sourceStride = sourceData.Stride;
			var resultStride = resultData.Stride;

			var sourceScan0 = sourceData.Scan0;
			var resultScan0 = resultData.Scan0;

			var resultPixelSize = resultStride / width;

			unsafe
			{
				for (var y = 0; y < height; y++)
				{
					var sourceRow = (byte*)sourceScan0 + (y * sourceStride);
					var resultRow = (byte*)resultScan0 + (y * resultStride);
					for (var x = 0; x < width; x++)
					{
						//var v = (byte)(0.3 * sourceRow[x * resultPixelSize + 2] + 0.59 * sourceRow[x * resultPixelSize + 1] +
						//   0.11 * sourceRow[x * resultPixelSize]);
						var v = (byte)((sourceRow[x * resultPixelSize + 2] + sourceRow[x * resultPixelSize + 1] + sourceRow[x * resultPixelSize]) / 3);

						if (v != 255)
						{
							v = 0;
						}
						else
						{
							v = 255;
						}
						resultRow[x * resultPixelSize] = v;
						resultRow[x * resultPixelSize + 1] = v;
						resultRow[x * resultPixelSize + 2] = v;
					}
				}
			}

			source.UnlockBits(sourceData);
			result.UnlockBits(resultData);
			return result;
		}
		 */

        private void ParseDocument(Document tdoc)
        {
            if (tdoc != null)
            {
                Variable typeDoc = tdoc.FindVariable("type_doc");
                if (typeDoc != null)
                {
                    String td = typeDoc.RealValue.ToString(CultureInfo.InvariantCulture);
                    switch (td)
                    {
                        #region парсим документ и експортируем 1 4 -1 3

                        case "1":
                        case "4":
                            {
                                if (tdoc.GetFragments3D().Count > 0)
                                {
                                    foreach (Fragment frag in tdoc.GetFragments())
                                    {
                                        //Boolean isFind = false;
                                        //foreach (FragmentVariableValue frag_root_var in frag.VariableValues)
                                        //{
                                        //    if (frag_root_var.Name == "type_doc")
                                        //    {
                                        //        is_find = true;
                                        //        break;
                                        //    }
                                        //}
                                        //if (is_find)
                                        //{
                                        Document docFrag = frag.GetFragmentDocument(true);//OpenPart(); // GetFragmentDocument(true);
                                        if (docFrag != null)
                                        {
                                            Variable typeDocFrag = docFrag.FindVariable("type_doc");
                                            if (typeDocFrag != null)
                                            {
                                                if (typeDocFrag.RealValue == 3)
                                                {
                                                    ExportDxfnoTFlex(docFrag);
                                                    //if (docFrag != null) docFrag.Close();
                                                    continue;
                                                }
                                                if (typeDocFrag.RealValue == -1)
                                                {
                                                    ExportDxfnoTFlex(docFrag);
                                                    foreach (Fragment fragVs in docFrag.GetFragments())
                                                    {
                                                        //foreach (
                                                        //    FragmentVariableValue frag_var in fragVs.VariableValues)
                                                        //{
                                                        //    if (frag_var.Name == "type_doc")
                                                        //    {
                                                        Document docFragVs = fragVs.GetFragmentDocument(true); //OpenPart();
                                                                                                               // GetFragmentDocument(true);
                                                        ParseDocument(docFragVs);
                                                        //docFragVs.Close();
                                                        //    }
                                                        //}
                                                    }
                                                }
                                            }
                                        }
                                        //if (docFrag != null) docFrag.Close();
                                        //}
                                    }
                                    ExportDxfnoTFlex(tdoc);
                                }
                                else
                                {
                                    ExportDxfnoTFlex(tdoc);
                                }
                            }
                            break;

                        #endregion парсим документ и експортируем 1 4 -1 3

                        #region во всех остальных случаях считаем что мы нашли и парсим

                        default:
                            { ExportDxfnoTFlex(tdoc); }
                            break;

                            #endregion во всех остальных случаях считаем что мы нашли и парсим
                    }
                }
            }
        }

        private void ExportDxfnoTFlex(Document tdoc)
        {
            var exp = new CExportDXF(_mDirectorySave, _mInvNum, cb_layer.SelectedValue.ToString());
            exp.GetDXFMemoryStream(tdoc, false);
        }

        private Int32 GetCountSelectedElements()
        {
            Int32 i = 0;
            if (cTflexFileDataGrid.Items != null)
            {
                i += cTflexFileDataGrid.Items.Cast<CTflexFile>().Count(tff => tff.Check);
            }
            if (cTflexFileDataGrid.Items != null) _mCountAll = cTflexFileDataGrid.Items.Count;
            //pb_process.Maximum = i;
            return i;
        }

        private void WindowClosed(object sender, EventArgs e)
        {
            RegistryKey test = Registry.CurrentUser.OpenSubKey(@"Software\TSM Utility", RegistryKeyPermissionCheck.ReadWriteSubTree) ??
                               Registry.CurrentUser.CreateSubKey(@"Software\TSM Utility");
            if (_initDir != null)
            {
                if (test != null) test.SetValue("Initial Directory", _initDir);
            }
            if (test != null) test.Close();
            _api.Terminate();
          //APILoader17.Terminate();

            //////MApiTflexLoader.Terminate();
            //TFlex.Application.BOMSectionsDatabase = null;
            //TFlex.Application.TerminateAllCommands();
            //TFlex.Application.ExitSession();
            //_setup = null;
        }

        private void BlockUnBlockControls()
        {
            b_check_all.IsEnabled = !b_check_all.IsEnabled;
            b_check_selected.IsEnabled = !b_check_selected.IsEnabled;
            b_exit.IsEnabled = !b_exit.IsEnabled;
            b_open_dir.IsEnabled = !b_open_dir.IsEnabled;
            b_start.IsEnabled = !b_start.IsEnabled;
            cTflexFileDataGrid.IsEnabled = !cTflexFileDataGrid.IsEnabled;

            //b_check_all.Dispatcher.Invoke(new MethodInvoker(delegate { b_check_all.IsEnabled = !b_check_all.IsEnabled; }));
            //b_check_selected.Dispatcher.Invoke(new MethodInvoker(delegate { b_check_selected.IsEnabled = !b_check_selected.IsEnabled; }));
            //b_exit.Dispatcher.Invoke(new MethodInvoker(delegate { b_exit.IsEnabled = !b_exit.IsEnabled; }));
            //b_open_dir.Dispatcher.Invoke(new MethodInvoker(delegate { b_open_dir.IsEnabled = !b_open_dir.IsEnabled; }));
            //b_start.Dispatcher.Invoke(new MethodInvoker(delegate { b_start.IsEnabled = !b_start.IsEnabled; }));
            //cTflexFileDataGrid.Dispatcher.Invoke(new MethodInvoker(delegate { cTflexFileDataGrid.IsEnabled = !cTflexFileDataGrid.IsEnabled; }));
        }

        #region variable

        private Int32 _mCountSelected;
        private Int32 _mCountAll;
        private Boolean _mIsScepUpdate;
        private Boolean _mIsExport;
        private Boolean _mIsExportDxf;

        // private Boolean _mIsExportDWG;
        private const String MStatus = "Выбранно {0} из {1} Обработка {2} действие {3} ({4} из {0})";

        private Boolean _mIsThStop;
        private Boolean _mIsLinkReplase;
        private Boolean _mIsCrypt;

        // private System.Threading.Thread _th;
        private String _mDirectorySave;

        //private Int32 _mProp = 200;
        private readonly List<CTflexFile> _tflexList = new List<CTflexFile>();

        private String _mInvNum;

        //   protected AutocadExportFileVersionType type_dxf = AutocadExportFileVersionType.efACAD2000;
        private Boolean _is_insert_value = false;

        #endregion variable

        private void BtClearClick(object sender, RoutedEventArgs e)
        {
            _tflexList.Clear();
            cTflexFileDataGrid.DataContext = null;
        }

        /*
                private void CbExportPngClick(object sender, RoutedEventArgs e)
                {
                    //   b_select_save_dir.IsEnabled = cb_export_png.IsChecked.Value;
                }
        */

        private void BSelectSaveDirClick(object sender, RoutedEventArgs e)
        {
            var fbd = new System.Windows.Forms.FolderBrowserDialog();

            if (fbd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            _mDirectorySave = fbd.SelectedPath;
            m_hlink.Inlines.Clear();
            m_hlink.Inlines.Add(_mDirectorySave);
            m_hlink.NavigateUri = new Uri(_mDirectorySave);
        }

        private void HyperlinkClick(object sender, RoutedEventArgs e)
        {
            var source = sender as Hyperlink;
            if (source != null)
            {
                System.Diagnostics.Process.Start(source.NavigateUri.ToString());
            }
        }

        private void TbLinkReplaceClick(object sender, RoutedEventArgs e)
        {
            if (tb_link_replace.IsChecked != null) _mIsLinkReplase = tb_link_replace.IsChecked.Value;
        }

        private void TbSpecUpdateClick(object sender, RoutedEventArgs e)
        {
            if (tb_spec_update.IsChecked != null) _mIsScepUpdate = tb_spec_update.IsChecked.Value;
        }

        private void TbExportClick(object sender, RoutedEventArgs e)
        {
            if (tb_export.IsChecked != null) _mIsExport = tb_export.IsChecked.Value;
            if (_mIsExport) System.Windows.MessageBox.Show("Отключите используемые экранные слои во фрагментах (н-р: сварной шов.grb)", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            // sliderCoef.IsEnabled = _mIsExport;
            b_select_save_dir.IsEnabled = _mIsExport;
            tb_crypt.IsEnabled = _mIsExport;
            tbMonohrom.IsEnabled = _mIsExport;
            //if (m_is_export)
            //{
            //    tb_prop.IsEnabled = true;
            //    b_select_save_dir.IsEnabled = true;
            //    tb_crypt.IsEnabled = true;
            //}
            //else
            //{
            //    tb_prop.IsEnabled = false;
            //    b_select_save_dir.IsEnabled = false;
            //    tb_crypt.IsEnabled = false;
            //}
        }

        private void TbCryptClick(object sender, RoutedEventArgs e)
        {
            if (tb_crypt.IsChecked != null) _mIsCrypt = tb_crypt.IsChecked.Value;
        }

        private void TbExportDxfClick(object sender, RoutedEventArgs e)
        {
            if (tb_export_dxf.IsChecked != null) _mIsExportDxf = tb_export_dxf.IsChecked.Value;
            b_select_save_dir.IsEnabled = _mIsExportDxf;
            tb_inv_num.IsEnabled = _mIsExportDxf;
            //cb_version.IsEnabled = m_is_export_dxf;
            //cb_autocad.IsEnabled = m_is_export_dxf;
        }

        //private void SliderCoefValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        //{
        //    if (label1 != null) label1.Content = (Int32)e.NewValue;
        //}

        private void BtHelpClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var currentPath = System.Windows.Forms.Application.StartupPath;
                System.Diagnostics.Process.Start(currentPath + @"\HelpDoc\Пакетная обработка файлов.chm");
            }
            catch
            {
                System.Windows.MessageBox.Show("Файл справки не удалось запустить", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region управление видимостью колонок

        private void cbName_Checked(object sender, RoutedEventArgs e)
        {
            if (fileNameColumn != null) fileNameColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbName_Unchecked(object sender, RoutedEventArgs e)
        {
            if (fileNameColumn != null) fileNameColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbType_Checked(object sender, RoutedEventArgs e)
        {
            if (docTypeColumn != null) docTypeColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbType_Unchecked(object sender, RoutedEventArgs e)
        {
            if (docTypeColumn != null) docTypeColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbFormat_Checked(object sender, RoutedEventArgs e)
        {
            if (formatColumn != null) formatColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbFormat_Unchecked(object sender, RoutedEventArgs e)
        {
            if (formatColumn != null) formatColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void chMassa_Checked(object sender, RoutedEventArgs e)
        {
            if (massaColumn != null) massaColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void chMassa_Unchecked(object sender, RoutedEventArgs e)
        {
            if (massaColumn != null) massaColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbArea_Checked(object sender, RoutedEventArgs e)
        {
            if (areaColumn != null) areaColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbArea_Unchecked(object sender, RoutedEventArgs e)
        {
            if (areaColumn != null) areaColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbMaterial_Checked(object sender, RoutedEventArgs e)
        {
            if (materialColumn != null) materialColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbMaterial_Unchecked(object sender, RoutedEventArgs e)
        {
            if (materialColumn != null) materialColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbSech_Checked(object sender, RoutedEventArgs e)
        {
            if (sechMkmdColumn != null) sechMkmdColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbSech_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sechMkmdColumn != null) sechMkmdColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbSize_Checked(object sender, RoutedEventArgs e)
        {
            if (fileSizeColumn != null) fileSizeColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbSize_Unchecked(object sender, RoutedEventArgs e)
        {
            if (fileSizeColumn != null) fileSizeColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbPath_Checked(object sender, RoutedEventArgs e)
        {
            if (fullFileNameColumn != null) fullFileNameColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbPath_Unchecked(object sender, RoutedEventArgs e)
        {
            if (fullFileNameColumn != null) fullFileNameColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbComment_Checked(object sender, RoutedEventArgs e)
        {
            if (commentColumn != null) commentColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbComment_Unchecked(object sender, RoutedEventArgs e)
        {
            if (commentColumn != null) commentColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        #endregion управление видимостью колонок

        private void tb_tflex_var_Checked(object sender, RoutedEventArgs e)
        {
            if (tb_tflex_var.IsChecked != null) _is_insert_value = tb_tflex_var.IsChecked.Value;
        }

        private void cbList_Checked(object sender, RoutedEventArgs e)
        {
            if (listColumn != null) listColumn.Visibility = System.Windows.Visibility.Visible;
        }

        private void cbList_Unchecked(object sender, RoutedEventArgs e)
        {
            if (listColumn != null) listColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbPoz_Unchecked(object sender, RoutedEventArgs e)
        {
            if (pozColumn != null) pozColumn.Visibility = System.Windows.Visibility.Hidden;
        }

        private void cbPoz_Checked(object sender, RoutedEventArgs e)
        {
            if (pozColumn != null) pozColumn.Visibility = System.Windows.Visibility.Visible;
        }
    }
}