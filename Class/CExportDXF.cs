using AppLinkReplace.Class;
using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Linq;
using TFlex.Model;
using TFlex.Model.Model2D;

namespace CTFlex.ExportDXF
{
    /// <summary>
    /// Класс предназанчен для экспорта страницы "contour" документа T-Flex CAD в векторный формат *.DXF для дальнейшей обработки технологами ЧПУ
    /// </summary>
    public class CExportDXF
    {
        /// <summary>
        /// Конструктор класса CExportDXF экспорт контуров в векторный формат *.DXF
        /// </summary>
        /// <param name="directoryDestination">Директория назначения куда будет помещен результат экспорта</param>
        /// <param name="invNum">Инвентарный номер</param>
        public CExportDXF(String directoryDestination, String invNum, String layer = "FREI")
        {
            if (invNum == null)
            {
                throw new ArgumentNullException("invNum", "Не указан инвентарный номер");
            }
            _mInvNum = invNum;
            if (directoryDestination == null)
            {
                throw new ArgumentNullException("directoryDestination", "Не указана директория назначения");
            }
            if (!Directory.Exists(directoryDestination))
            {
                throw new ArgumentException("DirectoryDestination", "Указанная директория не существует");
            }
            _mDirectorySave = directoryDestination;
            _mlayer = layer;
        }

        #region пересчет геометрии

        /*
                /// <summary>
                /// Пересчет координат какихто дал Ефремов Андрей Николаевич
                /// </summary>
                /// <param name="ang">Угол</param>
                /// <param name="x">Координата X</param>
                /// <param name="y">Координата Y</param>
                private void AffineTr(double ang, ref double x, ref double y)
                {
                    ang = -(Math.PI * ang / 180.0);
                    double x2 = x;
                    x = Math.Round(x * Math.Cos(ang) - y * Math.Sin(ang), 3);
                    y = Math.Round(x2 * Math.Sin(ang) + y * Math.Cos(ang), 3);
                }
        */

        /// <summary>
        /// Пересчет угла
        /// </summary>
        /// <param name="xc">Точка центра X</param>
        /// <param name="yc">Точка центра Y</param>
        /// <param name="x">Точка X</param>
        /// <param name="y">Точка Y</param>
        /// <returns>Скоректированный угол</returns>
        private static double DefAngle(double xc, double yc, double x, double y)
        {
            var dx = Math.Round(x - xc, 10);
            var dy = Math.Round(y - yc, 10);

            double ang;

            while (true)
            {
                if (dx == 0)
                {
                    ang = y < yc ? 270 : 90;
                    break;
                }
                if (dy == 0)
                {
                    ang = x < xc ? 180 : 0;
                    break;
                }
                ang = Math.Round(Math.Atan(dy / dx) * (180 / Math.PI), 10);
                if (x > xc)
                {
                    if (y < yc)
                    {
                        ang = 360 + ang;
                    }
                }
                else
                {
                    ang = y < yc ? 180 + Math.Abs(ang) : 180 - Math.Abs(ang);
                }
                break;
            }
            return ang;
        }

        #endregion пересчет геометрии

        #region variable

        private readonly String _mDirectorySave;
        private readonly String _mInvNum;
        private readonly String _mlayer;

        #endregion variable

        /// <summary>
        /// Метод экспорта документа в *.DXF контура
        /// </summary>
        /// <param name="doc">Документ T-Flex CAD</param>
        /// <param name="isPlugin">Признак сделан для плагина если true то отрезаются некоторые функции связанные с созданием директорий</param>
        /// <returns>Результат успешности выполения операции при успешном true неудача false</returns>
        public Boolean GetDXFMemoryStream(Document doc, Boolean isPlugin)
        {
            if (doc == null)
                return false;

            Page pageExp = null;

            #region затычка для -1 типа

            var varDocType = doc.FindVariable("type_doc");
            var isTypeBed = false;

            #endregion затычка для -1 типа

            if (isPlugin)
            {
                foreach (var conPage in doc.GetPages().Where(conPage => conPage.Name == "contour"))
                {
                    pageExp = conPage;
                    break;
                }
                if (pageExp == null)
                    pageExp = doc.ActivePage;
            }
            else
            {
                foreach (var countrPage in doc.GetPages())
                {
                    if (countrPage.Name == "contour")
                    {
                        pageExp = countrPage;
                        break;
                    }
                    if (varDocType != null && varDocType.RealValue == -1 && countrPage.Name.ToLower() == "план")
                    {
                        pageExp = countrPage;
                        isTypeBed = true;
                        break;
                    }
                }
            }
            var isCircle = false;
            var fileExport = "";
            if (pageExp != null)
            {
                var activeScale = pageExp.Scale.Value;

                #region формируем структуру DXF файла

                var nfi = new CultureInfo("en-US", false).NumberFormat;
                nfi.NumberDecimalSeparator = ".";
                // MemoryStream ms = new MemoryStream();

                var swDXF = new StreamWriter(_mDirectorySave + @"\temp.dxf");
                swDXF.WriteLine(" 0\r\nSECTION\r\n 2\r\nHEADER\r\n  9\r\n$DIMTAD\r\n 70\r\n     1\r\n 0\r\nENDSEC");
                swDXF.WriteLine(" 0\r\nSECTION\r\n 2\r\nTABLES");
                swDXF.WriteLine(" 0\r\nTABLE\r\n 2\r\nLTYPE\r\n 70\r\n1\r\n 0\r\nENDTAB");
                // swDXF.WriteLine(" 0\r\nTABLE\r\n 2\r\nLAYER\r\n 70\r\n1\r\n 0\r\nLAYER\r\n 2\r\n0\r\n 70\r\n0\r\n 62\r\n7\r\n 6\r\nCONTINUOUS\r\n 0\r\nENDTAB");
                swDXF.WriteLine(
                    $" 0\r\nTABLE\r\n 2\r\nLAYER\r\n 70\r\n2\r\n 0\r\nLAYER\r\n 2\r\n0\r\n 70\r\n0\r\n 62\r\n7\r\n 6\r\nCONTINUOUS\r\n 0\r\nLAYER\r\n 2\r\n{_mlayer}\r\n 70\r\n0\r\n 62\r\n3\r\n 6\r\nCONTINUOUS\r\n 0\r\nENDTAB");
                swDXF.WriteLine(" 0\r\nTABLE\r\n 2\r\nSTYLE\r\n 70\r\n1\r\n 0\r\nSTYLE\r\n 2\r\nSTANDARD");
                swDXF.WriteLine(
                    " 70\r\n0\r\n 40\r\n0.500\r\n 41\r\n1.000\r\n 50\r\n30.000\r\n 71\r\n0\r\n 42\r\n0.500\r\n 3\r\ntxt\r\n 4\r\n\r\n 0\r\nENDTAB");
                swDXF.WriteLine(" 0\r\nENDSEC\r\n 0\r\nSECTION\r\n 2\r\nENTITIES");

                var htElements = new Hashtable();

                #region пишем в dxf толщину

                //закоментировал для спец выпуска
                var thickness = 0.0;
                var varThickness = doc.FindVariable("s");
                if (varThickness != null) thickness = varThickness.RealValue;
                var th = " 0\r\n";
                th += "TEXT\r\n";
                th += " 5\r\n";
                th += "89\r\n";
                th += " 8\r\n";
                th += "FREI\r\n";
                th += " 62\r\n";
                th += "    15\r\n";
                th += " 10\r\n";
                th += "2.12\r\n";
                th += " 20\r\n";
                th += "2.58\r\n";
                th += " 30\r\n";
                th += "0.0\r\n";
                th += " 40\r\n";
                th += "1.0\r\n";
                th += " 1\r\n";
                th += String.Format("Thickness in mm: {0}", thickness);
                swDXF.WriteLine(th);

                #endregion пишем в dxf толщину

                #region обработка фрагмента (фрагменты значки в имени которых содержится _чпу.grb)

                // закоментировал для спец версии

                foreach (var fr in doc.GetFragments())
                {
                    var fl = fr.FileLink;
                    //var buf = fr.FullFilePath.ToLowerInvariant();
                    var buf = fl.FullFilePath.ToLowerInvariant();
                    if (fr.Visible == false || !buf.Contains("_чпу.grb"))
                        continue;

                    if (fr.Page.Name != pageExp.Name)
                        continue;

                    var doc_fr = fr.GetFragmentDocument(true);

                    var am = fr.Transformation;

                    foreach (var ln in doc_fr.GetOutlines())
                    {
                        if (ln.Visible == false || ln.PatternName != "CONTINUOUS")
                            continue;

                        //         if (ln.Page.Name != pageExp.Name)
                        //           continue;

                        #region Line

                        if (ln.Geometry is LineGeometry)
                        {
                            var lg = (TFlex.Model.Model2D.LineGeometry)ln.Geometry;

                            var x1 = lg.X1;
                            var y1 = lg.Y1;
                            var x2 = lg.X2;
                            var y2 = lg.Y2;
                            am.ToWCS(ref x1, ref y1);
                            am.ToWCS(ref x2, ref y2);
                            x1 /= pageExp.Scale.Value;
                            y1 /= pageExp.Scale.Value;
                            x2 /= pageExp.Scale.Value;
                            y2 /= pageExp.Scale.Value;

                            //           if (htElements.ContainsKey(lg.X1.ToString() + lg.X2.ToString() + lg.Y1.ToString() + lg.Y2.ToString()))
                            if (htElements.ContainsKey(x1.ToString() + x2.ToString() + y1.ToString() + y2.ToString()))
                            {
                                continue;
                            }
                            swDXF.WriteLine($" 0\r\nLINE\r\n 8\r\n{_mlayer}\r\n 62\r\n3\r\n 6\r\nCONTINUOUS");
                            swDXF.WriteLine(
                                " 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 11\r\n{2}\r\n 21\r\n{3}\r\n 31\r\n0.000",
                                //            lg.X1.ToString("G", nfi), lg.Y1.ToString("G", nfi), lg.X2.ToString("G", nfi), lg.Y2.ToString("G", nfi));
                                x1.ToString("G", nfi), y1.ToString("G", nfi), x2.ToString("G", nfi),
                                y2.ToString("G", nfi));
                            //           htElements.Add(lg.X1.ToString() + lg.X2.ToString() + lg.Y1.ToString() + lg.Y2.ToString(), lg);
                        }

                        #endregion

                    }
                }

                #endregion обработка фрагмента (фрагменты значки в имени которых содержится _чпу.grb)

                foreach (var ln in doc.GetOutlines())
                {
                    if (ln.Visible == false || ln.PatternName != "CONTINUOUS")
                        continue;
                    if (ln.Page.Name != pageExp.Name)
                        continue;
                    //if(ln.Layer.Name != "Основной")
                    //    continue;
                    //if (ln.Layer.Name != "Основной")
                    //    continue;
                    if (ln.Layer.Screen) continue;

                    #region Line

                    if (ln.Geometry is LineGeometry)
                    // if (ln.GeometryType == ObjectGeometryType.Line)
                    {
                        var lg = (TFlex.Model.Model2D.LineGeometry)ln.ModelGeometry;
                        if (
                            htElements.ContainsKey(lg.X1.ToString() + lg.X2.ToString() + lg.Y1.ToString() +
                                                   lg.Y2.ToString()))
                        {
                            continue;
                        }
                        swDXF.WriteLine(" 0\r\nLINE\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                        swDXF.WriteLine(
                            " 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 11\r\n{2}\r\n 21\r\n{3}\r\n 31\r\n0.000",
                            lg.X1.ToString("G", nfi), lg.Y1.ToString("G", nfi), lg.X2.ToString("G", nfi),
                            lg.Y2.ToString("G", nfi));
                        htElements.Add(lg.X1.ToString() + lg.X2.ToString() + lg.Y1.ToString() + lg.Y2.ToString(), lg);
                    }

                    #endregion Line

                    #region Circle

                    if (ln.Geometry is CircleGeometry)
                    {
                        var cg = (CircleGeometry)ln.ModelGeometry;
                        if (htElements.ContainsKey(cg.CenterX.ToString() + cg.CenterY.ToString() + cg.Radius.ToString()))
                        {
                            continue;
                        }
                        else
                        {
                            swDXF.WriteLine(" 0\r\nCIRCLE\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                            swDXF.WriteLine(" 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 40\r\n{2}",
                                cg.CenterX.ToString("F", nfi), cg.CenterY.ToString("F", nfi),
                                cg.Radius.ToString("F", nfi));
                            htElements.Add(cg.CenterX.ToString() + cg.CenterY.ToString() + cg.Radius.ToString(), cg);
                            isCircle = true;
                        }
                    }

                    #endregion Circle

                    #region arc

                    //if (ln.Geometry is CircleArcGeometry)
                    //{
                    //    var ag = (CircleArcGeometry)ln.ModelGeometry;
                    //    if (htElements.ContainsKey(ag.CenterX.ToString() + ag.CenterY.ToString() + ag.EndX.ToString() + ag.EndY.ToString() + ag.Radius.ToString()))
                    //    { continue; }
                    //    else
                    //    {
                    //        double a1 = DefAngle(ag.CenterX, ag.CenterY, ag.StartX, ag.StartY);
                    //        double a2 = DefAngle(ag.CenterX, ag.CenterY, ag.EndX, ag.EndY);
                    //        if (a2 < a1) a2 += 360;
                    //        //sw_DXF.WriteLine(" 0\r\nARC\r\n 8\r\n0\r\n 6\r\nCONTINUOUS\r\n 62\r\n{0}", LN.Color.IntValue);
                    //        swDXF.WriteLine(" 0\r\nARC\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                    //        swDXF.WriteLine(" 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 40\r\n{2}\r\n 50\r\n{3}\r\n 51\r\n{4}",
                    //        ag.CenterX.ToString("G", nfi), ag.CenterY.ToString("G", nfi), ag.Radius.ToString("G", nfi), a1.ToString("G", nfi), a2.ToString("G", nfi));
                    //        htElements.Add(ag.CenterX.ToString() + ag.CenterY.ToString() + ag.EndX.ToString() + ag.EndY.ToString() + ag.Radius.ToString(), ag);
                    //    }
                    //}
                    if (ln.Geometry is CircleArcGeometry)
                    {
                        var ag = (CircleArcGeometry)ln.ModelGeometry;
                        var a1 = DefAngle(ag.CenterX, ag.CenterY, ag.StartX, ag.StartY);
                        var a2 = DefAngle(ag.CenterX, ag.CenterY, ag.EndX, ag.EndY);
                        if (Math.Abs(a1 - a2) < 0.00001) continue;

                        if (

                            htElements.ContainsKey(ag.CenterX.ToString() + ag.CenterY.ToString() + ag.EndX.ToString() +
                                                   ag.EndY.ToString() + ag.Radius.ToString()))
                        {
                            continue;
                        }
                        else
                        {
                            if (a2 < a1) a2 += 360;

                            double maxR = 10000;
                            double maxDLT = 0.5;

                            if (ag.Radius > maxR)
                            {
                                double ang = Math.Acos((ag.Radius - maxDLT) / ag.Radius) * 2;

                                double x1 = ag.StartX;
                                double y1 = ag.StartY;

                                double x2 = 0;
                                double y2 = 0;

                                a1 = a1 / 180 * Math.PI;
                                a2 = a2 / 180 * Math.PI;
                                double dA = (a2 - a1);
                                int NN = (int)((dA) / ang) + 1;
                                ang = dA / NN;

                                double cur_ang = a1 + ang;
                                int n = 1;
                                while (n < NN)
                                {
                                    x2 = ag.CenterX + Math.Cos(cur_ang) * ag.Radius;
                                    y2 = ag.CenterY + Math.Sin(cur_ang) * ag.Radius;

                                    swDXF.WriteLine(" 0\r\nLINE\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                                    swDXF.WriteLine(
                                        " 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 11\r\n{2}\r\n 21\r\n{3}\r\n 31\r\n0.000",
                                        x1.ToString("G", nfi), y1.ToString("G", nfi), x2.ToString("G", nfi),
                                        y2.ToString("G", nfi));

                                    cur_ang += ang;
                                    x1 = x2;
                                    y1 = y2;
                                    n++;
                                }
                                swDXF.WriteLine(" 0\r\nLINE\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                                swDXF.WriteLine(
                                    " 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 11\r\n{2}\r\n 21\r\n{3}\r\n 31\r\n0.000",
                                    x1.ToString("G", nfi), y1.ToString("G", nfi), ag.EndX.ToString("G", nfi),
                                    ag.EndY.ToString("G", nfi));
                            }
                            else
                            {
                                //sw_DXF.WriteLine(" 0\r\nARC\r\n 8\r\n0\r\n 6\r\nCONTINUOUS\r\n 62\r\n{0}", LN.Color.IntValue);
                                swDXF.WriteLine(" 0\r\nARC\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                                swDXF.WriteLine(
                                    " 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 40\r\n{2}\r\n 50\r\n{3}\r\n 51\r\n{4}",
                                    ag.CenterX.ToString("G", nfi), ag.CenterY.ToString("G", nfi),
                                    ag.Radius.ToString("G", nfi), a1.ToString("G", nfi), a2.ToString("G", nfi));
                            }
                            htElements.Add(
                                ag.CenterX.ToString() + ag.CenterY.ToString() + ag.EndX.ToString() + ag.EndY.ToString() +
                                ag.Radius.ToString(), ag);
                        }
                    }

                    #endregion arc

                    #region ellipse & ellipsArc

                    if (ln.Geometry is EllipseGeometry ||
                        ln.Geometry is EllipseArcGeometry)
                    {
                        var pln = (TFlex.Drawing.Polyline)ln.GeometryAsPolyline;

                        var x1 = pln.get_X(0) / activeScale;
                        var y1 = pln.get_Y(0) / activeScale;

                        var N = pln.get_PointCount(0);
                        for (var n = 1; n < N; n++)
                        {
                            var x2 = pln.get_X(n) / activeScale;
                            var y2 = pln.get_Y(n) / activeScale;

                            swDXF.WriteLine(" 0\r\nLINE\r\n 8\r\n0\r\n 6\r\nCONTINUOUS");
                            swDXF.WriteLine(
                                " 10\r\n{0}\r\n 20\r\n{1}\r\n 30\r\n0.000\r\n 11\r\n{2}\r\n 21\r\n{3}\r\n 31\r\n0.000",
                                x1.ToString("G", nfi), y1.ToString("G", nfi), x2.ToString("G", nfi),
                                y2.ToString("G", nfi));
                            htElements.Add(x1.ToString() + x2.ToString() + y1.ToString() + y2.ToString(), ln);

                            x1 = x2;
                            y1 = y2;
                        }
                    }

                    #endregion ellipse & ellipsArc

                    #endregion формируем структуру DXF файла

                    #region poliline

                    // ObjectGeometryType.Undefined
                    if (ln.Geometry is PolylineGeometry)
                    {
                        var pl = ln.Geometry as PolylineGeometry;
                        if (pl != null)
                        {
                        }
                    }

                    #endregion poliline

                    //#region ellipse
                    //if (ln.GeometryType == ObjectGeometryType.Ellipse)
                    //{
                    //}
                    //#endregion
                    //#region ellipsearc
                    //if (ln.GeometryType == ObjectGeometryType.EllipseArc)
                    //{
                    //}
                    //#endregion
                }

                swDXF.WriteLine(" 0\r\nENDSEC\r\n 0\r\nEOF");
                swDXF.Close();

                #region получаем имя файла транслит

                String fileNameTr = null;
                var varTypeDoc = doc.FindVariable("type_doc");
                if (varTypeDoc == null)
                {
                    return false;
                }
                var varMarka = doc.FindVariable("$marka");

                var varPoz = doc.FindVariable("$poz");

                if (varPoz != null && varTypeDoc.RealValue != 2)
                {
                    fileNameTr = "p" + varPoz.TextValue;
                    if (isCircle)
                    {
                        fileNameTr += "+";
                    }
                }
                else
                {
                    if (varMarka != null && varMarka.TextValue != "")
                    {
                        fileNameTr = CTransliteration.Front(varMarka.TextValue);
                        if (isCircle)
                        {
                            fileNameTr += "+";
                        }
                    }
                }
                if (fileNameTr != null)
                {
                    fileNameTr = fileNameTr.Replace('/', '_');
                }

                #region затычка временная

                if (isTypeBed)
                {
                    var fi1 = new FileInfo(doc.FileName);
                    fileNameTr = fi1.Exists ? fi1.Name : "ошибка обратитесть к ВАХРУШЕВУ";
                }

                #endregion затычка временная

                #endregion получаем имя файла транслит

                #region усадка

                //var usL = doc.FindVariable("us_l");
                //var usB = doc.FindVariable("us_b");
                //var us = false;
                //if (usL != null && usB != null)
                //{
                //    if (usL.RealValue > 0)
                //    {
                //        us = true;
                //    }
                //    else
                //    {
                //        if (usB.RealValue > 0)
                //        {
                //            us = true;
                //        }
                //    }
                //}

                #endregion усадка

                #region переменные для структуры директорий

                //const string subDirNoCircle = "без отверстий";
                //const string subDirCircle = "с отверстием";
                //const string sunDitNoUs = "без усадки";
                //const string sunDitUs = "c усадкой";
                const string subDxf = "dxf";

                #endregion переменные для структуры директорий

                if (!isPlugin)
                {
                    #region создаем рабочие директории

                    if (!Directory.Exists(_mDirectorySave + @"\" + subDxf))
                    {
                        Directory.CreateDirectory(_mDirectorySave + @"\" + subDxf);
                    }
                    //if (!Directory.Exists(_mDirectorySave + @"\" + subDirCircle))
                    //{
                    //    Directory.CreateDirectory(_mDirectorySave + @"\" + subDirCircle);
                    //}
                    //if (!Directory.Exists(_mDirectorySave + @"\" + subDirNoCircle))
                    //{
                    //    Directory.CreateDirectory(_mDirectorySave + @"\" + subDirNoCircle);
                    //}
                    //if (!Directory.Exists(_mDirectorySave + @"\" + subDirCircle + @"\" + sunDitNoUs))
                    //{
                    //    Directory.CreateDirectory(_mDirectorySave + @"\" + subDirCircle + @"\" + sunDitNoUs);
                    //}
                    //if (!Directory.Exists(_mDirectorySave + @"\" + subDirCircle + @"\" + sunDitUs))
                    //{
                    //    Directory.CreateDirectory(_mDirectorySave + @"\" + subDirCircle + @"\" + sunDitUs);
                    //}
                    //if (!Directory.Exists(_mDirectorySave + @"\" + subDirNoCircle + @"\" + sunDitNoUs))
                    //{
                    //    Directory.CreateDirectory(_mDirectorySave + @"\" + subDirNoCircle + @"\" + sunDitNoUs);
                    //}
                    //if (!Directory.Exists(_mDirectorySave + @"\" + subDirNoCircle + @"\" + sunDitUs))
                    //{
                    //    Directory.CreateDirectory(_mDirectorySave + @"\" + subDirNoCircle + @"\" + sunDitUs);
                    //}

                    #endregion создаем рабочие директории

                    fileExport = _mDirectorySave + @"\" + subDxf + @"\" + _mInvNum +
                                 fileNameTr;
                    //if (isCircle)
                    //{
                    //    if (us)
                    //    {
                    //        fileExport = _mDirectorySave + @"\" + subDirCircle + @"\" + sunDitUs + @"\" + _mInvNum +
                    //                     fileNameTr;
                    //    }
                    //    else
                    //    {
                    //        fileExport = _mDirectorySave + @"\" + subDirCircle + @"\" + sunDitNoUs + @"\" + _mInvNum +
                    //                     fileNameTr;
                    //    }
                    //}
                    //else
                    //{
                    //    if (us)
                    //    {
                    //        fileExport = _mDirectorySave + @"\" + subDirNoCircle + @"\" + sunDitUs + @"\" + _mInvNum +
                    //                     fileNameTr;
                    //    }
                    //    else
                    //    {
                    //        fileExport = _mDirectorySave + @"\" + subDirNoCircle + @"\" + sunDitNoUs + @"\" + _mInvNum +
                    //                     fileNameTr;
                    //    }
                    //}
                }
                else
                {
                    fileExport = _mDirectorySave + @"\" + _mInvNum + fileNameTr;
                }
                var fi = new FileInfo(_mDirectorySave + @"\temp.dxf");
                if (fi.Exists)
                {
                    if (File.Exists(fileExport + ".dxf"))
                        return false;
                    fi.MoveTo(fileExport + ".dxf");
                }
                return true;
            }
            return false;
        }
    }
}