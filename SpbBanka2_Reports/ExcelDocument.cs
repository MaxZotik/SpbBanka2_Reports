using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;

namespace SpbBanka2_Reports
{
    class ExcelDocument
    {
        public static void MakeExcelReport(
            List<double[]> limitsList,  // Уставки

            // Вибропараметры (МИН-МАКС-СРЕДН)
            List<double[]> VA_Values,   // виброускорение
            List<double[]> VV_Values,   // виброскорость

            // Для построения графиков (Тренды)
            List<DataTable> VA_DataTables,  // виброускорение
            List<DataTable> VV_DataTables)  // виброскорость
        {
            try
            {
                //Thread.Sleep(5000);                
                EventLog.Log("Начало работы с Excel-документом;");

                Excel.Application excelApp = new Excel.Application();

                excelApp.DisplayAlerts = false;

                Excel._Workbook excelWorkbook;
                excelWorkbook = excelApp.Workbooks.Open(Path.modelPath);

                // заполнение листов данными
                InfoSheet       (ref excelWorkbook);                                // ИНФО
                LimitsSheet     (ref excelWorkbook, limitsList);                    // УСТАВКИ
                ParametersSheet (ref excelWorkbook, VA_Values, VV_Values);          // ПАРАМЕТРЫ

                    // удаление и создание папки с изобрадениями графиков, чтобы не было ошибок
                    Path.DelCat(Path.VA_GraphDelCreatePath);
                    Path.DelCat(Path.VV_GraphDelCreatePath);
                    Path.CreateCat(Path.VA_GraphDelCreatePath);
                    Path.CreateCat(Path.VV_GraphDelCreatePath);
                    
                    EventLog.Log("Папки VA и VV пересозданы успешно;");


                TrendsSheet(ref excelWorkbook, VA_DataTables, VV_DataTables);  // ТРЕНДЫ

                // удаление листа, который использовался для построения графиков
                Excel.Worksheet tempSheet = excelWorkbook.Worksheets.get_Item(5);
                tempSheet.Delete();


                // сохранение файла
                Save(ref excelWorkbook);

                // удаление Excel-процесса 
                excelWorkbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                EventLog.Log("Excel-файл успешно сформирован.", false);
            }
            catch (Exception ex)
            {
                Program.ChangeCheck("2");
                EventLog.Log(ex.ToString(), false);

                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                if (proc1.Length > 0) proc1[0].Kill();

                //MessageBox.Show("Ошибка при формировании Excel-документа! \n\n" + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                /*MessageBox.Show(
                            "Ошибка при формировании Excel-документа! \n\n" + ex.ToString(), "Ошибка", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);*/
            }
            finally
            {
                // удаление Excel-процесса
                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                if (proc1.Length > 0) proc1[0].Kill();
            }
        }
        
        // работа с листом "ИНФО"
        private static void InfoSheet(ref Excel._Workbook excelWorkbook)
        { 
            try
            {
                Excel.Worksheet sheet_Info = excelWorkbook.Worksheets.get_Item(1);

                sheet_Info.Cells[2, "C"] = Config.configInfo[3];    // ячейка C2 = дата и время формирования отчета
                sheet_Info.Cells[5, "E"] = Config.configInfo[0];    // ячейка E5 = дата и время начала периода отчета
                sheet_Info.Cells[5, "H"] = Config.configInfo[1];    // ячейка H5 = дата и время конца периода отчета

                EventLog.Log("Лист ИНФО \tзаполнен успешно;");
            }
            catch (Exception ex)
            {
                Program.ChangeCheck("2");
                EventLog.Log(ex.ToString());

                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                proc1[0].Kill();

                /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(),
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
            }

        }

        // работа с листом "УСТАВКИ"
        private static void LimitsSheet(ref Excel._Workbook excelWorkbook, List<double[]> limitsList)
        {
            try
            {
                Excel.Worksheet sheet_Info = excelWorkbook.Worksheets.get_Item(2);

                double[] temp = new double[8];  // 4 состояния точки ("Авария", "Опасность", "Предупреждение", "Норма") на каждой из 2 полос (4 * 2 = 8)
                string[] letters = { "D", "E", "F", "G", "H", "I", "J", "K", "L", "M" };    // буквы стобцов, в которые записываются данные
                int
                    pointIndex,         // для использования элементов массива pointsArray
                    row = 4,            // в отчете номер строки, куда записываются данные, начинается с 4-й
                    columnIndex = 0;    // для использования элементов массива 

                for (int i = 0; i < Config.pointsArray.Length; i++)
                {
                    row = 4;

                    // для перехода к следующему столбцу
                    pointIndex = Convert.ToInt32(Config.pointsArray[i]);
                    columnIndex = pointIndex;

                    temp = limitsList.ElementAt(i); // массив 4-х значений состояний точки на каждой из полос

                    for (int j = 0; j < temp.Length; j++)
                    {
                        sheet_Info.Cells[row, letters[columnIndex]] = Math.Round(temp[j], 2);
                        row++;

                        // для перехода к таблице следующей полосы
                        if (row == 8)
                        {
                            row += 3;
                        }
                    }
                }    
                            
                EventLog.Log("Лист УСТАВКИ \tзаполнен успешно;");
            }
            catch (Exception ex)
            {
                Program.ChangeCheck("2");
                EventLog.Log(ex.ToString());

                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                proc1[0].Kill();

                /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(),
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
            }            
        }

        // работа с листом "ПАРАМЕТРЫ"
        private static void ParametersSheet(ref Excel._Workbook excelWorkbook, List<double[]> VA_Values, List<double[]> VV_Values)
        {
            try
            {
                Excel.Worksheet sheet_Info = excelWorkbook.Worksheets.get_Item(3);

                // МИН / МАКС / СРЕДН на каждой из полос 
                double[] tempVA = new double[1];  // виброускорения
                double[] tempVV = new double[2];  // виброскорости

                string[] letters = { "C", "D", "E" };    // буквы стобцов, в которые записываются данные
                int
                    tempIndex = 0,      // для использования элементов коллекций VA_Values и VV_Values
                    pointIndex,         // для использования элементов массива pointsArray
                    rowVA = 5,          // в отчете номер строки, куда записываются данные по виброускорению, начинается с 5-й
                    rowVV = 25,         // в отчете номер строки, куда записываются данные по виброскорости, начинается с 19-й
                    columnIndex = 0;    // для использования элементов массива letters

                for (int _point = 0; _point < Config.pointsArray.Length; _point++)
                {
                    // для перехода к следующему столбцу
                    pointIndex = Convert.ToInt32(Config.pointsArray[_point]);
                    rowVA = 5 + pointIndex;
                    rowVV = 25 + pointIndex;                   


                    for (int _MinMaxAvg = 0; _MinMaxAvg < 3; _MinMaxAvg++)
                    {
                        // массив МИН-МАКС-СРЕДН-значений для 1 точки
                        tempVA = VA_Values.ElementAt(tempIndex); // по виброускорению
                        tempVV = VV_Values.ElementAt(tempIndex); // по виброскорости
                        tempIndex++;

                        columnIndex = _MinMaxAvg;

                        if(tempVA[0] == -1 || tempVV[0] == -1 ) // нет данных по точке
                        {
                            sheet_Info.Cells[rowVA, letters[columnIndex]] = "—";
                            sheet_Info.Cells[rowVV, letters[columnIndex]] = "—";
                        }
                        else
                        {
                            sheet_Info.Cells[rowVA, letters[columnIndex]] = Math.Round(tempVA[0], 2);
                            sheet_Info.Cells[rowVV, letters[columnIndex]] = Math.Round(tempVV[0], 2);
                        }
                        
                    }
                }

                EventLog.Log("Лист ПАРАМЕТРЫ \tзаполнен успешно;");
            }
            catch (Exception ex)
            {
                Program.ChangeCheck("2");
                EventLog.Log(ex.ToString());

                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                proc1[0].Kill();

                /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(),
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
            }            
        }

        // работа с листом "ТРЕНДЫ"
        private static void TrendsSheet(ref Excel._Workbook excelWorkbook, List<DataTable> VA_DataTables, List<DataTable> VV_DataTables)
        {
            try
            {
                Excel.Worksheet tempSheet = excelWorkbook.Worksheets.get_Item(5);

                // проход по всем полосам точки
                for (int point = 0; point < Config.pointsArray.Length; point++)
                {
                    //if(Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] == Program.nullPoints.ElementAt(point)) !!!
                    if(Program.nullPoints.Contains(Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])]))
                    {
                        // графики не делаются
                    }
                    else
                    {
                        // виброускорение
                        PreparingForChart(0, Convert.ToInt32(Config.pointsArray[point]), VA_DataTables[point], tempSheet, Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + " Виброускорение 10 - 5000 Гц", Path.VA_GraphSavePath);

                        // виброскорость
                        PreparingForChart(1, Convert.ToInt32(Config.pointsArray[point]), VV_DataTables[point], tempSheet, Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + " Виброскорость 10 - 1000 Гц", Path.VV_GraphSavePath);
                    }
                    
                }
                EventLog.Log("Изображения графиков созданы и помещены в папки VA и VV;");

                Excel.Worksheet trendsSheet = excelWorkbook.Worksheets.get_Item(4);
                //100, 50, 680, 480
                float
                    left = 40,
                    top = 0,
                    width = 600,
                    height = 480;

                // вставка изображений графиков
                for (int point = 0; point < Config.pointsArray.Length; point++)
                {
                    try
                    {
                        top = 0;

                        string
                            VApath = Path.VA_GraphSavePath + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + "_",
                            VVpath = Path.VV_GraphSavePath + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + "_";

                        // виброускорение
                        trendsSheet.Shapes.AddPicture(VApath + Points.stripsNames[0] + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height); top += height + 27;

                        // виброскорость
                        trendsSheet.Shapes.AddPicture(VVpath + Points.stripsNames[1] + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height); top += height + 27;
                        
                        left += width + 85;
                        EventLog.Log("Изображения графиков добавленый для точки " + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + ";");
                    }
                    catch
                    {
                        // нет данных...
                    }
                    
                }
                EventLog.Log("Изображения графиков помещены в отчет;");


                EventLog.Log("Лист ТРЕНДЫ \tзаполнен успешно;");
            }            
            catch (Exception ex)
            {
                Program.ChangeCheck("2");
                EventLog.Log(ex.ToString());

                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                proc1[0].Kill();

                //MessageBox.Show("Ошибка! \n" + ex.ToString());

                /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(), 
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
            }

            // создание массива данных DateTime и СреднЗначение
            void PreparingForChart(
                int stripIndex,         // индекс полосы
                int pointIndex,         // индекс точки
                DataTable DataTables,   // данные из БД (Средн. значения ИЛИ просто значения, если был выбран секундный архив)
                Excel.Worksheet sheet,  // временный лист, куда будут заноситься значения, по которым и будут строиться графики
                string seriesName,      // название графика
                string savePath)        // путь, куда будет сохранена картинка графика
            {
                try
                {
                    if (DataTables.Rows.Count <= 0) return;

                    string table = Program.table;   //TablePeriod.GetTable(Config.startDT, Config.endDT);   // таблица, из котороый производился забор данных

                    if (DataTables.Rows.Count > 0)
                    {
                        // двумерный массив для данных DateTime и Val / Avg 
                        object[,] valueMatrix = new object[DataTables.Rows.Count, 2];

                        for (int j = 0; j < DataTables.Rows.Count; j++)
                        {
                            valueMatrix[j, 0] = DataTables.Rows[j].ItemArray[0]; // колонка DateTime

                            if (table == "CnlData")
                                valueMatrix[j, 1] = DataTables.Rows[j].ItemArray[2]; // колонка Val
                            else
                                valueMatrix[j, 1] = DataTables.Rows[j].ItemArray[4]; // колонка Avg
                        }

                        int rowCount = DataTables.Rows.Count;
                        int columnCount = 2;

                        Excel.Range range = (Excel.Range)sheet.Cells[1, 1];
                        range = range.get_Resize(rowCount, columnCount);
                        range.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, valueMatrix);

                        GetChart(
                            stripIndex,
                            pointIndex,
                            sheet,
                            "A1", "A" + DataTables.Rows.Count.ToString(),
                            "B1", "B" + DataTables.Rows.Count.ToString(),
                            seriesName, savePath, DataTables.Rows.Count);
                    }

                }
                catch (Exception ex)
                {
                    Program.ChangeCheck("2");
                    EventLog.Log(ex.ToString());

                    Process[] proc1 = Process.GetProcessesByName("EXCEL");
                    proc1[0].Kill();

                    /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(),
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
                }
            }

            // строит график и сохраняет картинку этого графика
            void GetChart(
                int stripIndex,         // индекс полосы
                int pointIndex,         // индекс точки
                Excel.Worksheet sheet,  // временный лист, куда будут заноситься значения, по которым и будут строиться графики
                string firstCellX, string lastCellX,    // координаты ячеек, по которым будет строиться график (ось X)
                string firstCellY, string lastCellY,    // координаты ячеек, по которым будет строиться график (ось Y)
                string seriesName,                      // название графика
                string savePath,                        // путь, куда будет сохранена картинка графика
                int valuesCount)
            {
                double
                    left = 100,
                    top = 50,
                    width = 600,
                    height = 480;

                try
                {
                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = xlCharts.Add(left, top, width, height);
                    Excel.Chart chart = myChart.Chart;
                    Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);

                    chart.ChartType = Excel.XlChartType.xlLine;
                    chart.HasLegend = false;

                    Excel.Axis axis = chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    axis.CategoryType = Excel.XlCategoryType.xlCategoryScale;

                    var xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    xAxis.TickLabelSpacingIsAuto = true;
                    xAxis.TickLabels.Font.Size = 11;
                    xAxis.TickLabels.Orientation = Excel.XlTickLabelOrientation.xlTickLabelOrientationUpward;

                    
                    const int X_AXIS_LABEL_COUNT = 12, MINIMAL_X_AXIS_LABEL_COUNT = 4;
                    if (valuesCount / X_AXIS_LABEL_COUNT >= MINIMAL_X_AXIS_LABEL_COUNT) // при 4 минимальное количество точек для графика - 48
                    {
                        xAxis.TickLabelSpacing = valuesCount / X_AXIS_LABEL_COUNT; // 12 значений времени, отображаемых на графике
                        xAxis.TickMarkSpacing = valuesCount / X_AXIS_LABEL_COUNT;
                    }

                    var yAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    yAxis.HasTitle = false;
                    yAxis.TickLabels.Font.Size = 11;

                    Excel.Series series = seriesCollection.NewSeries();
                    series.XValues = sheet.get_Range(firstCellX, lastCellX);
                    series.Values = sheet.get_Range(firstCellY, lastCellY);
                    series.Name = seriesName;
                    series.Format.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle;
                    series.Format.Line.ForeColor.RGB = Color.FromArgb(255, 0, 0).ToArgb();
                    series.Format.Line.Weight = 1.25f;

                    chart.ChartTitle.Font.Size = 14;

                    savePath += Points.pointsNames[pointIndex] + "_" + Points.stripsNames[stripIndex] + ".png";
                    chart.Export(savePath, "PNG", false);

                    // EventLog.Log("График для " + Points.pointsNames[pointIndex] + " на полосе " + Points.stripsNames[stripIndex] + " успешно сохранен;");
                }
                catch (Exception ex)
                {
                    Program.ChangeCheck("2");
                    EventLog.Log(ex.ToString());

                    Process[] proc1 = Process.GetProcessesByName("EXCEL");
                    proc1[0].Kill();

                    /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(),
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
                }
            }
        
        
        }

        // сохранение отчета
        private static void Save(ref Excel._Workbook _excelWorkbook)
        {
            string save = "";
            try
            {
                if (Directory.Exists(Path.saveRepPath) == false)
                {
                    Directory.CreateDirectory(Path.saveRepPath);
                }

                DateTime dt = Config.reportGeneration;
                string savePart =
                    (dt.Day >= 10 ? dt.Day.ToString() : "0" + dt.Day.ToString()) + "." +
                    (dt.Month >= 10 ? dt.Month.ToString() : "0" + dt.Month.ToString()) + "." +
                    dt.Year.ToString() + "_" + 
                    (dt.Hour >= 10 ? dt.Hour.ToString() : "0" + dt.Hour.ToString()) + "." +
                    (dt.Minute >= 10 ? dt.Minute.ToString() : "0" + dt.Minute.ToString()) + "." +
                    (dt.Second >= 10 ? dt.Second.ToString() : "0" + dt.Second.ToString()) + ".xlsx";
                //save += Path.saveRepPath + "Report " + Config.configInfo[3].Replace(':', '.').Replace(' ', '_') + ".xlsx";
                save += Path.saveRepPath + "Report " + savePart;

                _excelWorkbook.SaveAs(save);

                EventLog.Log("Отчет сохранен;");
            }
            catch (Exception ex)
            {
                Program.ChangeCheck("2");
                EventLog.Log(ex.ToString());

                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                proc1[0].Kill();

                /*MessageBox.Show(
                           "Ошибка! \n" + ex.ToString(),
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);*/
            }
        }
    }
}
