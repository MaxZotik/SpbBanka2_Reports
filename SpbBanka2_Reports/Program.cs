using Scada.Data.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace SpbBanka2_Reports
{
    class Program
    {
        // коллекция точек, в которых нет данных
        public static List<string> nullPoints = new List<string>();
        public static string table = "";

        static void Main(string[] args)
        {
            bool test = false;  // если true - программа в режиме отладки

            try
            {               
                EventLog.Log("Начало работы программы;", true);


                // начало анимации в SCADA
                ChangeCheck("1");

                #region// УСТАВКИ
                // для поиска нужного .dat-файла
                DateTime dt = DateTime.Now;

                string currDayMinuteFileName = "";
                if (test)
                {
                    currDayMinuteFileName = "m220419.dat";   
                    //currDayMinuteFileName = "m240918.dat";
                }
                else
                {
                    currDayMinuteFileName = "m" + dt.Year.ToString().Substring(2, 2) + (dt.Month < 10 ? "0" + dt.Month.ToString() : dt.Month.ToString()) + (dt.Day < 10 ? "0" + dt.Day.ToString() : dt.Day.ToString()) + ".dat";
                }
                

                // SCADA-библиотека для получения данных из .dat-файла
                SrezTableLight currDaySnapshotTable = new SrezTableLight();
                SrezAdapter newSrezAdapter = new SrezAdapter();

                if (test)
                {
                    newSrezAdapter.FileName = @"C:\SCADA\ArchiveDAT\" + currDayMinuteFileName;  // TEST
                }
                else
                {
                    newSrezAdapter.FileName = @"C:\SCADA\ArchiveDAT\Min\" + currDayMinuteFileName;
                }

                newSrezAdapter.Fill(currDaySnapshotTable);


                const int LIMITS_COUNT = 8;   // количество уставок для 1 точки (4 уставки * 2 полосs)

                List<double[]> limitsList = new List<double[]>();   // коллекция уставок для точек

                // добавление уставок для выбранных точек
                for (int i = 0; i < Config.pointsArray.Length; i++)
                {
                    double[] limitsArray = new double[LIMITS_COUNT];    // временный массив для записи уставок
                    SetLimits(i, limitsArray, currDaySnapshotTable);

                    limitsList.Add(limitsArray);
                }
                EventLog.Log("Захват данных по уставкам: успешно;");
                #endregion

                #region// ВИБРОПАРАМЕТРЫ
                table = TablePeriod.GetTable(Config.startDT, Config.endDT);   // подбираем нужную таблицу для забора данных

                if(table == "Error")
                {
                    /*MessageBox.Show(
                        "Ошибка выбора архива!",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);*/

                    EventLog.Log("Ошибка выбора архива.");

                    Environment.Exit(0);
                }
                else if(table == "ErrorNoData")
                {
                    /*MessageBox.Show(
                        "Недостаточно данных!",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);*/

                    EventLog.Log("Недостаточно данных.");

                    ChangeCheck("2");
                    Environment.Exit(0);
                }
                EventLog.Log("Выбранная таблица в БД - " + table);


                // в коллекциях будут храниться номера нужных каналов 
                List<int> VV_Channels = new List<int>();
                List<int> VA_Channels = new List<int>();
                for (int i = 0; i < Config.pointsArray.Length; i++)
                {
                    // добавление нужных каналов виброускорения
                    VA_Channels.Add(Points.Parameters_VA_Strip_10_5000  [Convert.ToInt32(Config.pointsArray[i])]);

                    // добавление нужных каналов виброскорости
                    VV_Channels.Add(Points.Parameters_VV_Strip_10_1000  [Convert.ToInt32(Config.pointsArray[i])]);         
                }



                // для хранения таблиц по которым в дальнйшем будут строиться графики
                List<DataTable> VA_DataTables = new List<DataTable>();
                List<DataTable> VV_DataTables = new List<DataTable>();

                // для хранения мин, макс и средн. показателей точки
                List<double[]> VA_Values = new List<double[]>();
                List<double[]> VV_Values = new List<double[]>();

                SqlConnection connection = new SqlConnection(Path.connectionString);

                bool dataIsOK = false;   // полнота данных в БД
                for (int point = 0, VAIndex = 0, VVIndex = 0; point < Config.pointsArray.Length; point++, VAIndex++ , VVIndex++)
                {

                    // для забора таблиц, нужных для построения графиков
                    // виброускорение
                    string querryTable_10_5000_VA = Querry("*", table, VA_Channels[VAIndex], Config.startDT, Config.endDT);
                    // виброскорость
                    string querryTable_VV = Querry("*", table, VV_Channels[VVIndex], Config.startDT, Config.endDT);
                    

                    // виброускорение
                    ListTableFill(querryTable_10_5000_VA,  connection, ref VA_DataTables);
                    // виброскорость
                    ListTableFill(querryTable_VV,  connection, ref VV_DataTables);


                    #region// MIN, MAX и AVG
                    // для вычисления MIN, MAX и AVG
                    for (int minMaxAvg = 0; minMaxAvg < 3; minMaxAvg++)
                    {                        
                        string param = "";  // для вставки в запрос
                        if (table=="CnlData")
                        {
                            if      (minMaxAvg == 0) param = "MIN(Val)";
                            else if (minMaxAvg == 1) param = "MAX(Val)";
                            else                     param = "AVG(Val)";
                        }
                        else
                        {
                            if      (minMaxAvg == 0) param = "MIN(Min)";
                            else if (minMaxAvg == 1) param = "MAX(Max)";
                            else                     param = "AVG(Avg)";
                        }

                        string
                            // виброускорение
                            querry_10_5000_VA  = Querry(param, table, VA_Channels[VAIndex],  Config.startDT, Config.endDT),
                            // виброскорость
                            querry_VV  = Querry(param, table, VV_Channels[VVIndex],      Config.startDT, Config.endDT),

                            // для отправки в лог запроса при нехватки значений
                            tempQ = querry_10_5000_VA;

                        // в массив поочередно будут записыватсья результаты MIN-MAX - AVG запросов для точки
                        double[] tempVA = new double[1] { -1 };   
                        double[] tempVV = new double[1] { -1 };

                        // виброускорение
                        SqlCommand sqlCommand_10_5000   = new SqlCommand(querry_10_5000_VA, connection);
                        // виброскорость
                        SqlCommand sqlCommand_VV   = new SqlCommand(querry_VV, connection);


                        try
                        {
                            connection.Open();
                            // виброускорение
                            tempQ = querry_10_5000_VA;  tempVA[0] = Convert.ToDouble(sqlCommand_10_5000.ExecuteScalar());
                            // виброскорость
                            tempQ = querry_VV;          tempVV[0] = Convert.ToDouble(sqlCommand_VV.ExecuteScalar());
                            connection.Close();

                            dataIsOK = true;
                        }
                        catch   // если данных не будет
                        {
                            try { connection.Close(); } catch { };

                            if(nullPoints.Count == 0)
                            {
                                nullPoints.Add(Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])]);
                                EventLog.Log("Нехватка данных по точке " + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + ":\n" + tempQ + "\n");
                            }
                            else
                            {
                                if(nullPoints.ElementAt(nullPoints.Count - 1) != Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])])
                                {
                                    nullPoints.Add(Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])]);
                                    EventLog.Log("Нехватка данных по точке " + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + ":\n" + tempQ + "\n");
                                }
                            }

                            tempVA = new double[1] { -1 };
                            tempVV = new double[1] { -1 };
                        }                        

                        VA_Values.Add(tempVA);
                        VV_Values.Add(tempVV);
                    }                    
                    #endregion
                }
                EventLog.Log("Захват данных по вибропараметрам: успешно;");
                #endregion

                if(dataIsOK)
                {
                    if (nullPoints.Count > 0)
                    {
                        string messageNotFullData = "";
                        for (int i = 0; i < nullPoints.Count; i++)
                        {
                            if(i + 1 == nullPoints.Count) messageNotFullData += nullPoints.ElementAt(i) + " ";
                            else messageNotFullData += nullPoints.ElementAt(i) + "; ";
                        }

                        /*MessageBox.Show(
                            "Отсутствуют данные по некоторым точкам (" + messageNotFullData + ") за выбранный период времени!\n\nОтчет не будет заполнен данными по этим точкам.", 
                            "Внимание!", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);*/
                    }

                    ExcelDocument.MakeExcelReport(limitsList, VA_Values, VV_Values, VA_DataTables, VV_DataTables);

                    // завершение анимации в SCADA
                    ChangeCheck("0");
                }
                else
                {
                    ChangeCheck("2");
                    /*MessageBox.Show(
                        "Отсутствуют данные по всем точкам за выбранный период времени!", 
                        "Ошибка",
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Error, 
                        MessageBoxDefaultButton.Button1, 
                        MessageBoxOptions.DefaultDesktopOnly);*/

                    EventLog.Log("ОШИБКА!!  Отсутствие данных в БД.", false);
                    ChangeCheck("2");
                    Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                ChangeCheck("2");

                EventLog.Log(ex.ToString());    // запись логов

                try
                {
                    // завершение процесса Excel
                    Process[] process = Process.GetProcessesByName("EXCEL");
                    process[0].Kill();
                }
                catch
                {
                    EventLog.Log("Все Excel-процессы завершены.");    
                }
            }            
        }

        // меняет значение в текстовом документе Check.txt для изменения отображения статуса создания отчета в SCADA
        public static void ChangeCheck(string status)
        {
            // status = 0    отчет сформирован
            // status = 1    отчет в процессе формирования
            // status = 2    ошибка формирования отчета
            using (StreamWriter sw = new StreamWriter(Path.checkPath, false, Encoding.Default))
            {
                sw.WriteLine(status); 
            }
        }

        // запись уставок для 1 точки
        public static double[] SetLimits(int pointIndex, double[] limitsArray, SrezTableLight currDaySnapshotTable)
        {
            // ВУ 10-5000 Гц
            limitsArray[0] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VA_Strip_10_5000_Crash   [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;
            limitsArray[1] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VA_Strip_10_5000_Danger  [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;
            limitsArray[2] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VA_Strip_10_5000_Warning [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;
            limitsArray[3] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VA_Strip_10_5000_Norm    [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;

                // ВС 10-1000 Гц
            limitsArray[4] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VV_Strip_10_1000_Crash    [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;
            limitsArray[5] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VV_Strip_10_1000_Danger   [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;
            limitsArray[6] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VV_Strip_10_1000_Warning  [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;
            limitsArray[7] = currDaySnapshotTable.SrezList.Values[currDaySnapshotTable.SrezList.Count - 1].CnlData[Points.Lim_VV_Strip_10_1000_Norm     [Convert.ToInt32(Config.pointsArray[pointIndex])]].Val;

            return limitsArray;
        }

        // для написания запросов, с помощью которых будут забираться данные для графиков, а также мин-макс-средн. (зависит от parameter)
        public static string Querry(string parameter, string table, int CnlNum, DateTime start, DateTime end)
        {
            if (table == "CnlData") // если нужно взять данные из секундных измерений, где статус может быть = 0
            {
                return "SELECT " + parameter + " FROM " + table + " WHERE CnlNum = " + CnlNum.ToString() +
                    " AND (DateTime BETWEEN '" + start.ToString("yyyy-MM-dd") + "T" + start.ToString("HH:mm:ss") + ".000'" +    // начало периода
                    " AND '" + end.ToString("yyyy-MM-dd") + "T" + end.ToString("HH:mm:ss") + ".000')" +
                    " AND Stat <> 0";                          // конец периода
            }
            else
            {
                return "SELECT " + parameter + " FROM " + table + " WHERE CnlNum = " + CnlNum.ToString() +
                    " AND (DateTime BETWEEN '" + start.ToString("yyyy-MM-dd") + "T" + start.ToString("HH:mm:ss") + ".000'" +    // начало периода
                    " AND '" + end.ToString("yyyy-MM-dd") + "T" + end.ToString("HH:mm:ss") + ".000')";                          // конец периода
            }
        }

        // для заполнения коллекции данными, нужными для построения графиков
        public static void ListTableFill(string querry, SqlConnection connection, ref List<DataTable> DataTables)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            DataTable dataTable = new DataTable();

            dataAdapter.SelectCommand = new SqlCommand(querry, connection);
            dataAdapter.Fill(dataTable);
            DataTables.Add(dataTable);
            //dataTable.Clear();
        }
    }
}
