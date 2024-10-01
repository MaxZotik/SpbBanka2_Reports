using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpbBanka2_Reports
{
    class TablePeriod
    {
        /*
            класс нужен для определения таблицы в БД, к которой нужно нужно обращаться для забора данных по точкам

            необходимо минимум 5 значений в таблице, приоритет таблиц:
            CnlData     [наиболее приоритетная]
            HourData 
            DailyData 
            WeeklyData  [наименее приоритетная]
         */


        public static string GetTable(DateTime start, DateTime end)
        {            
            return GetTablePlease();
        }

        

        // опредление таблицы по приоритетности
        static string GetTablePlease()
        {
            try
            {
                string 
                    mianTable = "",     // если не для всех точек есть данные в каждой из таблиц, в эту переменную будет помещено имя таблицы, в которой есть данные для большинства точек
                    currentTable = "";  // таблица для промежуточного выбора

                int tableCount = -1;
                string[] tables = new string[4] { "CnlData", "HourData", "DailyData", "WeeklyData" };

                // в коллекциях будут храниться номера нужных каналов
                List<int> VV_Channels = new List<int>();
                List<int> VA_Channels = new List<int>();
                for (int i = 0; i < Config.pointsArray.Length; i++)
                {
                    // добавление нужных каналов виброускорения
                    VA_Channels.Add(Points.Parameters_VA_Strip_10_5000[Convert.ToInt32(Config.pointsArray[i])]);

                    // добавление нужных каналов виброскорости
                    VV_Channels.Add(Points.Parameters_VV_Strip_10_1000[Convert.ToInt32(Config.pointsArray[i])]);
                }

                SqlConnection connection = new SqlConnection(Path.connectionString);

                bool dataIsOK = true;   // полнота данных в БД для одной точки
                double[] recordsAmount = new double[2] { 0, 0 }; // количество записей из БД (должно быть больше 5)
                int 
                    tableWithData = 0,
                    maxTableWithData = 0;

                for (int tablesCount = 0; tablesCount < 4; tablesCount++) // проход по каждой таблице
                {
                    currentTable = TableVariant();

                    for (int point = 0, VAIndex = 0, VVIndex = 0; point < Config.pointsArray.Length; point++, VAIndex++, VVIndex++)
                    {
                        dataIsOK = true;
                        for (int j = 0; j < recordsAmount.Length; j++)
                        {
                            recordsAmount[j] = 0;
                        }

                        string param = "COUNT(*)";  // для вставки в запрос

                        string
                            // виброускорение
                            querry_VA = Querry(param, currentTable, VA_Channels[VAIndex], Config.startDT, Config.endDT),
                            // виброскорость
                            querry_VV  = Querry(param, currentTable, VV_Channels[VVIndex], Config.startDT, Config.endDT),

                            // для отправки в лог запроса при нехватки значений
                            tempQ = querry_VA;

                        // виброускорение
                        SqlCommand sqlCommand_VA = new SqlCommand(querry_VA, connection);
                        // виброскорость
                        SqlCommand sqlCommand_VV  = new SqlCommand(querry_VV, connection);

                        try
                        {
                            connection.Open();
                            // виброускорение
                            tempQ = querry_VA; recordsAmount[0]    = Convert.ToDouble(sqlCommand_VA.ExecuteScalar());
                            // виброскорость
                            tempQ = querry_VV; recordsAmount[1]    = Convert.ToDouble(sqlCommand_VV.ExecuteScalar());
                            connection.Close();
                        }
                        catch (Exception exx)   // если данных не будет
                        {
                            try { connection.Close(); } catch { };
                            EventLog.Log("Ошибка формирования запроса (точка " + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + "):\n" + tempQ + "\n" + exx.ToString());

                            return "Error";
                        }

                        for (int j = 0; j < recordsAmount.Length; j++)
                        {
                            if (recordsAmount[j] < 5) dataIsOK = false;
                        }

                        if (dataIsOK)
                            tableWithData++;    // есть минимум 5 записей в приоритетнейшей таблице для одной точки
                        else
                            EventLog.Log(
                                "Маленькое количество записей на полосах для точки " + Points.pointsNames[Convert.ToInt32(Config.pointsArray[point])] + "\tв таблице " + tables[tablesCount] +
                                "\tВУ 10...5000Гц\t= " + recordsAmount[0] +
                                "\tВС            \t= " + recordsAmount[1]);
                    }

                    if (maxTableWithData < tableWithData)
                    {
                        maxTableWithData = tableWithData;
                        mianTable = currentTable;
                    }

                    if (tableWithData == Config.pointsArray.Length) return currentTable;   // если для всех точек есть значения в таблице
                }

                if (mianTable != "") return mianTable;
                else return "ErrorNoData";

                // формирование текста запроса
                string Querry(string parameter, string _table, int CnlNum, DateTime start, DateTime end)
                {
                    if (_table == "CnlData") // если нужно взять данные из секундных измерений, где статус может быть = 0
                    {
                        return "SELECT " + parameter + " FROM " + _table + " WHERE CnlNum = " + CnlNum.ToString() +
                            " AND (DateTime BETWEEN '" + start.ToString("yyyy-MM-dd") + "T" + start.ToString("HH:mm:ss") + ".000'" +    // начало периода
                            " AND '" + end.ToString("yyyy-MM-dd") + "T" + end.ToString("HH:mm:ss") + ".000')" +                         // конец периода
                            " AND Stat <> 0";
                    }
                    else
                    {
                        return "SELECT " + parameter + " FROM " + _table + " WHERE CnlNum = " + CnlNum.ToString() +
                            " AND (DateTime BETWEEN '" + start.ToString("yyyy-MM-dd") + "T" + start.ToString("HH:mm:ss") + ".000'" +    // начало периода
                            " AND '" + end.ToString("yyyy-MM-dd") + "T" + end.ToString("HH:mm:ss") + ".000')";                          // конец периода
                    }
                }

                // подбор таблицы
                string TableVariant()
                {
                    tableCount++;
                    return tables[tableCount];
                }
            }
            catch(Exception eee)    // если данных не будет
            {                
                EventLog.Log("Ошибка при подборе таблицы:\n" + eee.ToString());

                /*MessageBox.Show(
                        "Ошибка выбора архива.",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);*/

                return "Error";
            }                       
        }
    }
}
