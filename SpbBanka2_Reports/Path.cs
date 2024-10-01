using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpbBanka2_Reports
{
    class Path
    {
        public static string
            // для записи логов
            logPath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Program\Log.txt",

            // для изменения отображения статуса формирования отчета в SCADA
            checkPath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Check.txt",

            // для забора даты начала и конца периода в отчете, времени создания отчета и нужных точек
            configInfoPath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\ConfigurationInfo.txt",

            // путь к шаблону для отчета
            modelPath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Program\RepModel.xlsx",

            // пути для сохранения и забора картинок графиков
            VA_GraphSavePath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Reports\Graph\VA\",
            VV_GraphSavePath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Reports\Graph\VV\",

            // пути для удаления и создания папок для графиков
            VA_GraphDelCreatePath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Reports\Graph\VA\",
            VV_GraphDelCreatePath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Reports\Graph\VV\",

            // путь для сохранения отчета
            saveRepPath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Reports\",

            // для строки подключения к БД
            dbInfoPath = @"C:\SCADA\SCADA Reports\SPB_Banka_2\Custom Reports\Program\db_name.txt";

        public static string[] dbInfo = File.ReadAllLines(dbInfoPath); // 0 - имя сервера, 1 - имя базы

        // строка подключения к БД
        public static string connectionString = "Data Source=" + dbInfo[0] + ";Initial Catalog=" + dbInfo[1] + ";Integrated Security=True";

        // удаление папки с изображениями графиков
        public static void DelCat(string dirName)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(dirName);
            if (dirInfo.Exists)
            {
                dirInfo.Delete(true);
            }
        }

        // создание папки для изображений графиков
        public static void CreateCat(string dirName)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(dirName);
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
        }
    }
}
