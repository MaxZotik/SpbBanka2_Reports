using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpbBanka2_Reports
{
    class Config
    {
        public static string[] configInfo = File.ReadAllLines(Path.configInfoPath); // 0 - начало периода; 1 - конец периода; 2 - индексы выбранных точек; 3 - дата и время формирования отчета

        static string pointsIndexes = configInfo[2].Substring(0, configInfo[2].Length - 1); // убирается пробел после последней точки
        public static string[] pointsArray = pointsIndexes.Split(' '); // массив с индексами выбранных точек

        public static DateTime
            startDT = Convert.ToDateTime(configInfo[0]),            // начало периода
            endDT = Convert.ToDateTime(configInfo[1]),              // конец периода
            reportGeneration = Convert.ToDateTime(configInfo[3]);   // дата и время формирования отчета
    }
}
