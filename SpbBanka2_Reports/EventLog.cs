using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpbBanka2_Reports
{
    class EventLog
    {
        static string logPath = Path.logPath; // путь к текстовому файлу, куда будут записываться логи

        // метод для логов; нужен для проверки корректности работы программы
        public static void Log(string mes)
        {
            try
            {
                if (!File.Exists(logPath))
                {
                    File.Create(logPath).Close();
                }

                File.AppendAllText(logPath, DateTime.Now.ToString() + "\t\t" + mes + Environment.NewLine);
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, DateTime.Now.ToString() + " - " + ex.ToString() + "\t\t" + Environment.NewLine);
                //MessageBox.Show("Ошибка в создании (записи) логов в журнал событий.");
            }
        }

        public static void Log(string mes, bool begin)
        {
            try
            {
                string separator = "==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==--==";

                if (!File.Exists(logPath))
                {
                    File.Create(logPath).Close();
                }

                if (begin)
                    File.AppendAllText(logPath, separator + "\n\n" + DateTime.Now.ToString() + "\t\t" + mes + Environment.NewLine);
                else
                    File.AppendAllText(logPath, DateTime.Now.ToString() + "\t\t" + mes + Environment.NewLine + "\n\n");

            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, DateTime.Now.ToString() + " - " + ex.ToString() + "\t\t" + Environment.NewLine);
                //MessageBox.Show("Ошибка в создании (записи) логов в журнал событий.");
            }
        }
    }
}
