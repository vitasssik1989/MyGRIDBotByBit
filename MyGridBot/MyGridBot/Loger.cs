using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class Loger
    {
        public static async Task WriteToFile(string text)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(@"..\\..\\..\\..\\Work\\Loger.txt", true))
                {
                    writer.WriteLine(DateTime.Now.ToString() + " " + text);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(" Не удалось записать в файл: " + ex.Message);
            }
        }
    }
}
