using ClosedXML.Excel;
using CryptoExchange.Net.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class SettingStart
    {
        static string _path = @"..\\..\\..\\..\\Work\\Setting.xlsx";
        public static string APIkey { get; set; }
        public static string APIsecret { get; set; }

        public static List<string> SymbolList;

        public static void Start()
        {
            Console.WriteLine(" Открываю ексель Setting.xlsx в папке Work");

            while (true)
            {
                try
                {
                    while (true)
                    {
                        using (var workbook = new XLWorkbook(_path))
                        {
                            var sheet = workbook.Worksheet(1);

                            if (sheet.Cell(1, 3).IsEmpty() && sheet.Cell(2, 3).IsEmpty())
                            {
                                Console.WriteLine(" Укажите APIkey и APIsecret и нажмите ENTER");
                                Console.ReadLine();
                                workbook.Dispose();
                                continue;
                            }
                            APIkey = sheet.Cell(1, 3).Value.ToString();
                            APIsecret = sheet.Cell(2, 3).Value.ToString();
                            break;
                        }
                    }
                    break;
                }
                catch
                {
                    Console.WriteLine(" Не смог открыть ексель Setting.xlsx в папке Work\n" +
                                      " Проверь не открыта ли ексель или есть ли доступ");
                    Thread.Sleep(10000);
                }
            }
        }
        public static void UpdateSymbolList()
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine();
            Console.WriteLine(" Копирую все торговые пары");
            while (true)
            {
                try
                {
                    SymbolList = new List<string>();
                    using (var workbook = new XLWorkbook(_path))
                    {
                        var sheet = workbook.Worksheet(1);
                        for (int i = 2; i < 500; i++)
                        {
                            if (sheet.Cell(i, 1).IsEmpty() != true)
                            {
                                SymbolList.Add(sheet.Cell(i, 1).Value.ToString());
                            }
                            else { break; }
                        }
                    }
                    break;
                }
                catch
                {
                    Console.WriteLine(" Не смог открыть ексель Setting.xlsx в папке Work\n" +
                                      " Проверь не открыта ли ексель или есть ли доступ");
                    Thread.Sleep(10000);
                }
            }
        }
    }
}
