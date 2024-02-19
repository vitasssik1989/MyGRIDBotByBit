using Bybit.Net.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class NewExcel
    {
        static string Path { get; set; } = @"..\\..\\..\\..\\Work";
        public static async Task TradingPairAsync(string tradingPair, BybitRestClient bybitRestClient)
        {
            string path = @"..\\..\\..\\..\\Work\\ШАБЛОН.xlsx";
            WebCallResult<System.Collections.Generic.IEnumerable<Bybit.Net.Objects.Models.Spot.v3.BybitSpotSymbolV3>> symbolData;
            while (true)
            {
                try
                {
                    symbolData = await bybitRestClient.SpotApiV3.ExchangeData.GetSymbolsAsync();
                    if (symbolData.ResponseStatusCode == System.Net.HttpStatusCode.OK)
                    {
                        break;
                    }
                }
                catch
                {

                }
            }
            foreach (var symbol in symbolData.Data)
            {
                if (symbol.Alias == tradingPair)
                {
                    while (true)
                    {
                        try
                        {
                            using (var workbook = new XLWorkbook(path))
                            {
                                var sheet = workbook.Worksheet(1);
                                string formatCommaPrice = FormatZeroСomma(symbol.PricePrecision);
                                string formatCommaBase = FormatZeroСomma(symbol.BasePrecision);
                                sheet.Cell(2, 15).Value = ValueAfterComma(symbol.BasePrecision);
                                //decimal minQty = Math.Round(symbol.MinOrderQuantity + (symbol.MinOrderQuantity / 100 * 0.1m),formatCommaBase.Length-2);
                                decimal minQty = symbol.MinOrderQuantity;
                                sheet.Cell(2, 8).Value = minQty;
                                while (Convert.ToDecimal(sheet.Cell(2, 7).Value) < symbol.MinOrderQuantity)
                                {
                                    minQty += symbol.BasePrecision;
                                    sheet.Cell(2, 8).Value = minQty;
                                }
                                for (int i = 2; i <= 5001; i++)
                                {
                                    sheet.Cell(i, 1).Value = 0;
                                    sheet.Cell(i, 4).Value = 0;
                                    sheet.Cell(i, 5).Value = 0;
                                    sheet.Cell(i, 6).Value = 0;


                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=D{i}=1")
                                    .Fill.SetBackgroundColor(XLColor.FromHtml("#ffa770"));
                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=A{i}=1")
                                    .Fill.SetBackgroundColor(XLColor.FromHtml("#a8ffc5"));
                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=Q{i}=1")
                                   .Fill.SetBackgroundColor(XLColor.FromHtml("#FF2400"));

                                    sheet.Cell(i, 9).Style.NumberFormat.Format = "0.00000000000000000000";
                                    sheet.Cell(i, 10).Style.NumberFormat.Format = "0.00000000000000000000";
                                    sheet.Cell(i, 14).Style.NumberFormat.Format = "0.000000000000000000000000000000";

                                    sheet.Cell(i, 2).Style.NumberFormat.Format = formatCommaPrice;//PricePrecision
                                    sheet.Cell(i, 3).Style.NumberFormat.Format = formatCommaPrice;//PricePrecision
                                    sheet.Cell(i, 2).Value = 0;
                                    sheet.Cell(i, 3).Value = 0;

                                    sheet.Cell(i, 7).Style.NumberFormat.Format = formatCommaBase; //BasePrecision
                                    sheet.Cell(i, 8).Style.NumberFormat.Format = formatCommaBase; //BasePrecision

                                    sheet.Cell(i, 8).Value = minQty;
                                }
                                workbook.SaveAs($@"..\\..\\..\\..\\Work\\{tradingPair}.xlsx");
                            }
                            break;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                            Console.WriteLine(" Не смог открыть ексель ШАБЛОН в папке Work");
                        }
                    }
                    break;
                }
            }
        }
        public static async Task Setka(string tradingPair, BybitRestClient bybitRestClient)
        {
            WebCallResult<System.Collections.Generic.IEnumerable<Bybit.Net.Objects.Models.Spot.v3.BybitSpotSymbolV3>> symbolData;
            while (true)
            {
                try
                {
                    symbolData = await bybitRestClient.SpotApiV3.ExchangeData.GetSymbolsAsync();
                    if (symbolData.ResponseStatusCode == System.Net.HttpStatusCode.OK)
                    {
                        break;
                    }
                }
                catch
                {

                }
            }
            foreach (var symbol in symbolData.Data)
            {
                if (symbol.Alias == tradingPair)
                {
                    while (true)
                    {
                        try
                        {

                            using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\Work\\{tradingPair}.xlsx"))
                            {
                                var sheet = workbook.Worksheet(1);
                                Console.WriteLine();
                                Console.WriteLine($" Введите максимальную цену и нажмите ENTER\n" +
                                                  $" Пример ввода: {symbol.PricePrecision} ");
                                decimal haigPrice = Kultura(Console.ReadLine());

                                Console.WriteLine($" Введите шаг цены и нажмите ENTER\n" +
                                                  $" Пример ввода: {symbol.PricePrecision}");
                                decimal priceStep = Kultura(Console.ReadLine());
                                while (true)
                                {
                                    if (priceStep < symbol.PricePrecision || CountTrailingZerosAfterDecimal(priceStep.ToString()) > CountTrailingZerosAfterDecimal(symbol.PricePrecision.ToString()))
                                    {
                                        Console.WriteLine($" Неверно указано кол-во символов: \n" +
                                                          $" Пример: {symbol.PricePrecision} \n" +
                                                          $" Вы ввели: {priceStep}\n" +
                                                          $" Введите шаг цены и нажмите ENTER");
                                        priceStep = Kultura(Console.ReadLine());
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                Console.WriteLine($" Введите процент продажи от цены покупки и нажмите ENTER\n" +
                                                  $" После запятой не больше 3 цифр\n" +
                                                  $" Пример ввода: 2,125");
                                decimal precent = Kultura(Console.ReadLine());
                                sheet.Cell(2, 2).Value = haigPrice;
                                sheet.Cell(2, 3).Value = MyСalculation(haigPrice, precent, symbol.PricePrecision.ToString());
                                for (int i = 3; i <= 5001; i++)
                                {
                                    if (haigPrice > 0)
                                    {
                                        haigPrice -= priceStep;
                                        sheet.Cell(i, 2).Value = haigPrice;
                                        sheet.Cell(i, 3).Value = MyСalculation(haigPrice, precent, symbol.PricePrecision.ToString());
                                    }
                                    else { break; }
                                }
                                workbook.Save();
                            }
                            break;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                            Console.WriteLine($" Не смог открыть ексель {tradingPair}.xlsx в папке Work");
                        }
                    }
                    break;
                }
            }
        }
        public static void SortBuySell()
        {
            List<decimal> sortBS = new List<decimal>();
            foreach (var excelSort in SettingStart.SymbolList)
            {
                try
                {
                    using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\Work\\{excelSort}.xlsx"))
                    {
                        var sheet = workbook.Worksheet(1);
                        if (!sheet.Cell(7, 16).IsEmpty())
                        {
                            if (Convert.ToInt32(sheet.Cell(7, 16).Value) == 1)
                            {
                                Console.WriteLine();
                                Console.Write(" Сортировка.Пара: ");
                                Console.ForegroundColor = ConsoleColor.White;
                                Console.Write($"{excelSort}");
                                Console.ForegroundColor = ConsoleColor.Blue;
                                Console.WriteLine();

                                //Сортировка ордеров на продажу

                                Console.Write(" Сортирую ордера! Тип: ");
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.Write($"Sell");
                                Console.ForegroundColor = ConsoleColor.Blue;
                                Console.WriteLine();
                                for (int s = 0; s < 2; s++)
                                {
                                    if (s == 0)
                                    {
                                        int flag = 0;
                                        for (int i = 2; i <= 5001; i++)
                                        {
                                            if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                            {
                                                if (flag == 0)
                                                {
                                                    flag = 1;
                                                }
                                                else { flag = 2; }
                                            }

                                            if (flag > 0)
                                            {
                                                if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                {
                                                    sortBS.Add(Convert.ToDecimal(sheet.Cell(i, 8).Value));
                                                }
                                            }
                                            if (flag == 2) { break; }
                                        }
                                        if (sortBS.Count > 0)
                                        {
                                            sortBS.Sort((a, b) => b.CompareTo(a));
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Нет ордеров! Тип: Sell");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        int sortIndex = 0;
                                        int flag = 0;
                                        for (int i = 2; i <= 5001; i++)
                                        {

                                            if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                            {
                                                if (flag == 0)
                                                {
                                                    flag = 1;
                                                }
                                                else { flag = 2; }
                                            }

                                            if (flag > 0)
                                            {
                                                if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                {
                                                    sheet.Cell(i, 8).Value = sortBS[sortIndex];
                                                    sortIndex++;
                                                }
                                            }
                                            if (flag == 2) { break; }
                                        }
                                        workbook.Save();
                                        Console.Write(" Кол-во ордеров: ");
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.Write($"{sortBS.Count}");
                                        Console.ForegroundColor = ConsoleColor.Blue;
                                        Console.WriteLine();
                                        sortBS.Clear();
                                    }
                                }

                                //Сортировка ордеров на покупку
                                Console.Write(" Сортирую ордера! Тип: ");
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.Write($"Buy");
                                Console.ForegroundColor = ConsoleColor.Blue;
                                Console.WriteLine();
                                for (int s = 0; s < 2; s++)
                                {
                                    if (s == 0)
                                    {
                                        int flag = 0;
                                        for (int i = 5001; i >= 2; i--)
                                        {
                                            if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                            {
                                                if (flag == 0)
                                                {
                                                    flag = 1;
                                                }
                                                else { flag = 2; }
                                            }

                                            if (flag > 0)
                                            {
                                                if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 0 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                {
                                                    sortBS.Add(Convert.ToDecimal(sheet.Cell(i, 8).Value));
                                                }
                                            }
                                            if (flag == 2) { break; }
                                        }
                                        if (sortBS.Count > 0)
                                        {
                                            sortBS.Sort((a, b) => b.CompareTo(a));
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Нет ордеров! Тип: Buy");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        int sortIndex = 0;
                                        int flag = 0;
                                        for (int i = 5001; i >= 2; i--)
                                        {
                                            if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                            {
                                                if (flag == 0)
                                                {
                                                    flag = 1;
                                                }
                                                else { flag = 2; }
                                            }

                                            if (flag > 0)
                                            {
                                                if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 0 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                {
                                                    sheet.Cell(i, 8).Value = sortBS[sortIndex];
                                                    sortIndex++;
                                                }
                                            }
                                            if (flag == 2) { break; }
                                        }
                                        workbook.Save();
                                        Console.Write(" Кол-во ордеров: ");
                                        Console.ForegroundColor = ConsoleColor.Green;
                                        Console.Write($"{sortBS.Count}");
                                        Console.ForegroundColor = ConsoleColor.Blue;
                                        Console.WriteLine();
                                        sortBS.Clear();
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(" Не смог открыть ексель");
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }

            }
        }
        static string FormatZeroСomma(decimal PricePrecision)
        {
            string priceP = PricePrecision.ToString();
            string result = "";
            foreach (var item in priceP)
            {

                if (item == '1')
                {
                    result += '0';
                    continue;
                }
                if (item == ',')
                {
                    result += '.';
                    continue;
                }
                result += '0';
            }
            return result;
        }
        static int ValueAfterComma(decimal BasePrecision)
        {
            string a = BasePrecision.ToString();
            int result = 0;
            if (a.Length == 1)
            {
                return 0;
            }
            return a.Length - 2;
        }
        static decimal Kultura(string kultyra)
        {
            decimal result = 0;
            if (decimal.TryParse(kultyra.Replace(',', '.'), out decimal H))
            {
                result = H;
            }
            else
            {
                if (decimal.TryParse(kultyra.Replace('.', ','), out decimal Hh))
                {
                    result = Hh;
                }
            }
            return result;
        }
        static decimal MyСalculation(decimal price, decimal precent, string PricePrecision)
        {
            return Math.Round(price + (price / 100 * precent), PricePrecision.Length - 2);
        }
        public static int CountTrailingZerosAfterDecimal(string input)
        {
            // Находим индекс разделителя (запятой или точки)
            int decimalIndex = input.IndexOfAny(new char[] { ',', '.' });

            // Если запятая не найдена или она в конце строки, возвращаем 0
            if (decimalIndex == -1 || decimalIndex == input.Length - 1)
                return 0;

            // Считаем количество нулей после запятой
            int zeroCount = 0;
            for (int i = input.Length - 1; i > decimalIndex; i--)
            {
                if (input[i] == '.' && input[i] == ',')
                {
                    break;
                }
                zeroCount++;
            }

            return zeroCount;
        }
    }
}
