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
        //string path = @"..\\..\\..\\..\\Work\\ШАБЛОН.xlsx";

        //using (var workbook = new XLWorkbook(path))
        //{
        //    var sheet = workbook.Worksheet(1);
        //    decimal price = 0.0005000m;
        //    decimal stepprice = 0.0000001m;

        //    for(int i = 2; i <= 5001; i++)
        //    {
        //        sheet.Cell(i,2).Value=price;
        //        price-=stepprice;
        //        sheet.Cell(i,2).Style.NumberFormat.Format = "0.0000000";
        //        sheet.Cell(i, 3).Style.NumberFormat.Format = "0.0000000";
        //        sheet.Cell(i,5).Style.NumberFormat.Format = "0.000";
        //        sheet.Cell(i, 6).Style.NumberFormat.Format = "0.000";
        //        sheet.Cell(i, 7).Style.NumberFormat.Format = "0.0000000000";
        //        sheet.Cell(i, 8).Style.NumberFormat.Format = "0.0000000000";
        //    }
        //    workbook.Save();
        //}
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

                                Console.WriteLine($" Введите процент и нажмите ENTER\n" +
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
    }
}
