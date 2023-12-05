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
                                    sheet.Cell(i,8).Value = symbol.MinOrderQuantity+symbol.BasePrecision;
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
        static string FormatZeroСomma(decimal PricePrecision)
        {
            string priceP = PricePrecision.ToString();
            string result = "";
            foreach (var item in priceP)
            {

                if(item == '1')
                {
                    result += '0';
                    continue;
                }
                if(item == ',')
                {
                    result += '.';
                    continue;
                }
                result += '0';
            }
            return  result;
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
    }
}
