using Bybit.Net.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Authentication;
using CryptoExchange.Net.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class ResultTrade
    {
        public static ulong Buy { get; set; } = 0;
        public static ulong Sell { get; set; } = 0;
        static decimal TotalBalanceUSDT { get; set; } = 0;
        static int Copy { get; set; } = 0;

        public static async Task Balance(BybitRestClient bybitRestClient)
        {
            Copy++;
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine(" Делаю запрос баланса");
            WebCallResult<IEnumerable<Bybit.Net.Objects.Models.Spot.BybitSpotBalance>> balance = null;
            while (true)
            {
                balance = await bybitRestClient.SpotApiV3.Account.GetBalancesAsync();
                if (balance.Error == null) { break; }
                else
                {
                    Console.WriteLine($" Ошибка при запросе баланса \n" +
                                      $" {balance.Error.Code} {balance.Error.Message}");
                    Console.ReadLine();
                }
            }

            TotalBalanceUSDT = 0;
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine();
            foreach (var Symbol in SettingStart.SymbolList)
            {
                //USDT
                string asset = "";
                for (int i = 0; i < Symbol.Length - 4; i++)
                {
                    asset += Symbol[i];
                }
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{Symbol}.xlsx"))
                        {
                            var sheet = workbook.Worksheet(1);

                            TotalBalanceUSDT += Convert.ToDecimal(sheet.Cell(1, 9).Value);
                            foreach (var coin in balance.Data)
                            {
                                if (coin.Asset == asset)
                                {
                                    if (coin.Total < Convert.ToDecimal(sheet.Cell(1, 10).Value))
                                    {
                                        Console.WriteLine($" Монета:{asset} меньше в наличии чем в ексель");
                                        Console.ReadLine();
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Монета: {asset} Профит: {coin.Total - Convert.ToDecimal(sheet.Cell(1, 10).Value)}");
                                        break;
                                    }
                                }
                            }
                        }
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {Symbol}.xlsx метод Balance");
                        Thread.Sleep(10000);
                    }
                }
            }
            foreach (var coin in balance.Data)
            {

                if (coin.Asset == "USDT")
                {
                    if (coin.Total < TotalBalanceUSDT)
                    {
                        Console.WriteLine(" USDT на счете меньше чем нужно для сетки");
                        Console.ReadLine();
                    }
                    else
                    {
                        Console.WriteLine($" Монета: USDT Профит: {coin.Total - TotalBalanceUSDT}");
                        break;
                    }

                }
            }
            Console.WriteLine( );
            Console.WriteLine($" Сделки: Buy: {Buy} Sell: {Sell}");
            if(Copy > 200)
            {
                Copy = 0;
                CopyTable.Copy(@"..\\..\\..\\..\\Work", @"..\\..\\..\\..\\WorkCopy");
            }
        }

        public static void TimerRevers(int seconds)
        {
            Console.ForegroundColor= ConsoleColor.White;
            Console.WriteLine();
            Console.WriteLine(" Можно остановить консоль и редактировать все ексель файлы");
            var dateTime = DateTime.Now;
            DateTime dt = dateTime.AddSeconds(-seconds);
            while (dateTime >= dt)
            {
                var ticks = (dateTime - dt).Ticks;
                Console.WriteLine(new DateTime(ticks).ToString("      HH:mm:ss"));
                Thread.Sleep(850);
                dt = dt.AddSeconds(1);
            }
        }
        public static async Task<BybitRestClient> TestTimeSpan(BybitRestClient bybitRestClient)
        {
            int t = 5;
            WebCallResult<IEnumerable<Bybit.Net.Objects.Models.Spot.BybitSpotBalance>> balance = null;
            while (true)
            {
                balance = await bybitRestClient.SpotApiV3.Account.GetBalancesAsync();
                if (balance.Error == null) { return bybitRestClient; }
                else if(balance.Error.Code == 10002 )
                {
                    t += 5;
                    bybitRestClient = new BybitRestClient(options =>
                    {
                        options.SpotOptions.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                        options.ReceiveWindow = TimeSpan.FromSeconds(t);
                    });
                }
            }
        }
    }
}
