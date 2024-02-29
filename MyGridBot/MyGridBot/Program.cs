using Bybit.Net.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Authentication;
using CryptoExchange.Net.Objects.Options;
using System.Globalization;

namespace MyGridBot
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var dateTime = DateTime.Now;
            Console.Title = "BoViGridBot V2.2";
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(" Начинаю работу");
            SettingStart.Start();
            Console.WriteLine();
            Console.WriteLine(" Какой у Вас аккаунт единый или стандартный?\n" +
                " Если стандарнтый введите 0 и нажмите ENTER\n" +
                " Если единый нажмите 1 и нажмите ENTER");
            if (Console.ReadLine()=="0")
            {
                Console.Title = "BoViGridBot V2.2 Стандартный";
                BybitRestClient bybitRestClient = new BybitRestClient(options =>
                {
                    options.SpotOptions.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                    options.RequestTimeout = TimeSpan.FromSeconds(10);
                    options.ReceiveWindow = TimeSpan.FromSeconds(10);
                    options.AutoTimestamp = true;
                    options.TimestampRecalculationInterval = TimeSpan.FromMinutes(30);
                });

                await SettingStart.StartNewExelAsync(bybitRestClient);

                SettingStart.UpdateSymbolList();
                await ResultTrade.Balance(bybitRestClient, dateTime);

                while (true)
                {
                    await Trader.Buy(bybitRestClient);
                    await Trader.Sell(bybitRestClient);
                    await ResultTrade.Balance(bybitRestClient, dateTime);
                    await ResultTrade.TimerReversAsync(5, bybitRestClient);
                    SettingStart.UpdateSymbolList();
                }
            }
            else
            {
                Console.Title = "BoViGridBot V2.2 Единый";
                BybitRestClient bybitRestClient = new BybitRestClient(options =>
                {
                    options.V5Options.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                    options.RequestTimeout = TimeSpan.FromSeconds(10);
                    options.ReceiveWindow = TimeSpan.FromSeconds(10);
                    options.V5Options.AutoTimestamp = true;
                    options.TimestampRecalculationInterval = TimeSpan.FromMinutes(30);
                });

                await SettingStart.StartNewExelAsync(bybitRestClient);
                SettingStart.UpdateSymbolList();

                await ResultTrade.BalanceUnified(bybitRestClient, dateTime);
                while (true)
                {
                    await Trader.BuyUnified(bybitRestClient);
                    await Trader.SellUnified(bybitRestClient);
                    await ResultTrade.BalanceUnified(bybitRestClient, dateTime);
                    await ResultTrade.TimerReversAsync(5, bybitRestClient);
                    SettingStart.UpdateSymbolList();
                }
            }
        }
    }
}