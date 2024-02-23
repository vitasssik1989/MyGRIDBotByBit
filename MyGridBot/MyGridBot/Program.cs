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
            Console.Title = "BoViGridBot V2.0";
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(" Начинаю работу");
            SettingStart.Start();
            
            BybitRestClient bybitRestClient = new BybitRestClient(options =>
            {
                options.SpotOptions.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                options.ReceiveWindow = TimeSpan.FromSeconds(10);
                options.TimestampRecalculationInterval = TimeSpan.FromSeconds(3);
                options.AutoTimestamp = true;
            });

            await SettingStart.StartNewExelAsync(bybitRestClient);

            SettingStart.UpdateSymbolList();
            await ResultTrade.Balance(bybitRestClient,dateTime);

            while (true)
            {
                await Trader.Buy(bybitRestClient);
                await Trader.Sell(bybitRestClient);
                await ResultTrade.Balance(bybitRestClient, dateTime);
                await ResultTrade.TimerReversAsync(5,bybitRestClient);
                SettingStart.UpdateSymbolList();
            }
        }
    }
}