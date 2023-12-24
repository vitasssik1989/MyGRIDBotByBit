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
            ResultTrade.TimerRevers(5);
            var dateTime = DateTime.Now;
            Console.Title = "BoViGridBot V1.2";
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
            //bybitRestClient = await ResultTrade.TestTimeSpan(bybitRestClient);
            
            Console.WriteLine(" Хотите создать новую ексель для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response.ToUpper() == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания новой ексель под торговую пару
                    await NewExcel.TradingPairAsync(tradingPair.ToUpper(), bybitRestClient);

                    Console.WriteLine(" Хотите ещё создать ексель для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }

            Console.WriteLine(" Создать ли Вам сетку для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response.ToUpper() == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания сетки под торговую пару
                    await NewExcel.Setka(tradingPair.ToUpper(), bybitRestClient);

                    Console.WriteLine(" Хотите ещё создать сетку для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }

            SettingStart.UpdateSymbolList();
            await ResultTrade.Balance(bybitRestClient,dateTime);

            while (true)
            {
                await Trader.Buy(bybitRestClient);
                await Trader.Sell(bybitRestClient);
                await ResultTrade.Balance(bybitRestClient, dateTime);
                ResultTrade.TimerRevers(5);
                SettingStart.UpdateSymbolList();
            }
        }
    }
}