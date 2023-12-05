using Bybit.Net.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Authentication;

namespace MyGridBot
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            Console.Title = "MyGridBot V1.0";
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(" Начинаю работу");
            SettingStart.Start();

            BybitRestClient bybitRestClient = new BybitRestClient(options =>
            {
                options.SpotOptions.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                options.ReceiveWindow = TimeSpan.FromSeconds(5);
            });

            
            //var result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
            //    (
            //        symbol: "LUNCUSDT",
            //        side: Bybit.Net.Enums.OrderSide.Sell,
            //        type: Bybit.Net.Enums.OrderType.Limit,
            //        price: 0.00000001m,
            //        quantity:0.227m,
            //        timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
            //    ) ;

            Console.WriteLine(" Хотите создать новую ексель для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания новой ексель под торговую пару
                    await NewExcel.TradingPairAsync(tradingPair, bybitRestClient);

                    Console.WriteLine(" Хотите ещё создать ексель для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }


            SettingStart.UpdateSymbolList();


            while (true)
            {
                await Trader.Buy(bybitRestClient);
                await Trader.Sell(bybitRestClient);
                await ResultTrade.Balance(bybitRestClient);
                SettingStart.UpdateSymbolList();
                ResultTrade.TimerRevers(5);
            }

        }
    }
}