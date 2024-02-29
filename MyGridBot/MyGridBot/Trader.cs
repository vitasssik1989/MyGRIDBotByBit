using Bybit.Net.Clients;
using Bybit.Net.Objects.Models.Spot;
using ClosedXML.Excel;
using CryptoExchange.Net.CommonObjects;
using CryptoExchange.Net.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class Trader
    {

        public static async Task Buy(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Buy <<<<<<<<<<<<");

            BybitSpotOrderBookEntry Ask = null;

            foreach (var BuySymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {BuySymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{BuySymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                            int grafic = 0;
                            
                            //Трейлинг
                            if (!sheet.Cell(6, 15).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 15).Value);
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 15).Value);
                                if ( precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Ask.Price)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Ask.Price)
                                    {
                                        if (strategPrice + (strategPrice / 100 * precent) <= Ask.Price)
                                        {
                                            sheet.Cell(4, 15).Value = Ask.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 2; i <= 5001; i++)
                            {
                                grafic = i;
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) == 1)
                                { continue; }
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 0)
                                {
                                    if (Ask.Price < Convert.ToDecimal(sheet.Cell(i, 2).Value) && Ask.Quantity > Convert.ToDecimal(sheet.Cell(i, 11).Value))
                                    {
                                        //Реинвестирование
                                        if (Convert.ToInt32(sheet.Cell(i, 5).Value) == 1)
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 11).Value)}\n" +
                                                              $" Реинвестиция: ДА");

                                            if (await BuyResult(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 11).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 8).Value = Convert.ToDecimal(sheet.Cell(i, 11).Value);
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                break;
                                            }

                                        }
                                        else
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 8).Value)}\n" +
                                                              $" Реинвестиция: НЕТ");

                                            if (await BuyResult(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 8).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                break;
                                            }

                                        }

                                        await Task.Delay(100);
                                        Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на покупку");
                                        break;
                                    }
                                }
                                if (i == 5001 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.WriteLine(" Закончилась сетка на покупку");
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                            Grafic.Write(grafic);
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {BuySymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }
            }
        }
        public static async Task Sell(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Sell <<<<<<<<<<<<");

            BybitSpotOrderBookEntry Bid = null;

            foreach (var SellSymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {SellSymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{SellSymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                            //Трейлинг
                            if (!sheet.Cell(6, 16).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 16).Value);//0.5
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 16).Value);//0
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;//0.0001245
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Bid.Price)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Bid.Price)
                                    {       
                                        if (strategPrice - (strategPrice / 100 * precent) >= Bid.Price)
                                        {
                                            sheet.Cell(4, 16).Value = Bid.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else 
                                    {
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break; 
                                    }
                                }
                            }
                            for (int i = 5001; i >= 2; i--)
                            {
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) != 2)
                                {
                                    if (Bid.Price > Convert.ToDecimal(sheet.Cell(i, 3).Value) && Bid.Quantity > Convert.ToDecimal(sheet.Cell(i, 7).Value))
                                    {
                                        Console.WriteLine();
                                        Console.WriteLine($" Продажа Торговой Пары: {SellSymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 3).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 7).Value)}");
                                        if (await SellResult(bybitRestClient, SellSymbol, Convert.ToDecimal(sheet.Cell(i, 3).Value), Convert.ToDecimal(sheet.Cell(i, 7).Value)))
                                        {
                                            Console.WriteLine(" Заявка исполнилась");
                                            sheet.Cell(i, 4).Value = 0;
                                            save = true;
                                            await Task.Delay(200);
                                        }
                                        else
                                        {
                                            Console.WriteLine(" Заявка не исполнилась");
                                            break;
                                        }
                                        await Task.Delay(100);
                                        Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на продажу");
                                        break;
                                    }
                                }
                                if (i == 2 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.WriteLine(" Закончилась сетка на продажу");
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {SellSymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }

            }
        }

        static async Task<bool> BuyResult(BybitRestClient bybitRestClient, string BuySymbol, decimal price, decimal quantity)
        {
            bool resltBuy = true;
            try
            {
                WebCallResult<BybitSpotOrderPlaced> result = null;
                WebCallResult<Bybit.Net.Objects.Models.Spot.v3.BybitSpotOrderV3> resultOrderBuy = null;
                try
                {
                    result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
                               (
                                   symbol: BuySymbol,
                                   side: Bybit.Net.Enums.OrderSide.Buy,
                                   type: Bybit.Net.Enums.OrderType.Limit,
                                   price: price,
                                   quantity: quantity,
                                   timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} стр 312");
                    await Loger.WriteToFile($"{ex.Message} стр 312");
                    Console.ReadLine();
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderBuy = await bybitRestClient.SpotApiV3.Trading.GetOrderAsync(clientOrderId: result.Data.ClientOrderId);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 327");
                            await Loger.WriteToFile($"{ex.Message} стр 327");
                            Console.ReadLine();
                        }

                        if (resultOrderBuy.Error == null)
                        {
                            if (resultOrderBuy.Data.Status == Bybit.Net.Enums.OrderStatus.Filled)
                            {
                                resltBuy = true;
                                break;
                            }
                            else if (resultOrderBuy.Data.Status == Bybit.Net.Enums.OrderStatus.Canceled)
                            {
                                resltBuy = false;
                                break;
                            }
                            await Task.Delay(2000);
                            continue;
                        }
                        else if (resultOrderBuy.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            await Loger.WriteToFile($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            Console.ReadLine();
                        }
                    }

                }
                else if (result.Error.Code == 12193)
                {
                    while (true)
                    {
                        BybitSpotOrderBookEntry Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                        try
                        {
                            result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
                                            (
                                                symbol: BuySymbol,
                                                side: Bybit.Net.Enums.OrderSide.Buy,
                                                type: Bybit.Net.Enums.OrderType.Limit,
                                                price: Ask.Price,
                                                quantity: quantity,
                                                timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                            );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 380");
                            await Loger.WriteToFile($"{ex.Message} стр 380");
                            Console.ReadLine();
                        }
                        if (result.Error == null)
                        {
                            while (true)
                            {
                                try
                                {
                                    resultOrderBuy = await bybitRestClient.SpotApiV3.Trading.GetOrderAsync(clientOrderId: result.Data.ClientOrderId);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message} стр 394");
                                    await Loger.WriteToFile($"{ex.Message} стр 394");
                                    Console.ReadLine();
                                }

                                if (resultOrderBuy.Error == null)
                                {
                                    if (resultOrderBuy.Data.Status == Bybit.Net.Enums.OrderStatus.Filled)
                                    {
                                        resltBuy = true;
                                        break;
                                    }
                                    else if (resultOrderBuy.Data.Status == Bybit.Net.Enums.OrderStatus.Canceled)
                                    {
                                        resltBuy = false;
                                        break;
                                    }
                                    await Task.Delay(1000); continue;
                                }
                                else if (resultOrderBuy.Error.Code == 10002)
                                {
                                    await Task.Delay(2000);
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                                    await Loger.WriteToFile($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message} Клас Trader стр 422");
                                    Console.WriteLine(" Клас Trader стр 422");
                                    Console.ReadLine();
                                }
                            }
                        }
                        else if (result.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                            if (result.Error.Code == 12193)
                            {
                                continue;
                            }
                            Console.ReadLine();
                        }
                        break;
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltBuy = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message} Клас Trader стр 451");
                    Console.WriteLine(" Клас Trader стр 451");
                    Console.ReadLine();
                }
                if (resltBuy == true)
                {
                    ResultTrade.Buy++;
                }
            }
            catch
            {
                //resltBuy = true;
                IEnumerable<Bybit.Net.Objects.Models.Spot.v3.BybitSpotOrderV3> histori = null;
                while (histori == null)
                {
                    try
                    {
                        try { histori = (await bybitRestClient.SpotApiV3.Trading.GetOrdersAsync(BuySymbol, limit: 1)).Data; }
                        catch (Exception ex) { await Loger.WriteToFile($"{ex.Message} стр 469"); await Task.Delay(1000); }
                        if (histori == null) { await Task.Delay(300); continue; }
                        foreach (var orderHistory in histori)
                        {
                            if (orderHistory.Side == Bybit.Net.Enums.OrderSide.Buy && orderHistory.Status == Bybit.Net.Enums.OrderStatus.Filled)
                            {
                                resltBuy = true;
                            }
                            else
                            {
                                resltBuy = false;
                            }
                        }
                    }
                    catch { }

                }
            }

            return resltBuy;
        }
        static async Task<bool> SellResult(BybitRestClient bybitRestClient, string SellSymbol, decimal price, decimal quantity)
        {
            bool resltSell = true;
            try
            {
                WebCallResult<BybitSpotOrderPlaced> result = null;
                WebCallResult<Bybit.Net.Objects.Models.Spot.v3.BybitSpotOrderV3> resultOrderSell = null;
                try
                {
                    result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
                                (
                                    symbol: SellSymbol,
                                    side: Bybit.Net.Enums.OrderSide.Sell,
                                    type: Bybit.Net.Enums.OrderType.Limit,
                                    price: price,
                                    quantity: quantity,
                                    timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                 );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} стр 511");
                    await Loger.WriteToFile($"{ex.Message} стр 511");
                    Console.ReadLine();
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderSell = await bybitRestClient.SpotApiV3.Trading.GetOrderAsync(clientOrderId: result.Data.ClientOrderId);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 526");
                            await Loger.WriteToFile($"{ex.Message} стр 526");
                            Console.ReadLine();
                        }

                        if (resultOrderSell.Error == null)
                        {
                            if (resultOrderSell.Data.Status == Bybit.Net.Enums.OrderStatus.Filled)
                            {
                                resltSell = true;
                                break;
                            }
                            else if (resultOrderSell.Data.Status == Bybit.Net.Enums.OrderStatus.Canceled)
                            {
                                resltSell = false;
                                break;
                            }
                            await Task.Delay(1000); continue;
                        }
                        else if (resultOrderSell.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                            await Loger.WriteToFile($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message} Клас Trader стр 553");
                            Console.WriteLine(" Клас Trader стр 553");
                            Console.ReadLine();
                        }
                    }

                }
                else if (result.Error.Code == 12194)
                {
                    while (true)
                    {
                        BybitSpotOrderBookEntry Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                        try
                        {
                            result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
                                            (
                                                symbol: SellSymbol,
                                                side: Bybit.Net.Enums.OrderSide.Sell,
                                                type: Bybit.Net.Enums.OrderType.Limit,
                                                price: Bid.Price,
                                                quantity: quantity,
                                                timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                            );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 579");
                            await Loger.WriteToFile($"{ex.Message} стр 579");
                            Console.ReadLine();
                        }
                        if (result.Error == null)
                        {
                            while (true)
                            {
                                try
                                {
                                    resultOrderSell = await bybitRestClient.SpotApiV3.Trading.GetOrderAsync(clientOrderId: result.Data.ClientOrderId);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message} стр 593");
                                    await Loger.WriteToFile($"{ex.Message} стр 593");
                                    Console.ReadLine();
                                }

                                if (resultOrderSell.Error == null)
                                {
                                    if (resultOrderSell.Data.Status == Bybit.Net.Enums.OrderStatus.Filled)
                                    {
                                        resltSell = true;
                                        break;
                                    }
                                    else if (resultOrderSell.Data.Status == Bybit.Net.Enums.OrderStatus.Canceled)
                                    {
                                        resltSell = false;
                                        break;
                                    }
                                    await Task.Delay(1000); continue;
                                }
                                else if (resultOrderSell.Error.Code == 10002)
                                {
                                    await Task.Delay(2000);
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                                    await Loger.WriteToFile($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message} Клас Trader стр 620");
                                    Console.WriteLine(" Клас Trader стр 620");
                                    Console.ReadLine();
                                }
                            }
                        }
                        else if (result.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}"); 
                            await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message}");
                            if (result.Error.Code == 12194)
                            {
                                continue;
                            }
                            Console.ReadLine();
                        }
                        break;
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltSell = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message} Клас Trader стр 651");
                    Console.WriteLine(" Клас Trader стр 651");
                    Console.ReadLine();
                }
                if (resltSell == true)
                {
                    ResultTrade.Sell++;
                }
            }
            catch
            {
                //resltSell = true;
                IEnumerable<Bybit.Net.Objects.Models.Spot.v3.BybitSpotOrderV3> histori = null;
                while (histori == null)
                {
                    try
                    {
                        try { histori = (await bybitRestClient.SpotApiV3.Trading.GetOrdersAsync(SellSymbol, limit: 1)).Data; }
                        catch (Exception ex) { await Loger.WriteToFile($"{ex.Message} стр 698"); await Task.Delay(1000); }
                        if (histori == null) { await Task.Delay(300); continue; }
                        foreach (var orderHistory in histori)
                        {
                            if (orderHistory.Side == Bybit.Net.Enums.OrderSide.Sell && orderHistory.Status == Bybit.Net.Enums.OrderStatus.Filled)
                            {
                                resltSell = true;
                            }
                            else
                            {
                                resltSell = false;
                            }
                        }
                    }
                    catch { }
                }
            }
            return resltSell;
        }

        static async Task<BybitSpotOrderBookEntry> AskPriceQuantity(BybitRestClient bybitRestClient, string BuySymbol)
        {
            WebCallResult<BybitSpotOrderBook> orderBookData = null;
            while (true)
            {
                try
                {
                    orderBookData = await bybitRestClient.SpotApiV3.ExchangeData.GetOrderBookAsync(BuySymbol);

                    if (orderBookData.Error != null)
                    {
                        Console.WriteLine($" Не получил данные по стакану AskPriceQuantity\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                        await Loger.WriteToFile($" Не получил данные по стакану AskPriceQuantity\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану AskPriceQuantity\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                    await Loger.WriteToFile($" Не получил данные по стакану AskPriceQuantity\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                    Thread.Sleep(1000);
                }
                if (orderBookData.Data.Asks.First().Price <= 0 && orderBookData.Data.Asks.First().Quantity <= 0)
                {
                    continue;
                }
                break;
            }
            var Ask = orderBookData.Data.Asks.First();
            Console.WriteLine($" AskPrice: {Ask.Price} AskQuantity: {Ask.Quantity}");
            return orderBookData.Data.Asks.First();
        }
        static async Task<BybitSpotOrderBookEntry> BidPriceQuantity(BybitRestClient bybitRestClient, string SellSymbol)
        {
            WebCallResult<BybitSpotOrderBook> orderBookData = null;
            while (true)
            {
                try
                {
                    orderBookData = await bybitRestClient.SpotApiV3.ExchangeData.GetOrderBookAsync(SellSymbol);
                    if (orderBookData.Error != null)
                    {
                        Console.WriteLine($" Не получил данные по стакану BidPriceQuantity\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                        await Loger.WriteToFile($" Не получил данные по стакану BidPriceQuantity\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану BidPriceQuantity\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                    await Loger.WriteToFile($" Не получил данные по стакану BidPriceQuantity\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                    Thread.Sleep(1000);
                }
                if (orderBookData.Data.Bids.First().Price <= 0 && orderBookData.Data.Bids.First().Quantity <= 0)
                {
                    continue;
                }
                break;
            }
            var Bid = orderBookData.Data.Bids.First();
            Console.WriteLine($" BidPrice: {Bid.Price} BidQuantity: {Bid.Quantity}");
            return orderBookData.Data.Bids.First();
        }


        public static async Task BuyUnified(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Buy <<<<<<<<<<<<");

            BybitSpotOrderBookEntry Ask = null;

            foreach (var BuySymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {BuySymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{BuySymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                            int grafic = 0;

                            //Трейлинг
                            if (!sheet.Cell(6, 15).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 15).Value);
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 15).Value);
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Ask.Price)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Ask.Price)
                                    {
                                        if (strategPrice + (strategPrice / 100 * precent) <= Ask.Price)
                                        {
                                            sheet.Cell(4, 15).Value = Ask.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 2; i <= 5001; i++)
                            {
                                grafic = i;
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) == 1)
                                { continue; }
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 0)
                                {
                                    if (Ask.Price < Convert.ToDecimal(sheet.Cell(i, 2).Value) && Ask.Quantity > Convert.ToDecimal(sheet.Cell(i, 11).Value))
                                    {
                                        //Реинвестирование
                                        if (Convert.ToInt32(sheet.Cell(i, 5).Value) == 1)
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 11).Value)}\n" +
                                                              $" Реинвестиция: ДА");

                                            if (await BuyResultUnified(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 11).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 8).Value = Convert.ToDecimal(sheet.Cell(i, 11).Value);
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                break;
                                            }

                                        }
                                        else
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 8).Value)}\n" +
                                                              $" Реинвестиция: НЕТ");

                                            if (await BuyResultUnified(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 8).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                break;
                                            }

                                        }

                                        await Task.Delay(100);
                                        Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на покупку");
                                        break;
                                    }
                                }
                                if (i == 5001 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.WriteLine(" Закончилась сетка на покупку");
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                            Grafic.Write(grafic);
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {BuySymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }
            }
        }
        public static async Task SellUnified(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Sell <<<<<<<<<<<<");

            BybitSpotOrderBookEntry Bid = null;

            foreach (var SellSymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {SellSymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{SellSymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                            //Трейлинг
                            if (!sheet.Cell(6, 16).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 16).Value);//0.5
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 16).Value);//0
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;//0.0001245
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Bid.Price)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Bid.Price)
                                    {
                                        if (strategPrice - (strategPrice / 100 * precent) >= Bid.Price)
                                        {
                                            sheet.Cell(4, 16).Value = Bid.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 5001; i >= 2; i--)
                            {
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) != 2)
                                {
                                    if (Bid.Price > Convert.ToDecimal(sheet.Cell(i, 3).Value) && Bid.Quantity > Convert.ToDecimal(sheet.Cell(i, 7).Value))
                                    {
                                        Console.WriteLine();
                                        Console.WriteLine($" Продажа Торговой Пары: {SellSymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 3).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 7).Value)}");
                                        if (await SellResultUnified(bybitRestClient, SellSymbol, Convert.ToDecimal(sheet.Cell(i, 3).Value), Convert.ToDecimal(sheet.Cell(i, 7).Value)))
                                        {
                                            Console.WriteLine(" Заявка исполнилась");
                                            sheet.Cell(i, 4).Value = 0;
                                            save = true;
                                            await Task.Delay(200);
                                        }
                                        else
                                        {
                                            Console.WriteLine(" Заявка не исполнилась");
                                            break;
                                        }
                                        await Task.Delay(100);
                                        Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на продажу");
                                        break;
                                    }
                                }
                                if (i == 2 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.WriteLine(" Закончилась сетка на продажу");
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {SellSymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }

            }
        }

        static async Task<bool> BuyResultUnified(BybitRestClient bybitRestClient, string BuySymbol, decimal price, decimal quantity)
        {
            bool resltBuy = true;
            try
            {
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitOrderId> result = null;
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitOrder>> resultOrderBuy = null;
                try
                {
                     result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                               (
                                   Bybit.Net.Enums.Category.Spot,
                                   symbol: BuySymbol,
                                   side: Bybit.Net.Enums.OrderSide.Buy,
                                   type: Bybit.Net.Enums.NewOrderType.Limit,
                                   price: price,
                                   quantity: quantity,
                                   timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} стр 1059");
                    await Loger.WriteToFile($"{ex.Message} стр 1060");
                    Console.ReadLine();
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                                resultOrderBuy = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                (
                                    category:Bybit.Net.Enums.Category.Spot,
                                    symbol:BuySymbol,
                                    clientOrderId:result.Data.ClientOrderId
                                );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 1079");
                            await Loger.WriteToFile($"{ex.Message} стр 1080");
                            Console.ReadLine();
                        }

                        if (resultOrderBuy.Error == null)
                        {
                            if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                            {
                                resltBuy = true;
                                break;
                            }
                            else if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                            {
                                resltBuy = false;
                                break;
                            }
                            await Task.Delay(2000);
                            continue;
                        }
                        else if (resultOrderBuy.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            await Loger.WriteToFile($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            Console.ReadLine();
                        }
                    }

                }
                else if (result.Error.Code == 170193)
                {
                    while (true)
                    {
                        BybitSpotOrderBookEntry Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                        try
                        {
                            result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                                    (
                                       Bybit.Net.Enums.Category.Spot,
                                       symbol: BuySymbol,
                                       side: Bybit.Net.Enums.OrderSide.Buy,
                                       type: Bybit.Net.Enums.NewOrderType.Limit,
                                       price: Ask.Price,
                                       quantity: quantity,
                                       timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                    );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 1133");
                            await Loger.WriteToFile($"{ex.Message} стр 1133");
                            Console.ReadLine();
                        }
                        if (result.Error == null)
                        {
                            while (true)
                            {
                                try
                                {
                                    resultOrderBuy = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                (
                                    category: Bybit.Net.Enums.Category.Spot,
                                    symbol: BuySymbol,
                                    clientOrderId: result.Data.ClientOrderId
                                );
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message} стр 1152");
                                    await Loger.WriteToFile($"{ex.Message} стр 1152");
                                    Console.ReadLine();
                                }

                                if (resultOrderBuy.Error == null)
                                {
                                    if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                                    {
                                        resltBuy = true;
                                        break;
                                    }
                                    else if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                                    {
                                        resltBuy = false;
                                        break;
                                    }
                                    await Task.Delay(1000); continue;
                                }
                                else if (resultOrderBuy.Error.Code == 10002)
                                {
                                    await Task.Delay(2000);
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                                    await Loger.WriteToFile($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message} Клас Trader стр 1179");
                                    Console.WriteLine(" Клас Trader стр 1179");
                                    Console.ReadLine();
                                }
                            }
                        }
                        else if (result.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                            await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message}");
                            if (result.Error.Code == 170193)
                            {
                                continue;
                            }
                            Console.ReadLine();
                        }
                        break;
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltBuy = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message} Клас Trader стр 1210");
                    Console.WriteLine(" Клас Trader стр 1210");
                    Console.ReadLine();
                }
                if (resltBuy == true)
                {
                    ResultTrade.Buy++;
                }
            }
            catch
            {
                //resltSell = true;
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitOrder>> histori = null;
                while (histori == null)
                {
                    try
                    {

                        histori = await bybitRestClient.V5Api.Trading.GetOrderHistoryAsync
                            (
                                Bybit.Net.Enums.Category.Spot,
                                BuySymbol,
                                limit: 1
                            );
                        if (histori == null) { await Task.Delay(300); continue; }
                        foreach (var orderHistory in histori.Data.List)
                        {
                            if (orderHistory.Side == Bybit.Net.Enums.OrderSide.Buy && orderHistory.Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                            {
                                resltBuy = true;
                            }
                            else
                            {
                                resltBuy = false;
                            }
                        }
                    }
                    catch { }
                }
            }

            return resltBuy;
        }
        static async Task<bool> SellResultUnified(BybitRestClient bybitRestClient, string SellSymbol, decimal price, decimal quantity)
        {
            bool resltSell = true;
            try
            {
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitOrderId> result = null;
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitOrder>> resultOrderSell = null;
                try
                {
                    result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                               (
                                   Bybit.Net.Enums.Category.Spot,
                                   symbol: SellSymbol,
                                   side: Bybit.Net.Enums.OrderSide.Sell,
                                   type: Bybit.Net.Enums.NewOrderType.Limit,
                                   price: price,
                                   quantity: quantity,
                                   timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} стр 1275");
                    await Loger.WriteToFile($"{ex.Message} стр 1275");
                    Console.ReadLine();
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderSell = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                (
                                    category: Bybit.Net.Enums.Category.Spot,
                                    symbol: SellSymbol,
                                    clientOrderId: result.Data.ClientOrderId
                                );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 1295");
                            await Loger.WriteToFile($"{ex.Message} стр 1295");
                            Console.ReadLine();
                        }

                        if (resultOrderSell.Error == null)
                        {
                            if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                            {
                                resltSell = true;
                                break;
                            }
                            else if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                            {
                                resltSell = false;
                                break;
                            }
                            await Task.Delay(1000); continue;
                        }
                        else if (resultOrderSell.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                            await Loger.WriteToFile($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}  Клас Trader стр 1322");
                            Console.WriteLine(" Клас Trader стр 1322");
                            Console.ReadLine();
                        }
                    }

                }
                else if (result.Error.Code == 170194)
                {
                    while (true)
                    {
                        BybitSpotOrderBookEntry Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                        try
                        {
                            result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                                (
                                    Bybit.Net.Enums.Category.Spot,
                                    symbol: SellSymbol,
                                    side: Bybit.Net.Enums.OrderSide.Sell,
                                    type: Bybit.Net.Enums.NewOrderType.Limit,
                                    price: Bid.Price,
                                    quantity: quantity,
                                    timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                 );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"{ex.Message} стр 1349");
                            await Loger.WriteToFile($"{ex.Message} стр 1349");
                            Console.ReadLine();
                        }
                        if (result.Error == null)
                        {
                            while (true)
                            {
                                try
                                {
                                    resultOrderSell = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                (
                                    category: Bybit.Net.Enums.Category.Spot,
                                    symbol: SellSymbol,
                                    clientOrderId: result.Data.ClientOrderId
                                );
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"{ex.Message} стр 1368");
                                    await Loger.WriteToFile($"{ex.Message} стр 1368");
                                    Console.ReadLine();
                                }

                                if (resultOrderSell.Error == null)
                                {
                                    if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                                    {
                                        resltSell = true;
                                        break;
                                    }
                                    else if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                                    {
                                        resltSell = false;
                                        break;
                                    }
                                    await Task.Delay(1000); continue;
                                }
                                else if (resultOrderSell.Error.Code == 10002)
                                {
                                    await Task.Delay(2000);
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                                    await Loger.WriteToFile($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}  Клас Trader стр 1395");
                                    Console.WriteLine(" Клас Trader стр 1395");
                                    Console.ReadLine();
                                }
                            }
                        }
                        else if (result.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                            await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message}");
                            if (result.Error.Code == 170194)
                            {
                                continue;
                            }
                            Console.ReadLine();
                        }
                        break;
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltSell = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    await Loger.WriteToFile($" {result.Error.Code} {result.Error.Message} Клас Trader стр 1426");
                    Console.WriteLine(" Клас Trader стр 1426");
                    Console.ReadLine();
                }
                if (resltSell == true)
                {
                    ResultTrade.Sell++;
                }
            }
            catch
            {
                //resltSell = true;
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitOrder>> histori = null;
                while (histori == null)
                {
                    try
                    {
                        
                        histori = await bybitRestClient.V5Api.Trading.GetOrderHistoryAsync
                            (
                                Bybit.Net.Enums.Category.Spot,
                                SellSymbol,
                                limit:1
                            );
                        if (histori == null) { await Task.Delay(300); continue; }
                        foreach (var orderHistory in histori.Data.List)
                        {  
                            if (orderHistory.Side == Bybit.Net.Enums.OrderSide.Sell && orderHistory.Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                            {
                                resltSell = true;
                            }
                            else
                            {
                                resltSell = false;
                            }
                        }
                    }
                    catch { }
                }
            }
            return resltSell;
        }
    }
}
