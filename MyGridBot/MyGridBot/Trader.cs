﻿using Bybit.Net.Clients;
using Bybit.Net.Objects.Models.Spot;
using ClosedXML.Excel;
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
                Ask = await AskPriceQuantity(bybitRestClient, BuySymbol);
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{BuySymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            for (int i = 2; i <= 5001; i++)
                            {
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

                                            if (await BuyResult(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 11).Value), Ask.Price))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 8).Value = Convert.ToDecimal(sheet.Cell(i, 11).Value);
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
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

                                            if (await BuyResult(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 8).Value), Ask.Price))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                break;
                                            }

                                        }

                                        Thread.Sleep(100);
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

                        }
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {BuySymbol}.xlsx");
                        Thread.Sleep(10000);
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
                Bid = await BidPriceQuantity(bybitRestClient, SellSymbol);
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{SellSymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            for (int i = 5001; i >= 2; i--)
                            {
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 1)
                                {
                                    if (Bid.Price > Convert.ToDecimal(sheet.Cell(i, 3).Value) && Bid.Quantity > Convert.ToDecimal(sheet.Cell(i, 7).Value))
                                    {
                                        Console.WriteLine();
                                        Console.WriteLine($" Продажа Торговой Пары: {SellSymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 3).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 7).Value)}");
                                        if (await SellResult(bybitRestClient, SellSymbol, Convert.ToDecimal(sheet.Cell(i, 3).Value), Convert.ToDecimal(sheet.Cell(i, 7).Value), Bid.Price))
                                        {
                                            Console.WriteLine(" Заявка исполнилась");
                                            sheet.Cell(i, 4).Value = 0;
                                            save = true;
                                        }
                                        else
                                        {
                                            Console.WriteLine(" Заявка не исполнилась");
                                            break;
                                        }
                                        Thread.Sleep(100);
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
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {SellSymbol}.xlsx");
                        Thread.Sleep(10000);
                    }
                }

            }
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
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану AskPriceQuantity\n" +
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
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану BidPriceQuantity\n" +
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

        static async Task<bool> BuyResult(BybitRestClient bybitRestClient, string BuySymbol, decimal price, decimal quantity, decimal priceEror)
        {
            bool resltBuy = true;
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
                Console.WriteLine($"{ex.Message} стр 285");
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
                        Console.WriteLine($"{ex.Message} стр 299");
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

                    }
                    else if (resultOrderBuy.Error.Code == 10002)
                    {
                        await Task.Delay(2000);
                        continue;
                    }
                    else
                    {
                        Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                        Console.ReadLine();
                    }
                }

            }
            else if (result.Error.Code == 12193)
            {
                try
                {
                    result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
                                    (
                                        symbol: BuySymbol,
                                        side: Bybit.Net.Enums.OrderSide.Buy,
                                        type: Bybit.Net.Enums.OrderType.Limit,
                                        price: priceEror,
                                        quantity: quantity,
                                        timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                    );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} стр 346");
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
                            Console.WriteLine($"{ex.Message} стр 359");
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
                            await Task.Delay(1000);
                        }
                        else if (resultOrderBuy.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            Console.WriteLine(" Клас Trader стр 374");
                            Console.ReadLine();
                        }
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltBuy = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    Console.ReadLine();
                }
            }
            else if (result.Error.Code == 10002)
            {
                resltBuy = false;
            }
            else
            {
                Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                Console.WriteLine(" Клас Trader стр 407");
                Console.ReadLine();
            }
            if (resltBuy == true)
            {
                ResultTrade.Buy++;
            }
            return resltBuy;
        }
        static async Task<bool> SellResult(BybitRestClient bybitRestClient, string SellSymbol, decimal price, decimal quantity, decimal priceEror)
        {
            bool resltSell = true;
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
                Console.WriteLine($"{ex.Message} стр 435");
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
                        Console.WriteLine($"{ex.Message} стр 449");
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
                        await Task.Delay(1000);
                    }
                    if (resultOrderSell.Error.Code == 10002)
                    {
                        await Task.Delay(2000);
                        continue;
                    }
                    else
                    {
                        Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                        Console.WriteLine(" Клас Trader стр 475");
                        Console.ReadLine();
                    }
                }

            }
            else if (result.Error.Code == 12194)
            {
                try
                {
                    result = await bybitRestClient.SpotApiV3.Trading.PlaceOrderAsync
                                    (
                                        symbol: SellSymbol,
                                        side: Bybit.Net.Enums.OrderSide.Sell,
                                        type: Bybit.Net.Enums.OrderType.Limit,
                                        price: priceEror,
                                        quantity: quantity,
                                        timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                    );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} стр 497");
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
                            Console.WriteLine($"{ex.Message} стр 510");
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
                            await Task.Delay(1000);
                        }
                        else if(resultOrderSell.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                            Console.WriteLine(" Клас Trader стр 536");
                            Console.ReadLine();
                        }
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltSell = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    Console.ReadLine();
                }
            }
            else if (result.Error.Code == 10002)
            {
                resltSell = false;
            }
            else
            {
                Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                Console.WriteLine(" Клас Trader стр 558");
                Console.ReadLine();
            }
            if (resltSell == true)
            {
                ResultTrade.Sell++;
            }

            return resltSell;
        }
    }
}
