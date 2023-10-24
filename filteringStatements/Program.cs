namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Data.SqlClient;
using filteringStatements;
using GraphQL.Validation;
using Microsoft.VisualBasic;

class Program
{
    static void Main()
    {
        // Подключение методов с других файлов //
        filteringStatements.Variables variables = new filteringStatements.Variables();
        filteringStatements.Repository repository = new Repository();
        // Вырезание элементов готовых массивов для поиска одинаковых элеметов //
        // Локальные переменные для работы с массивами //
        int startIterationElementDB = 0;
        int startIterationElementDocument = 0;
        int limitOfIterationElementDB = 7;
        int limitOfIterationElementDocument = 7;

        // Счетчик пересечения дат в массивх //
        int intersection = 0;

        // Переменная регистрирующая ответ пользователя в консоли //
        string? messageUser = null;

        // Главный блок кода с детьми //
        try
        {
            // Блок работы с файлом //
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filteringStatements.Variables.path)))
            {
                // Получаем доступ к рабочему листу в книге
                ExcelWorksheet worksheet = package.Workbook.Worksheets["TDSheet"];
                worksheet.View.ShowHeaders = true;
                // Фильтруем столбецы в xslx файле и заливаем данные в объект //
                for (int i = 3; !string.IsNullOrEmpty(worksheet.Cells[i, 4].Text); i++)
                {
                    // Объявляем переменные //
                    object period;
                    object dt;
                    object kt;
                    object debet;
                    object debetSum;
                    object kredit;
                    object kreditSum;
                    // Вытаскиваем переменные из объекта //
                    period = worksheet.Cells[i, 1].Value;
                    dt = worksheet.Cells[i, 3].Value;
                    kt = worksheet.Cells[i, 4].Value;
                    debet = worksheet.Cells[i, 5].Value;
                    debetSum = worksheet.Cells[i, 6].Value;
                    kredit = worksheet.Cells[i, 8].Value;
                    kreditSum = worksheet.Cells[i, 9].Value;
                    // Трансформируем данные //
                    string eptyPeriod = (string)period;
                    string eptyDt = (string)dt;
                    string eptyKt = (string)kt;
                    //Console.WriteLine(eptyKt);
                    string eptyDebet = (string)debet;
                    double eptyDebetSum = Convert.ToDouble(debetSum);
                    string eptyKredit = (string)kredit;
                    double eptyKreditSum = Convert.ToDouble(debetSum);
                    // Проверяем элементы на значение null //
                    if (eptyPeriod != null && eptyDt != null && eptyKt != null && eptyDebet != null && eptyDebetSum != null && eptyKredit != null && eptyKreditSum != null)
                    {
                        // Поиск через регулярные выражения //
                        Regex regex = new(filteringStatements.Variables.consractFirstOption);
                        Match clippingOfKt = regex.Match(eptyKt);
                        if (clippingOfKt.Value == "" || clippingOfKt.Value == " " || clippingOfKt.Value == "   " || clippingOfKt.Value == null)
                        {
                            // Добавляем в массив элемент не прошедший фильтрацию //
                            filteringStatements.Variables.arrayOfErrors.Add(eptyPeriod);
                            filteringStatements.Variables.arrayOfErrors.Add(eptyDt);
                            filteringStatements.Variables.arrayOfErrors.Add(eptyKt);
                            filteringStatements.Variables.arrayOfErrors.Add(eptyDebet);
                            if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01")
                            {
                                filteringStatements.Variables.arrayOfErrors.Add("-" + eptyDebetSum.ToString());
                            }
                            else
                            {
                                filteringStatements.Variables.arrayOfErrors.Add(eptyDebetSum.ToString());
                            }
                            filteringStatements.Variables.arrayOfErrors.Add(eptyKredit);
                            filteringStatements.Variables.arrayOfErrors.Add(eptyKreditSum.ToString());
                        }
                        else
                        {
                            // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                            Regex numberSearcher = new(filteringStatements.Variables.contractNumber);
                            Regex dateSearcher = new(filteringStatements.Variables.dataStandart);
                            Match clippingDateOfEptyKt = dateSearcher.Match(clippingOfKt.ToString());
                            Match clippingNumberOfEptyKt = numberSearcher.Match(clippingOfKt.ToString());
                            if (clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == null || clippingDateOfEptyKt.Value == "" || clippingDateOfEptyKt.Value == " " || clippingDateOfEptyKt.Value == "   " || clippingDateOfEptyKt.Value == null)
                            {
                                // Добавляем в массив элемент не прошедший фильтрацию //
                                filteringStatements.Variables.arrayOfErrors.Add(eptyPeriod);
                                filteringStatements.Variables.arrayOfErrors.Add(eptyDt);
                                filteringStatements.Variables.arrayOfErrors.Add(eptyKt);
                                filteringStatements.Variables.arrayOfErrors.Add(eptyDebet);
                                if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01")
                                {
                                    filteringStatements.Variables.arrayOfErrors.Add("-" + eptyDebetSum.ToString());
                                }
                                else
                                {
                                    filteringStatements.Variables.arrayOfErrors.Add(eptyDebetSum.ToString());
                                }
                                filteringStatements.Variables.arrayOfErrors.Add(eptyKredit);
                                filteringStatements.Variables.arrayOfErrors.Add(eptyKreditSum.ToString());
                            }
                            else
                            {
                                // Заливаем элемент в объект с валидными данными //
                                filteringStatements.Variables.arrayOfFilteredData.Add(clippingNumberOfEptyKt.Value);
                                if (clippingDateOfEptyKt.ToString().Trim().IndexOf(" ") == -1)
                                {
                                    filteringStatements.Variables.arrayOfFilteredData.Add(clippingDateOfEptyKt.ToString().Trim());
                                }
                                else
                                {
                                    filteringStatements.Variables.arrayOfFilteredData.Add(clippingDateOfEptyKt.ToString().Trim().Replace(" ", "."));
                                }
                                filteringStatements.Variables.arrayOfFilteredData.Add(eptyDebet);
                                filteringStatements.Variables.arrayOfFilteredData.Add(eptyKredit);
                                filteringStatements.Variables.arrayOfFilteredData.Add(eptyDebetSum.ToString());
                                filteringStatements.Variables.arrayOfFilteredData.Add(eptyPeriod);
                                filteringStatements.Variables.arrayOfFilteredData.Add(eptyKt);
                            }
                        }
                    }
                    else
                    {
                        // Добавляем в массив элемент не прошедший фильтрацию //
                        filteringStatements.Variables.arrayOfErrors.Add(eptyPeriod);
                        filteringStatements.Variables.arrayOfErrors.Add(eptyDt);
                        filteringStatements.Variables.arrayOfErrors.Add(eptyKt);
                        filteringStatements.Variables.arrayOfErrors.Add(eptyDebet);
                        if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01")
                        {
                            filteringStatements.Variables.arrayOfErrors.Add("-" + eptyDebetSum.ToString());
                        }
                        else
                        {
                            filteringStatements.Variables.arrayOfErrors.Add(eptyDebetSum.ToString());
                        }
                        filteringStatements.Variables.arrayOfErrors.Add(eptyKredit);
                        filteringStatements.Variables.arrayOfErrors.Add(eptyKreditSum.ToString());
                    }

                }
                // Цикл для прохода по массиву ошибок //
                while (variables.counter < variables.move)
                {
                    // Начало среза //
                    int skip = 0;

                    // Конец среза //
                    int take = 7;

                    // Проходимся циклом  //
                    for (int n = 0; n < filteringStatements.Variables.arrayOfErrors.Count; n += 7)
                    {
                        if (filteringStatements.Variables.arrayOfErrors.Count < 7) { break; };
                        // Срезаем ошибочный элемент с листа длиною 7 и проходимся по нему //
                        var list = filteringStatements.Variables.arrayOfErrors.Skip(skip).Take(take).ToList();

                        // Проходимся цмклом по-вырезанному элементу //
                        for (int i = 0; i < list.Count; i++)
                        {
                            // Задаем условия оплаты или приема средств //
                            if (list[5] == "62.02" || list[5] == "62.01")
                            {
                                // Выбираем i элемент //
                                if (i == 2)
                                {
                                    // contract - массив условий //
                                    filteringStatements.Variables.contract.ForEach((item) =>
                                    {
                                        // Фильтрация //
                                        Regex regex = new(item);
                                        Match clippingElement = regex.Match(list[2]);

                                        // Провекра clippingElement на удачную сортировку //
                                        if (clippingElement.ToString() != "" && clippingElement.ToString() != " " && clippingElement != null)
                                        {
                                            // Провекра clippingElementFilterLevelTwo на удачную сортировку //
                                            Match clippingElementFilterLevelTwo = regex.Match(list[2]);

                                            if (clippingElementFilterLevelTwo.ToString() != "" && clippingElementFilterLevelTwo.ToString() != " " && clippingElementFilterLevelTwo != null)
                                            {
                                                // Провекра clippingElementFilterLevelThree на удачную сортировку //
                                                Match clippingElementFilterLevelThree = regex.Match(list[2]);

                                                if (clippingElementFilterLevelThree.ToString() != "" && clippingElementFilterLevelThree.ToString() != " " && clippingElementFilterLevelThree != null)
                                                {
                                                    // Раскладываем полученный элемент на два и фильтруем каждый отдельно //
                                                    Regex numberSearcher = new(filteringStatements.Variables.contractNumber);
                                                    Regex dateSearcher = new(filteringStatements.Variables.dataStandart);
                                                    Match clippingElementNumber = numberSearcher.Match(clippingElementFilterLevelThree.ToString());
                                                    Match clippingElementDate = dateSearcher.Match(clippingElementFilterLevelThree.ToString());

                                                    if (clippingElementNumber.ToString() != "" && clippingElementNumber.ToString() != " " && clippingElementNumber != null && clippingElementDate.ToString() != "" && clippingElementDate.ToString() != " " && clippingElementDate != null)
                                                    {
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(clippingElementNumber.Value);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(clippingElementDate.Value);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[3]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[5]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[6]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[0]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[2]);

                                                        // Удаление элемента прошедшего фильтрацию из массива ошибок //
                                                        filteringStatements.Variables.arrayOfErrors.RemoveRange(n, 7);

                                                        // Смещение индекса назад на длину элемента //
                                                        n -= 7;
                                                    }
                                                }
                                            }
                                        }
                                    });
                                }
                            }
                            else
                            {
                                // Выбираем i элемент //
                                if (i == 1)
                                {
                                    // contract - массив условий //
                                    filteringStatements.Variables.contract.ForEach((item) =>
                                    {
                                        // Фильтрация //
                                        Regex regex = new(item);
                                        Match clippingElement = regex.Match(list[1]);

                                        // Провекра clippingElement на удачную сортировку //
                                        if (clippingElement.ToString() != "" && clippingElement.ToString() != " " && clippingElement != null)
                                        {
                                            // Провекра clippingElementFilterLevelTwo на удачную сортировку //
                                            Match clippingElementFilterLevelTwo = regex.Match(list[1]);

                                            if (clippingElementFilterLevelTwo.ToString() != "" && clippingElementFilterLevelTwo.ToString() != " " && clippingElementFilterLevelTwo != null)
                                            {
                                                // Провекра clippingElementFilterLevelThree на удачную сортировку //
                                                Match clippingElementFilterLevelThree = regex.Match(list[1]);

                                                if (clippingElementFilterLevelThree.ToString() != "" && clippingElementFilterLevelThree.ToString() != " " && clippingElementFilterLevelThree != null)
                                                {
                                                    // Раскладываем полученный элемент на два и фильтруем каждый отдельно //
                                                    Regex numberSearcher = new(filteringStatements.Variables.contractNumber);
                                                    Regex dateSearcher = new(filteringStatements.Variables.dataStandart);
                                                    Match clippingElementNumber = numberSearcher.Match(clippingElementFilterLevelThree.ToString());
                                                    Match clippingElementDate = dateSearcher.Match(clippingElementFilterLevelThree.ToString());

                                                    if (clippingElementNumber.ToString() != "" && clippingElementNumber.ToString() != " " && clippingElementNumber != null && clippingElementDate.ToString() != "" && clippingElementDate.ToString() != " " && clippingElementDate != null)
                                                    {
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(clippingElementNumber.Value);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(clippingElementDate.Value);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[3]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[5]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add("-" + list[6]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[0]);
                                                        filteringStatements.Variables.arrayOfFilteredData.Add(list[1]);

                                                        // Удаление элемента прошедшего фильтрацию из массива ошибок //
                                                        filteringStatements.Variables.arrayOfErrors.RemoveRange(n, 7);

                                                        // Смещение индекса назад на длину элемента //
                                                        n -= 7;
                                                    }
                                                }
                                            }
                                        }
                                    });
                                }
                            }
                        }
                        // Обнуляем массив //
                        list = new List<string> { };

                        // Сдвигаем интервал для вырезания следующего элемента //
                        skip += 7;
                        take += 7;
                    }
                    // Увеличеваем шаг счетчика //
                    variables.counter++;

                    // Обнуление индексов //
                    skip = 0;
                    take = 7;
                }

                Repository.ContentFromTableBuh();
            }
        }
        // Код работающий при ошибке //
        catch (Exception er)
        {
            Console.WriteLine(er.Message);
        }
        // Код работатющий в любом случше //
        finally
        {
            // Устанавливаем количество элементов с базы //
            if (filteringStatements.Variables.arrayBuhFromDataBase.Count > 7000000)
            {
                // Проход по массиву фильтрованных элементов //
                for (int i = 0; i < 1; i++)
                {
                    // Проход по старому массиву с сервера //
                    for (int k = 0; k < 1; k++)
                    {
                        // Стадия выполнения сценариев //
                        // Проверка на пересечение дат //
                        for (int j = 1; j < filteringStatements.Variables.arrayBuhFromDataBase.Count; j += 7)
                        {
                            for (int n = 1; n < filteringStatements.Variables.arrayOfFilteredData.Count; n += 7)
                            {
                                if (Convert.ToDateTime(filteringStatements.Variables.arrayBuhFromDataBase[j]).ToString() == Convert.ToDateTime(filteringStatements.Variables.arrayOfFilteredData[n]).ToString())
                                {
                                    // Фиксируем пересечение дат //
                                    intersection++;
                                    // Вывод информации для пользователя //
                                    Console.WriteLine($"Пересечение дат нового элемента: {filteringStatements.Variables.cutOutDataFromNewArray[i]}, {filteringStatements.Variables.cutOutDataFromNewArray[i + 1]}, {filteringStatements.Variables.cutOutDataFromNewArray[i + 2]}, {filteringStatements.Variables.cutOutDataFromNewArray[i + 3]}, {filteringStatements.Variables.cutOutDataFromNewArray[i + 4]} | и старого элемента: | {filteringStatements.Variables.cutOutDataFromOldArray[k].ToString().Trim()}, {filteringStatements.Variables.cutOutDataFromOldArray[k + 1].ToString().Trim()}, {filteringStatements.Variables.cutOutDataFromOldArray[k + 2].ToString().Trim()}, {filteringStatements.Variables.cutOutDataFromOldArray[k + 3].ToString().Trim()}, {filteringStatements.Variables.cutOutDataFromOldArray[k + 4].ToString().Trim()}");
                                    // Условия выбора для пользователя //
                                    Console.Write("Введите_да_(Заменить все пересечения на новые элементы),_нет_(Оставить старые элементы):");
                                    // Получение ответа от пользователя //
                                    messageUser = Console.ReadLine();
                                    // Сценарий ДА //
                                    if (messageUser == "да")
                                    {
                                        Console.WriteLine("Сценарий _да_");
                                        break;
                                        //if (Convert.ToDateTime(filteringStatements.Variables.arrayOfFilteredData[i]).ToString() == filteringStatements.Variables.arrayBuhFromDataBase[k])
                                        //{
                                        //    Console.WriteLine("Стадия 01...");
                                        //    for (int dlt = 0; dlt < 7; dlt++)
                                        //    {
                                        //        filteringStatements.Variables.arrayBuhFromDataBase.RemoveRange(dlt, 7);
                                        //    }
                                            //using (SqlConnection connection = new SqlConnection(connectionString: filteringStatements.Variables.connectionString))
                                            //{
                                            //    try
                                            //    {
                                            //        Console.WriteLine("Подключение к серверу");
                                            //        connection.Open();
                                            //        Console.WriteLine("Стадия 1...");
                                            //        using (SqlCommand command = new SqlCommand("TRUNCATE TABLE buh", connection))
                                            //        {
                                            //            command.ExecuteNonQuery();
                                            //        }
                                            //        Console.WriteLine("Стадия 2...");
                                            //        // Загрузка данных в главную таблицу //
                                            //        Repository.AccessingAtServerForTableBuh(filteringStatements.Variables.queryTableBuh, connection, filteringStatements.Variables.arrayBuhFromDataBase);
                                            //        Console.WriteLine("Стадия 3...");
                                            //        // Загрузка данных в главную таблицу //
                                            //        Repository.AccessingAtServerWithFilteredData(filteringStatements.Variables.queryTableBuh, connection, filteringStatements.Variables.arrayOfFilteredData);
                                            //        Console.WriteLine("Отключение от сервера");
                                            //        connection.Close();
                                            //    }
                                            //    catch (Exception er)
                                            //    {
                                            //        Console.WriteLine(er.Message);
                                            //    }
                                            //}
                                        //}
                                    }
                                    // Сценарий НЕТ //
                                    else if (messageUser == "нет")
                                    {
                                        Console.WriteLine("Сценарий _нет_");
                                        break;
                                        //Console.WriteLine("Стадия 0...");
                                        //if (Convert.ToDateTime(filteringStatements.Variables.arrayOfFilteredData[5]).ToString() == filteringStatements.Variables.arrayBuhFromDataBase[5])
                                        //{
                                        //    for (int dlt = 0; dlt < 7; dlt++)
                                        //    {
                                        //        filteringStatements.Variables.arrayOfFilteredData.RemoveRange(dlt, 7);
                                        //    }
                                        //    using (SqlConnection connection = new SqlConnection(filteringStatements.Variables.connectionString))
                                        //    {
                                        //        try
                                        //        {
                                        //            Console.WriteLine("Подключение к серверу");
                                        //            connection.Open();
                                        //            Console.WriteLine("Стадия 1...");
                                        //            // Загрузка данных в главную таблицу //
                                        //            Repository.AccessingAtServerWithFilteredData(filteringStatements.Variables.queryTableBuh, connection, filteringStatements.Variables.arrayOfFilteredData);
                                        //            Console.WriteLine("Отключение от сервера");
                                        //            connection.Close();
                                        //        }
                                        //        catch (Exception er)
                                        //        {
                                        //            Console.WriteLine(er.Message);
                                        //        }
                                        //    }
                                        //}
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Cценарий {messageUser} отсутствует");
                                        break;
                                    }
                                }
                            }
                        }
                        if (messageUser == "да")
                        {
                            if (filteringStatements.Variables.arrayOfFilteredData.Count < filteringStatements.Variables.arrayBuhFromDataBase.Count)
                            {
                                for (int adb = 0; adb < filteringStatements.Variables.arrayBuhFromDataBase.Count; adb += 7)
                                {
                                    for (int ad = 0; adb < filteringStatements.Variables.arrayOfFilteredData.Count; ad += 7)
                                    {
                                        // Обработка массива из БД //
                                        for (int iterationOfElementDB = startIterationElementDB; iterationOfElementDB < limitOfIterationElementDB; iterationOfElementDB++)
                                        {
                                            // Заливаем элемент в массив //
                                            filteringStatements.Variables.cutOutDataFromOldArray.Add(filteringStatements.Variables.arrayBuhFromDataBase[iterationOfElementDB]);
                                            if (iterationOfElementDB == limitOfIterationElementDB)
                                            {
                                                // Вырезаем элемент из родительского массива //
                                                filteringStatements.Variables.cutOutDataFromOldArray.RemoveRange(startIterationElementDB, limitOfIterationElementDB);
                                            }
                                        }

                                        // Обработка массива из файла //
                                        for (int iterationOfElementDocument = startIterationElementDocument; iterationOfElementDocument < limitOfIterationElementDocument; iterationOfElementDocument++)
                                        {
                                            // Заливаем элемент в массив //
                                            filteringStatements.Variables.cutOutDataFromNewArray.Add(filteringStatements.Variables.arrayOfFilteredData[iterationOfElementDocument]);
                                            if (iterationOfElementDocument == limitOfIterationElementDocument)
                                            {
                                                // Вырезаем элемент из родительского массива //
                                                filteringStatements.Variables.cutOutDataFromNewArray.RemoveRange(startIterationElementDocument, limitOfIterationElementDocument);
                                            }
                                        }
                                        // Сравниваем элементы массивов //
                                        // Если они совпадают о дате, то я заливаю новый элемент вресто старого после чего отчищаю их //
                                        // Если они не совпадают, то я заливаю их в новый массив, отчищаю два массива для обрезки и после всей проверки в самом конце заливаю в главный массив //
                                        
                                    }
                                }
                            }
                            else
                            {
                                for (int el = 1; el < filteringStatements.Variables.arrayOfFilteredData.Count; el += 7)
                                {
                                    if (Convert.ToDateTime(filteringStatements.Variables.cutOutDataFromNewArray[1]).ToString() == Convert.ToDateTime(filteringStatements.Variables.cutOutDataFromOldArray[1]).ToString())
                                    {

                                    }
                                }
                            }
                        }
                        else if (messageUser == "нет")
                        {

                        }
                        else
                        {
                            Console.WriteLine("None");
                        }
                    }
                }
            }
            if (intersection == 0)
            {
                // Пересечения дат нет //
                Console.WriteLine("Пересечения дат нет");
                using (SqlConnection connection = new SqlConnection(filteringStatements.Variables.connectionString))
                {
                    try
                    {
                        Console.WriteLine("Подключение к серверу");
                        connection.Open();
                        Console.WriteLine("Стадия 1...");
                        // Загрузка данных в таблицу с ошибками //
                        Repository.AccessingServerOfTableTrash(filteringStatements.Variables.queryTableTrash, connection, filteringStatements.Variables.arrayOfErrors);
                        Console.WriteLine("Стадия 2...");
                        // Загрузка данных в главную таблицу //
                        Repository.AccessingAtServerWithFilteredData(filteringStatements.Variables.queryTableBuh, connection, filteringStatements.Variables.arrayOfFilteredData);
                        Console.WriteLine("Отключение от сервера");
                        connection.Close();
                    }
                    catch (Exception er)
                    {
                        Console.WriteLine(er.Message);
                    }
                }
            }
            else
            {
                // Таблица пуста //
                Console.WriteLine("Сценарий _таблица_пуста_");

                using (SqlConnection connection = new SqlConnection(filteringStatements.Variables.connectionString))
                {
                    try
                    {
                        Console.WriteLine("Подключение к серверу");
                        connection.Open();
                        Console.WriteLine("Стадия 1...");
                        // Загрузка данных в таблицу с ошибками //
                        Repository.AccessingServerOfTableTrash(filteringStatements.Variables.queryTableTrash, connection, filteringStatements.Variables.arrayOfErrors);
                        Console.WriteLine("Стадия 2...");
                        // Загрузка данных в главную таблицу //
                        Repository.AccessingAtServerWithFilteredData(filteringStatements.Variables.queryTableBuh, connection, filteringStatements.Variables.arrayOfFilteredData);
                        Console.WriteLine("Отключение от сервера");
                        connection.Close();
                    }
                    catch (Exception er)
                    {
                        Console.WriteLine(er.Message);
                    }
                }
            }
            Console.WriteLine("Done");
            Console.WriteLine(filteringStatements.Variables.arrayOfFilteredData.Count / 7);
            Console.WriteLine(filteringStatements.Variables.arrayOfErrors.Count / 7);
            Console.WriteLine(filteringStatements.Variables.arrayBuhFromDataBase.Count / 7);
        }
    }
}