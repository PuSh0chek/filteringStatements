namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Data;
using System.Collections;

class Program
{
    static void Main()
    {
        // Регулярные выражения //
        // Первая проверка(стандарт) //
        string consractFirstOption = @"(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2,4})";
        
        // Вторая проверка для нахождения исключений в ошибочных элементах и переноса их в главный массив с корректными данными //
        string contractFirstOptionNoS = @"\d{1,4}-[в,к,В,К]{1,2}от\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionTO = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*то\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionOOF = @"№\s\d{1,4}-{1,2}[в,к,В,К]{1,2}\/[О,Ф]{3}-{1,2}\d{5,7}\/[С]-[К,к,А,а,В,в]{1,4}\s{1,2}от\s{1,2}\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionOOFnumberWithOutS = @"№\d{1,4}-{1,2}[в,к,В,К]{1,2}\/[О,Ф]{3}-{1,2}\d{5,7}\/[С]-[К,к,А,а,В,в]{1,4}\s{1,2}от\s{1,2}\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionSlash = @"№\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\/\d{1,3}\s*от\s*\d{1,2}.\d{1,2}.\d{1,4}";
        //string contractFirstOptionOnlyYeacr = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}г.";
        //string contractFirstOptionOnlyYeacrS = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}\sг.";
        string contractTwoOption = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*тот\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractTwoOptionOTI = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*оти\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractThreeOption = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFourOption = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{0,2}\s*\d{1,2}.\d{1,2}.\d{1,4}";
        //string contractFourOptionOnOneNumberWithFrom = @"№\s*\d{1}\s{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{1,4}}";
        //string contractFourOptionOnOneNumber = @"\s№\d\s\d{1,2}.\d{1,2}.\d{1,4}\s*";
        //string contractFourOptionOnOneNumberWithFromSOGL = @"Соглашение{0,1}\s*№{1}\s*\d{1}\s*\s*от{1}\s*\d{1,2}.\d{1,2}.\d{1,4}";
        
        // Вариант с регулярнаым выражением для проверки числа //
        string contractNumber = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}";
        
        // Вариант с регулярнаым выражением для проверки даты // 
        string dataStandart = @"\s\d{2}.\d{2}.\d{2,4}";
        string dataStandartNoS = @"от\d{2}.\d{2}.\d{2,4}";
        //string dataOnlyYear = @"\s\d{4}г{1}.{1}\s";
        string dataOnlyFullYeacrWithS = @"\s\d{2}.\d{2}.\d{2,4}\sг{1}.{1}\s";
        
        // Массивы //
        
        // Массив с вариантами регулярных выражений для второй проверки главного массива с данными //
        List<string> contract = new List<string>() { consractFirstOption, contractTwoOption, contractThreeOption, contractFourOption, contractTwoOptionOTI, contractFirstOptionSlash, contractFirstOptionTO, contractFirstOptionOOF, contractFirstOptionNoS, contractFirstOptionOOFnumberWithOutS };
        
        // Массив с вариантами регулярных функций для проверки даты //
        List<string> date = new List<string>() { dataStandart, dataStandartNoS, dataOnlyFullYeacrWithS };

        // Массив с данными удачно прошедшими фильтрацию //
        List<string> arrayOfFilteredData = new List<string>() {  };

        // Массив с данными НЕ удачно прошедшими фильтрацию //
        List<string> arrayOfErrors = new List<string>() { };

        // Массив с данными фильтрованных ошибок прошедшими фильтрацию //
        List<string> arrayOfFilteredErrors = new List<string>() { };

        // Массив с данными из БД //
        List<string> arrayBuhFromDataBase = new List<string>() { };

        try
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo("E:\\filteringStatements\\filteringStatements\\doc.xlsx")))
            {
                // Получаем доступ к рабочему листу в книге
                ExcelWorksheet worksheet = package.Workbook.Worksheets["TDSheet"];

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
                    string eptyDebet = (string)debet;
                    double eptyDebetSum = Convert.ToDouble(debetSum);
                    string eptyKredit = (string)kredit;
                    double eptyKreditSum = Convert.ToDouble(debetSum);
                    // Проверяем элементы на значение null //
                    if (eptyPeriod != null && eptyDt != null && eptyKt != null && eptyDebet != null && eptyDebetSum != null && eptyKredit != null && eptyKreditSum != null)
                    {
                        // Поиск через регулярные выражения //
                        Regex regex = new(consractFirstOption);
                        Match clippingOfKt = regex.Match(eptyKt);
                        if (clippingOfKt.Value == "" || clippingOfKt.Value == " " || clippingOfKt.Value == "   " || clippingOfKt.Value == null)
                        {
                            // Добавляем в массив элемент не прошедший фильтрацию //
                            arrayOfErrors.Add(eptyPeriod);
                            arrayOfErrors.Add(eptyDt);
                            arrayOfErrors.Add(eptyKt);
                            arrayOfErrors.Add(eptyDebet);
                            arrayOfErrors.Add(eptyDebetSum.ToString());
                            arrayOfErrors.Add(eptyKredit);
                            arrayOfErrors.Add(eptyKreditSum.ToString());
                        }
                        else
                        {
                            // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                            Regex numberSearcher = new(contractNumber);
                            Regex dateSearcher = new(dataStandart);
                            Match clippingDateOfEptyKt = dateSearcher.Match(clippingOfKt.ToString());
                            Match clippingNumberOfEptyKt = numberSearcher.Match(clippingOfKt.ToString());
                            if (clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == null || clippingDateOfEptyKt.Value == "" || clippingDateOfEptyKt.Value == " " || clippingDateOfEptyKt.Value == "   " || clippingDateOfEptyKt.Value == null)
                            {
                                // Добавляем в массив элемент не прошедший фильтрацию //
                                arrayOfErrors.Add(eptyPeriod);
                                arrayOfErrors.Add(eptyDt);
                                arrayOfErrors.Add(eptyKt);
                                arrayOfErrors.Add(eptyDebet);
                                arrayOfErrors.Add(eptyDebetSum.ToString());
                                arrayOfErrors.Add(eptyKredit);
                                arrayOfErrors.Add(eptyKreditSum.ToString());
                            }
                            else
                            {
                                // Заливаем элемент в объект с валидными данными //
                                arrayOfFilteredData.Add(clippingNumberOfEptyKt.Value);
                                if (clippingDateOfEptyKt.ToString().Trim().IndexOf(" ") == -1)
                                {
                                    arrayOfFilteredData.Add(clippingDateOfEptyKt.ToString().Trim());
                                }
                                else
                                {
                                    arrayOfFilteredData.Add(clippingDateOfEptyKt.ToString().Trim().Replace(" ", "."));
                                }
                                arrayOfFilteredData.Add(eptyDebet);
                                arrayOfFilteredData.Add(eptyKredit);
                                arrayOfFilteredData.Add(eptyDebetSum.ToString());
                                arrayOfFilteredData.Add(eptyPeriod);
                                arrayOfFilteredData.Add(eptyKt);
                            }
                        }
                    }
                    else
                    {
                        // Добавляем в массив элемент не прошедший фильтрацию //
                        arrayOfErrors.Add(eptyPeriod);
                        arrayOfErrors.Add(eptyDt);
                        arrayOfErrors.Add(eptyKt);
                        arrayOfErrors.Add(eptyDebet);
                        arrayOfErrors.Add(eptyDebetSum.ToString());
                        arrayOfErrors.Add(eptyKredit);
                        arrayOfErrors.Add(eptyKreditSum.ToString());
                    }

                }
                // Счетчик //
                int counter = 0;

                while (counter < 1)
                {
                    // Начало среза //
                    int skip = 0;

                    // Конец среза //
                    int take = 7;

                    // Проходимся циклом  //
                    for (int n = 0; n < arrayOfErrors.Count; n += 7)
                    {
                        if(arrayOfErrors.Count < 7) { break; };
                        // Срезаем ошибочный элемент с листа длиною 7 и проходимся по нему //
                        var list = arrayOfErrors.Skip(skip).Take(take).ToList();

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
                                    contract.ForEach((item) =>
                                    {
                                        // Фильтрация //
                                        Regex regex = new(item);
                                        Match clippingElement = regex.Match(list[2]);

                                        // Провекра clippingElement на удачную сотрировку //
                                        if (clippingElement.ToString() != "" && clippingElement.ToString() != " " && clippingElement != null)
                                        {
                                            // Провекра clippingElementFilterLevelTwo на удачную сотрировку //
                                            Match clippingElementFilterLevelTwo = regex.Match(list[2]);

                                            if (clippingElementFilterLevelTwo.ToString() != "" && clippingElementFilterLevelTwo.ToString() != " " && clippingElementFilterLevelTwo != null)
                                            {
                                                // Провекра clippingElementFilterLevelThree на удачную сотрировку //
                                                Match clippingElementFilterLevelThree = regex.Match(list[2]);

                                                if (clippingElementFilterLevelThree.ToString() != "" && clippingElementFilterLevelThree.ToString() != " " && clippingElementFilterLevelThree != null)
                                                {
                                                    // Раскладываем полученный элемент на два и фильтруем каждый отдельно //
                                                    Regex numberSearcher = new(contractNumber);
                                                    Regex dateSearcher = new(dataStandart);
                                                    Match clippingElementNumber = numberSearcher.Match(clippingElementFilterLevelThree.ToString());
                                                    Match clippingElementDate = dateSearcher.Match(clippingElementFilterLevelThree.ToString());

                                                    if (clippingElementNumber.ToString() != "" && clippingElementNumber.ToString() != " " && clippingElementNumber != null && clippingElementDate.ToString() != "" && clippingElementDate.ToString() != " " && clippingElementDate != null)
                                                    {
                                                        arrayOfFilteredData.Add(clippingElementNumber.Value);
                                                        arrayOfFilteredData.Add(clippingElementDate.Value);
                                                        arrayOfFilteredData.Add(list[3]);
                                                        arrayOfFilteredData.Add(list[5]);
                                                        arrayOfFilteredData.Add(list[6]);
                                                        arrayOfFilteredData.Add(list[0]);
                                                        arrayOfFilteredData.Add(list[2]);

                                                        // Удаление элемента прошедшего фильтрацию из массива ошибок //
                                                        arrayOfErrors.RemoveRange(n, 7);

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
                                    contract.ForEach((item) =>
                                    {
                                        // Фильтрация //
                                        Regex regex = new(item);
                                        Match clippingElement = regex.Match(list[1]);

                                        // Провекра clippingElement на удачную сотрировку //
                                        if (clippingElement.ToString() != "" && clippingElement.ToString() != " " && clippingElement != null)
                                        {
                                            // Провекра clippingElementFilterLevelTwo на удачную сотрировку //
                                            Match clippingElementFilterLevelTwo = regex.Match(list[1]);

                                            if (clippingElementFilterLevelTwo.ToString() != "" && clippingElementFilterLevelTwo.ToString() != " " && clippingElementFilterLevelTwo != null)
                                            {
                                                // Провекра clippingElementFilterLevelThree на удачную сотрировку //
                                                Match clippingElementFilterLevelThree = regex.Match(list[1]);

                                                if (clippingElementFilterLevelThree.ToString() != "" && clippingElementFilterLevelThree.ToString() != " " && clippingElementFilterLevelThree != null)
                                                {
                                                    // Раскладываем полученный элемент на два и фильтруем каждый отдельно //
                                                    Regex numberSearcher = new(contractNumber);
                                                    Regex dateSearcher = new(dataStandart);
                                                    Match clippingElementNumber = numberSearcher.Match(clippingElementFilterLevelThree.ToString());
                                                    Match clippingElementDate = dateSearcher.Match(clippingElementFilterLevelThree.ToString());
                                                    
                                                    if (clippingElementNumber.ToString() != "" && clippingElementNumber.ToString() != " " && clippingElementNumber != null && clippingElementDate.ToString() != "" && clippingElementDate.ToString() != " " && clippingElementDate != null)
                                                    {
                                                        arrayOfFilteredData.Add(clippingElementNumber.Value);
                                                        arrayOfFilteredData.Add(clippingElementDate.Value);
                                                        arrayOfFilteredData.Add(list[3]);
                                                        arrayOfFilteredData.Add(list[5]);
                                                        arrayOfFilteredData.Add("-" + list[6]);
                                                        arrayOfFilteredData.Add(list[0]);
                                                        arrayOfFilteredData.Add(list[1]);

                                                        // Удаление элемента прошедшего фильтрацию из массива ошибок //
                                                        arrayOfErrors.RemoveRange(n, 7);

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
                    counter++;
                    
                    // Обнуление индексов //
                    skip = 0;
                    take = 7;
                }

                // Данные для подключения к серверу //
                string connectionString = "Data Source= rvdk-svr-6091, 1500;Initial Catalog= TechConditions;Integrated Security=SSPI;";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    try
                    {
                        // Подключение //
                        connection.Open();
                        Console.WriteLine("Connection successfully opened");
                        string query = "SELECT * FROM buh;";
                        SqlCommand command = new SqlCommand(query, connection);
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            // Вытаскиваем таблицу в массив для проверки //
                            while (reader.Read())
                            {
                                // Обработка результатов запроса
                                string numberСontract = reader["dog"].ToString();
                                string dateСontract = reader["datadog"].ToString();
                                string dt = reader["dt"].ToString();
                                string kt = reader["kt"].ToString();
                                string summ = reader["summ"].ToString();
                                string periodDate = reader["datepl"].ToString();
                                string textDt = reader["text"].ToString();
                                arrayBuhFromDataBase.Add(numberСontract);
                                arrayBuhFromDataBase.Add(dateСontract);
                                arrayBuhFromDataBase.Add(dt);
                                arrayBuhFromDataBase.Add(kt);
                                arrayBuhFromDataBase.Add(summ);
                                arrayBuhFromDataBase.Add(periodDate);
                                arrayBuhFromDataBase.Add(textDt);
                            }
                        }
                        connection.Close();
                    }
                    catch (Exception er)
                    {
                        Console.WriteLine(er.Message);
                    }
                }
            }
        }

        catch (Exception er)
        {
            Console.WriteLine(er.Message);
        }

        finally
        {
            // Отчистка таблицы БД //
            //using (var comm = connection.CreateCommand())
            //{
            //    comm.CommandText = "TRUNCATE TABLE buh";
            //    comm.ExecuteNonQuery();
            //}
            for (int i = 0; i < 1; i++)
            {
                for(int k = 0; k < 1; k++)
                {
                    if(arrayBuhFromDataBase.Count > 7)
                    {
                        string pattern = @"\b0\:00\:00\b";
                        string connectionString = "Data Source= rvdk-svr-6091, 1500;Initial Catalog= TechConditions;Integrated Security=SSPI;";
                        string queryTableBuh = "INSERT INTO buh (dog, datadog, dt, kt, summ, datepl, text) VALUES (@dog, @datadog, @dt, @kt, @summ, @datepl, @text)";
                        string queryTableTrash = "INSERT INTO trash (dog, datadog, dt, kt, summ, datepl, text) VALUES (@dog, @datadog, @dt, @kt, @summ, @datepl, @text)";

                        Regex dateSearcher = new(pattern);
                        Match clippingElementDate = dateSearcher.Match(arrayBuhFromDataBase[5].ToString());
                        string qwer = Regex.Replace(arrayBuhFromDataBase[5], dateSearcher.ToString(), "");
                        Console.WriteLine(qwer);
                        Console.WriteLine(arrayOfFilteredData[5].ToString());
                        Console.WriteLine(arrayOfFilteredData[5].ToString() == qwer.Trim());
                        if (arrayOfFilteredData[5].ToString() == qwer.Trim())
                        {
                            Console.WriteLine($"Заменить элемент {arrayOfFilteredData[5]}, {arrayOfFilteredData[4]}, {arrayOfFilteredData[3]}, {arrayOfFilteredData[2]}, {arrayOfFilteredData[1]} | на | {qwer.Trim()}, {arrayBuhFromDataBase[4].ToString().Trim()}, {arrayBuhFromDataBase[3].ToString().Trim()}, {arrayBuhFromDataBase[2].ToString().Trim()}, {arrayBuhFromDataBase[1].ToString().Trim()}?");
                            Console.Write("Введите_да_(Заменит на новый),_нет_(Оставит старый),_залить_новый_(Заменит старый массив на новый):");
                            string? messageUser = Console.ReadLine();


                            if (messageUser == "да")
                            {
                                Console.WriteLine("Сценарий _да_");

                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    try
                                    {
                                        connection.Open();

                                        connection.Close();
                                    }
                                    catch (Exception er)
                                    {
                                        Console.WriteLine(er.Message);
                                    }
                                }
                            }
                            else if (messageUser == "нет")
                            {
                                Console.WriteLine("Сценарий _нет_");

                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    try
                                    {
                                        continue;
                                    }
                                    catch (Exception er)
                                    {
                                        Console.WriteLine(er.Message);
                                    }
                                }
                            }
                            else if (messageUser == "залить новый")
                            {
                                Console.WriteLine("Сценарий _залить_новый_");

                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    try
                                    {
                                        connection.Open();

                                        using (var comm = connection.CreateCommand())
                                        {
                                            comm.CommandText = "TRUNCATE TABLE buh";
                                            comm.ExecuteNonQuery();
                                        }

                                        using (var comm = connection.CreateCommand())
                                        {
                                            comm.CommandText = "TRUNCATE TABLE trash";
                                            comm.ExecuteNonQuery();
                                        }

                                        for (int f = 0; f < arrayOfFilteredData.Count; f += 7)
                                        {
                                            using (SqlCommand command = new SqlCommand(queryTableBuh, connection))
                                            {
                                                command.Parameters.AddWithValue("@dog", arrayOfFilteredData[f]);
                                                command.Parameters.AddWithValue("@datadog", arrayOfFilteredData[f + 1]);
                                                command.Parameters.AddWithValue("@dt", arrayOfFilteredData[f + 2]);
                                                command.Parameters.AddWithValue("@kt", arrayOfFilteredData[f + 3]);
                                                command.Parameters.AddWithValue("@summ", arrayOfFilteredData[f + 4]);
                                                command.Parameters.AddWithValue("@datepl", arrayOfFilteredData[f + 5]);
                                                command.Parameters.AddWithValue("@text", arrayOfFilteredData[f + 6]);
                                                int rowsAffected = command.ExecuteNonQuery();
                                            }
                                        }
                                        for (int f = 0; f < arrayOfErrors.Count; f += 7)
                                        {
                                            using (SqlCommand command = new SqlCommand(queryTableTrash, connection))
                                            {
                                                command.Parameters.AddWithValue("@dog", arrayOfErrors[f]);
                                                command.Parameters.AddWithValue("@datadog", arrayOfErrors[f + 1]);
                                                command.Parameters.AddWithValue("@dt", arrayOfErrors[f + 2]);
                                                command.Parameters.AddWithValue("@kt", arrayOfErrors[f + 3]);
                                                command.Parameters.AddWithValue("@summ", arrayOfErrors[f + 4]);
                                                command.Parameters.AddWithValue("@datepl", arrayOfErrors[f + 5]);
                                                command.Parameters.AddWithValue("@text", arrayOfErrors[f + 6]);
                                                int rowsAffected = command.ExecuteNonQuery();
                                            }
                                        }
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
                                Console.WriteLine($"Cценарий {messageUser} отсутствует");
                            }
                        }
                    } else
                    {
                        Console.WriteLine("Сценарий _таблица_пуста_");
                        // Данные для подключения к серверу //
                        string connectionString = "Data Source= rvdk-svr-6091, 1500;Initial Catalog= TechConditions;Integrated Security=SSPI;";

                        string queryTableBuh = "INSERT INTO buh (dog, datadog, dt, kt, summ, datepl, text) VALUES (@dog, @datadog, @dt, @kt, @summ, @datepl, @text)";
                        string queryTableTrash = "INSERT INTO trash (dog, datadog, dt, kt, summ, datepl, text) VALUES (@dog, @datadog, @dt, @kt, @summ, @datepl, @text)";

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            try
                            {
                                connection.Open();

                                for (int f = 0; f < arrayOfErrors.Count; f += 7)
                                {
                                    using (SqlCommand command = new SqlCommand(queryTableTrash, connection))
                                    {
                                        command.Parameters.AddWithValue("@dog", DBNull.Value);
                                        command.Parameters.AddWithValue("@datadog", DBNull.Value);
                                        command.Parameters.AddWithValue("@dt", arrayOfErrors[f + 3]);
                                        command.Parameters.AddWithValue("@kt", arrayOfErrors[f + 5]);
                                        command.Parameters.AddWithValue("@summ", arrayOfErrors[f + 4]);
                                        command.Parameters.AddWithValue("@datepl", arrayOfErrors[f]);
                                        if(arrayOfErrors[f + 3] == "62.02" || arrayOfErrors[f + 3] == "62.01")
                                        {
                                            command.Parameters.AddWithValue("@text", arrayOfErrors[f + 1]);
                                        } else
                                        {
                                            command.Parameters.AddWithValue("@text", arrayOfErrors[f + 2]);
                                        }
                                        int rowsAffected = command.ExecuteNonQuery();
                                    }
                                }

                                for (int f = 0; f < arrayOfFilteredData.Count; f+=7)
                                {
                                    using (SqlCommand command = new SqlCommand(queryTableBuh, connection))
                                    {
                                        command.Parameters.AddWithValue("@dog", arrayOfFilteredData[f]);
                                        command.Parameters.AddWithValue("@datadog", arrayOfFilteredData[f + 1]);
                                        command.Parameters.AddWithValue("@dt", arrayOfFilteredData[f + 2]);
                                        command.Parameters.AddWithValue("@kt", arrayOfFilteredData[f + 3]);
                                        command.Parameters.AddWithValue("@summ", arrayOfFilteredData[f + 4]);
                                        command.Parameters.AddWithValue("@datepl", arrayOfFilteredData[f + 5]);
                                        command.Parameters.AddWithValue("@text", arrayOfFilteredData[f + 6]);
                                        int rowsAffected = command.ExecuteNonQuery();
                                    }
                                }
                                connection.Close();
                            }
                            catch (Exception er)
                            {
                                Console.WriteLine(er.Message);
                            }
                        }
                    }
                    
                }
            }
            Console.WriteLine("Done");
            Console.WriteLine(arrayOfFilteredData.Count / 7);
            Console.WriteLine(arrayOfErrors.Count / 7);
            Console.WriteLine(arrayBuhFromDataBase.Count / 7);
        }
    }
}