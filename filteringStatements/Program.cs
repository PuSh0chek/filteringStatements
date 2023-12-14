namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Data.SqlClient;
using filteringStatements;

class Program
{
    public static void Main()
    {
        // Подключение методов с других файлов //
        filteringStatements.Variables variables = new filteringStatements.Variables();
        filteringStatements.Repository repository = new Repository();

        // Счетчик пересечения дат в массивх //
        int intersection = 0;

        // Переменная регистрирующая ответ пользователя в консоли //
        string? messageUser = null;

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        try
        {
            // Блок работы с файлом //
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filteringStatements.Variables.path)))
            {
                // Получаем доступ к рабочему листу в книге
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
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
                    string eptyPeriod = Convert.ToDateTime(period).ToString();
                    string eptyDt = (string)dt;
                    string eptyKt = (string)kt;
                    string eptyDebet = (string)debet;
                    double eptyDebetSum = Convert.ToDouble(debetSum);
                    string eptyKredit = (string)kredit;
                    double eptyKreditSum = Convert.ToDouble(debetSum);

                    // Поиск через регулярные выражения //
                    Regex regex = new(filteringStatements.Variables.consractFirstOption);
                    Regex numberSearcher = new(filteringStatements.Variables.contractNumber);
                    Regex dateSearcher = new(filteringStatements.Variables.dataStandart);
                    Match Kt = regex.Match(eptyKt);
                    Match Dt = regex.Match(eptyDt);

                    if (eptyDebet.ToString().Contains("62") && eptyKredit.ToString().Contains("62"))
                    {
                        if (Kt.Value == "" || Kt.Value == " " || Kt.Value == "   " || Kt.Value == null || Dt.Value == "" || Dt.Value == " " || Dt.Value == "   " || Dt.Value == null)
                        {
                            if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01" && eptyKredit.ToString() == "62.02" || eptyKredit.ToString() == "62.01")
                            {
                                // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                                Match clippingOfDt = regex.Match(eptyDt);
                                Match clippingOfKt = regex.Match(eptyKt);
                                Match clippingDateOfEptyDt = dateSearcher.Match(clippingOfDt.ToString());
                                Match clippingNumberOfEptyDt = numberSearcher.Match(clippingOfDt.ToString());
                                Match clippingDateOfEptyKt = dateSearcher.Match(clippingOfKt.ToString());
                                Match clippingNumberOfEptyKt = numberSearcher.Match(clippingOfKt.ToString());

                                if (clippingNumberOfEptyKt.Value == null || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyDt.Value == "" || clippingNumberOfEptyDt.Value == " " || clippingNumberOfEptyDt.Value == "   " || clippingNumberOfEptyDt.Value == null || clippingDateOfEptyDt.Value == "" || clippingDateOfEptyDt.Value == " " || clippingDateOfEptyDt.Value == "   " || clippingDateOfEptyDt.Value == null)
                                {
                                    if(clippingNumberOfEptyKt != null || clippingNumberOfEptyDt != null)
                                    {
                                        if (clippingNumberOfEptyKt.Value != "" && clippingNumberOfEptyDt.Value == "")
                                        {
                                            // Добавляем в массив (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                            WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                            // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                            WorkThisContent.LoadElementInArrayOfContent(filteringStatements.Variables.number, eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                        }
                                        else if (clippingNumberOfEptyKt.Value == "" && clippingNumberOfEptyDt.Value != "")
                                        {
                                            // Добавляем в массив (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                            WorkThisContent.LoadElementInArrayOfErrors(filteringStatements.Variables.number, eptyPeriod, eptyKt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                            // Заливаем элемент в объект (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                            WorkThisContent.LoadElementInArrayOfContent(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, clippingNumberOfEptyDt, clippingDateOfEptyDt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                        }
                                        else
                                        {
                                            // Добавляем в массив (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                            WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                            // Добавляем в массив (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                            WorkThisContent.LoadElementInArrayOfErrors(filteringStatements.Variables.number, eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                        }
                                    }
                                    else if (clippingNumberOfEptyKt.Value != "" && clippingNumberOfEptyDt.Value == "")
                                    {
                                        // Добавляем в массив (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                        WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                        // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                        WorkThisContent.LoadElementInArrayOfContent(filteringStatements.Variables.number, eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                    }
                                    else if (clippingNumberOfEptyKt.Value != "" && clippingNumberOfEptyDt.Value == "")
                                    {
                                        // Добавляем в массив (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                        WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                        // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                        WorkThisContent.LoadElementInArrayOfContent(filteringStatements.Variables.number, eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                    }
                                    else
                                    {
                                        // Добавляем в массив (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                        WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                        // Добавляем в массив (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) элемент не прошедший фильтрацию //
                                        WorkThisContent.LoadElementInArrayOfErrors(filteringStatements.Variables.number, eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                    }
                                }
                                else
                                {
                                    // Заливаем элемент в объект (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                    WorkThisContent.LoadElementInArrayOfContent(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, clippingNumberOfEptyDt, clippingDateOfEptyDt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                    // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                    WorkThisContent.LoadElementInArrayOfContent(filteringStatements.Variables.number, eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                            }
                            else
                            {
                                // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                                Match clippingOfDt = regex.Match(eptyDt);
                                Match clippingOfKt = regex.Match(eptyKt);
                                Match clippingDateOfEptyDt = dateSearcher.Match(clippingOfDt.ToString());
                                Match clippingNumberOfEptyDt = numberSearcher.Match(clippingOfDt.ToString());
                                Match clippingDateOfEptyKt = dateSearcher.Match(clippingOfKt.ToString());
                                Match clippingNumberOfEptyKt = numberSearcher.Match(clippingOfKt.ToString());

                                if (clippingNumberOfEptyKt.Value == null || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyDt.Value == "" || clippingNumberOfEptyDt.Value == " " || clippingNumberOfEptyDt.Value == "   " || clippingNumberOfEptyDt.Value == null || clippingDateOfEptyDt.Value == "" || clippingDateOfEptyDt.Value == " " || clippingDateOfEptyDt.Value == "   " || clippingDateOfEptyDt.Value == null)
                                {
                                    // Добавляем в массив элемент (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) не прошедший фильтрацию //
                                    WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                    // Добавляем в массив элемент (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) не прошедший фильтрацию //
                                    WorkThisContent.LoadElementInArrayOfErrors(filteringStatements.Variables.number, eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                                else
                                {
                                    // Заливаем элемент в объект (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                    WorkThisContent.LoadElementInArrayOfContent(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, clippingNumberOfEptyDt, clippingDateOfEptyDt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                    // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                    WorkThisContent.LoadElementInArrayOfContent(filteringStatements.Variables.number, eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                            }
                        }
                        else
                        {
                            // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                            Match clippingOfDt = regex.Match(eptyDt);
                            Match clippingOfKt = regex.Match(eptyKt);
                            Match clippingDateOfEptyDt = dateSearcher.Match(clippingOfDt.ToString());
                            Match clippingNumberOfEptyDt = numberSearcher.Match(clippingOfDt.ToString());
                            Match clippingDateOfEptyKt = dateSearcher.Match(clippingOfKt.ToString());
                            Match clippingNumberOfEptyKt = numberSearcher.Match(clippingOfKt.ToString());

                            if (clippingNumberOfEptyKt.Value == null || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyDt.Value == "" || clippingNumberOfEptyDt.Value == " " || clippingNumberOfEptyDt.Value == "   " || clippingNumberOfEptyDt.Value == null || clippingDateOfEptyDt.Value == "" || clippingDateOfEptyDt.Value == " " || clippingDateOfEptyDt.Value == "   " || clippingDateOfEptyDt.Value == null)
                            {
                                // Заливаем элемент в объект (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                WorkThisContent.LoadElementInArrayOfErrors(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                WorkThisContent.LoadElementInArrayOfErrors(filteringStatements.Variables.number, eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                            }
                            else
                            {
                                // Заливаем элемент в объект (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                WorkThisContent.LoadElementInArrayOfContent(-(filteringStatements.Variables.number), eptyPeriod, eptyDt, clippingNumberOfEptyDt, clippingDateOfEptyDt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                                // Заливаем элемент в объект (ПОЛОЖИТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
                                WorkThisContent.LoadElementInArrayOfContent(filteringStatements.Variables.number, eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                            }
                        }
                    }
                    else
                    {
                        if (Kt.Value == "" || Kt.Value == " " || Kt.Value == "   " || Kt.Value == null)
                        {
                            if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01")
                            {
                                Regex regexDt = new(filteringStatements.Variables.consractFirstOption);
                                Match clippingOfDt = regexDt.Match(eptyDt);
                                // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                                Match clippingDateOfEptyDt = dateSearcher.Match(clippingOfDt.ToString());
                                Match clippingNumberOfEptyDt = numberSearcher.Match(clippingOfDt.ToString());

                                if (clippingNumberOfEptyDt.Value == "" || clippingNumberOfEptyDt.Value == " " || clippingNumberOfEptyDt.Value == "   " || clippingNumberOfEptyDt.Value == null || clippingDateOfEptyDt.Value == "" || clippingDateOfEptyDt.Value == " " || clippingDateOfEptyDt.Value == "   " || clippingDateOfEptyDt.Value == null)
                                {
                                    // Добавляем в массив элемент не прошедший фильтрацию //
                                    WorkThisContent.LoadElementInArrayOfErrorsAlternativeEvent(eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                                else
                                {
                                    // Заливаем элемент в объект с валидными данными //
                                    WorkThisContent.LoadElementInArrayOfContentAlternativeEvent(eptyPeriod, eptyDt, clippingNumberOfEptyDt, clippingDateOfEptyDt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                            }
                            else
                            {
                                // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                                Match clippingDateOfEptyKt = dateSearcher.Match(Kt.ToString());
                                Match clippingNumberOfEptyKt = numberSearcher.Match(Kt.ToString());

                                if (clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == null || clippingDateOfEptyKt.Value == "" || clippingDateOfEptyKt.Value == " " || clippingDateOfEptyKt.Value == "   " || clippingDateOfEptyKt.Value == null)
                                {
                                    // Добавляем в массив элемент не прошедший фильтрацию //
                                    WorkThisContent.LoadElementInArrayOfErrorsAlternativeEvent(eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                                else
                                {
                                    // Заливаем элемент в объект с валидными данными //
                                    WorkThisContent.LoadElementInArrayOfContentAlternativeEvent(eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                                }
                            }
                        }
                        else
                        {
                            // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                            Match clippingDateOfEptyKt = dateSearcher.Match(Kt.ToString());
                            Match clippingNumberOfEptyKt = numberSearcher.Match(Kt.ToString());

                            if (clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == null || clippingDateOfEptyKt.Value == "" || clippingDateOfEptyKt.Value == " " || clippingDateOfEptyKt.Value == "   " || clippingDateOfEptyKt.Value == null)
                            {
                                // Добавляем в массив элемент не прошедший фильтрацию //
                                WorkThisContent.LoadElementInArrayOfErrorsAlternativeEvent(eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                            }
                            else
                            {
                                // Заливаем элемент в объект с валидными данными //
                                WorkThisContent.LoadElementInArrayOfContentAlternativeEvent(eptyPeriod, eptyKt, clippingNumberOfEptyKt, clippingDateOfEptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);
                            }
                        }
                    }
                }
            }

            // Выгружаем из БД данные buh и trash //
            Repository.ContentFromTable("SELECT *,format(datepl,'dd.MM.yyyy 0:00:00') as dateplString FROM buh;", filteringStatements.Variables.arrayBuhFromDataBase);
            Repository.ContentFromTable("SELECT *,format(datepl,'dd.MM.yyyy 0:00:00') as dateplString FROM trash;", filteringStatements.Variables.arrayTrashFromDataBase);
        }

        // Код работающий при ошибке //
        catch (Exception er)
        {
            Console.WriteLine(er.Message);
        }
        // Код работатющий в любом случше //
        finally
        {
            if(filteringStatements.Variables.arrayBuhFromDataBase.Count > 0 || filteringStatements.Variables.arrayTrashFromDataBase.Count > 0)
            {
                int newIntersection = WorkThisContent.FixIntersection(intersection, filteringStatements.Variables.arrayBuhFromDataBase, filteringStatements.Variables.arrayOfFilteredData, filteringStatements.Variables.arrayTrashFromDataBase, filteringStatements.Variables.arrayOfErrors);
                if (newIntersection == 0)
                {
                    // Пересечения дат нет //
                    Console.WriteLine("Пересечения дат нет");
                    // Загрузка результатов в БД //
                    filteringStatements.Repository.LoadFilteredContentFromFileInDataBase();
                }
                else
                {
                    Console.WriteLine("Обнаружены пересечения");
                    Console.WriteLine("Слушаем решение пользователя");
                    var messageForContent = WorkThisContent.GroupMetodsForCheckAnswerFromUser(messageUser, filteringStatements.Variables.arrayBuhFromDataBase, filteringStatements.Variables.arrayOfFilteredData, filteringStatements.Variables.arrayTrashFromDataBase, filteringStatements.Variables.arrayOfErrors);
                    filteringStatements.WorkThisContent.FilteringResponses(messageForContent, filteringStatements.Variables.arrayBuhFromDataBase, filteringStatements.Variables.arrayOfFilteredData, filteringStatements.Variables.arrayTrashFromDataBase, filteringStatements.Variables.arrayOfErrors);
                }
            } else 
            {
                // Пересечения дат нет //
                Console.WriteLine("Таблица в базе данных пуста");
                // Загрузка результатов в БД //
                filteringStatements.Repository.LoadFilteredContentFromFileInDataBase();
            }
            
            Console.WriteLine("Done");
            Console.WriteLine(filteringStatements.Variables.arrayOfFilteredData.Count / 7);
            Console.WriteLine(filteringStatements.Variables.arrayOfErrors.Count / 7);
            Console.WriteLine(filteringStatements.Variables.arrayBuhFromDataBase.Count / 7);
        }
    }
}