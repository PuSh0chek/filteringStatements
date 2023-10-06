namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Data;

class Program
{
    // Объект с данными прошедшими фильтрацию //
    public class verifiedData
    {
        public string Period { get; set; }
        public string ContractNumber { get; set; }
        public string ContractDate { get; set; }
        public string Debet { get; set; }
        public string DebetSum { get; set; }
        public string Kredit { get; set; }
        public string KreditSum { get; set; }
        public string Dt { get; set; }
        public string Kt { get; set; }
        public verifiedData(string period, string contractNumber, string contractDate, string debet, string debetSum, string kredit, string kreditSum, string dt, string kt)
        {
            Period = period;
            ContractNumber = contractNumber;
            ContractDate = contractDate;
            Debet = debet;
            DebetSum = debetSum;
            Kredit = kredit;
            KreditSum = kreditSum;
            Dt = dt;
            Kt = kt;
        }
    }
    // Объект с данными не прошедших фильтрацию первого уровня //
    public class erroneousData
    {
        public string Period { get; set; }
        public string Dt { get; set; }
        public string Kt { get; set; }
        public string Debet { get; set; }
        public string DebetSum { get; set; }
        public string Kredit { get; set; }
        public string KreditSum { get; set; }
        public erroneousData(string period, string dt, string kt, string debet, string debetSum, string kredit, string kreditSum)
        {
            Period = period;
            Dt = dt;
            Kt = kt;
            Debet = debet;
            DebetSum = debetSum;
            Kredit = kredit;
            KreditSum = kreditSum;
        }
    }
    // Объект с данными не прошедшими двух уровневую фильтрацию и оставленных для исправления вручную //
    public class trash
    {
        public string Period { get; set; }
        public string Dt { get; set; }
        public string Kt { get; set; }
        public string Debet { get; set; }
        public string DebetSum { get; set; }
        public string Kredit { get; set; }
        public string KreditSum { get; set; }
        public trash(string period, string dt, string kt, string debet, string debetSum, string kredit, string kreditSum)
        {
            Period = period;
            Dt = dt;
            Kt = kt;
            Debet = debet;
            DebetSum = debetSum;
            Kredit = kredit;
            KreditSum = kreditSum;
        }
    }
    static void Main()
    {
        // Подключение //

        // Регулярные выражения //
        // Первая проверка(стандарт) //
        string consractFirstOption = @"(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.20\d{2})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}г.\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}г.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.20\d{2})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}г.\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}г.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.20\d{2})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}г.\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}Г.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.20\d{2})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}г.\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}\SГ.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.20\d{2})|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4}\sг.)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2}г.\s)|(\d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}г.\s)|(№\s\d{1}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2})|(№\s\d{4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4}г)";
        // Вторая проверка для нахождения исключений в ошибочных элементах и переноса их в главный массив с корректными данными //
        string contractFirstOptionNoS = @"\d{1,4}-[в,к,В,К]{1,2}от\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionTO = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*то\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionOOF = @"№\s\d{1,4}-{1,2}[в,к,В,К]{1,2}\/[О,Ф]{3}-{1,2}\d{5,7}\/[С]-[К,к,А,а,В,в]{1,4}\s{1,2}от\s{1,2}\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionOOFnumberWithOutS = @"№\d{1,4}-{1,2}[в,к,В,К]{1,2}\/[О,Ф]{3}-{1,2}\d{5,7}\/[С]-[К,к,А,а,В,в]{1,4}\s{1,2}от\s{1,2}\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionSlash = @"№\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\/\d{1,3}\s*от\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFirstOptionOnlyYeacr = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}г.";
        string contractFirstOptionOnlyYeacrS = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{4}\sг.";
        string contractTwoOption = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*тот\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractTwoOptionOTI = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*оти\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractThreeOption = @"\s*№\s*\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFourOption = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{0,2}\s*\d{1,2}.\d{1,2}.\d{1,4}";
        string contractFourOptionOnOneNumberWithFrom = @"№\s*\d{1}\s{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{1,4}}";
        string contractFourOptionOnOneNumber = @"\s№\d\s\d{1,2}.\d{1,2}.\d{1,4}\s*";
        string contractFourOptionOnOneNumberWithFromSOGL = @"Соглашение{0,1}\s*№{1}\s*\d{1}\s*\s*от{1}\s*\d{1,2}.\d{1,2}.\d{1,4}";
        // Вариант с регулярнаым выражением для проверки числа //
        string contractNumber = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}";
        // Вариант с регулярнаым выражением для проверки даты // 
        string dataStandartNoS = @"от\d{2}.\d{2}.\d{2,4}";
        string dataStandart = @"\s\d{2}.\d{2}.\d{2,4}";
        string dataOnlyYear = @"\s\d{4}г{1}.{1}\s";
        string dataOnlyFullYeacrWithS = @"\s\d{2}.\d{2}.\d{2,4}\sг{1}.{1}\s";
        // Вариант с запасным регулярным выражением //
        string scoreNumberOne = @"(?:50|51|57)";
        string amount = @"d+";
        string scoreNumberTwo = @"d+";
        // Массивы //
        // Массив с вариантами регулярных выражений для второй проверки главного массива с данными //
        List<string> contract = new List<string>() { consractFirstOption, contractTwoOption, contractThreeOption, contractFourOption, contractTwoOptionOTI, contractFirstOptionOnlyYeacr, contractFourOptionOnOneNumberWithFrom, contractFourOptionOnOneNumber, contractFourOptionOnOneNumberWithFromSOGL, contractFirstOptionSlash, contractFirstOptionTO, contractFirstOptionOOF, contractFirstOptionNoS, contractFirstOptionOnlyYeacrS, contractFirstOptionOOFnumberWithOutS };
        // Массив с вариантами регулярных функций для проверки даты //
        List<string> date = new List<string>() { dataStandart, dataOnlyYear, dataStandartNoS, dataOnlyFullYeacrWithS };
        // Массив с данными удачно прошедшими фильтрацию //
        List<string> arrayOfFilteredData = new List<string>() {  };
        // Массив с данными НЕ удачно прошедшими фильтрацию //
        List<string> arrayOfErrors = new List<string>() { };
        // Массив с данными НЕ удачно прошедшими фильтрацию //
        List<string> arrayOfErrors2 = new List<string>() { };
        // Массив с данными фильтрованных ошибок прошедшими фильтрацию //
        List<string> arrayOfFilteredErrors = new List<string>() { };
        // Массив с данными фильтрованных ошибок прошедшими фильтрацию //
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
                    string eptyDebetSum = "1";
                    string eptyKredit = (string)kredit;
                    string eptyKreditSum = "1";
                    // Проверяем элементы на значение null //
                    if(eptyPeriod != null && eptyDt != null && eptyKt != null && eptyDebet != null && eptyDebetSum != null && eptyKredit != null && eptyKreditSum != null)
                    {   
                        // Поиск через регулярные выражения //
                        Regex regex = new(consractFirstOption);
                        Match clippingOfKt = regex.Match(eptyKt);
                        if(clippingOfKt.Value == "" || clippingOfKt.Value == " " || clippingOfKt.Value == "   " || clippingOfKt.Value == null)
                        {
                            // Добавляем в массив элемент не прошедший фильтрацию //
                            arrayOfErrors.Add(eptyPeriod);
                            arrayOfErrors.Add(eptyDt);
                            arrayOfErrors.Add(eptyKt);
                            arrayOfErrors.Add(eptyDebet);
                            arrayOfErrors.Add(eptyDebetSum);
                            arrayOfErrors.Add(eptyKredit);
                            arrayOfErrors.Add(eptyKreditSum);
                        } else
                        {
                            // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                            Regex numberSearcher = new(contractNumber);
                            Regex dateSearcher = new(dataStandart);
                            Match clippingDateOfEptyKt = dateSearcher.Match(eptyKt);
                            Match clippingNumberOfEptyKt = numberSearcher.Match(eptyKt);
                            if (clippingNumberOfEptyKt.Value == "" || clippingNumberOfEptyKt.Value == " " || clippingNumberOfEptyKt.Value == "   " || clippingNumberOfEptyKt.Value == null || clippingDateOfEptyKt.Value == "" || clippingDateOfEptyKt.Value == " " || clippingDateOfEptyKt.Value == "   " || clippingDateOfEptyKt.Value == null)
                            {
                                // Добавляем в массив элемент не прошедший фильтрацию //
                                arrayOfErrors.Add(eptyPeriod);
                                arrayOfErrors.Add(eptyDt);
                                arrayOfErrors.Add(eptyKt);
                                arrayOfErrors.Add(eptyDebet);
                                arrayOfErrors.Add(eptyDebetSum);
                                arrayOfErrors.Add(eptyKredit);
                                arrayOfErrors.Add(eptyKreditSum);
                            }
                            else
                            {
                                // Заливаем элемент в объект с валидными данными //
                                //verifiedData verifiedData = new verifiedData(eptyPeriod, clippingNumberOfEptyKt.Value, clippingDateOfEptyKt.Value, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum, eptyDt, eptyKt); //
                                arrayOfFilteredData.Add(clippingNumberOfEptyKt.Value);
                                arrayOfFilteredData.Add(clippingDateOfEptyKt.Value);
                                arrayOfFilteredData.Add(eptyDebet);
                                arrayOfFilteredData.Add(eptyKredit);
                                arrayOfFilteredData.Add(eptyDebetSum);
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
                        arrayOfErrors.Add(eptyDebetSum);
                        arrayOfErrors.Add(eptyKredit);
                        arrayOfErrors.Add(eptyKreditSum);
                    }

                }

                // Начало среза //
                int skip = 0;
                // Конец среза //
                int take = 7;

                // Проходимся циклом  //
                for (int n = 0; n < arrayOfErrors.Count; n += 7)
                {
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
                                                if(clippingElementNumber.ToString() != "" && clippingElementNumber.ToString() != " " && clippingElementNumber != null && clippingElementDate.ToString() != "" && clippingElementDate.ToString() != " " && clippingElementDate != null)
                                                {
                                                    arrayOfFilteredData.Add(clippingElementNumber.Value);
                                                    arrayOfFilteredData.Add(clippingElementDate.Value);
                                                    arrayOfFilteredData.Add(list[3]);
                                                    arrayOfFilteredData.Add(list[5]);
                                                    arrayOfFilteredData.Add(list[6]);
                                                    arrayOfFilteredData.Add(list[0]);
                                                    arrayOfFilteredData.Add(list[2]);
                                                    
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
                                                    arrayOfFilteredData.Add("-"+list[6]);
                                                    arrayOfFilteredData.Add(list[0]);
                                                    arrayOfFilteredData.Add(list[1]);
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
            }
            // Данные для подключения к серверу //
            string connectionString = "Data Source= rvdk-svr-6091, 1500;Initial Catalog= TechConditions;Integrated Security=SSPI;";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    // Подключение //
                    connection.Open();
                    Console.WriteLine("Connection successfully opened.");
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
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        catch (Exception er)
        {
            Console.WriteLine(er.Message);
        }
        finally
        {
            Console.WriteLine("Done");
            Console.WriteLine(arrayOfFilteredData.Count/7);
            Console.WriteLine(arrayOfErrors.Count/7);
            Console.WriteLine(arrayOfErrors2.Count/7);
            Console.WriteLine(arrayBuhFromDataBase.Count/7);
        }
    }
}