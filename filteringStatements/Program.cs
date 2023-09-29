namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Xsl;
using System.Linq;
using System.Transactions;
using System.Data.Common;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using Microsoft.Extensions.FileSystemGlobbing;

class Program
{
    // Объект с данными прошедшими фильтрацию //
    public class verifiedData
    {
        private string epty2ContractString;
        private string epty3ContractString;
        private string epty5ContractString;
        private string epty6ContractString;
        private string epty7ContractString;
        private string epty8ContractString;

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
                    // Вытаскиваем переенные из объекта //
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
                        Regex regex = new Regex(consractFirstOption);
                        Match clippingOfKt = regex.Match(eptyKt);
                        // Если проверку элемент не прошел заливаем его в объект trash, иначе в объект с валидными данными //
                        if(clippingOfKt.Value == "" || clippingOfKt.Value == " " || clippingOfKt.Value == "   " || clippingOfKt.Value == null)
                        {
                            trash trash = new trash(eptyPeriod, eptyDt, eptyKt, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum);

                        } else
                        {
                            // Поиск через регулярные выражения для деления прошлого элемента на две составляющие //
                            Regex numberSearcher = new Regex(contractNumber);
                            Regex dateSearcher = new Regex(dataStandart);
                            Match clippingNumberOfEptyKt = numberSearcher.Match(eptyKt);
                            Match clippingDateOfEptyKt = dateSearcher.Match(eptyKt);
                            // Заливаем элемент в объект с валидными данными //
                            verifiedData employee = new verifiedData(eptyPeriod, clippingDateOfEptyKt.Value, clippingNumberOfEptyKt.Value, eptyDebet, eptyDebetSum, eptyKredit, eptyKreditSum, eptyDt, eptyKt) ;
                        }
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
            Console.WriteLine("Done");
        }
    }
}