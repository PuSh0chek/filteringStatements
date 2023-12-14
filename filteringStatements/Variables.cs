using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace filteringStatements
{
    public class Variables
    {
        // Путь к файлу с данными//
        public static string path = "E:\\Рабочие проекты\\filteringStatements\\filteringStatements\\MainDoc.xlsx";

        // РЕГУЛЯРНЫЕ ВЫРАЖЕНИЯ //
        // Первая проверка(стандарт) //
        public static string consractFirstOption = @"(\d{1,4}\s*-{1,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{2,4})|(\d{1,4}\s*-{1,2}\s*[в,к,В,К]{1,2}\s*\d{1,2}.\d{1,2}.\d{2,4})";

        // Вторая проверка для нахождения исключений в ошибочных элементах и переноса их в главный массив с корректными данными //
        public static string contractFirstOptionNoS = @"\d{1,4}-[в,к,В,К]{1,2}от\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractFirstOptionTO = @"\d{1,4}\s*\-{1,2}\s*[в,к,В,К]{1,2}\s*то\s*\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractFirstOptionOOF = @"№\s\d{1,4}-{1,2}[в,к,В,К]{1,2}\/[О,Ф]{3}-{1,2}\d{5,7}\/[С]-[К,к,А,а,В,в]{1,4}\s{1,2}от\s{1,2}\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractFirstOptionOOFnumberWithOutS = @"№\d{1,4}-{1,2}[в,к,В,К]{1,2}\/[О,Ф]{3}-{1,2}\d{5,7}\/[С]-[К,к,А,а,В,в]{1,4}\s{1,2}от\s{1,2}\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractFirstOptionSlash = @"№\d{1,4}\s*\-{1,2}\s*[в,к,В,К]{1,2}\/\d{1,3}\s*от\s*\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractTwoOption = @"\s*№\s*\d{1,4}\s*\-{1,2}\s*[в,к,В,К]{1,2}\s*тот\s*\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractTwoOptionOTI = @"\s*№\s*\d{1,4}\s*\-{1,2}\s*[в,к,В,К]{1,2}\s*оти\s*\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractThreeOption = @"\s*№\s*\d{1,4}\s*\-{1,2}\s*[в,к,В,К]{1,2}\s*\d{1,2}.\d{1,2}.\d{2,4}";
        public static string contractFourOption = @"\d{1,4}\s*\-{1,2}\s*[в,к,В,К]{1,2}\s*\d{1,2}.\d{1,2}.\d{2,4}";

        public static int number = 1;

        // Вариант с регулярнаым выражением для проверки числа //
        public static string contractNumber = @"\d{1,4}\s*\-{0,2}\s*[в,к,В,К]{1,2}";

        // Вариант с регулярнаым выражением для проверки даты // 
        public static string dataStandart = @"\s\d{2}.\d{2}.\d{2,4}";
        public static string dataStandartNoS = @"\d{2}.\d{2}.\d{2,4}";
        public static string dataOnlyFullYeacrWithS = @"\s\d{2}.\d{2}.\d{2,4}\sг{1}.{1}\s";

        // Паттерн даты //
        public static string pattern = @"\b0\:00\:00\b";

        // Замена перехода на новую строку подчеркиваниями //
        public static string underscores = @"__";

        // МАССИВЫ //
        // Массив с вариантами регулярных выражений для второй проверки главного массива с данными //
        public static List<string> contract = new List<string>() { consractFirstOption, contractTwoOption, contractThreeOption, contractFourOption, contractTwoOptionOTI, contractFirstOptionSlash, contractFirstOptionTO, contractFirstOptionOOF, contractFirstOptionNoS, contractFirstOptionOOFnumberWithOutS };

        // Массив с вариантами регулярных функций для проверки даты //
        public static List<string> date = new List<string>() { dataStandart, dataStandartNoS, dataOnlyFullYeacrWithS };

        // Массив с данными удачно прошедшими фильтрацию //
        public static List<string> arrayOfFilteredData = new List<string>() { };

        // Массив с данными НЕ удачно прошедшими фильтрацию //
        public static List<string> arrayOfErrors = new List<string>() { };

        // Массив с данными ошибок прошедших фильтрацию //
        public static List<string> arrayOfFilteredErrors = new List<string>() { };

        // Массив с данными из БД (только правельные) //
        public static List<string> arrayBuhFromDataBase = new List<string>() { };

        // Массив с данными из БД (только ошибки) //
        public static List<string> arrayTrashFromDataBase = new List<string>() { };

        // РАБОТА С СЕРВЕРОМ //
        // Данные для подключения к серверу //
        public static string connectionString = "Data Source= rvdk-svr-6091, 1500;Initial Catalog= TechConditions;Integrated Security=SSPI;";

        // Запросы к таблицам //
        public static string queryTableBuh = "INSERT INTO buh2 (dog, datadog, dt, kt, summ, datepl, text) VALUES (@dog, @datadog, @dt, @kt, @summ, @datepl, @text)";
        public static string queryTableTrash = "INSERT INTO trash2 (dog, datadog, dt, kt, summ, datepl, text) VALUES (@dog, @datadog, @dt, @kt, @summ, @datepl, @text)";

        // Переменная определяющая таблицу //
        public static string tableBuh = "buh2";
        public static string tableTrash = "trash2";
    }
}
