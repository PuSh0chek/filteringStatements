namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        string objectForFilter = "02.11.2022, 166-К,  20.06.22, 62.02, 150783, 51, -150783, Южный филиал ПАО \"Промсвязьбанк\" (Расчетный) Основной 700\nПоступления - выполнение действий по подготовке ВиК к подключению, РЖД (ОАО)\n№166-К от 20.06.22 о подключ. к с-ме водоотвед. пр-кт Сельмаш, 1а\nПоступление на расчетный счет 00000003723 от 30.06.2022 23:59:59";
        string contractFirstOption = @"\s\d{2}.\d{2}.\d{2,4}";
        Regex regex = new Regex(objectForFilter);

        MatchCollection match = regex.Matches(contractFirstOption);
        Console.WriteLine(match.Count);
    }
}