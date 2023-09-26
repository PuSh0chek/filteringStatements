namespace FilterStatements;
using System;
using System.IO;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        string objectForFilter = "Олейников Руслан Витальевич\n№153-В от 2017г. о подключ. к с-ме водоснабж. ул.37 Линия, 95\nПоступление наличных 00000001264 от 16.03.2018 18:57:28";
        string contractFirstOption = @"d{1,4}\s*-{0,2}\s*[в,к,В,К]{1,2}\s*от\s*\d{1,2}.\d{1,2}.\d{4})";


    }
}