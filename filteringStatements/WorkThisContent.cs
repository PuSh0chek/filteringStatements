using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using filteringStatements;
using FilterStatements;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace filteringStatements
{
    public class WorkThisContent
    {
        // Группа методов для проверки совпадений по дате в контенте и уведомления пользователя о фиксации первого совпадения //
        public static string GroupMetodsForCheckAnswerFromUser(string messageUser, List<string> arrayParent, List<string> arrayСompared, List<string> arrayParentEr, List<string> arrayСomparedEr)
        {
            // Стадия выполнения сценариев //
            for (int j = 5; j < arrayParent.Count; j += 7)
            {
                for (int n = 5; n < arrayСomparedEr.Count; n += 7)
                {
                    // Проверяем равны ли элементы //
                    if (arrayParent[j].ToString() == arrayСomparedEr[n - 5].ToString()) return Answer(messageUser, arrayParent, arrayСomparedEr, j, n);
                }
            }
            for (int j = 5; j < arrayParentEr.Count; j += 7)
            {
                for (int n = 0; n < arrayСompared.Count; n += 7)
                {
                    // Проверяем равны ли элементы //
                    if (arrayParentEr[j].ToString() == arrayСompared[n].ToString()) return Answer(messageUser, arrayParentEr, arrayСompared, j, n);
                }
            }
            for (int j = 5; j < arrayParent.Count; j += 7)
            {
                for (int n = 5; n < arrayСompared.Count; n += 7)
                {
                    // Проверяем равны ли элементы //
                    if (arrayParent[j].ToString() == arrayСompared[n].ToString()) return Answer(messageUser, arrayParent, arrayСompared, j, n);
                }
            }
            for (int j = 5; j < arrayParentEr.Count; j += 7)
            {
                for (int n = 5; n < arrayСomparedEr.Count; n += 7)
                {
                    // Проверяем равны ли элементы //
                    if (arrayParentEr[j].ToString() == arrayСomparedEr[n].ToString()) return Answer(messageUser, arrayParentEr, arrayСomparedEr, j, n);
                }
            }
            return messageUser;
        }
        // Метод для проверки файта пересечения даты (хотябы одной) //
        public static int FixIntersection(int intersection, List<string> arrayParent, List<string> arrayСompared, List<string> arrayParentEr, List<string> arrayСomparedEr)
        {
            // Стадия выполнения сценариев //
            int iterationParentTrueArrayThisErrayContent = GetIntersection(intersection, -5, arrayParent, arrayСomparedEr);
            int iterationParentErrorsArrayThisTrueContent = GetIntersection(intersection, 0, arrayParentEr, arrayСompared);
            int iterationParentTrueArrayThisTrueContent = GetIntersection(intersection, 0, arrayParent, arrayСompared);
            int iterationParentErrorsArrayThisErrorsContent = GetIntersection(intersection, -5, arrayParentEr, arrayСomparedEr);
            // Стадия проверки условий //
            if(iterationParentTrueArrayThisErrayContent != 0) return iterationParentTrueArrayThisErrayContent;
            else if (iterationParentErrorsArrayThisTrueContent != 0) return iterationParentErrorsArrayThisTrueContent;
            else if (iterationParentTrueArrayThisTrueContent != 0) return iterationParentTrueArrayThisTrueContent;
            else if (iterationParentErrorsArrayThisErrorsContent != 0) return iterationParentErrorsArrayThisErrorsContent;
            return intersection;
        }
        // Метод для получения пересечений дат. При присутствии пересечения задает новое значение переменной равное единице //
        public static int GetIntersection(int intersection, int indexControl, List<string> arrayParent, List<string> arrayСompared)
        {
            for (int j = 5; j < arrayParent.Count; j += 7)
            {
                for (int n = 5; n < arrayСompared.Count; n += 7)
                {
                    // Проверяем равны ли элементы //
                    if (arrayParent[j].ToString() == arrayСompared[n + indexControl].ToString()) return intersection = 1;
                }
            }
            return intersection;
        }
        // Получаем ответ от пользователя //
        public static string Answer(string messageUser, List<string> arrayParent, List<string> arrayСompared, int j, int n)
        {
            // Вывод информации для пользователя //
            Console.WriteLine($"Пересечение дат элементов: {arrayParent[j]} c данными {arrayParent[j - 5]}, {arrayParent[j - 4]}, {arrayParent[j - 3]}, {arrayParent[j - 2]}, {arrayParent[j - 1]} | и нового элемента: | {arrayСompared[n]} c данными {arrayСompared[n - 5]}, {arrayСompared[n - 4]}, {arrayСompared[n - 3]}, {arrayСompared[n - 2]}, {arrayСompared[n - 1]}");
            // Условия выбора для пользователя //
            Console.Write("Введите да (Заменить все старые пересекающиеся элементы на новые), нет (Оставить старые элементы):");
            // Получение ответа от пользователя //
            messageUser = Console.ReadLine();
            if (messageUser == "да") return messageUser;
            else if (messageUser == "нет") return messageUser;
            else
            {
                Console.WriteLine($"Cценарий {messageUser} отсутствует");
                return messageUser;
            }
        }
        // Группирование группы циклов для обработки бд и массивов //
        public static void GroupBlockСyclesForProcessingContent(string messageUser, List<string> arrayParent, List<string> arrayСompared, List<string> arrayParentEr, List<string> arrayСomparedEr)
        {
            // Группа методов с циклами для работы с контентом //
            UpdateContentArrayAndDB(messageUser, filteringStatements.Variables.tableTrash, 0, arrayParentEr, arrayСomparedEr);
            UpdateContentArrayAndDB(messageUser, filteringStatements.Variables.tableTrash, 5, arrayParentEr, arrayСompared);
            UpdateContentArrayAndDB(messageUser, filteringStatements.Variables.tableBuh, 0, arrayParent, arrayСomparedEr);
            UpdateContentArrayAndDB(messageUser, filteringStatements.Variables.tableBuh, 5, arrayParent, arrayСompared);

            // Загрузка результатов в БД //
            using SqlConnection connection = new(Variables.connectionString);
            try
            {
                connection.Open();
                Repository.AccessingAtServerWithFilteredData(Variables.queryTableBuh, connection, arrayСompared);
                Repository.AccessingServerOfTableTrash(Variables.queryTableTrash, connection, arrayСomparedEr);
                connection.Close();
            }
            catch (Exception er)
            {
                Console.WriteLine(er.Message);
            }
        }
        // Работа с массивами и бд(удаление из бд при messageUser == да и удаление из массива элемента messageUser == нет) // 
        public static void UpdateContentArrayAndDB(string messageUser, string table, int indexControl, List<string> arrayParent, List<string> arrayChildren)
        {
            for (int j = 0; j < arrayParent.Count; j += 7)
            {
                for (int n = 0; n < arrayChildren.Count; n += 7)
                {
                    if (arrayParent[j + 5].ToString() == arrayChildren[n + indexControl].ToString())
                    {
                        if (messageUser == "да")
                        {
                            filteringStatements.Repository.RequestToDeleteAnItemThatMatchesTheDate($"DELETE FROM {table} WHERE datepl = '{arrayChildren[n + indexControl].Replace("0:00:00", "").Replace(" ", "")}'");
                        }
                        else if (messageUser == "нет")
                        {
                            arrayChildren.RemoveRange(n, 7);
                        }
                    }
                }
            }
        }
        public static void FilteringResponses(string message, List<string> arrayParent, List<string> arrayСompared, List<string> arrayParentEr, List<string> arrayСomparedEr)
        {
            Console.WriteLine("Работа метода FilteringResponses после получения ответа от пользователя ");
            // Функция родитель(содержит много дочерних функция для работы с фильтрацией и работой с бд, массивами) //
            // Версия с вырезанием старых пересеченных элементов по датам и загрузкой нового контента //
            if (message == "да") GroupBlockСyclesForProcessingContent(message, arrayParent, arrayСompared, arrayParentEr, arrayСomparedEr);
            // Функция родитель(содержит много дочерних функция для работы с фильтрацией и работой с бд, массивами) //
            // Версия с вырезанием новых пересеченных элементов по датам и загрузкой старого контента //
            else if (message == "нет") GroupBlockСyclesForProcessingContent(message, arrayParent, arrayСompared, arrayParentEr, arrayСomparedEr);
            else Console.WriteLine("Ответ от пользователя не равен да/нет");
        }
        // Доавить новый элемент в массив с ошибками //
        public static void LoadElementInArrayOfErrors(int number, string eptyPeriod, string eptyDt, string eptyKt, string eptyDebet, double eptyDebetSum, string eptyKredit, double eptyKreditSum)
        {
            Variables.arrayOfErrors.Add(eptyPeriod);
            Variables.arrayOfErrors.Add(eptyDt);
            Variables.arrayOfErrors.Add(eptyKt);
            Variables.arrayOfErrors.Add(eptyDebet);
            Variables.arrayOfErrors.Add((number * Convert.ToDouble(eptyDebetSum)).ToString());
            Variables.arrayOfErrors.Add(eptyKredit);
            Variables.arrayOfErrors.Add(eptyKreditSum.ToString());
        }
        // Доавить новый элемент в фильтрованный массив //
        public static void LoadElementInArrayOfContent(int number, string eptyPeriod, string epty, System.Text.RegularExpressions.Match clippingNumberOfEpty, System.Text.RegularExpressions.Match clippingDateOfEpty, string eptyDebet, double eptyDebetSum, string eptyKredit, double eptyKreditSum)
        {
            // Заливаем элемент в объект (ОТРИЦАТЕЛЬНЫЙ ДЕБЕТ) с валидными данными //
            // Указываем EPTY в зависимости от предъявленных требований (DT или KT) //
            Variables.arrayOfFilteredData.Add(clippingNumberOfEpty.Value);
            if (clippingDateOfEpty.ToString().Trim().IndexOf(" ") == -1)
            {
                // Указываем EPTY в зависимости от предъявленных требований (DT или KT) //
                Variables.arrayOfFilteredData.Add(clippingDateOfEpty.ToString().Trim());
            }
            else
            {
                // Указываем EPTY в зависимости от предъявленных требований (DT или KT) //
                Variables.arrayOfFilteredData.Add(clippingDateOfEpty.ToString().Trim().Replace(" ", "."));
            }
            Variables.arrayOfFilteredData.Add(eptyDebet);
            Variables.arrayOfFilteredData.Add(eptyKredit);
            Variables.arrayOfFilteredData.Add((number * Convert.ToDouble(eptyDebetSum)).ToString());
            Variables.arrayOfFilteredData.Add(eptyPeriod);
            // Указываем EPTY в зависимости от предъявленных требований (DT или KT) //
            Variables.arrayOfFilteredData.Add(epty);
        }
        // Доавить новый элемент в массив с ошибками (Нижний альтернативный уровень проверки элементов файла) //
        public static void LoadElementInArrayOfErrorsAlternativeEvent(string eptyPeriod, string eptyDt, string eptyKt, string eptyDebet, double eptyDebetSum, string eptyKredit, double eptyKreditSum)
        {
            // Добавляем в массив элемент не прошедший фильтрацию //
            Variables.arrayOfErrors.Add(eptyPeriod);
            Variables.arrayOfErrors.Add(eptyDt);
            Variables.arrayOfErrors.Add(eptyKt);
            Variables.arrayOfErrors.Add(eptyDebet);
            if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01")
            {
                Variables.arrayOfErrors.Add(Convert.ToDouble(-eptyDebetSum).ToString());
            }
            else
            {
                Variables.arrayOfErrors.Add(Convert.ToDouble(eptyDebetSum).ToString());
            }
            Variables.arrayOfErrors.Add(eptyKredit);
            Variables.arrayOfErrors.Add(Convert.ToDouble(eptyKreditSum).ToString());
        }
        // Доавить новый элемент в фильтрованный массив (Нижний альтернативный уровень проверки элементов файла) //
        public static void LoadElementInArrayOfContentAlternativeEvent (string eptyPeriod, string epty, System.Text.RegularExpressions.Match clippingNumberOfEpty, System.Text.RegularExpressions.Match clippingDateOfEpty, string eptyDebet, double eptyDebetSum, string eptyKredit, double eptyKreditSum)
        {
            try {
                Variables.arrayOfFilteredData.Add(clippingNumberOfEpty.Value);
                if (clippingDateOfEpty.ToString().Trim().IndexOf(" ") == -1)
                {
                    Variables.arrayOfFilteredData.Add(clippingDateOfEpty.ToString());
                }
                else
                {
                    Variables.arrayOfFilteredData.Add(clippingDateOfEpty.ToString().Trim().Replace(" ", "."));
                }
                Variables.arrayOfFilteredData.Add(eptyDebet);
                Variables.arrayOfFilteredData.Add(eptyKredit);
                if (eptyDebet.ToString() == "62.02" || eptyDebet.ToString() == "62.01")
                {
                    Variables.arrayOfFilteredData.Add(Convert.ToDouble(-eptyDebetSum).ToString());
                }
                else
                {
                    Variables.arrayOfFilteredData.Add(Convert.ToDouble(eptyDebetSum).ToString());
                }
                Variables.arrayOfFilteredData.Add(eptyPeriod);
                Variables.arrayOfFilteredData.Add(epty);
            } catch (Exception e)
            {
                Console.WriteLine (e.ToString());
            }
        }
    }
}