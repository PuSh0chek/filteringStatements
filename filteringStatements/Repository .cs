using GraphQL.Types.Relay.DataObjects;
using GraphQL.Validation;
using Microsoft.VisualBasic;
using OfficeOpenXml.Drawing.Slicer.Style;
using System;
using System.Data.SqlClient;
using System.Runtime.CompilerServices;

namespace filteringStatements
{
    public class Repository
    {
        filteringStatements.Variables variables = new Variables();
        public static void AccessingAtServerForTableBuh(string queryTableBuh, SqlConnection connection, List<string> arrayBuhFromDataBase)
        {
             try
            {
                for (int f = 0; f < arrayBuhFromDataBase.Count; f += 7)
                {
                    using (SqlCommand command = new SqlCommand(queryTableBuh, connection))
                    {
                        command.Parameters.AddWithValue("@dog", arrayBuhFromDataBase[f]);
                        command.Parameters.AddWithValue("@datadog", arrayBuhFromDataBase[f + 1]);
                        command.Parameters.AddWithValue("@dt", arrayBuhFromDataBase[f + 2]);
                        command.Parameters.AddWithValue("@kt", arrayBuhFromDataBase[f + 3]);
                        command.Parameters.AddWithValue("@summ", Convert.ToDouble(arrayBuhFromDataBase[f + 4]));
                        command.Parameters.AddWithValue("@datepl", arrayBuhFromDataBase[f + 5]);
                        command.Parameters.AddWithValue("@text", arrayBuhFromDataBase[f + 6].Replace("\n", "__"));
                        int rowsAffected = command.ExecuteNonQuery();
                    }
                }
            } catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        // Залить данные из массива в БД таблицу Buh //
        public static void AccessingAtServerWithFilteredData(string queryTableBuh, SqlConnection connection, List<string> array)
        {
            try
            {
                for (int f = 0; f < array.Count; f += 7)
                {
                    using (SqlCommand command = new SqlCommand(queryTableBuh, connection))
                    {
                        command.Parameters.AddWithValue("@dog", array[f]);
                        command.Parameters.AddWithValue("@datadog", array[f + 1]);
                        command.Parameters.AddWithValue("@dt", array[f + 2]);
                        command.Parameters.AddWithValue("@kt", array[f + 3]);
                        command.Parameters.AddWithValue("@summ", Convert.ToDouble(array[f + 4]));
                        command.Parameters.AddWithValue("@datepl", Convert.ToDateTime(array[f + 5]));
                        command.Parameters.AddWithValue("@text", array[f + 6].Replace("\n", "__"));
                        int rowsAffected = command.ExecuteNonQuery();
                    }
                }
            } catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        // Залить данные из массива в БД таблицу Trash //
        public static void AccessingServerOfTableTrash(string queryTableTrash, SqlConnection connection, List<string> arrayOfErrors)
        {
            if(arrayOfErrors.Count > 0)
            {
                for (int f = 0; f < arrayOfErrors.Count; f += 7)
                {
                    using (SqlCommand command = new SqlCommand(queryTableTrash, connection))
                    {
                        command.Parameters.AddWithValue("@dog", DBNull.Value);
                        command.Parameters.AddWithValue("@datadog", DBNull.Value);
                        command.Parameters.AddWithValue("@dt", arrayOfErrors[f + 3]);
                        command.Parameters.AddWithValue("@kt", arrayOfErrors[f + 5]);
                        command.Parameters.AddWithValue("@summ", Convert.ToDouble(arrayOfErrors[f + 4].ToString()));
                        command.Parameters.AddWithValue("@datepl", Convert.ToDateTime(arrayOfErrors[f]));
                        if(arrayOfErrors[f + 3] == "62.02" && arrayOfErrors[f + 5] == "62.01")
                        {
                            if (Convert.ToDouble(arrayOfErrors[f + 4]) < 0)
                            {
                                command.Parameters.AddWithValue("@text", (arrayOfErrors[f + 1].Replace("\n", "__")).ToString());
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@text", (arrayOfErrors[f + 2].Replace("\n", "__").ToString()));
                            }
                        } 
                        else
                        {
                            if (arrayOfErrors[f + 3] == "62.02" || arrayOfErrors[f + 3] == "62.01")
                            {
                                command.Parameters.AddWithValue("@text", (arrayOfErrors[f + 1].Replace("\n", "__")).ToString());
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@text", (arrayOfErrors[f + 2].Replace("\n", "__").ToString()));
                            }
                        }
                        int rowsAffected = command.ExecuteNonQuery();
                    }
                }
            }
        }
        // Выгрузить таблицу в массив //
        public static void ContentFromTable(string query, List<string> array) 
        {
            using (SqlConnection connection = new SqlConnection(Variables.connectionString))
            {
                try
                {
                    // Подключение //
                    connection.Open();
                    Console.WriteLine("Connection successfully opened");
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
                            string periodDate = reader["dateplString"].ToString();
                            string textDt = reader["text"].ToString().Replace("\n", "__");
                            array.Add(numberСontract);
                            array.Add(dateСontract);
                            array.Add(dt);
                            array.Add(kt);
                            array.Add(summ);
                            array.Add(periodDate);
                            array.Add(textDt);
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
        // Загрузить отфильтрованный контент в БД //
        public static void LoadFilteredContentFromFileInDataBase()
        {
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
        // Удаляем из БД совпадения по датам //
        public static void RequestToDeleteAnItemThatMatchesTheDate(string query)
        {
            using (SqlConnection connection = new SqlConnection(Variables.connectionString))
            {
                try
                {
                    // Подключение //
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        int rowsAffected = command.ExecuteNonQuery();
                    };
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
