using GraphQL.Types.Relay.DataObjects;
using GraphQL.Validation;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace filteringStatements
{
    public class Repository
    {
        filteringStatements.Variables variables = new Variables();
        public static void AccessingAtServerForTableBuh(string queryTableBuh, SqlConnection connection, List<string> arrayBuhFromDataBase)
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
        }
        public static void AccessingAtServerWithFilteredData(string queryTableBuh, SqlConnection connection, List<string> arrayOfFilteredData)
        {
            for (int f = 0; f < arrayOfFilteredData.Count; f += 7)
            {
                using (SqlCommand command = new SqlCommand(queryTableBuh, connection))
                {
                    command.Parameters.AddWithValue("@dog", arrayOfFilteredData[f]);
                    command.Parameters.AddWithValue("@datadog", arrayOfFilteredData[f + 1]);
                    command.Parameters.AddWithValue("@dt", arrayOfFilteredData[f + 2]);
                    command.Parameters.AddWithValue("@kt", arrayOfFilteredData[f + 3]);
                    command.Parameters.AddWithValue("@summ", Convert.ToDouble(arrayOfFilteredData[f + 4]));
                    command.Parameters.AddWithValue("@datepl", arrayOfFilteredData[f + 5]);
                    command.Parameters.AddWithValue("@text", arrayOfFilteredData[f + 6].Replace("\n", "__"));
                    int rowsAffected = command.ExecuteNonQuery();
                }
            }
        }
        public static void AccessingServerOfTableTrash(string queryTableTrash, SqlConnection connection, List<string> arrayOfErrors)
        {
            for (int f = 0; f < arrayOfErrors.Count; f += 7)
            {
                using (SqlCommand command = new SqlCommand(queryTableTrash, connection))
                {
                    command.Parameters.AddWithValue("@dog", DBNull.Value);
                    command.Parameters.AddWithValue("@datadog", DBNull.Value);
                    command.Parameters.AddWithValue("@dt", arrayOfErrors[f + 3]);
                    command.Parameters.AddWithValue("@kt", arrayOfErrors[f + 5]);
                    command.Parameters.AddWithValue("@summ", Convert.ToDouble(arrayOfErrors[f + 4]));
                    command.Parameters.AddWithValue("@datepl", arrayOfErrors[f]);
                    if (arrayOfErrors[f + 3] == "62.02" || arrayOfErrors[f + 3] == "62.01")
                    {
                        command.Parameters.AddWithValue("@text", arrayOfErrors[f + 1].Replace("\n", "__"));
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@text", arrayOfErrors[f + 2].Replace("\n", "__"));
                    }
                    int rowsAffected = command.ExecuteNonQuery();
                }
            }
        }
        public static void ContentFromTableBuh() 
        {
            using (SqlConnection connection = new SqlConnection(Variables.connectionString))
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
                            string textDt = reader["text"].ToString().Replace("\n", "__");
                            Variables.arrayBuhFromDataBase.Add(numberСontract);
                            Variables.arrayBuhFromDataBase.Add(dateСontract);
                            Variables.arrayBuhFromDataBase.Add(dt);
                            Variables.arrayBuhFromDataBase.Add(kt);
                            Variables.arrayBuhFromDataBase.Add(summ);
                            Variables.arrayBuhFromDataBase.Add(periodDate);
                            Variables.arrayBuhFromDataBase.Add(textDt);
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
