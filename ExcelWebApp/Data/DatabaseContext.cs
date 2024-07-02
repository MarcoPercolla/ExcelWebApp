using System;
using System.Collections.Generic;
using System.Data;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace ExcelWebApp.Data
{
    public class DatabaseContext
    {
        private readonly string _connectionString;

        public DatabaseContext(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        private MySqlConnection GetConnection()
        {
            return new MySqlConnection(_connectionString);
        }

        public void ProcessExcelFile(string filePath)

        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var connection = GetConnection())
            {
                connection.Open();

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var sheet in package.Workbook.Worksheets)
                    {
                        switch (sheet.Name)
                        {
                            case "Utenti":
                                CreateTableIfNotExistsUtenti(connection, sheet);
                                InsertOrUpdateUtenti(connection, sheet);
                                break;
                            case "Indirizzi":
                                CreateTableIfNotExistsIndirizzi(connection, sheet);
                                InsertOrUpdateIndirizzi(connection, sheet);
                                break;
                            case "Abbonamenti":
                                CreateTableIfNotExistsAbbonamenti(connection, sheet);
                                InsertOrUpdateAbbonamenti(connection, sheet);
                                break;
                        }
                    }
                }
            }
        }

        private void CreateTableIfNotExistsUtenti(MySqlConnection conn, ExcelWorksheet sheet)
        {
            string createTableSql = "CREATE TABLE IF NOT EXISTS Utenti (" +
                                    "id INT PRIMARY KEY, " +
                                    "nome VARCHAR(255), " +
                                    "cognome VARCHAR(255), " +
                                    "eta INT, " +
                                    "indirizzo INT, " +
                                    "abbonamento INT);";

            using (MySqlCommand cmd = new MySqlCommand(createTableSql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertOrUpdateUtenti(MySqlConnection conn, ExcelWorksheet sheet)
        {
            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                var values = new List<object>();
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    values.Add(sheet.Cells[row, col].Text);
                }

                string sql = "INSERT INTO Utenti (id, nome, cognome, eta, indirizzo, abbonamento) " +
                             "VALUES (@id, @nome, @cognome, @eta, @indirizzo, @abbonamento) " +
                             "ON DUPLICATE KEY UPDATE nome = VALUES(nome), cognome = VALUES(cognome), eta = VALUES(eta), " +
                             "indirizzo = VALUES(indirizzo), abbonamento = VALUES(abbonamento);";

                using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", values[0]);
                    cmd.Parameters.AddWithValue("@nome", values[1]);
                    cmd.Parameters.AddWithValue("@cognome", values[2]);

                    int eta;
                    if (int.TryParse(values[3]?.ToString(), out eta))
                    {
                        cmd.Parameters.AddWithValue("@eta", eta);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@eta", DBNull.Value);
                    }

                    cmd.Parameters.AddWithValue("@indirizzo", values[4]);

                    int abbonamento;
                    if (int.TryParse(values[5]?.ToString(), out abbonamento))
                    {
                        cmd.Parameters.AddWithValue("@abbonamento", abbonamento);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@abbonamento", DBNull.Value);
                    }

                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void CreateTableIfNotExistsIndirizzi(MySqlConnection conn, ExcelWorksheet sheet)
        {
            string createTableSql = "CREATE TABLE IF NOT EXISTS Indirizzi (" +
                                    "id INT PRIMARY KEY, " +
                                    "citta VARCHAR(255), " +
                                    "via VARCHAR(255), " +
                                    "civico INT);";

            using (MySqlCommand cmd = new MySqlCommand(createTableSql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertOrUpdateIndirizzi(MySqlConnection conn, ExcelWorksheet sheet)
        {
            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                var values = new List<object>();
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    values.Add(sheet.Cells[row, col].Text);
                }

                string sql = "INSERT INTO Indirizzi (id, citta, via, civico) VALUES (@id, @citta, @via, @civico) " +
                             "ON DUPLICATE KEY UPDATE citta = VALUES(citta), via = VALUES(via), civico = VALUES(civico);";

                using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", values[0]);
                    cmd.Parameters.AddWithValue("@citta", values[1]);
                    cmd.Parameters.AddWithValue("@via", values[2]);
                    cmd.Parameters.AddWithValue("@civico", values[3]);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void CreateTableIfNotExistsAbbonamenti(MySqlConnection conn, ExcelWorksheet sheet)
        {
            string createTableSql = "CREATE TABLE IF NOT EXISTS Abbonamenti (" +
                                    "id INT PRIMARY KEY, " +
                                    "nome VARCHAR(255));";

            using (MySqlCommand cmd = new MySqlCommand(createTableSql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertOrUpdateAbbonamenti(MySqlConnection conn, ExcelWorksheet sheet)
        {
            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                var values = new List<object>();
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    values.Add(sheet.Cells[row, col].Text);
                }

                string sql = "INSERT INTO Abbonamenti (id, nome) VALUES (@id, @nome) " +
                             "ON DUPLICATE KEY UPDATE nome = VALUES(nome);";

                using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", values[0]);
                    cmd.Parameters.AddWithValue("@nome", values[1]);

                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
