using System;
using System.Data.SQLite;
using ClosedXML.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace sqlToExcel
{
    public class SqlToExcelConverter
    {
        private const int BatchSize = 1000;

        public void ConvertSqlFileToExcel(string sqlFilePath, string excelFilePath)
        {
            ValidateInputPaths(sqlFilePath, excelFilePath);

            try
            {
                using var connection = InitializeInMemoryDatabase(sqlFilePath);
                var tables = GetAllTables(connection);

                if (!tables.Any())
                {
                    throw new InvalidOperationException("Nessuna tabella trovata nel database");
                }

                EnsureExcelFileNotLocked(excelFilePath);

                using var workbook = new XLWorkbook();
                foreach (var table in tables)
                {
                    ExportTableToExcel(table, connection, workbook);
                    Console.WriteLine($"Tabella {table} esportata con successo");
                }

                workbook.SaveAs(excelFilePath);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Errore durante la conversione: {ex.Message}", ex);
            }
        }

        private void ValidateInputPaths(string sqlFilePath, string excelFilePath)
        {
            if (string.IsNullOrWhiteSpace(sqlFilePath))
                throw new ArgumentException("Il percorso del file SQL non può essere vuoto");

            if (string.IsNullOrWhiteSpace(excelFilePath))
                throw new ArgumentException("Il percorso del file Excel non può essere vuoto");

            if (!File.Exists(sqlFilePath))
                throw new FileNotFoundException("File SQL non trovato", sqlFilePath);

            var excelDir = Path.GetDirectoryName(excelFilePath);
            if (!Directory.Exists(excelDir))
                throw new DirectoryNotFoundException($"Directory di destinazione non trovata: {excelDir}");
        }

        private void EnsureExcelFileNotLocked(string excelFilePath)
        {
            if (File.Exists(excelFilePath))
            {
                try
                {
                    using var fs = File.Open(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                catch (IOException)
                {
                    throw new IOException("Il file Excel di destinazione è già aperto in un altro processo");
                }
            }
        }

        private SQLiteConnection InitializeInMemoryDatabase(string sqlFilePath)
        {
            var connection = new SQLiteConnection("Data Source=:memory:");

            try
            {
                connection.Open();
                string sqlContent = File.ReadAllText(sqlFilePath);
                string[] sqlCommands = Regex.Split(sqlContent, @";(\r?\n|$)")
                                          .Where(cmd => !string.IsNullOrWhiteSpace(cmd))
                                          .ToArray();

                foreach (string command in sqlCommands)
                {
                    if (!string.IsNullOrWhiteSpace(command))
                    {
                        using var cmd = new SQLiteCommand(command, connection);
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (SQLiteException ex)
                        {
                            throw new SQLiteException($"Errore nell'esecuzione del comando SQL: {ex.Message}\nComando: {command}", ex);
                        }
                    }
                }

                return connection;
            }
            catch
            {
                connection.Dispose();
                throw;
            }
        }

        private List<string> GetAllTables(SQLiteConnection connection)
        {
            var tables = new List<string>();
            using var cmd = new SQLiteCommand(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';",
                connection);
            using var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                tables.Add(reader.GetString(0));
            }

            return tables;
        }

        private void ExportTableToExcel(string tableName, SQLiteConnection connection, XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add(tableName);

            using var command = new SQLiteCommand($"SELECT * FROM [{tableName}]", connection);
            using var reader = command.ExecuteReader();

            for (int i = 0; i < reader.FieldCount; i++)
            {
                worksheet.Cell(1, i + 1).Value = reader.GetName(i);
            }

            int row = 2;
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    if (!reader.IsDBNull(i))
                    {
                        switch (reader.GetFieldType(i).Name.ToLower())
                        {
                            case "int32":
                            case "int64":
                                worksheet.Cell(row, i + 1).Value = reader.GetInt64(i);
                                break;
                            case "double":
                            case "decimal":
                                worksheet.Cell(row, i + 1).Value = reader.GetDouble(i);
                                break;
                            case "datetime":
                                worksheet.Cell(row, i + 1).Value = reader.GetDateTime(i);
                                break;
                            case "boolean":
                                worksheet.Cell(row, i + 1).Value = reader.GetBoolean(i);
                                break;
                            default:
                                worksheet.Cell(row, i + 1).Value = reader.GetString(i);
                                break;
                        }
                    }
                }
                row++;

                if (row % BatchSize == 0)
                {
                    worksheet.Range(row - BatchSize, 1, row - 1, reader.FieldCount).Style
                            .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                }
            }

            var tableRange = worksheet.Range(1, 1, row - 1, reader.FieldCount);
            var table = tableRange.CreateTable();

            worksheet.Row(1).Style
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.LightGray);

            worksheet.Columns().AdjustToContents(1, 100);
        }
    }
}