using Newtonsoft.Json;
using Npgsql;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace BARS_GrupTest
{
    class Program
    {
        public static Thread MThread = new Thread(MainFunction) { IsBackground = true };
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("-----------------------------------------------" + Environment.NewLine +
                                  "Выберите действие:" + Environment.NewLine +
                                  "0 - Остановить выполнение программы и выйти." + Environment.NewLine +
                                  "1 - Запустить выгрузку." + Environment.NewLine +
                                  "2 - Установить таймаут для повтора выгрузки.(в минутах)." + Environment.NewLine +
                                  "-----------------------------------------------"
                                  );
                var Key = Console.ReadLine();
                switch (Key)
                {
                    case "0":
                        return;
                        break;
                    case "1":
                        MThread.Start();
                        break;
                    case "2":
                        Console.WriteLine("Введите данные:");
                        var number = Console.ReadLine();
                        int result;
                        if (Int32.TryParse(number.Trim(), out result))
                        {
                            Config.Timeout = result;
                            Console.WriteLine("Новые значение установлено.");
                        }
                        else
                        {
                            Console.WriteLine("Введены не корректные данные.");
                        }
                        break;
                }
            }
        }

        private static void MainFunction()
        {
            while (true)
            {
                Console.WriteLine("Начинаем выгрузку.");
                var ListOfBases = new List<BaseInfo>();
                CheckAndCreateDirectory();

                // Получаем конфигурационные файлы для подключения к БД
                DirectoryInfo dir = new DirectoryInfo(Path.Combine(Helper.CurrentDirectopy, "Databases"));

                if (dir.GetFiles().Length == 0)
                {
                    Console.WriteLine("Не найдено ни одного конфигурационного файла с данными для подключения к БД.");
                    Console.ReadKey();
                    return;
                }

                foreach (var item in dir.GetFiles())
                {
                    using (var stream = new FileStream(item.FullName, FileMode.Open, FileAccess.Read))
                    {
                        // Получение данныых для подключения к БД
                        string jsonFromFile = Helper.GetJsonFromFileStream(stream);
                        DBInfo dbinfo = null;


                        try
                        {
                            dbinfo = JsonConvert.DeserializeObject<DBInfo>(jsonFromFile);
                        }
                        catch
                        {
                            Console.WriteLine($"Не удалось загрузить файл конфигурации.");
                            Console.WriteLine($"Не корректный формат файла: {item.Name}" + Environment.NewLine);
                            continue;
                        }

                        if (dbinfo != null)
                        {
                            GetDatabaseInfo(ListOfBases, item, dbinfo);
                        }
                    }
                }

                // Добавление данных в Google Sheets
                GoogleSheetsClass.SendToGoogleSpreadsheets(ListOfBases);

                Thread.Sleep(Config.Timeout * 60000);
            }
        }

        private static void GetDatabaseInfo(List<BaseInfo> ListOfBases, FileInfo item, DBInfo dbinfo)
        {
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection($"Server={dbinfo.server};User Id={dbinfo.user}; Password={dbinfo.password};")) // "Server=localhost;User Id=postgres; Password=123456;"
                {
                    conn.Open();
                    NpgsqlCommand command = CreateQuery(conn);

                    // Получаем информацию и сохраняем в список
                    NpgsqlDataReader dr = command.ExecuteReader();

                    while (dr.Read())
                    {
                        ListOfBases.Add(new BaseInfo { server = dbinfo.server, dataBase = dr[0].ToString(), totalSize = dr[1].ToString(), diskSize = dbinfo.diskSize });
                    }

                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при подключении к базе данных: {ex.Message}");
                Console.WriteLine($"Проверьте данные в файле \"{item.Name}\"");
            }
        }

        private static NpgsqlCommand CreateQuery(NpgsqlConnection conn)
        {
            return new NpgsqlCommand(@"SELECT d.datname AS Name, " +
                                                                                      @"CASE WHEN pg_catalog.has_database_privilege(d.datname, 'CONNECT') " +
                                                                                      @"THEN pg_catalog.pg_database_size(d.datname) " +
                                                                                      "ELSE 0 " +
                                                                                      //@"THEN pg_catalog.pg_size_pretty(pg_catalog.pg_database_size(d.datname)) " +
                                                                                      //@"ELSE 'No Access' " +
                                                                                      "END AS SIZE " +
                                                                                      "FROM pg_catalog.pg_database d " +
                                                                                      "ORDER BY " +
                                                                                      @"CASE WHEN pg_catalog.has_database_privilege(d.datname, 'CONNECT') " +
                                                                                      "THEN pg_catalog.pg_database_size(d.datname) " +
                                                                                      "ELSE NULL END DESC", conn);
        }

        private static void CheckAndCreateDirectory()
        {
            if (!Directory.Exists(Path.Combine(Helper.CurrentDirectopy, "Databases")))
            {
                Directory.CreateDirectory(Path.Combine(Helper.CurrentDirectopy, "Databases"));
            }
            if (!Directory.Exists(Path.Combine(Helper.CurrentDirectopy, "GoogleAPI")))
            {
                Directory.CreateDirectory(Path.Combine(Helper.CurrentDirectopy, "GoogleAPI"));
            }
            if (!Directory.Exists(Path.Combine(Helper.CurrentDirectopy, "Spreadsheets")))
            {
                Directory.CreateDirectory(Path.Combine(Helper.CurrentDirectopy, "Spreadsheets"));
            }
        }
    }
}
