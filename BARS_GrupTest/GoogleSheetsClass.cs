using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace BARS_GrupTest
{
    internal static class GoogleSheetsClass
    {
        internal static void SendToGoogleSpreadsheets(List<BaseInfo> ListOfDataBases)
        {
            if (ListOfDataBases.Count == 0) return;

            // Получаем данные для подключения к GoogleAPI. Необходимо "Идентификатор клиента OAuth 2.0"
            DirectoryInfo dir = new DirectoryInfo(Path.Combine(Helper.CurrentDirectopy, "GoogleAPI"));

            foreach (var item in dir.GetFiles())
            {
                var Client_id = String.Empty;
                var SpreadsheetId = String.Empty;

                // Чтение и десериализация данных GoogleAPI                   
                string jsonFromFile = Helper.GetJsonFromFile(item.FullName);

                try
                {
                    GoogleAuthСredentials GAС = JsonConvert.DeserializeObject<GoogleAuthСredentials>(jsonFromFile);
                    if (GAС.installed == null)
                    {
                        Console.WriteLine($"Не корректный формат файла: {item.Name}");
                        continue;
                    }
                    else
                        Client_id = GAС.installed.client_id;
                }
                catch
                {
                    Console.WriteLine("Ошибка при получение данных для подключения к GoogleAPI");
                    Console.WriteLine($"Не корректный формат файла: {item.Name}");
                    continue;
                }

                string credPath = $"{Client_id}_Token.json";
                UserCredential credential = null;
                SheetsService service = null;
                Spreadsheet Spreadsheet = null;

                try
                {
                    credential = CreateCredential(item.FullName, credPath);

                    // Создание сервиса Google Sheets API
                    service = CreateService(credential);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при подключение к Google API: {ex.Message}");
                    // TODO Логирование
                    continue;
                }

                Spreadsheet = GetSpreadSheet(item, Client_id, ref SpreadsheetId, credPath, service);

                if (Spreadsheet != null)
                {
                    // Save SpreadSheetID
                    SaveSpreadSheetID(Client_id, Spreadsheet);

                    // Заполняем таблицу значениями
                    List<Request> requests = FillSpreadSheet(ListOfDataBases, Spreadsheet);

                    // Отправляем данные в Google
                    SaveChanges(SpreadsheetId, service, requests);
                    Console.WriteLine("Данные выгружены в файл.");
                }
                else
                {
                    Console.WriteLine($"Не удалось получить доступ к файлу в Google Sheets. client_id: {Client_id}");
                    continue;
                }
            }
        }

        private static void SaveSpreadSheetID(string client_id, Spreadsheet spreadsheet)
        {
            DirectoryInfo dir = new DirectoryInfo(Path.Combine(Helper.CurrentDirectopy, "Spreadsheets"));

            foreach (var item in dir.GetFiles())
            {
                string jsonFromFile = Helper.GetJsonFromFile(item.FullName);

                try
                {
                    var DBS = JsonConvert.DeserializeObject<DBSpreadsheet>(jsonFromFile);

                    if (DBS == null) continue;

                    if (DBS.client_id != null && DBS.client_id == client_id)
                    {
                        DBS.spreadsheetId = spreadsheet.SpreadsheetId;
                        DBS.spreadsheetName = spreadsheet.Properties.Title;
                        var jsonToFile = JsonConvert.SerializeObject(DBS);
                        Helper.SaveJsonToFile(item.FullName, jsonFromFile);
                        return;
                    }
                }
                catch { continue; }
            }

            var Spr = new DBSpreadsheet();
            Spr.client_id = client_id;
            Spr.spreadsheetId = spreadsheet.SpreadsheetId;
            Spr.spreadsheetName = spreadsheet.Properties.Title;
            var json = JsonConvert.SerializeObject(Spr);
            Helper.SaveJsonToFile(Path.Combine(dir.FullName, client_id) + "_SS.json", json);
        }

        private static void SaveChanges(string SpreadsheetId, SheetsService service, List<Request> requests)
        {
            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            var response = service.Spreadsheets.BatchUpdate(busr, SpreadsheetId).Execute();
        }

        private static List<Request> FillSpreadSheet(List<BaseInfo> ListOfDataBases, Spreadsheet Spreadsheet)
        {
            List<Request> requests = new List<Request>();
            var nextSheetId = Spreadsheet.Sheets.Max(c => c.Properties.SheetId) + 1;
            var Servers = ListOfDataBases.Select(c => c.server).Distinct().ToList();
            var FreeSpaces = CalculateFreeSpace(ListOfDataBases, Servers);
            var sheetId = 0;
            var row = 1;

            foreach (var serv in Servers)
            {
                var SelectedList = ListOfDataBases.Where(c => c.server == serv).ToList();
                sheetId = AddSheet(requests, serv, SelectedList, nextSheetId, Spreadsheet);
                if (sheetId == nextSheetId) nextSheetId++;
                AddValue(requests, sheetId, 0, 0, "Сервер");
                AddValue(requests, sheetId, 1, 0, "База данных");
                AddValue(requests, sheetId, 2, 0, "Размер в ГБ");
                AddValue(requests, sheetId, 3, 0, "Дата обновления");

                foreach (var dataBase in SelectedList)
                {
                    AddValue(requests, sheetId, 0, row, dataBase.server);
                    AddValue(requests, sheetId, 1, row, dataBase.dataBase);
                    AddValue(requests, sheetId, 2, row, ConvertToGB(dataBase.totalSize));
                    AddValue(requests, sheetId, 3, row, DateTime.Now.ToString("dd.MM.yyyy"));
                    row++;
                }

            }

            // Вывод итогов                    
            row = 1;
            sheetId = AddSheet(requests, "FreeSpace", Servers, nextSheetId, Spreadsheet);
            AddValue(requests, sheetId, 0, 0, "Сервер");
            AddValue(requests, sheetId, 1, 0, "Свободно");
            AddValue(requests, sheetId, 2, 0, "Размер в ГБ");
            AddValue(requests, sheetId, 3, 0, "Дата обновления");
            foreach (var space in FreeSpaces)
            {
                AddValue(requests, sheetId, 0, row, space.server);
                AddValue(requests, sheetId, 1, row, "Свободно");
                AddValue(requests, sheetId, 2, row, Math.Round(space.free, 3).ToString());
                AddValue(requests, sheetId, 3, row, DateTime.Now.ToString("dd.MM.yyyy"));
                row++;
            }

            DeleteSheets(requests, Spreadsheet);
            return requests;
        }

        private static Spreadsheet GetSpreadSheet(FileInfo item, string Client_id, ref string SpreadsheetId, string credPath, SheetsService service)
        {

            // Поиск файла по данным в папке Spreadsheet 
            // TODO Переделать
            Spreadsheet Spreadsheet = GetSpreadSheetFromFile(Client_id, service, credPath);
            if (Spreadsheet != null)
            {
                SpreadsheetId = Spreadsheet.SpreadsheetId;
            }

            if (SpreadsheetId == null)
            {
                // Поиск по имени ранее созданного файла в Google Drive
                // При отсутствие разрешения: "Google Drive API" - "../auth/drive" возвращает null                        
                SpreadsheetId = GoogleDriveClass.FindFileInGoogleDrive(item.FullName, credPath);
            }

            // Если нашли id файла - получаем файл, иначе создаем новый
            if (!String.IsNullOrEmpty(SpreadsheetId))
            {
                try
                {
                    Spreadsheet = service.Spreadsheets.Get(SpreadsheetId).Execute();
                }
                catch (Exception ex)
                {
                    // Возможно файла с таким id не существует, создаем новый
                    Console.WriteLine(ex.Message);
                    SpreadsheetId = null;
                }
            }


            if (String.IsNullOrEmpty(SpreadsheetId))
            {
                Spreadsheet = service.Spreadsheets.Create(new Spreadsheet { Properties = new SpreadsheetProperties() { Locale = "ru_RU", Title = Config.GoogleFileName } }).Execute();
                if (Spreadsheet != null)
                    SpreadsheetId = Spreadsheet.SpreadsheetId;
            }

            return Spreadsheet;
        }

        private static string ConvertToGB(string totalSize)
        {
            double size;
            if (double.TryParse(totalSize, out size))
            {
                return Math.Round(size / 1073741824, 3).ToString();
            }
            return totalSize;
        }

        private static void AddValue(List<Request> requests, int? sheetId, int column, int row, string value)
        {
            requests.Add(new Request
            {
                UpdateCells = new UpdateCellsRequest
                {
                    Start = new GridCoordinate
                    {
                        SheetId = sheetId,
                        ColumnIndex = column,
                        RowIndex = row
                    },
                    Rows = new List<RowData> {
                                new RowData {
                                    Values = new List<CellData> {
                                        new CellData{UserEnteredValue = new ExtendedValue {
                                            StringValue = value
                                        }}}}},
                    Fields = "userEnteredValue"
                }
            });
        }

        private static int AddSheet<T>(List<Request> requests, string serv, IList<T> SelectedList, int? sheetId, Spreadsheet spreadsheet)
        {
            // Если есть такая страница, то очищаем его, если нет - добавляем новую
            var mathingSheet = spreadsheet.Sheets.Where(c => c.Properties.Title == serv).FirstOrDefault();
            if (mathingSheet != null)
            {
                return ClearSheet(requests, spreadsheet, mathingSheet);
            }
            else
            {
                requests.Add(new Request
                {
                    AddSheet = new AddSheetRequest
                    {
                        Properties = new SheetProperties
                        {
                            SheetType = "GRID",
                            Title = serv,
                            SheetId = sheetId,
                            GridProperties = new GridProperties
                            {
                                ColumnCount = 4,
                                //RowCount = SelectedList.Count + 1
                            }
                        }
                    }
                });
                return sheetId.Value;
            }
        }

        private static int ClearSheet(List<Request> requests, Spreadsheet spreadsheet, Sheet mathingSheet)
        {
            requests.Add(new Request
            {
                UpdateCells = new UpdateCellsRequest
                {
                    Range = new GridRange { SheetId = mathingSheet.Properties.SheetId },
                    Fields = "*"
                }
            });
            spreadsheet.Sheets.Remove(mathingSheet);
            return mathingSheet.Properties.SheetId.Value;
        }

        private static void DeleteSheets(List<Request> requests, Spreadsheet spreadsheet)
        {
            foreach (var spr in spreadsheet.Sheets)
                requests.Add(new Request
                { DeleteSheet = new DeleteSheetRequest { SheetId = spr.Properties.SheetId } });
        }

        // Расчет свободного места на диске
        private static List<BaseInfo> CalculateFreeSpace(List<BaseInfo> listOfBases, List<string> Servers)
        {
            var result = new List<BaseInfo>();

            foreach (var serv in Servers)
            {
                double sum = 0;
                double diskSize = 0;
                foreach (var b in listOfBases.Where(c => c.server == serv))
                {
                    double size;
                    if (double.TryParse(b.totalSize, out size))
                    {
                        sum = sum + size;
                    }
                    diskSize = b.diskSize;
                }
                result.Add(new BaseInfo { server = serv, free = diskSize - (sum / 1073741824) });
            }
            return result;
        }

        private static SheetsService CreateService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = Config.ApplicationName,
            });
        }

        private static UserCredential CreateCredential(string FileName, string credPath)
        {
            using (var stream = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                        GoogleClientSecrets.Load(stream).Secrets,
                                                    Scopes,
                                                    "user",
                                                    CancellationToken.None,
                                                    new FileDataStore(credPath, true)).Result;
                return credential;
            }
        }

        private static Spreadsheet GetSpreadSheetFromFile(string client_id, SheetsService service, string credPath)
        {
            DirectoryInfo dir = new DirectoryInfo(Path.Combine(Helper.CurrentDirectopy, "Spreadsheets"));

            foreach (var item in dir.GetFiles())
            {
                string jsonFromFile = Helper.GetJsonFromFile(item.FullName);

                try
                {
                    var DBS = JsonConvert.DeserializeObject<DBSpreadsheet>(jsonFromFile);

                    if (DBS == null) continue;

                    if (DBS.client_id != null && DBS.client_id == client_id)
                    {
                        // Сперва ищем по spreadsheetId
                        if (DBS.spreadsheetId != null && DBS.spreadsheetId.Length > 30)
                        {
                            try
                            {
                                var Spreadsheet = service.Spreadsheets.Get(DBS.spreadsheetId).Execute();
                                if (Spreadsheet != null)
                                    return Spreadsheet;
                            }
                            catch (Exception ex)
                            {
                                // Возможно файла с таким id не существует
                                // TODO В зависимости от вида исключения очищать поле spreadsheetId                                
                            }
                        }
                        // Потом ищем по имени файла
                        if (!String.IsNullOrWhiteSpace(DBS.spreadsheetName))
                        {
                            var SpreadsheetId = GoogleDriveClass.FindFileInGoogleDrive(DBS.spreadsheetName, credPath);
                            if (SpreadsheetId != null)
                            {
                                var Spreadsheet = service.Spreadsheets.Get(SpreadsheetId).Execute();
                                if (Spreadsheet != null)
                                    return Spreadsheet;
                            }
                        }
                    }
                }
                catch { continue; }
            }
            return null;
        }
    }
}
