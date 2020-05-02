using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace BARS_GrupTest
{
    internal static class GoogleDriveClass
    {
        internal static string FindFileInGoogleDrive(string fileName, string credPath)
        {
            // У аккаунта должны быть разрешение: "Google Drive API" - "../auth/drive"            
            try
            {
                UserCredential credential = CreateCredential(fileName, credPath);
                // Создание сервиса Drive API 
                DriveService service = CreateService(credential);

                // Установка параметров запроса
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.PageSize = 100;
                listRequest.Fields = "files(id, name)";
                listRequest.Q = ("trashed = false");

                // Перебор файлов, при нахождение файла с соответствующим именем возврат id файла
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;

                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        if (file.Name.Trim().ToLower() == Config.GoogleFileName.Trim().ToLower())
                            return file.Id;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка Google Drive: {ex.Message}");
                // TODO Добавить логирование. 
            }

            return null;
        }

        private static DriveService CreateService(UserCredential credential)
        {
            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = Config.ApplicationName,
            });
        }

        private static UserCredential CreateCredential(string fileName, string credPath)
        {
            using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                string[] Scopes = { DriveService.Scope.DriveFile };
                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                                GoogleClientSecrets.Load(stream).Secrets,
                                Scopes,
                                "user",
                                CancellationToken.None,
                                new FileDataStore(credPath, true)).Result;
            }
        }
    }
}
