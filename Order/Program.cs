using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Amazon;
using Amazon.S3;
using Amazon.S3.Transfer;
using RestSharp;
using Google.Cloud.Storage.V1;
using System.Data;
using System.Data.SqlClient;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Data.Odbc;
using Order.System.Io;
using System.Threading;
using iText.Signatures;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;
using System.Net.Http;
using iText.Commons.Utils;
using System.Security.Permissions;


namespace Order
{
    internal class Program
    {
        private static string connectionString = $"Server=win-8db95k1hu90\\sqlexpress;Database=VStest;Integrated Security=True;";

        // SQL-запрос
        private static string query = "SELECT [id],[modelId],[documentTypeId],[siteId],[status],[localPath],[link] FROM [DocumentDraft] WHERE [status] IS NULL AND [documentTypeId] = 2";
        private static int wordCount = 0;
        private static Timer timer;
        private static int[] id;
        private static int[] modelId;
        private static int[] documentTypeId;
        private static int[] siteId;
        private static string[] status;
        private static string[] localPath;
        private static string[] link;

        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            timer = new Timer(Check, null, 0, 10000); 
            Console.ReadLine();
        }
        public static string SanitizeFileName(string fileName)
        {
            fileName = fileName.TrimStart('\\');
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string sanitized = Regex.Replace(nameWithoutExtension, @"[\s\p{P}-[_\-]]", "_");
            return sanitized + extension;
        }
        static string ExtractTextFromPdf(string filePath)
        {
            StringBuilder text = new StringBuilder();

            using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(filePath)))
            {
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    PdfPage page = pdfDoc.GetPage(i);
                    string pageText = PdfTextExtractor.GetTextFromPage(page);
                    text.Append(pageText);
                }
            }

            return text.ToString();
        }
        static int CountWords(string text)
        {
            // Удаляем знаки пунктуации и лишние пробелы
            string cleanedText = Regex.Replace(text, @"[\p{P}]?\s", " ").Trim();

            // Разделяем текст на слова
            string[] words = cleanedText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            // Возвращаем количество слов
            return words.Length;
        }
        private static async void Check(object state)
        {
            
            // Создание подключения
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                
                try
                {
                    // Открываем соединение
                    await connection.OpenAsync();

                    // Создаем команду для выполнения SQL-запроса
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        id = new int[dbStringCount()];
                        modelId = new int[dbStringCount()];
                        documentTypeId = new int[dbStringCount()];
                        siteId = new int[dbStringCount()];
                        status = new string[dbStringCount()];
                        localPath = new string[dbStringCount()];
                        link = new string[dbStringCount()];
                        // Выполняем запрос и получаем данные
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            int i = 0;
                            // Чтение данных
                            while (reader.Read() && i < dbStringCount())
                            {
                                
                                id[i] = reader.GetInt32(0);
                                modelId[i] = reader.GetInt32(1);
                                documentTypeId[i] = reader.GetInt32(2);
                                siteId[i] = reader.GetInt32(3);
                                try { status[i] = reader.GetString(4); } catch { status[i] = "-1"; }
                                try { localPath[i] = reader.GetString(5); } catch { localPath[i] = "-1"; }
                                try { link[i] = reader.GetString(6); } catch { link[i] = "-1"; }
                                i++;
                            }

                        }
                        
                    }
                    for (int i = 0; i < id.Length; i++)
                    {
                        wordCount = 0;
                        bool isProcessed = false;
                        Console.WriteLine($"ID: {id[i]}; model: {modelId[i]}; siteId: {siteId[i]}; status: {status[i]}; localPath: {localPath[i]}");
                        SqlCommand status_work = new SqlCommand($"UPDATE [DocumentDraft] SET [status] = 1 WHERE [status] IS NULL AND [documentTypeId] = 2 AND [id] = {id[i]}", connection);
                        SqlCommand status_processed = new SqlCommand($"UPDATE [DocumentDraft] SET [status] = 2 WHERE [status] = 1 AND [documentTypeId] = 2 AND [id] = {id[i]}", connection);
                        SqlCommand status_not_access = new SqlCommand($"UPDATE [DocumentDraft] SET [status] = 5 WHERE [status] = 1 AND [documentTypeId] = 2 AND [id]={id[i]}", connection);
                        SqlCommand status_not_loaded = new SqlCommand($"UPDATE [DocumentDraft] SET [status] = 6 WHERE [status] = 1 AND [documentTypeId] = 2 AND [id]={id[i]}", connection);
                        SqlCommand status_not_open = new SqlCommand($"UPDATE [DocumentDraft] SET [status] = 7 WHERE [status] = 1 AND [documentTypeId] = 2 AND [id]={id[i]}", connection);

                        status_work.ExecuteNonQuery();
                        if (localPath[i] != "-1")
                        {
                            if (IsPdfFileAccessible(localPath[i], id[i]) == 1)
                            {
                                FileInfo file = new FileInfo(localPath[i]);
                                string text = ExtractTextFromPdf(localPath[i]);
                                wordCount = CountWords(text);
                                file.Rename(SanitizeFileName(localPath[i]));
                                localPath[i] = file.FullName;
                                isProcessed = true;
                            }
                            else if (IsPdfFileAccessible(localPath[i], id[i]) == 2) status_not_access.ExecuteNonQuery();
                            else status_not_open.ExecuteNonQuery();
                        }
                        else if (link[i] != "-1")
                        {
                            string fileName = GetFileNameFromUrl(link[i]);
                            string savePath = $"D:\\Sources\\{siteId[i]}\\{fileName}";
                            bool isDownlodaded = await DownloadFileAsync(link[i], savePath, id[i]);
                            if (isDownlodaded)
                            {
                                if (IsPdfFileAccessible(savePath, id[i]) == 1)
                                {
                                    FileInfo file = new FileInfo(savePath);
                                    string text = ExtractTextFromPdf(savePath);
                                    wordCount = CountWords(text);
                                    file.Rename(SanitizeFileName(savePath));
                                    savePath = file.FullName;
                                    isProcessed = true;
                                }
                                else if (IsPdfFileAccessible(savePath, id[i]) == 2) status_not_access.ExecuteNonQuery();
                                else status_not_open.ExecuteNonQuery();
                            }
                            else status_not_loaded.ExecuteNonQuery();
                        }
                        else status_not_access.ExecuteNonQuery();
                        if (isProcessed)
                        {
                            SqlCommand query_modelId_out = new SqlCommand($"UPDATE [Document] SET [modelId]={modelId[i]} WHERE [id]={id[i]}", connection);
                            SqlCommand query_localPath_out = new SqlCommand($"UPDATE [Document] SET [localPath]='{localPath[i]}' WHERE [id]={id[i]}", connection);
                            SqlCommand query_words_out = new SqlCommand($"UPDATE [Document] SET [words]={wordCount} WHERE [id]={id[i]}", connection);
                            status_processed.ExecuteNonQuery();
                            query_modelId_out.ExecuteNonQuery();
                            query_localPath_out.ExecuteNonQuery();
                            query_words_out.ExecuteNonQuery();
                        }
                        Console.WriteLine(isProcessed);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }
        }
        static int IsPdfFileAccessible(string filePath, int id)
        {
            // Проверка существования файла
            if (!File.Exists(filePath))
            {
                
                return 2; // нет доступа
            }

            try
            {
                // Попытка открыть PDF-файл с помощью iTextSharp
                using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(filePath)))
                {
                    // Проверка, содержит ли файл хотя бы одну страницу
                    if (pdfDoc.GetNumberOfPages() > 0)
                    {
                        return 1; //всё в норме
                    }
                    else
                    {
                        return 2; //нет доступа
                    }
                }
            }
            catch 
            {               
                return 3; //не открывается
            }
        }
        static async Task<bool> DownloadFileAsync(string fileUrl, string savePath, int id)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                try
                {
                    // Выполняем GET-запрос для загрузки файла
                    HttpResponseMessage response = await httpClient.GetAsync(fileUrl, HttpCompletionOption.ResponseHeadersRead);

                    // Проверяем, успешен ли запрос
                    response.EnsureSuccessStatusCode();

                    // Читаем содержимое файла
                    using (Stream contentStream = await response.Content.ReadAsStreamAsync(),
                                 fileStream = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        await contentStream.CopyToAsync(fileStream);
                    }

                    return true;
                }
                catch 
                {
                    return false;
                }
            }
        }
        public static int dbStringCount()
        {

            string query = "SELECT COUNT(*) FROM [DocumentDraft] WHERE [status] IS NULL AND [documentTypeId] = 2";
            int count = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                
                // Создаем команду для выполнения SQL-запроса
                using (SqlCommand command = new SqlCommand(query, connection))
                    {
                    connection.Open();
                    count = (int)command.ExecuteScalar();
                    }
            }
            
            return count;
        }
        static string GetFileNameFromUrl(string url)
        {
            // Убираем параметры запроса (всё, что после "?")
            int queryIndex = url.IndexOf('?');
            if (queryIndex != -1)
            {
                url = url.Substring(0, queryIndex);
            }

            // Извлекаем имя файла из последнего сегмента URL
            int lastSlashIndex = url.LastIndexOf('/');
            if (lastSlashIndex != -1)
            {
                return url.Substring(lastSlashIndex + 1);
            }

            return url; // Если слешей нет, возвращаем весь URL
        }
    }
    namespace System.Io
    {
        public static class ExtendedMethod
        {
            public static void Rename(this FileInfo fileinfo, string SanitizeFileName)
            {
                fileinfo.MoveTo(fileinfo.Directory.FullName + "\\" + SanitizeFileName);
            }
        }
    }
}
