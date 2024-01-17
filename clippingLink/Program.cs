using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.ComponentModel;
using clippingLink;
using System.Net;

namespace clippingLink;
class Program 
{
    public static void Main()
    {
        GetContent();
    }
    public static void GetContent()
    {
        Data data = new Data();
        string filePath = $"{Environment.CurrentDirectory}\\document.xlsx";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        FileInfo file = new(filePath);
        using (ExcelPackage package = new(file))
        {
            // Получение ссылки на рабочий лист
            ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

            // Проход по строкам и столбцам для извлечения данных
            int rowCount = workSheet.Dimension.Rows;
            int colCount = workSheet.Dimension.Columns;
            for (int row = 2; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col += 6)
                {
                    data.LogLink = (string?)workSheet.Cells[row, 5].Value;
                    data.ShortLink = (GetData(data.LogLink)).ToString();
                }
            }
        }
    }
    public static void ProcessingData()
    {

    }
    public static async Task<string> GetData(string LongLink)
    {
        string apiUrl = "https://clck.ru/"; // замените на URL вашего API

        using (HttpClient client = new HttpClient())
        {
            // Отправка GET-запроса с переменной в виде строки
            HttpResponseMessage response = await client.GetAsync($"{apiUrl}?inputString={LongLink}");

            // Проверка успешности запроса
            if (response.IsSuccessStatusCode)
            {
                string jsonResponse = await response.Content.ReadAsStringAsync();

                // Вывод результата
                return jsonResponse;
            }
            else
            {
                return "Ошибка при выполнении запроса";
            }
        }
    }
}
class Data
{
    public string? LogLink { get; set; }
    public string? ShortLink { get; set; }
}
