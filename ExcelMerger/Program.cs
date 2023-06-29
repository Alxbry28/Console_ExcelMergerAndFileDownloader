using ClosedXML.Excel;
using System.Net;
using System;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;
// See https://aka.ms/new-console-template for more information

const string excelToMerge = "ortracking_IPS_SC_";
const string folderPath = @"C:\Users\AQArellano\Desktop\KP-63689_06-27-2023";
const string folderToSaveImage = "Image";
DirectoryInfo directoryOrigin = new DirectoryInfo(folderPath);
DirectoryInfo[] directories = directoryOrigin.GetDirectories();

foreach (DirectoryInfo directory in directories)
{
    Console.WriteLine($"Folder: {directory.Name}");
    FileInfo[] files = directory.GetFiles();

    string csvFile = files[0].FullName;
    string excelFile = files[2].FullName;

    List<Dictionary<string, string[]>> list = new List<Dictionary<string, string[]>>();
    using (var reader = new StreamReader(csvFile))
    {
        Console.WriteLine($"Reading {csvFile}.  Loading....");
        int counter = 0;
        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            string[] values = line!.Split(',');

            Dictionary<string, string[]> keyValuePairs = new Dictionary<string, string[]>();
            string[] datas = new string[] { values[0], values[1], values[2] };
            keyValuePairs.Add(values[0], datas);
            list.Add(keyValuePairs);
            counter++;

            string? downloadMessage = await downloadURLFile(values[2], directory.FullName, folderToSaveImage, counter);
            Console.WriteLine(downloadMessage);

        }

        Console.WriteLine($"Done Reading {csvFile} and downloaded files: {counter}");
        Console.WriteLine();
    }

    // Merge to ortracking_IPS_SC_
    using (XLWorkbook wb = new XLWorkbook(excelFile))
    {
        var ws = wb.Worksheet(1);
        var range = ws.RangeUsed();
        int rowCount = range.RowCount();
        var columnReference = "B";
        var columnToFill = "P";

        Console.WriteLine($"Reading {excelFile}. Loading....");
        for (int i = 2; i < rowCount; i++)
        {
            string cellToRead = $"{columnReference}{i}";
            var refData = ws.Cell(cellToRead).GetValue<string>();

            Dictionary<string, string[]>? valuePairs = list.Where(x => x.ContainsKey(refData)).FirstOrDefault();
            string data = valuePairs![refData][2];

            string cellToFill = $"{columnToFill}{i}";
            ws.Cell(cellToFill).Value = data;
            wb.Save();

        }
        Console.WriteLine("Done Saving to " + excelFile);
        list = null;
    }

    Console.WriteLine();
}

async Task<string?> downloadURLFile(string urlFile, string path, string folder, int? counter = null)
{
    using (WebClient client = new WebClient())
    {
        Uri uriFile = new Uri(urlFile);
        string filename = Path.GetFileName(uriFile.AbsolutePath);
        string folderPathToSave = $@"{path}\{folder}";
        if (!Directory.Exists(folderPathToSave)) Directory.CreateDirectory(folderPathToSave);

        string fullFilename = $@"{folderPathToSave}\{filename}";
        await client.DownloadFileTaskAsync(uriFile, fullFilename);
        string message = (counter is null) ? $"done downloading {filename}" : $"{counter}: done downloading {filename}";
        if (File.Exists(fullFilename)) return message;
        return null;
    }
}

List<Dictionary<string, string[]>> readCSVFile(string idReference, string csvFilePath)
{
    List<Dictionary<string, string[]>> list = new List<Dictionary<string, string[]>>();
    using (var reader = new StreamReader(csvFilePath))
    {
        Console.WriteLine($"Reading {csvFilePath}.  Loading....");
        int counter = 0;
        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            string[] values = line!.Split(',');

            Dictionary<string, string[]> keyValuePairs = new Dictionary<string, string[]>();
            string[] datas = values;
            keyValuePairs.Add(values[0], datas);
            list.Add(keyValuePairs);
            counter++;
        }

        Console.WriteLine($"Done Reading {csvFilePath}");
        Console.WriteLine();
    }
    return list;
}

List<Dictionary<string, string[]>> readExcelFile(string idReference, string csvFilePath)
{
    List<Dictionary<string, string[]>> list = new List<Dictionary<string, string[]>>();
    using (var reader = new StreamReader(csvFilePath))
    {
        Console.WriteLine($"Reading {csvFilePath}.  Loading....");
        int counter = 0;
        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            string[] values = line!.Split(',');

            Dictionary<string, string[]> keyValuePairs = new Dictionary<string, string[]>();
            string[] datas = values;
            keyValuePairs.Add(values[0], datas);
            list.Add(keyValuePairs);
            counter++;
        }

        Console.WriteLine($"Done Reading {csvFilePath}");
        Console.WriteLine();
    }
    return list;
}