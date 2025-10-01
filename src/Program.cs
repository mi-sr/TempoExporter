using Microsoft.Extensions.Configuration;
using System.Globalization;
using Tempo.Exporter;

Console.WriteLine("Reading configuration...");
var config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("config.json", optional: false, reloadOnChange: true)
    .AddJsonFile("config.Development.json", optional: true)
    .Build();

string apiKey = config["Configuration:ApiKey"];
if (string.IsNullOrWhiteSpace(apiKey))
{
    Console.WriteLine("The API key is not defined. Please add the value in the 'config.json' file.");
    Console.ReadLine();
    return;
}

string accountId = config["Configuration:AccountId"];
if (string.IsNullOrWhiteSpace(accountId))
{
    Console.WriteLine("The account id is not defined. Please add the value in the 'config.json' file.");
    Console.ReadLine();
    return;
}

string folderPath = config["Configuration:TimesheetFolderPath"];
if (string.IsNullOrWhiteSpace(folderPath))
{
    Console.WriteLine("The folder path is not defined. Please add the value in the 'config.json' file.");
    Console.ReadLine();
    return;
}

DateTime from;
DateTime to;

while (true)
{
    Console.WriteLine("Please enter:");
    Console.WriteLine(" * (default) \"y\" to export the previous month");
    Console.WriteLine(" * The desired year/month (Format: yyyy-MM)");
    Console.WriteLine(" * \"x\" to exit");

    var input = Console.ReadLine();

    if (input.ToUpperInvariant() == "X")
    {
        return;
    }
    else if (input.ToUpperInvariant() == "Y" || string.IsNullOrWhiteSpace(input))
    {
        from = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(-1);
        to = from.AddMonths(1).AddDays(-1);
        break;
    }
    else if (DateTime.TryParseExact(input, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var yearMonth))
    {
        from = yearMonth;
        to = from.AddMonths(1).AddDays(-1);
        break;
    }
    else
    {
        Console.WriteLine("Unknown input. Please try again.");
        continue;
    }
}

var tempoClient = new TempoClient(accountId, apiKey);
var arbeitszeiten = await tempoClient.GetWorklogsAsync(from, to);

var excelFileSelector = new ConsoleFileSelector(folderPath);
var path = excelFileSelector.SelectExcelFile();

if (string.IsNullOrWhiteSpace(path))
{
    Console.WriteLine("No excel file has been selected. Aborting...");
    Console.ReadLine();
    return;
}

var timeSheet = new TimeSheet(path);
timeSheet.ExportWorkingTime(arbeitszeiten);

Console.WriteLine("Tempo time has been exported. Press Enter to exit.");
Console.ReadLine();
