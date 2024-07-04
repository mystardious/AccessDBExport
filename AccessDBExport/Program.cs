using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Access;

class Program
{
    private static void ExportQuery(string databaseLocation, string queryNameToExport, string locationToExportTo)
    {
        var application = new Application();
        application.OpenCurrentDatabase(databaseLocation);
        application.DoCmd.TransferSpreadsheet(AcDataTransferType.acExport, AcSpreadSheetType.acSpreadsheetTypeExcel12,
                                              queryNameToExport, locationToExportTo, true);
        application.CloseCurrentDatabase();
        application.Quit();
        Marshal.ReleaseComObject(application);
    }

    private static void ShowHelp()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  AccessDBExport --db <DatabaseLocation> --query <QueryName> --export <ExportLocation>");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --db      Path to the Access database file.");
        Console.WriteLine("  --query   Name of the query to export.");
        Console.WriteLine("  --export  Path to export the query result to an Excel file.");
        Console.WriteLine("  --help    Display this help message.");
    }

    public static void Main(string[] args)
    {
        string databaseLocation = null;
        string queryNameToExport = null;
        string locationToExportTo = null;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--db":
                    if (i + 1 < args.Length) databaseLocation = args[++i];
                    break;
                case "--query":
                    if (i + 1 < args.Length) queryNameToExport = args[++i];
                    break;
                case "--export":
                    if (i + 1 < args.Length) locationToExportTo = args[++i];
                    break;
                case "--help":
                    ShowHelp();
                    return;
                default:
                    Console.WriteLine($"Unknown option: {args[i]}");
                    ShowHelp();
                    return;
            }
        }

        if (string.IsNullOrEmpty(databaseLocation) || string.IsNullOrEmpty(queryNameToExport) || string.IsNullOrEmpty(locationToExportTo))
        {
            Console.WriteLine("Error: Missing required arguments.");
            ShowHelp();
            return;
        }

        try
        {
            ExportQuery(databaseLocation, queryNameToExport, locationToExportTo);
            Console.WriteLine("Query exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
