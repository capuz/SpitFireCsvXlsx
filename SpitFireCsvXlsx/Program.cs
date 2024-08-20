using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using MiniExcelLibs;
using Spectre.Console;

class Program
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        AnsiConsole.MarkupLine("[bold green]INICIO:[/]");
        string directorioCsv = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "csv");
        string archivoSalida = $"merge-{DateTime.Now:yyyy-MM-dd}.xlsx";

        LeerCsvYCrearXlsx(directorioCsv, archivoSalida);
        AnsiConsole.MarkupLine("[bold yellow]Presiona Enter para cerrar la consola...[/]");
        Console.ReadLine();
    }

    static void LeerCsvYCrearXlsx(string directorioCsv, string archivoSalida)
    {
        try
        {
            // Obtén la lista de archivos CSV en el directorio especificado
            var archivosCsv = Directory.GetFiles(directorioCsv, "*.csv");

            if (archivosCsv.Length == 0)
            {
                AnsiConsole.MarkupLine("[red]No se encontraron archivos CSV en el directorio especificado.[/]");
                return;
            }

            // Ordenar los archivos CSV por el número en el nombre del archivo
            var archivosCsvOrdenados = archivosCsv
                .Select(f => new
                {
                    FileName = f,
                    Number = ExtractNumberFromFileName(Path.GetFileName(f))
                })
                .OrderBy(x => x.Number)
                .Select(x => x.FileName)
                .ToList();

            // Lista para almacenar todas las filas de los CSV
            var todasLasFilas = new List<dynamic>();
            bool esPrimerArchivo = true;

            AnsiConsole.Progress()
                .Start(ctx =>
                {
                    var task = ctx.AddTask("[green]Procesando archivos CSV[/]", maxValue: archivosCsvOrdenados.Count);

                    foreach (var archivoCsv in archivosCsvOrdenados)
                    {
                        AnsiConsole.MarkupLine($"[blue]Merging {Path.GetFileName(archivoCsv)}[/]");

                        var configuration = new MiniExcelLibs.Csv.CsvConfiguration
                        {
                            Seperator = '\t' // Establecer el delimitador deseado, por ejemplo, punto y coma
                        };
              
                        var filas = MiniExcel.Query(archivoCsv, configuration: configuration).ToList();
                        var filasLimpias = filas.Select(row => CleanRow(row)).ToList();

                        if (esPrimerArchivo)
                        {
                            todasLasFilas.AddRange(filasLimpias);
                            esPrimerArchivo = false;
                        }
                        else
                        {
                            todasLasFilas.AddRange(filasLimpias.Skip(1));
                        }

                        task.Increment(1);
                        Thread.Sleep(130); // Pausa de 130 milisegundos para ver el progreso
                    }
                });

            MiniExcel.SaveAs(archivoSalida, todasLasFilas, printHeader: false, excelType: ExcelType.XLSX, overwriteFile: true);

            AnsiConsole.MarkupLine($"[bold green]El archivo Excel '{archivoSalida}' se ha creado exitosamente.[/]");
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[bold red]Error: {ex.Message}[/]");
        }
    }

    static long ExtractNumberFromFileName(string fileName)
    {
        var match = Regex.Match(fileName, @"\d+");
        return match.Success ? long.Parse(match.Value) : 0;
    }

    static IDictionary<string, object> CleanRow(IDictionary<string, object> row)
    {
        var cleanedRow = new Dictionary<string, object>();

        foreach (var kvp in row)
        {
            var cleanedValue = Regex.Replace(kvp.Value.ToString(), @"=""(.*?)""", "$1");
            cleanedRow[kvp.Key] = cleanedValue;
        }

        return cleanedRow;
    }
}
