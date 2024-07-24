using System.Text.RegularExpressions;
using MiniExcelLibs;


class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("INICIO");
        string directorioCsv = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "csv");
        string archivoSalida = $"merge-{DateTime.Now.ToString("yyyy-MM-dd")}.xlsx";

        LeerCsvYCrearXlsx(directorioCsv, archivoSalida);
        Console.WriteLine("Presiona Enter para cerrar la consola...");
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
                Console.WriteLine("No se encontraron archivos CSV en el directorio especificado.");
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

            foreach (var archivoCsv in archivosCsvOrdenados)
            {
                Console.WriteLine("Merging " + Path.GetFileName(archivoCsv));
                // Lee cada archivo CSV y convierte su contenido a una lista de diccionarios
                var configuration = new MiniExcelLibs.Csv.CsvConfiguration
                {
                    Seperator = '\t' // Establecer el delimitador deseado, por ejemplo, punto y coma
                };

                var filas = MiniExcel.Query(archivoCsv, configuration: configuration).ToList();

                // Limpiar las filas usando el método CleanRow
                var filasLimpias = filas.Select(row => CleanRow(row)).ToList();
                if (esPrimerArchivo)
                {
                    // Añade todas las filas incluyendo la cabecera del primer archivo
                    todasLasFilas.AddRange(filasLimpias);
                    esPrimerArchivo = false;
                }
                else
                {
                    // Añade las filas de los archivos siguientes, omitiendo la primera fila (cabecera)
                    todasLasFilas.AddRange(filas.Skip(1));
                }
            }


            MiniExcel.SaveAs(archivoSalida, todasLasFilas, printHeader: false, excelType: ExcelType.XLSX, overwriteFile: true);

            Console.WriteLine($"El archivo Excel '{archivoSalida}' se ha creado exitosamente.");
        }
        catch (Exception ex)
        {

            Console.WriteLine(ex);
        }
    }

    static long ExtractNumberFromFileName(string fileName)
    {
        // Extrae la primera secuencia numérica del nombre del archivo
        var match = Regex.Match(fileName, @"\d+");
        return match.Success ? long.Parse(match.Value) : 0;
    }
    static IDictionary<string, object> CleanRow(IDictionary<string, object> row)
    {
        var cleanedRow = new Dictionary<string, object>();

        foreach (var kvp in row)
        {
            // Limpia cada valor usando la expresión regular
            var cleanedValue = Regex.Replace(kvp.Value.ToString(), @"=""(.*?)""", "$1");
            cleanedRow[kvp.Key] = cleanedValue;

        }

        return cleanedRow;
    }

}
