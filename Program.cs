using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Mover_varios_archivos_a_uno_solo
{
    class Program
    {
        static string path = @"C:\Users\K90011729\Documents\Log\prueba\"; // No borrar el ultimo '\'
        static string pathFileText = @"C:\Users\K90011729\Documents\Log\prueba\"; // Ruta del archivo donde se guardara todo.
        static string pathCSV = pathFileText + "Datos.csv";
        static string pathExcel = pathFileText + "Datos.xlsx";

        static void Main(string[] args)
        {
            //Se escribe el rango de fechas de los archivos           
            DateTime FechaInicio = Convert.ToDateTime("01-Feb-22");
            DateTime FechaFin = Convert.ToDateTime("10-Feb-22");

            string pathFile;
            StreamReader Reader;
            StreamWriter Writer;
            string array;
            string contenidodelimitado;
            try
            {
                Writer = File.AppendText(pathCSV);
                string[] dato = new string[290];
                while (true)
                {
                    if (FechaInicio > FechaFin)
                        break;
                    try
                    {
                        //leo cada archivo
                        pathFile = path + FechaInicio.ToString("MM-dd-yy") + @"\Ford P552 HVAC.dat";       // Example: 02-01-22\Ford P552 HVAC.dat

                        Reader = File.OpenText(pathFile);
                        array = Reader.ReadLine();
                        while (array != null)
                        {
                            dato = array.Split('	');
                            contenidodelimitado = array.Replace('	', ',');
                            if (dato[0].Trim() != "Model No.")
                            {
                                Writer.WriteLine(contenidodelimitado); //No copia el encabezado de los archivos
                            }
                            array = Reader.ReadLine();
                        }
                        Reader.Close();

                        Console.WriteLine(pathFile + " Copiado");
                        FechaInicio = FechaInicio.AddDays(1);
                    }
                    catch (Exception ex)
                    {
                        if (ex is DirectoryNotFoundException) // Si no encuentra la carpeta, sigue
                        {                           
                            FechaInicio = FechaInicio.AddDays(1);
                            continue;
                        }
                    }
                }
                Writer.Close();
                Console.WriteLine("Proceso completado...");
                Console.WriteLine("Conversion csv a excel...");
                CsvToExcel(pathCSV, pathExcel);
                Console.WriteLine("Excel generado correctamente...");
                Console.WriteLine("Desea borrar el archivo csv? s/n");
                string borrarArchivo = Console.ReadLine();
                if (borrarArchivo == "s")
                    File.Delete(pathCSV);
                Console.WriteLine("Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        static void CsvToExcel(string csv, string xlsx)
        {
            Excel.Application xl = new Excel.Application();
            //Open Excel Workbook for conversion.
            Excel.Workbook wb = xl.Workbooks.Open(csv);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            //Select The UsedRange
            Excel.Range used = ws.UsedRange;
            //Autofit The Columns
            used.EntireColumn.AutoFit();
            //Save file as csv file
            wb.SaveAs(xlsx, 51);
            //Close the Workbook.
            wb.Close();
            //Quit Excel Application.
            xl.Quit();
        }
    }
}
