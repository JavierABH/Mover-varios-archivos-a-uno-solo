using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Mover_varios_archivos_a_uno_solo
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\K90011729\Documents\Log\DATOS TE0364-BRO\"; // No borrar el ultimo '\'
            string pathFileText = @"C:\Users\K90011729\Documents\Log\Datos.csv"; // Ruta del archivo donde se guardara todo.
            //Se escribe el rango de fechas de los archivos           
            DateTime FechaInicio = Convert.ToDateTime("01-Feb-22");
            DateTime FechaFin = Convert.ToDateTime("10-Sep-22");

            string pathFile;
            StreamReader Reader;
            StreamWriter Writer;
            string array;
            string contenidodelimitado;
            try
            {
                Writer = File.AppendText(pathFileText);
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
                            contenidodelimitado = array.Replace('	', ';');
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
                Console.WriteLine("Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
