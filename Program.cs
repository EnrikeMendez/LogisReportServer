using System;
using System.Data;
using ReportServer2022.code.querys;
using ReportServer2022.code.procesos;

namespace ReportServer2022
{
    class Program
    {
        static void Main(string[] args)
        {
            class_querys_generics obj_querys_generics = new class_querys_generics();

            string rep_id, command;
            string reporte_temporal;
            string Fecha_1, Fecha_2;
            string Carpeta, NombreArchivo, reporte_name;
            string[,] tab_archivos = new string[1, 5];
            string PARAM_1, PARAM_2, PARAM_3, PARAM_4;
            long days_deleted;


            try
            {
                rep_id = string.Empty;
                command = string.Empty;
                reporte_temporal = string.Empty;
                Fecha_1 = string.Empty;
                Fecha_2 = string.Empty;
                Carpeta = string.Empty;
                NombreArchivo = string.Empty;
                reporte_name = string.Empty;
                tab_archivos = new string[1, 5]; //1 fila por 5 columnas
                PARAM_1 = string.Empty;
                PARAM_2 = string.Empty;
                PARAM_3 = string.Empty;
                PARAM_4 = string.Empty;
                days_deleted = 0;

                if (args.Length == 2)
                {
                    /*****************************************************
                     * Obtener valores inciales para ejecutar el reporte *
                     *****************************************************/
                    rep_id = args[0]; //id_chron del reporte
                    reporte_temporal = args[1];
                    ImprimeConsola(string.Format("Parametros introducidos: {0} {1} \n", rep_id, reporte_temporal));


                    /* CONSULTA PARA OBTENER LA INFORMACIÓN DEL REPORTE */
                    obj_querys_generics.sub_reporte_ejecutar_consultar(rep_id, Int32.Parse(reporte_temporal));
                    obj_querys_generics.sub_numero_parametros_reporte_consultar(rep_id);


                    /* Comprueba que las varibles a consultar, obtengan la info desde la DB */
                    if (obj_querys_generics.bandera == true)
                    {
                        //Asigacion de valores a variables:
                        if (obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena != null)
                        {
                            if (obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Rows.Count > 0)
                            {
                                Fecha_1 = obj_querys_generics.Fecha_1;
                                Fecha_2 = obj_querys_generics.Fecha_2;
                                tab_archivos = obj_querys_generics.tab_archivos;
                                Carpeta = obj_querys_generics.Carpeta;

                                command = obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["COMMAND"].ToString();
                                NombreArchivo = obj_querys_generics.file_name;
                                reporte_name = obj_querys_generics.reporte_name;
                                days_deleted = obj_querys_generics.days_deleted;

                                foreach (DataColumn dc in obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Columns)
                                {
                                    if (dc.ColumnName.ToUpper().Equals("PARAM_1"))
                                    {
                                        PARAM_1 = obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["PARAM_1"].ToString();
                                    }
                                    else if (dc.ColumnName.ToUpper().Equals("PARAM_2"))
                                    {
                                        PARAM_2 = obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["PARAM_2"].ToString();
                                    }
                                    else if (dc.ColumnName.ToUpper().Equals("PARAM_3"))
                                    {
                                        PARAM_3 = obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["PARAM_3"].ToString();
                                    }
                                    else if (dc.ColumnName.ToUpper().Equals("PARAM_4"))
                                    {
                                        PARAM_4 = obj_querys_generics.tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["PARAM_4"].ToString();
                                    }
                                }
                            }
                            else
                            {
                                obj_querys_generics.bandera = false;
                            }
                        }
                        else
                        {
                            obj_querys_generics.bandera = false;
                        }
                    }



                    if (obj_querys_generics.bandera == false)
                    {
                        //En caso de error:
                        Fecha_1 = "";
                        Fecha_2 = "";
                        PARAM_2 = "";
                        tab_archivos[0, 0] = "ERROR01";
                        Carpeta = "ERROR02";

                        ImprimeConsola("Error: No se encontró información para generar el reporte.");
                    }

                    if (obj_querys_generics.bandera == true)
                    {
                        switch (command)
                        {
                            case "gsk_pedimientos":
                                class_trading_genera_GSK obj_trading_genera_GSK = new class_trading_genera_GSK();
                                obj_trading_genera_GSK.sub_trading_genera_GSK(Carpeta, NombreArchivo, tab_archivos, Fecha_1, Fecha_2, PARAM_2, rep_id, reporte_name, days_deleted);
                                obj_querys_generics.sub_libera_InProgress(rep_id);
                                break;

                            case "guias_disponibles":
                                class_guias_disponibles_genera obj_guias_disponibles = new class_guias_disponibles_genera();
                                obj_guias_disponibles.sub_guias_disponibles_genera(Carpeta, tab_archivos, Fecha_1, Fecha_2, days_deleted, rep_id);
                                obj_querys_generics.sub_libera_InProgress(rep_id);
                                break;

                            case "lotes_pendientes_nc":
                                class_nuis_pendientes_nc obj_lotes_p_nc = new class_nuis_pendientes_nc();
                                obj_lotes_p_nc.sub_nuis_pendientes_nc(Carpeta, tab_archivos, rep_id);
                                obj_querys_generics.sub_libera_InProgress(rep_id);
                                break;

                            default:
                                ImprimeConsola(string.Format("El reporte seleccionado no se encuentra dentro del catálogo."));
                                break;
                        }
                    }
                }
                else
                {
                    ImprimeConsola("No se ha introducido el número correcto de argumentos por linea de comandos.");

                    //********************** EJECUCION DE PROCESOS TEMPORALES SIN REGISTRO DE PROGRAMACION EN BASE DE DAT0S **********************
                    ImprimeConsola("Proceso local, no consulta parametros de la BD´s");

                    //Guias disponibles:
                    //class_guias_disponibles_genera obj_guias_disponibles_genera = new class_guias_disponibles_genera();

                    //Anexo 24 Tetrapack:
                    class_anexo24 obj_anexo24_tetrapack = new class_anexo24();
                    obj_anexo24_tetrapack.sub_anex24_tetrapack(-7);


                    //Expedientes Marelli Test:



                    //****************************************************************************************************************************
                }
            }
            catch (Exception ex)
            {
                string error;
                error = "Ocurrio un error inseperado: " + ex + "";
                ImprimeConsola(error);
            }
            finally
            {
                if (obj_querys_generics != null)
                {
                    obj_querys_generics.Dispose();
                    GC.SuppressFinalize(obj_querys_generics);
                }

                //ImprimeConsola("Proceso Terminado.");
                //Console.WriteLine("Presione una tecla para terminar.");
                //Console.ReadLine();
            }

        } // static void Main(string[] args)


        private static void ImprimeConsola(string texto)
        {
            LogisFunctions lf = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());
            texto = string.Format("{0} - {1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"), texto);
            Console.WriteLine(texto);
            lf.WriteLog(texto);
        }
    }
}