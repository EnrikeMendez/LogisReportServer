using System;
using System.Collections.Generic;
using System.IO;

namespace ReportServer2022.code.procesos
{
    public class class_guias_disponibles_genera
    {
        public string[,] tab_titulos;
        public string[,] tab_file = new string[0, 0];
        class_principal principal = new class_principal();
        LogisFunctions fun = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());

        public void sub_init_local_var()
        {
            tab_titulos = new string[7, 2];

            tab_titulos[0, 0] = "NUMERO DE CLIENTE";
            tab_titulos[1, 0] = "CLIENTE";
            tab_titulos[2, 0] = "TIPO";
            tab_titulos[3, 0] = "GUIAS TOTALES";
            tab_titulos[4, 0] = "GUIAS DISPONIBLES";
            tab_titulos[5, 0] = "GUIAS OCUPADAS";
            tab_titulos[6, 0] = "GUIAS CANCELADAS";

            tab_titulos[0, 1] = "mm/dd/yyyy hh:mm";
        }
        public void sub_guias_disponibles_genera(string Carpeta, string[,] tab_archivos, string Fecha_1, string Fecha_2, long days_deleted, string rep_id)
        //public void sub_guias_disponibles_genera()
        {
            string archivoXls = string.Empty, archivoZip = string.Empty, ruta = string.Empty;
            List<string> lstDestinatarios = new List<string>();
            List<string> lstArchivos = new List<string>();
            List<string> lstArchivosDelete = new List<string>();
            List<string> nameSheets = new List<string>();

            ImprimeConsola("Inicia proceso Guias Disponibles. " + "(" + rep_id + ")");

            try
            {
                ////******************************************* TEMP ****************************************************
                ////Proceso local, no consulta parametros de la BD´s
                //string Date = DateTime.Now.ToString("dd-MM-yyyy");
                //String Carpeta = AppDomain.CurrentDomain.BaseDirectory + "\\reportes\\web_reports\\GUIAS_DISPONIBLES\\";
                //String File_Name = "Guias_Disponibles_" + Date.ToString();
                //String[,] tab_archivos = new string[1, 5];
                //
                ////tab_archivos(0,i) > nombre del archivo
                ////tab_archivos(1,i) > nombre del reporte
                ////tab_archivos(2,i) > tamaño del archivo
                ////tab_archivos(3,i) > Hash MD5
                ////tab_archivos(4,i) > 1 o 0 (o si se olivida, vacio) (si se necesita o no un zip)
                ////tab_archivos(5,i) > tamaño del zip
                //tab_archivos[0, 0] = "Guias_Disponibles_" + Date.ToString();
                //tab_archivos[0, 1] = "Guias Disponibles LTL/CROSS DOCK";
                //tab_archivos[0, 4] = "1";
                //
                //
                //String Fecha_1 = Date.ToString();
                //String Fecha_2 = Date.ToString(); ;
                //
                ////String idCron;
                //String reporte_name = "Guias Disponibles LTL/CROSS DOCK";
                //int days_deleted = 7;
                //
                //String creado = string.Empty;
                //
                //
                //
                ////verificamos que la carpeta exista, si no la creamos:
                //class_xfunciones obj_xfunciones = new class_xfunciones();
                //if (Directory.Exists(Carpeta) == false)
                //{
                //    obj_xfunciones.sub_Create_Entire_Path(Carpeta);
                //    creado = "1";
                //}
                //else
                //{
                //    creado = "0";
                //}
                //******************************************* TEMP ****************************************************

                sub_init_local_var();

                //Arreglo de configuracion:
                tab_file = new string[3, tab_file.GetLength(1) + 1];
                /*tab_file[0, tab_file.GetLength(1) - 1] = Carpeta + File_Name;
                tab_file[1, tab_file.GetLength(1) - 1] = "Shipments";
                tab_file[2, tab_file.GetLength(1) - 1] = tab_titulos.GetLength(0) + "";*/
                tab_file[0, 0] = Carpeta + tab_archivos[0, 0].ToString();
                tab_file[1, 0] = "GuiasDisponibles";
                tab_file[2, 0] = tab_titulos.GetLength(0) + "";

                /*
                //Borra los archivos existentes en la carpeta del los procesos anteriores antes de comenzar a generar los nuevos:
                principal.obj_xfunciones.ftn_delete_old_file(Carpeta, File_Name);

                //Genera Excel:
                principal.obj_querys.sub_guias_disponibles_consultar();
                principal.obj_xfunciones.ftn_excel_simple2(tab_file, tab_titulos, principal.obj_querys.tbl_guias_disponibles_rpt.tabla_llena);

                //Genera ZIP:
                principal.obj_xfunciones.ftn_compress_file(Carpeta, File_Name);

                //Envia e-Mail:
                principal.obj_xfunciones.ftn_sendMailReports(reporte_name, tab_archivos, Carpeta, days_deleted);

                //Una vez enviado el correo, se eliminan los archivos del repositorio:
                principal.obj_xfunciones.ftn_delete_old_file(Carpeta, File_Name);
                */

                try
                {
                    ImprimeConsola("Consulta información. " + "(" + rep_id + ")");
                    principal.obj_querys.sub_guias_disponibles_consultar();

                    ImprimeConsola("Genera excel. " + "(" + rep_id + ")");

                    //Nombre Hoja 1
                    nameSheets.Add(tab_archivos[0, 0].ToString());

                    archivoXls = fun.DataTableToExcel(principal.obj_querys.tbl_guias_disponibles_rpt.tabla_llena, tab_archivos[0, 0].ToString(), nameSheets);
                    lstArchivosDelete.Add(archivoXls);
                    ruta = new FileInfo(archivoXls).DirectoryName;


                    //Verifica si se va a zipear el archivo y lo adjunta al mail:
                    if (tab_archivos[0, 4] == "1")
                    {
                        ImprimeConsola("Genera Zip. " + "(" + rep_id + ")");
                        archivoZip = fun.CompressFolder(ruta);
                        lstArchivos.Add(archivoZip);
                        lstArchivosDelete.Add(archivoZip);
                    }
                    else
                    {
                        lstArchivos.Add(archivoXls);
                    }

                    ImprimeConsola("Envía correo electrónico. " + "(" + rep_id + ")");
                    lstDestinatarios.Add("oscarrp@logis.com.mx");
                    lstDestinatarios.Add("margaritarg@logis.com.mx");
                    lstDestinatarios.Add("joseemv@logis.com.mx");
                    //lstDestinatarios.Add("julioecv@logis.com.mx");
                    lstDestinatarios.Add("abrahamrr@logis.com.mx");

                    //Envia correo:
                    fun.SendMail("Guias Disponibles", lstDestinatarios, lstArchivos);


                    //Borra los archivos generados:
                    fun.ftn_file_instant_delete(lstArchivosDelete);


                }
                catch (Exception ex)
                {
                    fun.WriteExceptionOnLog(ex);
                }
                finally
                {
                    if (!ruta.Equals(string.Empty))
                    {
                        fun.DeleteOldFile(ruta, -1);
                    }
                }
            }
            catch { }
            finally
            {
                if (archivoXls != null)
                {
                    archivoXls = string.Empty;
                    archivoXls = null;
                }
                if (archivoZip != null)
                {
                    archivoZip = string.Empty;
                    archivoZip = null;
                }
                if (ruta != null)
                {
                    ruta = string.Empty;
                    ruta = null;
                }
                if (lstDestinatarios != null)
                {
                    lstDestinatarios.Clear();
                    GC.SuppressFinalize(lstDestinatarios);
                }
                if (lstArchivos != null)
                {
                    lstArchivos.Clear();
                    GC.SuppressFinalize(lstArchivos);
                }
            }

            ImprimeConsola("Termina proceso Guias Disponibles. " + "(" + rep_id + ")");
        }
        private void ImprimeConsola(string texto)
        {
            LogisFunctions lf = null;

            try
            {
                lf = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());
                texto = string.Format("{0} - {1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"), texto);
                Console.WriteLine(texto);
                lf.WriteLog(texto);
            }
            catch { }
            finally
            {
                if (lf != null)
                {
                    GC.SuppressFinalize(lf);
                }
            }
        }
        /// <summary>
        /// Libera los recursos utilizados por el objeto.
        /// </summary>
        public void Dispose()
        {
            if (this.tab_titulos != null)
            {
                GC.SuppressFinalize(this.tab_titulos);
            }
            if (this.tab_file != null)
            {
                GC.SuppressFinalize(this.tab_file);
            }
            if (this.principal != null)
            {
                this.principal.Dispose();
                GC.SuppressFinalize(this.principal);
            }
            if (this.fun != null)
            {
                GC.SuppressFinalize(this.fun);
            }
            GC.Collect();
        }
    }
}