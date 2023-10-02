using System;
using System.Collections.Generic;
using System.IO;

namespace ReportServer2022.code.procesos
{
    public class class_trading_genera_GSK
    {
        public string[,] tab_titulos;
        public string[,] tab_file = new string[0, 0];
        class_principal principal = new class_principal();
        LogisFunctions fun = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());

        public void sub_init_local_var()
        {
            tab_titulos = new string[14, 2];

            tab_titulos[0, 0] = "SHIPMENT_NO";
            tab_titulos[1, 0] = "CARRIER";
            tab_titulos[2, 0] = "PLANNED_SHIPDATE";
            tab_titulos[3, 0] = "SHIP_DATE";
            tab_titulos[4, 0] = "PLANNED_DELIVERY_DATE";
            tab_titulos[5, 0] = "ORIGIN";
            tab_titulos[6, 0] = "ORIGIN_ADDRESS";
            tab_titulos[7, 0] = "ORIGIN_CITY";
            tab_titulos[8, 0] = "DESTINATION";
            tab_titulos[9, 0] = "DESTINATION_ADDRESS";
            tab_titulos[10, 0] = "DESTINATION_CITY";
            tab_titulos[11, 0] = "MODE_";
            tab_titulos[12, 0] = "SHIPMENT_LINE#";
            tab_titulos[13, 0] = "CREATION_DATE";

            tab_titulos[0, 1] = "mm/dd/yyyy hh:mm";
        }
        public void sub_trading_genera_GSK(string Carpeta, string File_Name, string[,] tab_archivos, string Fecha_1, string Fecha_2, string Empresa, string idCron, string reporte_name, long days_deleted)
        {
            string archivoXls = string.Empty, archivoZip = string.Empty, ruta = string.Empty;
            
            List<string> lstDestinatarios = new List<string>();
            List<string> lstArchivos = new List<string>();
            List<string> nameSheets = new List<string>();


            try
            {
                sub_init_local_var();

                //Arreglo de configuracion:
                tab_file = new string[3, tab_file.GetLength(1) + 1];
                /*tab_file[0, tab_file.GetLength(1) - 1] = Carpeta + File_Name;
                tab_file[1, tab_file.GetLength(1) - 1] = "Shipments";
                tab_file[2, tab_file.GetLength(1) - 1] = tab_titulos.GetLength(0) + "";*/
                tab_file[0, 0] = Carpeta + File_Name;
                tab_file[1, 0] = "Shipments";
                tab_file[2, 0] = tab_titulos.GetLength(0) + "";

                try
                {
                    ImprimeConsola("Consulta información.");
                    principal.obj_querys.sub_trading_genera_GSK_consultar();

                    ImprimeConsola("Genera excel.");

                    //Nombre Hoja 1
                    nameSheets.Add(File_Name);

                    archivoXls = fun.DataTableToExcel(principal.obj_querys.tbl_trading_genera_GSK_rpt.tabla_llena, File_Name, nameSheets);
                    ruta = new FileInfo(archivoXls).DirectoryName;

                    ImprimeConsola("Genera Zip.");
                    archivoZip = fun.CompressFolder(ruta);

                    ImprimeConsola("Envía correo electrónico.");
                    lstDestinatarios.Add("oscarrp@logis.com.mx");
                    //lstDestinatarios.Add("margaritarg@logis.com.mx");
                    //lstDestinatarios.Add("joseemv@logis.com.mx");
                    //lstDestinatarios.Add("julioecv@logis.com.mx");
                    //lstDestinatarios.Add("abrahamrr@logis.com.mx");
                    lstArchivos.Add(archivoZip);


                    //Envio de correo:
                    fun.SendMail(reporte_name, lstDestinatarios, lstArchivos);

                }
                catch (Exception ex)
                {
                    fun.WriteExceptionOnLog(ex);
                }
                finally
                {
                    if (!ruta.Equals(string.Empty))
                    {
                        fun.DeleteOldFile(ruta, 1);
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

            ImprimeConsola("Termina proceso " + reporte_name);
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
        }
    }
}