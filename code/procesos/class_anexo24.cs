using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ReportServer2022.code.include;

namespace ReportServer2022.code.procesos
{
    class class_anexo24
    {
        class_principal principal = new class_principal();
        LogisFunctions fun = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());
        
       
        //public void sub_anex24_tetrapack(string Carpeta, string[,] tab_archivos, string Fecha_1, string Fecha_2, long days_deleted, string rep_id)
        public void sub_anex24_tetrapack(int daysInfo) //TEMP
        {
            string archivoXls = string.Empty, archivoZip = string.Empty, ruta = string.Empty;
            List<string> lstDestinatarios = new List<string>();
            List<string> lstArchivos = new List<string>();

            List<string> lstArchivosDelete = new List<string>();
            List<string> nameSheets = new List<string>();

            fun.ImprimeConsola("Inicia proceso Anexo 24 Tetrapack");

            try
            {

                ////tab_archivos(0,i) > nombre del archivo
                ////tab_archivos(1,i) > nombre del reporte
                ////tab_archivos(2,i) > tamaño del archivo
                ////tab_archivos(3,i) > Hash MD5
                ////tab_archivos(4,i) > 1 o 0 (o si se olivida, vacio) (si se necesita o no un zip)
                ////tab_archivos(5,i) > tamaño del zip


                ////******************************************* TEMP ****************************************************
                ////Proceso local, no consulta parametros de la BD´s

                DateTime fec1 = DateTime.Now;
                DateTime fec2 = DateTime.Now;


                string File_Name = "Anexo_24_" + fec1.AddDays(daysInfo).ToString("dd-MM-yyyy") + "_to_" + fec2.AddDays(-1).ToString("dd-MM-yyyy") + "_cliente_14005";

                // *** REPROCESAR ***
                //string File_Name = "Anexo_24_12-09-2022_to_18-09-2022_cliente_14005";

                String[,] tab_archivos = new string[1, 5];


                tab_archivos[0, 0] = File_Name;
                tab_archivos[0, 1] = File_Name;
                tab_archivos[0, 4] = "1";

                //******************************************* TEMP ***************************************************
                try
                {
                    fun.ImprimeConsola("Consulta información Anexo24_14005");
                    principal.obj_querys.sub_Anexo24_tetrapack_14005(daysInfo);

                    fun.ImprimeConsola("Genera excel Anexo24_14005");


                    //Nombre de Hoja 1:
                    nameSheets.Add(fec1.AddDays(daysInfo).ToString("dd-MM-yyyy") + "_to_" + fec2.AddDays(-1).ToString("dd-MM-yyyy"));
                    
                    // *** REPROCESAR ***
                    //nameSheets.Add("12-09-2022_to_18-09-2022");


                    archivoXls = fun.DataTableToExcel(principal.obj_querys.tbl_anexo24_tetrapack_rpt.tabla_llena, tab_archivos[0, 0].ToString(), nameSheets);
                    lstArchivosDelete.Add(archivoXls);
                    ruta = new FileInfo(archivoXls).DirectoryName;


                    //Verifica si se va a zipear el archivo y lo adjunta al mail:
                    if (tab_archivos[0, 4] == "1")
                    {
                        fun.ImprimeConsola("Genera Zip Anexo24_14005");
                        archivoZip = fun.CompressFolder(ruta);
                        lstArchivos.Add(archivoZip);
                        lstArchivosDelete.Add(archivoZip);
                    }
                    else
                    {
                        lstArchivos.Add(archivoXls);
                    }

                    //contactos a recibir el reporte (pendiente consulta desde base de datos !!!):
                    fun.ImprimeConsola("Envía correo electrónico Anexo24_14005");
                    //lstDestinatarios.Add("oscarrp@logis.com.mx");
                    //lstDestinatarios.Add("margaritarg@logis.com.mx");
                    //lstDestinatarios.Add("joseemv@logis.com.mx");
                    //lstDestinatarios.Add("abrahamrr@logis.com.mx");
                    lstDestinatarios.Add("desarrollo_web@logis.com.mx");

                    //Envio de correo electronico: 
                    fun.SendMail(tab_archivos[0, 0].ToString(), lstDestinatarios, lstArchivos);


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

            //ImprimeConsola("Termina proceso Anexo 24. " + "(" + rep_id + ")");
            fun.ImprimeConsola("Termina proceso Anexo24_14005");

        }


        public void Dispose()
        {
            /*if (this.tab_file != null)
            {
                GC.SuppressFinalize(this.tab_file);
            }*/
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
