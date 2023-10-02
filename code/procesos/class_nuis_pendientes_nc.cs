using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportServer2022.code.procesos


{
    class class_nuis_pendientes_nc
    {
        public string[,] tab_file = new string[0, 0];
        class_principal principal = new class_principal();
        LogisFunctions fun = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());

        public void sub_nuis_pendientes_nc(string Carpeta, string[,] tab_archivos,  string rep_id)
        {

            string archivoXls = string.Empty, archivoZip = string.Empty, ruta = string.Empty;
            List<string> lstDestinatarios = new List<string>();
            List<string> lstArchivos = new List<string>();
            List<string> lstArchivosDelete = new List<string>();
            List<string> nameSheets = new List<string>();

            fun.ImprimeConsola("Inicia proceso nuis pedientes nota de credito " + "(" + rep_id + ")");

            //id_rep = 340


            //Arreglo de configuracion:
            tab_file = new string[3, tab_file.GetLength(1) + 1];
            /*tab_file[0, tab_file.GetLength(1) - 1] = Carpeta + File_Name;
            tab_file[1, tab_file.GetLength(1) - 1] = "Shipments";
            tab_file[2, tab_file.GetLength(1) - 1] = tab_titulos.GetLength(0) + "";*/
            tab_file[0, 0] = Carpeta + tab_archivos[0, 0].ToString();
            tab_file[1, 0] = "Lotes pedientes de NC";
            //tab_file[2, 0] = tab_titulos.GetLength(0) + "";

            try
            {
                fun.ImprimeConsola("Consulta información. " + "(" + rep_id + ")");
                principal.obj_querys.sub_nuis_pendientes_nc();


                fun.ImprimeConsola("Genera excel. " + "(" + rep_id + ")");

                //Nombre Hoja 1
                nameSheets.Add(tab_archivos[0, 0].ToString());

                archivoXls = fun.DataTableToExcel(principal.obj_querys.tbl_nuis_nc_rpt.tabla_llena, tab_archivos[0, 0].ToString(), nameSheets);
                lstArchivosDelete.Add(archivoXls);
                ruta = new FileInfo(archivoXls).DirectoryName;


                //Verifica si se va a zipear el archivo y lo adjunta al mail:
                if (tab_archivos[0, 4] == "1")
                {
                    fun.ImprimeConsola("Genera Zip. " + "(" + rep_id + ")");
                    archivoZip = fun.CompressFolder(ruta);
                    lstArchivos.Add(archivoZip);
                    lstArchivosDelete.Add(archivoZip);
                }
                else
                {
                    lstArchivos.Add(archivoXls);
                }

                fun.ImprimeConsola("Envía correo electrónico. " + "(" + rep_id + ")");
                lstDestinatarios.Add("oscarrp@logis.com.mx");
                lstDestinatarios.Add("margaritarg@logis.com.mx");
                lstDestinatarios.Add("joseemv@logis.com.mx");
                //lstDestinatarios.Add("julioecv@logis.com.mx");
                //lstDestinatarios.Add("abrahamrr@logis.com.mx");

                //Envia correo:
                fun.SendMail("LOTES PENDIENTES DE NC", lstDestinatarios, lstArchivos);


                //Borra los archivos generados:
                fun.ftn_file_instant_delete(lstArchivosDelete);

            }
            catch
            {

            }
            finally 
            { 
            
            }
            


        }

        

    }
}
