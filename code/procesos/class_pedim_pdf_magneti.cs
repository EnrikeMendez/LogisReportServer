using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Configuration;

namespace ReportServer2022.code.procesos
{
    class class_pedim_pdf_magneti
    {

        class_principal principal = new class_principal();
        int total;
        int c;
        string P;

        private string carpeta_generacion_base;
        private string carpeta_generacion;
        private string ClearFolder;
        private string ClearFolder1;
        private int cg_index;
        private int cgc_index;
        private string cgc_xml;


        string MiArchivo;
        string ClearFile;



        private string ConnString;
        private string[] ConnStringT;



        ProcessStartInfo info = null;


        private void sub_init_var()
        {
            carpeta_generacion_base = @"D:\files_ped\";

            //<add key="dbConnection" value="DATA SOURCE = 192.168.0.4:1521 / Orfeo2; PASSWORD = va4ncMC3P; USER ID = web_adm;" />
            /*
            ConnString = Split(Split(Db_link_orfeo.ConnectionString, ";")(2), "=")(1)   'User ID
            ConnString = ConnString & "/" & Split(Split(Db_link_orfeo.ConnectionString, ";")(1), "=")(1)    'Password
            ConnString = ConnString & "@" & Split(Split(Db_link_orfeo.ConnectionString, ";")(3), "=")(1)    'Data Source
            */
        }


        public void sub_pedim_pdf_magneti(string id_cron, string cliente, string fecha_ini, string fecha_fin, int days_deleted)
        {
            try
            {
                ImprimeConsola("Consulta información. " + "(" + id_cron + ")");
                principal.obj_querys.sub_sql_pedtos_marelli(cliente, fecha_ini, fecha_fin);
                ImprimeConsola("Se procesará la información del cliente " + cliente + " para el periodo del " + fecha_ini + " al " + fecha_fin);

                if (principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows.Count > 0)
                {
                    total = principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows.Count;
                    ImprimeConsola(DateTime.Now + " Se encontraron " + total + " registros para el cliente: " + cliente);
                }
                else
                {
                    ImprimeConsola(DateTime.Now + " No hay datos en la base de datos para el cliente: " + cliente);
                }


                for (int i = 0; i <= principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows.Count; i++)
                {
                    c = c + 1;
                    P = principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["SGEDOUCLEF"].ToString() +
                        principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["SGE_ADUANA_SECCION"].ToString() +
                        "-" +
                        principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["SGEPEDNUMERO"].ToString();

                    ImprimeConsola(DateTime.Now + " Procesando el Pedimento (" + c + "/" + total + ") " + P);

                    carpeta_generacion = carpeta_generacion_base;



                    if (Directory.Exists(carpeta_generacion) == false)
                    {
                        Directory.CreateDirectory(carpeta_generacion);
                    }

                    carpeta_generacion = carpeta_generacion + cliente + @"\";


                    if (Directory.Exists(carpeta_generacion) == false)
                    {
                        Directory.CreateDirectory(carpeta_generacion);
                    }

                    ClearFolder = carpeta_generacion;

                    carpeta_generacion = carpeta_generacion + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["PED_CARPETA"].ToString() + @"\";


                    if (Directory.Exists(carpeta_generacion) == false)
                    {
                        Directory.CreateDirectory(carpeta_generacion);
                    }


                    ClearFolder1 = carpeta_generacion;

                    //PDF de la cuenta de gastos AA --------------------------------------------------->
                    principal.obj_querys.sub_PDF_de_la_cuenta_de_gastos_AA(principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["FCTCLEF"].ToString());

                    cg_index = 1;
                    cgc_index = 1;

                    info = new ProcessStartInfo();

                    for (int j = 0; j <= principal.obj_querys.tbl_consulta_AA_Marelli.tabla_llena.Rows.Count; j++)
                    {
                        MiArchivo = carpeta_generacion + @"\" + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGEDOUCLEF"].ToString();
                        MiArchivo = MiArchivo + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGE_ADUANA_SECCION"].ToString();
                        MiArchivo = MiArchivo + "-";
                        MiArchivo = MiArchivo + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGEPEDNUMERO"].ToString();
                        MiArchivo = MiArchivo + "-CG";
                        MiArchivo = MiArchivo + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["FCTNUMERO"].ToString();
                        MiArchivo = MiArchivo + ".pdf";


                        info.UseShellExecute = true;
                        info.FileName = "RWCLI60.EXE";
                        //info.WorkingDirectory = @"d:\Logis_VB\";
                        info.WorkingDirectory = @"\\192.168.100.4\Logis_VB\";
                        //info.Arguments = "server = Rep60_LOGISWEBVM1 userid = " + ConnString + "" ;
                        info.Arguments = "server = Rep60_LOGISWEBVM1 userid = web_adm";
                        info.Arguments = info.Arguments + " desformat=pdf ";
                        info.Arguments = info.Arguments + " MI_FCTCLEF=" + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["FCTCLEF"].ToString();
                        info.Arguments = info.Arguments + " report=FACTURA_XX_CFDI_3.rdf ";
                        info.Arguments = info.Arguments + " desname='" + MiArchivo + "'";
                        info.Arguments = info.Arguments + " destype=file";

                        //Miniminiza la ventana del exe:       
                        //info.WindowStyle = ProcessWindowStyle.Minimized;

                        //oculta la venta del exe:
                        //info.WindowStyle = ProcessWindowStyle.Hidden;

                        Process.Start(info);


                        ClearFile = principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGEDOUCLEF"].ToString();
                        ClearFile = ClearFile + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGE_ADUANA_SECCION"].ToString();
                        ClearFile = ClearFile + "-";
                        ClearFile = ClearFile + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGEPEDNUMERO"].ToString();
                        ClearFile = ClearFile + "-CG";
                        ClearFile = ClearFile + principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["FCTNUMERO"].ToString();
                        ClearFile = ClearFile + ".pdf";


                        if (File.Exists(ClearFolder1 + ClearFile))
                        {
                            principal.obj_querys.sub_RegistraArchivoGenerado(ClearFolder, ClearFolder1, ClearFile, days_deleted);
                            ImprimeConsola(DateTime.Now + " Registrando el archivo: " + ClearFile);
                        }
                        else
                        {
                            ImprimeConsola(DateTime.Now + " El archivo no existe. ");
                        }


                        //'Genera archivos CG.pdf
                        //'CFDI de la cuenta de gastos AA

                        StreamWriter objStream = new StreamWriter(carpeta_generacion + @"\" +
                            principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGEDOUCLEF"].ToString() +
                            principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGE_ADUANA_SECCION"].ToString() +
                            "-" +
                            principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["SGEPEDNUMERO"].ToString() +
                            "-CG" +
                            principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["FCTNUMERO"].ToString() +
                            ".xml");

                        principal.obj_querys.sub_consulta_xml_CG(principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["FCTCLEF"].ToString());

                        //Pinta el XML de la CG
                        for (int c = 0; c <= principal.obj_querys.tbl_consulta_xml_cg_Marelli.tabla_llena.Rows.Count; c++)
                        {
                            objStream.WriteLine(principal.obj_querys.tbl_consulta_xml_cg_Marelli.tabla_llena.Rows[c]["COLUMN_VALUE"].ToString());
                        }
                        objStream.Close();
                        objStream.Dispose();


                        //Genera el archivo de la Cuenta de Gastos Consolidada XML
                        if (principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[j]["FCTDIVISA"].ToString() == "MXN")
                        {
                            //Genera archivos CGC.xml

                        }

                    }

                } // for i...
            }
            catch
            { }
            finally
            { }
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


        //Genera archivos CGC.xml

        private void sub_generacion_cgc_xml_alt(String Carpeta, String mi_fctclef, String mi_pedNumero, String mi_pedDouane, String mi_pedanio, String mi_seccion, String mi_folclave)
        {
            string Archivo;
            string Cuenta_Gastos;

            try
            {
                cgc_xml = "";

                //Va por la informacion
                principal.obj_querys.sub_SQL_FACTURA(mi_fctclef);


                //Si no hay informacion sale
                if (principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows.Count <= 0)
                {
                    principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Dispose();
                    return;
                }

                //Crea el encabezado
                //'cgc_xml = " <?xml version='1.0' encoding='utf - 8'?>" & vbCrLf
                cgc_xml = cgc_xml + "<PRINCIPAL_CONSOLIDADO> \n";
                cgc_xml = cgc_xml + "   <CUENTA> \n";

                //Recorre el detalle
                for (int i = 0; i <= principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows.Count; i++)
                {
                    Cuenta_Gastos = principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTNUMERO"].ToString();


                    cgc_xml = cgc_xml + "       <CUENTA_GASTOS_CONSOLIDADA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTNUMERO"].ToString() + "</CUENTA_GASTOS_CONSOLIDADA> \n";
                    cgc_xml = cgc_xml + "       <CUENTA_GASTOS_CONSOLIDADA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTNUMERO"].ToString() + "\n";
                    cgc_xml = cgc_xml + "       <ANTICIPO>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTTOTANTICIPO"].ToString() + "</ANTICIPO> \n";
                    cgc_xml = cgc_xml + "       <POLIZA_ANTICIPO>N/A</POLIZA_ANTICIPO> \n";
                    cgc_xml = cgc_xml + "       <UUID_CG_CN>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTUUID_UP"].ToString() + "</UUID_CG_CN> \n";
                    cgc_xml = cgc_xml + "       <RFC_COMPAÑIA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["CLIRFC"].ToString() + "</RFC_COMPAÑIA> \n";
                    cgc_xml = cgc_xml + "       <RFC_AGENCIA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["CLIRFC_EMP"].ToString() + "</RFC_AGENCIA> \n";
                    cgc_xml = cgc_xml + "       <CLAVE_CLIENTE>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTCLIENT"].ToString() + "</CLAVE_CLIENTE> \n";
                    cgc_xml = cgc_xml + "       <FECHA_CUENTA_GASTOS>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTDATEFACTURE"].ToString() + "</FECHA_CUENTA_GASTOS> \n";
                    cgc_xml = cgc_xml + "       <TIPO_OPERACION>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["TIPO_OP"].ToString() + " </TIPO_OPERACION> \n";


                    principal.obj_querys.sub_BASE_IVA(principal.obj_querys.tbl_consulta_pedtos_Marelli.tabla_llena.Rows[i]["FCTCLEF"].ToString());


                    if (principal.obj_querys.tbl_consulta_sql_base_iva_Marelli.tabla_llena.Rows.Count > 0)
                    {
                        for (int k = 0; k <= principal.obj_querys.tbl_consulta_sql_base_iva_Marelli.tabla_llena.Rows.Count; k++)
                        {
                            cgc_xml = cgc_xml + "   <BASE_IVA>" + principal.obj_querys.tbl_consulta_sql_base_iva_Marelli.tabla_llena.Rows[k]["BASE_IVA"].ToString() + "</BASE_IVA> \n";
                        }
                    }
                    else
                    {
                        cgc_xml = cgc_xml + "       <BASE_IVA>0</BASE_IVA> \n";
                    }

                    principal.obj_querys.tbl_consulta_sql_base_iva_Marelli.tabla_llena.Dispose();


                    cgc_xml = cgc_xml + "       <IVA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTIVA"].ToString() +"</IVA> \n";
                    cgc_xml = cgc_xml + "       <RETENCION_IVA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTRETENCIONFLETE"].ToString() + "</RETENCION_IVA> \n";
                    cgc_xml = cgc_xml + "       <TOTAL_CUENTA_GASTOS>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTTOTAL"].ToString() +  "</TOTAL_CUENTA_GASTOS> \n";
                    cgc_xml = cgc_xml + "       <CANCELAA>N/A</CANCELAA> \n";
                    cgc_xml = cgc_xml + "       <UUID_CUENTA_CANCELAA>N/A</UUID_CUENTA_CANCELAA> \n";

                    
                    cgc_xml = cgc_xml + "       <TIPO_CUENTA>";

                    if (principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTDIVISA"].ToString() == "MXN")
                    {
                        cgc_xml = cgc_xml + "MEXICANA </TIPO_CUENTA> \n";
                    }
                    else
                    {
                        cgc_xml = cgc_xml + "EXTRANGERA </TIPO_CUENTA> \n";
                    }
                    
                    cgc_xml = cgc_xml + "       <MONEDA>" + principal.obj_querys.tbl_consulta_sql_factura_Marelli.tabla_llena.Rows[i]["FCTDIVISA"].ToString() + "</MONEDA> \n";


                    

                }
            }
            catch
            { }
            finally
            { }


        }





    }
}
