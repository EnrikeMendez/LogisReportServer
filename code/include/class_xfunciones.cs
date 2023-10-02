using SpreadsheetLight;
using System;
using System.Data;
using System.Collections;
using System.IO.Compression;
using System.Globalization;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReportServer2022.code.include
{
    public class class_xfunciones
    {
        public string first_path = string.Empty, second_path = string.Empty;
        public string IP_servidor1 = string.Empty, IP_servidor2 = string.Empty;
        public string ini_path = string.Empty;
        private string sResult = string.Empty;
        private string sTempPath = Path.GetTempPath();

        public void sub_init_var()
        {
            ini_path = AppDomain.CurrentDomain.BaseDirectory;

            //En servidor:
            /*IP_servidor1 = "192.168.100.5";
            IP_servidor2 = "192.168.100.4";

            first_path = "\\\\" + IP_servidor1 + "\\reportes\\web_reports\\";
            second_path = "\\\\" + IP_servidor2 + "\\reportes\\web_reports\\";
            */

            //Local:
            first_path = ini_path + "reportes\\web_reports\\";
            //second_path = "\\\\" + IP_servidor2 + "\\reportes\\web_reports\\";
        }

        //Crea el directorio si no existe:
        public void sub_Create_Entire_Path(string Fullpath)
        {
            string[] FolderTree = new string[50];
            string CurrentPath;

            int n, Start;
            Start = -1;

            try
            {
                FolderTree = Fullpath.Split("\\".ToCharArray());

                if (FolderTree[0].Equals(""))
                {
                    CurrentPath = "\\" + FolderTree[2] + "\\" + FolderTree[3];
                    Start = 4;
                }
                else
                {
                    CurrentPath = FolderTree[0];
                    Start = 1;
                }

                if (Directory.Exists(CurrentPath) == false)
                {
                    // Break;
                }

                for (n = 1; n < FolderTree.GetLength(0); n++)
                {
                    CurrentPath = CurrentPath + "\\" + FolderTree[n];
                    if (Directory.Exists(CurrentPath) == false)
                    {
                        Directory.CreateDirectory(CurrentPath);
                    }
                }
            }
            catch { }
            finally
            {
                if (FolderTree != null)
                {
                    FolderTree = null;
                }
            }
        }
        public StreamWriter ftn_titulos_columnas_txt(string[,] tab_titulos, string Carpeta, string File_Name)
        {
            string Line_Buffer = "";
            int i;
            StreamWriter File_IO = new StreamWriter(Carpeta + File_Name, true);

            try
            {
                //poner los titulos
                for (i = 0; i < tab_titulos.GetLength(0); i++)
                {
                    Line_Buffer = Line_Buffer + tab_titulos[i, 0];
                    if (tab_titulos[i, 0] != "")
                    {
                        Line_Buffer = Line_Buffer + "|" + tab_titulos[i, 1];
                    }
                    Line_Buffer = Line_Buffer + "\n";
                }

                File_IO.WriteLine(Line_Buffer);
            }
            catch { }
            finally
            {
                if (Line_Buffer != null)
                {
                    Line_Buffer = string.Empty;
                    Line_Buffer = null;
                }
            }

            return File_IO;
        }
        public StreamWriter ftn_contenido_txt(StreamWriter File_IO, DataTable tabla_llena)
        {
            int i, j;
            //= Concentrado
            string Line_Buffer = "";

            try
            {
                for (i = 0; i < tabla_llena.Rows.Count; i++)
                {
                    for (j = 0; j < tabla_llena.Columns.Count; j++)
                    {
                        Line_Buffer = Line_Buffer + tabla_llena.Rows[i][j] + "" + "\n";
                    }
                }
                File_IO.WriteLine(Line_Buffer);
            }
            catch { }
            finally
            {
                if (Line_Buffer != null)
                {
                    Line_Buffer = string.Empty;
                    Line_Buffer = null;
                }
            }

            return File_IO;
        }
        public char ftn_LettreCol(int numcol)
        {
            char LettreCol = new char();

            try
            {
                if (numcol < 26)
                {
                    numcol = numcol + 65;
                    LettreCol = char.Parse(numcol + "");
                }
                else
                {
                    numcol = numcol + ((numcol / 26) + 64) + ((numcol % 26) + 65);
                    LettreCol = char.Parse(numcol + "");
                }
            }
            catch { }

            return LettreCol;
        }
        public void ftn_excel_simple2(string[,] tab_file, string[,] tab_titulos, DataTable td)
        {
            int i, j, k;
            bool cabeceras = false;

            //Las coordenadas de las celdas en excel comienzan en 1,1 y no en 0 como los arrays o DT
            int n_column, n_row;

            SLDocument My_Workbook = new SLDocument();
            SLStyle styleHeaders = My_Workbook.CreateStyle();
            SLStyle styleRowsDate = My_Workbook.CreateStyle();
            SLStyle styleRowsNumber = My_Workbook.CreateStyle();

            try
            {
                n_column = 1;
                n_row = 2; //se inicializa en 2 porque la primera fila pertenece a la cabecera, el DT comenzara a escribir informacion en la segunda fila

                //Estilos de cabecera:
                styleHeaders.Font.FontName = "Arial";
                styleHeaders.Font.FontSize = 8;
                styleHeaders.Font.Bold = true;
                styleHeaders.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                styleHeaders.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
                styleHeaders.Font.FontColor = System.Drawing.Color.White;
                styleHeaders.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, System.Drawing.Color.Black, System.Drawing.Color.Black);

                //Pinta Cabeceras columnas:
                for (i = 0; i < tab_titulos.GetLength(0); i++)
                {
                    //El primer parametro son las filas, el segundo son las columnas, y el tercero son los datos de la celda:
                    My_Workbook.SetCellValue(n_row - 1, n_column + i, " " + tab_titulos[i, 0] + " ");
                    //Se aplica estilo de las cabeceras;
                    My_Workbook.SetCellStyle(n_row - 1, n_column + i, styleHeaders);
                }

                //Estilos de filas con formato fecha:
                styleRowsDate.Font.FontName = "Arial";
                styleRowsDate.Font.FontSize = 8;
                styleRowsDate.Font.Bold = true;
                styleRowsDate.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                styleRowsDate.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top);
                if (tab_titulos.GetLength(1) > 0)
                {
                    for (i = 0; i < tab_titulos.GetLength(0); i++)
                    {
                        if (tab_titulos[i, 1].Contains("/dd"))
                        {
                            styleRowsDate.FormatCode = tab_titulos[i, 1];
                            break;
                        }
                    }
                }

                //Estilos de filas con formato Numero:
                styleRowsNumber.Font.FontName = "Arial";
                styleRowsNumber.Font.FontSize = 8;
                styleRowsNumber.Font.Bold = true;
                styleRowsNumber.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                styleRowsNumber.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top);
                styleRowsNumber.FormatCode = "0";

                //Pinta Filas datos --> pasa el DataTable al Excel: (index fila, index columna, DataTable, incluye cabecera):
                if (tab_file[2, 0] == null || tab_file[2, 0] == "0")
                {
                    cabeceras = true;
                }
                else
                {
                    cabeceras = false;
                }

                //Llena la hola de calculo:
                My_Workbook.ImportDataTable(n_row, n_column, td, cabeceras);
                My_Workbook.AutoFitColumn(0, td.Columns.Count + 1);

                //Aplica los estilos:
                for (i = 0; i < td.Columns.Count; i++)
                {
                    for (j = 0; j < td.Rows.Count; j++)
                    {
                        if (td.Rows[j][i].ToString().Contains("/22"))
                        {
                            //fechas:
                            My_Workbook.SetCellValue(n_row + j, n_column + i, DateTime.Parse(td.Rows[j][i].ToString()));
                            My_Workbook.SetColumnStyle(i + 1, styleRowsDate);
                            My_Workbook.AutoFitColumn(i + 1, i + 1);
                        }
                        else
                        {
                            My_Workbook.SetColumnStyle(i + 1, styleRowsNumber);
                        }
                    }
                }

                //Aplica el Freze a la primera fila (Encabezados):
                My_Workbook.FreezePanes(1, 0);

                //Renombra la primera hoja de calculo:
                if (tab_file[1, 0] != null)
                {
                    My_Workbook.RenameWorksheet(SLDocument.DefaultFirstSheetName, tab_file[1, 0]);
                }

                //Crea y nombra las siguientes hojas de calculo:
                //if (tab_file[1, 0] != "" && tab_file.GetLength(1) > 0)
                //{
                //    for (i = 0; i < tab_file.GetLength(1); i++)
                //    {
                //        if (tab_file[1, i] != null)
                //        {
                //            My_Workbook.AddWorksheet(tab_file[1, i]);
                //        }
                //    }
                //}

                //Guarda el archivo:
                My_Workbook.SaveAs(tab_file[0, 0] + ".xlsx");
                Console.WriteLine("Archivo Excel generado: " + tab_file[0, 0] + ".xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al generar Excel: " + ex.ToString());
            }
            finally
            {
                if (styleHeaders != null)
                {
                    GC.SuppressFinalize(styleHeaders);
                }
                if (styleRowsDate != null)
                {
                    GC.SuppressFinalize(styleRowsDate);
                }
                if (styleRowsNumber != null)
                {
                    GC.SuppressFinalize(styleRowsNumber);
                }
                if (My_Workbook != null)
                {
                    My_Workbook.Dispose();
                    GC.SuppressFinalize(My_Workbook);
                }
            }
        }

        //Comprime los archivos:
        public bool ftn_compress_file(string Carpeta, string file_name)
        {
            bool res = false;
            string phat = string.Empty;
            string phatZip = string.Empty;
            string phatZipTemp = string.Empty;

            try
            {
                phat = Carpeta;
                phatZip = Carpeta + file_name + ".zip";
                phatZipTemp = Carpeta + file_name + "_zip\\";


                if (File.Exists(phatZip))
                {
                    File.Delete(phatZip);
                }

                /*Para iniciar la comprecion, primero crea una carpeta temporal y copia el archivo(s), despues comprime esa carpeta,
                guarda el zip en la ruta de origen y borra la carpeta temporal*/

                try
                {
                    DirectoryInfo di = new DirectoryInfo(phat);
                    foreach (FileInfo file in di.GetFiles())
                    {
                        if (File.Exists(phat + file.ToString()))
                        {
                            if (Directory.Exists(phatZipTemp))
                            {
                                if (File.Exists(phatZipTemp + file.ToString())) { File.Delete(phatZipTemp + file.ToString()); }
                                Directory.Delete(phatZipTemp);
                            }
                            Directory.CreateDirectory(Carpeta + file_name + "_zip");
                            File.Copy(phat + file.ToString(), phatZipTemp + file.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error al copiar los archivos al phat temporal de comprecion: " + ex.ToString());
                }

                //Comprime:
                ZipFile.CreateFromDirectory(Carpeta + file_name + "_zip\\", phatZip);

                res = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al comprimir los archivos: " + ex.ToString());
                res = false;
            }
            finally
            {
                res = File.Exists(phatZip);

                if (phat != null)
                {
                    phat = string.Empty;
                    phat = null;
                }
                if (phatZip != null)
                {
                    phatZip = string.Empty;
                    phatZip = null;
                }
                if (phatZipTemp != null)
                {
                    phatZipTemp = string.Empty;
                    phatZipTemp = null;
                }
            }

            return res;
        }

        //Envia el correo:
        public string ftn_sendMailReports(string reportName, string[,] tab_archivos, string carpeta, int days_deleted)
        {
            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            //'' mandar el correo para decir que esta listo el reporte ''
            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            string mailServer = string.Empty;
            string from = string.Empty;
            string subject = string.Empty;
            string bodyHTML = string.Empty;
            string qa = string.Empty;

            ArrayList array_files = new ArrayList();
            ArrayList array_size_file = new ArrayList();
            ArrayList array_to = new ArrayList();

            MailMessage message = new MailMessage();
            SmtpClient client = new SmtpClient();

            try
            {
                qa = "";

                mailServer = "192.168.100.6";
                subject = "Report : < " + reportName + " " + qa + " > created.";
                from = "web-reports@logis.com.mx";

                array_to.Add("oscarrp@logis.com.mx");
                //array_to.Add("margaritarg@logis.com.mx");
                //array_to.Add("joseemv@logis.com.mx");
                //array_to.Add("abrahamrr@logis.com.mx");

                //Comprueba si existen los archivos adjuntos a enviar:
                if (tab_archivos[0, 4] == "1")//Si esta activado el envio del zip, solo guarda el zip al arreglo de archivos
                {
                    if (File.Exists(carpeta + tab_archivos[0, 0] + ".zip") == true)
                    {
                        array_files.Add(carpeta + tab_archivos[0, 0] + ".zip");
                    }
                }
                else
                {
                    if (File.Exists(carpeta + tab_archivos[0, 0] + ".xls") == true)
                    {
                        array_files.Add(carpeta + tab_archivos[0, 0] + ".xls");
                    }
                    if (File.Exists(carpeta + tab_archivos[0, 0] + ".xlsx") == true)
                    {
                        array_files.Add(carpeta + tab_archivos[0, 0] + ".xlsx");
                    }
                }

                //De quien, hacia quien, asunto, cuerpo:
                message = new MailMessage();

                //Agrega los contactos
                if (array_to.Count > 0)
                {
                    for (int t = 0; t < array_to.Count; t++)
                    {
                        message.To.Add(array_to[t].ToString());
                    }
                }

                message.From = new MailAddress(from, "Logis report server");
                message.Subject = subject;

                //Adjunta los archivos a enviar:
                for (int j = 0; j < array_files.Count; j++)
                {

                    message.Attachments.Add(new Attachment(array_files[j].ToString()));
                    //Obtiene el tamaño de en KB de cada archivo que se va a adjuntar:
                    array_size_file.Add(new FileInfo(array_files[j].ToString()).Length);
                }

                //body en HTML
                bodyHTML = this.ftn_display_mail("http://www.logiscomercioexterior.com.mx/", "", tab_archivos, days_deleted, array_files, array_size_file, qa);

                message.Body = bodyHTML;
                //Indica que el contenido del correo es en HTMl:
                message.IsBodyHtml = true;

                //FileInfo fileinfo = new FileInfo("dark.jpg");

                //Envia el correo:
                client = new SmtpClient(mailServer);
                //Agrega credenciales de servidor de correo si se requiere, si no, se deja en credenciales de red por defecto:
                client.Credentials = CredentialCache.DefaultNetworkCredentials;

                try
                {
                    //Envio e-mail ...
                    client.Send(message);

                    message.Dispose();
                    client.Dispose();

                    Console.WriteLine("Mail Enviado con exito");
                    return "";
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ocurrio un error al enviar el Mail: {0}" + ex.ToString());
                    return "";
                }

                // Detalle y propiedades de envio del archivo adjunto:
                //ContentDisposition cd = data.ContentDisposition;
                //Console.WriteLine("Content disposition");
                //Console.WriteLine(cd.ToString());
                //Console.WriteLine("File {0}", cd.FileName);
                //Console.WriteLine("Size {0}", cd.Size);
                //Console.WriteLine("Creation {0}", cd.CreationDate);
                //Console.WriteLine("Modification {0}", cd.ModificationDate);
                //Console.WriteLine("Read {0}", cd.ReadDate);
                //Console.WriteLine("Inline {0}", cd.Inline);
                //Console.WriteLine("Parameters: {0}", cd.Parameters.Count);
                //foreach (System.Collections.DictionaryEntry d in cd.Parameters)
                //{
                //    Console.WriteLine("{0} = {1}", d.Key, d.Value);
                //}
            }
            catch { }
            finally
            {
                if (mailServer != null)
                {
                    mailServer = string.Empty;
                    mailServer = null;
                }
                if (from != null)
                {
                    from = string.Empty;
                    from = null;
                }
                if (subject != null)
                {
                    subject = string.Empty;
                    subject = null;
                }
                if (bodyHTML != null)
                {
                    bodyHTML = string.Empty;
                    bodyHTML = null;
                }
                if (qa != null)
                {
                    qa = string.Empty;
                    qa = null;
                }
                if (array_files != null)
                {
                    array_files.Clear();
                    GC.SuppressFinalize(array_files);
                }
                if (array_size_file != null)
                {
                    array_size_file.Clear();
                    GC.SuppressFinalize(array_size_file);
                }
                if (array_to != null)
                {
                    array_to.Clear();
                    GC.SuppressFinalize(array_to);
                }
                if (message != null)
                {
                    message.Dispose();
                    GC.SuppressFinalize(message);
                }
                if (client != null)
                {
                    client.Dispose();
                    GC.SuppressFinalize(client);
                }
            }

            return "";
        }
        public string ftn_display_mail(string servidor, string warning_message, string[,] tab_archivos, int days_deleted, ArrayList array_files, ArrayList array_size_file, string qa, string adittional_info = "")
        {
            //'crea el cuerpo del correo
            //'- servidor : permite de poner las direccion de rede interna o foranea :
            //'             192.168.100.10 o www.logisconcept.com
            //'- warning_message : desplega el contenido del aviso si existe
            //'tab_archivos(0,i) > nombre del archivo
            //'tab_archivos(1,i) > nombre del reporte
            //'tab_archivos(2,i) > tamaño del archivo
            //'tab_archivos(3,i) > Hash MD5
            //'tab_archivos(4,i) > 1 o 0 (o si se olivida, vacio) (si se necesita o no un zip)
            //'tab_archivos(5,i) > tamaño del zip

            int i;
            bool Zip;
            bool Pdf;
            bool Excel;
            string display_mail = "";
            ArrayList array_ext = new ArrayList();


            try
            {
                Zip = false;
                Pdf = false;
                Excel = false;

                display_mail = "<html>";
                display_mail = display_mail + "<body>\n";
                display_mail = display_mail + "		<center>\n";
                display_mail = display_mail + "			<table style='width:400.0pt;height:45pt;mso-cellspacing:0cm;background:white;mso-yfti-tbllook:1184;mso-padding-alt:0cm 0cm 0cm 0cm; border: 1pt solid;border-spacing:0;'>\n";
                display_mail = display_mail + "				<tbody>\n";
                display_mail = display_mail + "					<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>\n";
                display_mail = display_mail + "						<td style='padding:0cm 0cm 0cm 0cm; background-color:#336699;'>\n";
                display_mail = display_mail + "							<img src='" + servidor + "/v2/images/logis.gif' alt='LOGIS' />\n";
                display_mail = display_mail + "						</td>\n";
                display_mail = display_mail + "					</tr>\n";
                display_mail = display_mail + "					<tr>\n";
                display_mail = display_mail + "						<td valign='bottom' style='background:#C69633;padding:2.25pt 2.25pt 2.25pt 2.25pt;height:18.75pt'>\n";
                display_mail = display_mail + "							<p>\n";
                display_mail = display_mail + "								<b>\n";
                display_mail = display_mail + "									<span style='font-size:12.0pt;font-family:+quot;Arial+quot;,sans-serif;mso-fareast-font-family:+quot;Times New Roman+quot;;color:white'>\n";
                display_mail = display_mail + "										Logis <span class='SpellE'>Report</span>Server" + qa + ":\n";
                display_mail = display_mail + "									</span>\n";
                display_mail = display_mail + "								</b>\n";
                display_mail = display_mail + "							</p>\n";
                display_mail = display_mail + "						</td>\n";
                display_mail = display_mail + "					</tr>\n";
                display_mail = display_mail + "					<tr>\n";
                display_mail = display_mail + "						<td  style='padding-top: 10pt;'>\n";
                display_mail = display_mail + "							<p>\n";
                display_mail = display_mail + "								<ul>\n";
                display_mail = display_mail + "									<li style='margin-bottom: 10pt;'>\n";
                display_mail = display_mail + "										<p>\n";
                display_mail = display_mail + "											<b>Report Name:</b>\n";
                display_mail = display_mail + "											" + tab_archivos[0, 1] + "\n";
                display_mail = display_mail + "										</p>\n";
                display_mail = display_mail + "									</li>\n";
                display_mail = display_mail + "									<li style='margin-bottom: 10pt;'>\n";
                display_mail = display_mail + "										<p>\n";
                display_mail = display_mail + "											<b>Date:</b>\n";
                display_mail = display_mail + "											" + DateTime.Now.ToString("dd/MM/yyyy hh:mm") + " \n";
                display_mail = display_mail + "										</p>\n";
                display_mail = display_mail + "									</li>\n";
                display_mail = display_mail + "									<li style='margin-bottom: 10pt;'>\n";
                display_mail = display_mail + "										<p>\n";
                //display_mail = display_mail + "											<b>Direct Link:</b>\n";
                //display_mail = display_mail + "											<a href='http://192.168.100.4/download.asp?id=" + replace(WEL_FIRMA,"'","") + "'>http://192.168.100.4/download.asp?id=" + replace(WEL_FIRMA,"'","") + "</a>\n";'					display_mail = display_mail + "											<b>Direct Link:</b>\n";

                //Configuracion para archivos adjuntos:
                if (array_files.Count != 0)
                {
                    display_mail = display_mail + "<b>Created File(s):</b>\n";
                    display_mail = display_mail + "<ol start='1'>\n";

                    for (int k = 0; k < array_files.Count; k++)
                    {

                        //Obtiene la extencion del los archivos:
                        array_ext.Add(Path.GetExtension(array_files[k].ToString()));

                        switch (array_ext[k].ToString())
                        {
                            case ".zip":
                                if (tab_archivos[0, 4] == "1")//Si se indico que se genere el zip en el arreglo de configuracion de archivos
                                {
                                    display_mail = display_mail + "<li><img src='" + servidor + "/images/winzip2.gif' alt='Zip' />\n";
                                    display_mail = display_mail + "&nbsp; " + tab_archivos[0, 0] + ".zip &nbsp; (" + this.ftn_conviertePeso("KB", array_size_file[k]) + " KB) </li>\n";
                                    Zip = true;
                                }
                                break;

                            case ".xls":
                                display_mail = display_mail + "<li> Excel <img src='" + servidor + "/v2/images/excel.gif' alt='Excel' />\n";
                                display_mail = display_mail + "&nbsp; " + tab_archivos[0, 0] + ".xls &nbsp; (" + this.ftn_conviertePeso("KB", array_size_file[k]) + " KB) </li>\n";
                                Excel = true;
                                break;

                            case ".xlsx":
                                display_mail = display_mail + "<li>  <img src='" + servidor + "/v2/images/excel.gif' alt='Excel' />\n";
                                display_mail = display_mail + "&nbsp;   " + tab_archivos[0, 0] + ".xlsx &nbsp; (" + this.ftn_conviertePeso("KB", array_size_file[k]) + " KB) </li>\n";
                                Excel = true;
                                break;

                            default:
                                break;
                        }
                    }
                }

                display_mail = display_mail + "											</ol>\n";
                display_mail = display_mail + "										</p>\n";
                display_mail = display_mail + "									</li>\n";
                display_mail = display_mail + "								</ul>\n";
                display_mail = display_mail + "							</p>\n";
                //display_mail = display_mail + "							<p>\n";
                //display_mail = display_mail + "								This report will be automatically deleted in " + daysDeleted + " days.<br/>\n";
                //display_mail = display_mail + "							</p>\n";
                display_mail = display_mail + "							<p>\n";
                display_mail = display_mail + "							</p>\n";
                display_mail = display_mail + "							<p style='margin-bottom: 20pt;'>\n";
                display_mail = display_mail + "								Regards. &nbsp; <br/> \n";
                display_mail = display_mail + "								<b>\n";
                display_mail = display_mail + "									Logis Reports Server.<br/>\n";
                display_mail = display_mail + "								</b>\n";
                display_mail = display_mail + "							</p>\n";
                display_mail = display_mail + "						</td>\n";
                display_mail = display_mail + "					</tr>\n";
                display_mail = display_mail + "					<tr>\n";
                display_mail = display_mail + "						<td valign='bottom' style='background:#C69633;padding:2.25pt 2.25pt 2.25pt 2.25pt;height:18.75pt'>\n";
                display_mail = display_mail + "							<p>\n";
                display_mail = display_mail + "								<b>\n";
                display_mail = display_mail + "									<span style='font-size:12.0pt;font-family:+quot;Arial+quot;,sans-serif;mso-fareast-font-family:+quot;Times New Roman+quot;;color:white'>\n";
                display_mail = display_mail + "										Help :\n";
                display_mail = display_mail + "									</span>\n";
                display_mail = display_mail + "								</b>\n";
                display_mail = display_mail + "							</p>\n";
                display_mail = display_mail + "						</td>\n";
                display_mail = display_mail + "					</tr>\n";
                //if WEL_ARCHIVO_XLS = true then
                //    display_mail = display_mail + "					<tr>\n";
                //display_mail = display_mail + "						<td>\n";
                //display_mail = display_mail + "							<p style='margin-top: 15pt;'>\n";
                //display_mail = display_mail + "								<img src='"+servidor+"/images/excel.gif' alt='Excel' />\n";
                //'							display_mail = display_mail + "								<b>+nbsp;&nbsp; Excel</b>: you will need office 2000 (and superior) or <a href='http://office.microsoft.com/downloads/2000/xlviewer.aspx'>XLViewer</a>.\n";
                //display_mail = display_mail + "								<b>&nbsp;&nbsp; Excel</b>: you will need office 2000 (and superior) or <a href='https://docs.microsoft.com/en-us/office/troubleshoot/excel/get-latest-excel-viewer'>XLViewer</a>.\n";
                //display_mail = display_mail + "							</p>\n";
                //display_mail = display_mail + "							<br/>\n";
                //display_mail = display_mail + "						</td>\n";
                //display_mail = display_mail + "					</tr>\n";
                //end if
                display_mail = display_mail + "					<tr>\n";
                display_mail = display_mail + "						<td> &nbsp;\n";

                if (Zip == true)
                {
                    display_mail = display_mail + "    <IMG SRC='" + servidor + "/images/pixel.gif' WIDTH='5' HEIGHT='20' align='bottom' alt=''>\n";
                    display_mail = display_mail + "    <IMG SRC='" + servidor + "/images/winzip2.gif' align='bottom' alt=''>&nbsp;- <b>Zip</b> : In order to reduce your download time, we compressed your report.\n";
                    display_mail = display_mail + "     To open it, you will need Winzip (<a href='http://www.winzip.com' class='link'>free trial</a>) or equivalent : 7-zip (<a href='http://www.7-zip.org' class='link'>free</a>).<br/> \n";
                }
                if (Excel == true)
                {
                    display_mail = display_mail + "    <FONT SIZE='2' FACE='Arial,Helvetica' COLOR='#000000'>\n";
                    display_mail = display_mail + "    <IMG SRC='" + servidor + "/images/pixel.gif' WIDTH='5' HEIGHT='1' alt=''>\n";
                    display_mail = display_mail + "    <IMG SRC='" + servidor + "/images/excel.gif' align='bottom' alt=''>&nbsp;- <b>Excel</b> : you will need office 2000 (and superior) or <a href='http://office.microsoft.com/downloads/2000/xlviewer.aspx' class='link'>XLViewer</a>.<br/>\n";
                }

                display_mail = display_mail + "							<hr/>\n";
                display_mail = display_mail + "							<br/>\n";
                display_mail = display_mail + "							This is a message automatically generated, please contact <a href='mailto:web-master@logis.com.mx'>web-master@logis.com.mx</a> for any question or to unsubscribe.\n";
                display_mail = display_mail + "						</td>\n";
                display_mail = display_mail + "					</tr>\n";
                display_mail = display_mail + "					<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>\n";
                display_mail = display_mail + "						<td style='padding:0cm 0cm 0cm 0cm; background-color:#336699;'>\n";
                display_mail = display_mail + "							&nbsp;\n";
                display_mail = display_mail + "						</td>\n";
                display_mail = display_mail + "					</tr>\n";
                display_mail = display_mail + "				</tbody>\n";
                display_mail = display_mail + "			</table>\n";
                display_mail = display_mail + "		</center>\n";
                display_mail = display_mail + "	</body>\n";
                display_mail = display_mail + "</html>";
            }
            catch { }
            finally
            {
                if (display_mail != null)
                {
                    display_mail = string.Empty;
                    display_mail = null;
                }
                if (array_ext != null)
                {
                    array_ext.Clear();
                    GC.SuppressFinalize(array_ext);
                }
            }

            return display_mail;
        }
        public decimal ftn_conviertePeso(string unidad, object dt)
        {
            decimal resConver = 0;

            try
            {
                if (Convert.ToInt32(dt) <= 0)
                {
                    return 0;
                }
                else
                {
                    switch (unidad)
                    {
                        case "KB":
                            resConver = Convert.ToInt32(dt) / 1024;
                            decimal.Round(resConver);
                            return resConver;
                        case "MB":
                            resConver = Convert.ToInt32(dt) / 1048576;
                            decimal.Round(resConver);
                            return resConver;
                        case "GB":
                            resConver = Convert.ToInt32(dt) / 1073741824;
                            decimal.Round(resConver);
                            return resConver;
                        default:
                            return 0;
                    }
                }
            }
            catch { }
            finally
            {
                GC.SuppressFinalize(resConver);
            }

            return 0;
        }

        //Se usa para crear el nombre del archivo con datos de fecha...
        public string ftn_filter_file_name(string Archivo, string date_1, string date_2)
        {
            //test:
            //Archivo = "Guias_Disponibles_%P";
            //date_1 = "04/01/2022";
            //date_2 = "04/30/2022";

            string[] new_date_1;
            string[] new_date_2;
            string filter_file_name = "";
            string new_customDate, new_customDate1, new_customDate2;

            DateTime localDate = new DateTime();
            DateTime customDate = new DateTime();
            DateTime customDate1 = new DateTime();
            DateTime customDate2 = new DateTime();
            string month = string.Empty;
            int day = 0;
            int year = 0;

            try
            {
                localDate = DateTime.Now;
                month = localDate.ToString("MMM");
                month = month.Replace(".", "");
                month = (CultureInfo.InvariantCulture.TextInfo.ToTitleCase(month));
                day = localDate.Day;
                year = localDate.Year;

                filter_file_name = Archivo.Replace("%M", month);
                filter_file_name = filter_file_name.Replace("%D", day + "");
                filter_file_name = filter_file_name.Replace("%Y", year + "");

                new_date_1 = date_1.Split("/".ToCharArray());
                new_date_2 = date_2.Split("/".ToCharArray());

                if (date_2 != "" && date_2 != date_1)
                {
                    customDate1 = new DateTime(Int32.Parse(new_date_1[2]), Int32.Parse(new_date_1[0]), Int32.Parse(new_date_1[1]));
                    new_customDate1 = customDate1.ToString("MMM-dd-yyyy");
                    new_customDate1 = new_customDate1.Replace(".", "");
                    new_customDate1 = (CultureInfo.InvariantCulture.TextInfo.ToTitleCase(new_customDate1));

                    customDate2 = new DateTime(Int32.Parse(new_date_2[2]), Int32.Parse(new_date_2[0]), Int32.Parse(new_date_2[1]));
                    new_customDate2 = customDate2.ToString("MMM-dd-yyyy");
                    new_customDate2 = new_customDate2.Replace(".", "");
                    new_customDate2 = (CultureInfo.InvariantCulture.TextInfo.ToTitleCase(new_customDate2));

                    filter_file_name = filter_file_name.Replace("%P", new_customDate1) + new_customDate2;
                }
                else
                {
                    new_date_1 = new String[3];

                    new_date_1[2] = year.ToString();
                    month = localDate.ToString("MM");
                    new_date_1[0] = month.ToString();
                    new_date_1[1] = day.ToString();

                    customDate1 = new DateTime(int.Parse(year.ToString()), int.Parse(month.ToString()), int.Parse(day.ToString()));
                    new_customDate1 = customDate1.ToString("MMM-dd-yyyy");
                    new_customDate1 = new_customDate1.Replace(".", "");
                    new_customDate1 = (CultureInfo.InvariantCulture.TextInfo.ToTitleCase(new_customDate1));

                    filter_file_name = filter_file_name.Replace("%P", new_customDate1);
                }

                customDate = new DateTime(Int32.Parse(new_date_1[2]), Int32.Parse(new_date_1[0]), Int32.Parse(new_date_1[1]));
                new_customDate = customDate.ToString("MMM-dd-yyyy");
                new_customDate = new_customDate.Replace(".", "");
                new_customDate = (CultureInfo.InvariantCulture.TextInfo.ToTitleCase(new_customDate));
                filter_file_name = filter_file_name.Replace("%p", new_customDate);

                //filter_file_name = Mid(FSO.GetBaseName(FSO.GetTempName), 4) & "-" & filter_file_name
            }
            catch { }
            finally
            {
                if (month != null)
                {
                    month = string.Empty;
                    month = null;
                }
                if (localDate != null)
                {
                    GC.SuppressFinalize(localDate);
                }
                if (customDate != null)
                {
                    GC.SuppressFinalize(customDate);
                }
                if (customDate1 != null)
                {
                    GC.SuppressFinalize(customDate1);
                }
                if (customDate2 != null)
                {
                    GC.SuppressFinalize(customDate2);
                }
                GC.SuppressFinalize(day);
                GC.SuppressFinalize(year);
            }

            return filter_file_name;
        }
        public string ftn_NVL(string str)
        {
            string NVL = "";

            try
            {
                if (str.Equals(null))
                {
                    NVL = "";
                }
                else
                {
                    NVL = str;
                }
            }
            catch
            {
                NVL = "";
            }

            return NVL;
        }
        public int ftn_NVL_num(string str)
        {
            int NVL_num = 0;

            try
            {
                if (str.Equals(null))
                {
                    NVL_num = 0;
                }
                else if (str.Equals(""))
                {
                    NVL_num = 0;
                }
                else
                {
                    NVL_num = Int32.Parse(str);
                }
            }
            catch
            {
                NVL_num = 0;
            }

            return NVL_num;
        }
        public string ftn_Left(string str, int count)
        {
            string Left = "";

            try
            {
                if (str.Equals(null))
                {
                    Left = "";
                }
                else if (str.Equals(""))
                {
                    Left = "";
                }
                else
                {
                    Left = str.Substring(0, count);
                }
            }
            catch
            {
                Left = "";
            }

            return Left;
        }
        public string ftn_Right(string str, int count)
        {
            string Right = "";

            try
            {
                if (str.Equals(null))
                {
                    Right = "";
                }
                else if (str.Equals(""))
                {
                    Right = "";
                }
                else
                {
                    Right = str.Substring(str.Length - count, count);
                }
            }
            catch
            {
                Right = "";
            }

            return Right;
        }
        public long getRandom()
        {
            int x = 0;
            int s = 0;
            Random r = null;
            long randomNumber = 0;

            try
            {
                for (x = 0; x <= 10; x++)
                {
                    s = Environment.TickCount;
                    r = new Random(s);
                    randomNumber = r.Next();
                }
            }
            catch
            {
                randomNumber += DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day;
                randomNumber += DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            }
            finally
            {
                if (r != null)
                {
                    GC.SuppressFinalize(r);
                }
            }

            return randomNumber;
        }

        //public void ftn_delete_old_file(string path)
        //{
        //    try
        //    {
        //        if (File.Exists(path))
        //        {
        //            File.Delete(path);
        //            Console.WriteLine("Archivo eliminado");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error al borrar los archivos: " + ex.ToString());
        //    }
        //}


        //Borra los archivos en la carpeta de generacion del reporte:
        public void ftn_delete_old_file(string Carpeta, string file_name)
        {
            string phatZipTemp = string.Empty;
            DirectoryInfo di = null;

            try
            {
                phatZipTemp = Carpeta + file_name + "_zip\\";
                di = new DirectoryInfo(Carpeta);

                foreach (FileInfo file in di.GetFiles())
                {
                    if (Directory.Exists(phatZipTemp))
                    {
                        if (File.Exists(phatZipTemp + file.ToString())) { File.Delete(phatZipTemp + file.ToString()); }
                        Directory.Delete(phatZipTemp);
                        Console.WriteLine("Carpeta de comprecion temporal eliminada");
                    }
                    if (File.Exists(Carpeta + file.ToString()))
                    {
                        file.Delete();
                        Console.WriteLine("Archivo " + file.ToString() + " eliminado");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al borrar los archivos: " + ex.ToString() + " \n\t " + phatZipTemp);
            }
            finally
            {
                if (phatZipTemp != null)
                {
                    phatZipTemp = string.Empty;
                    phatZipTemp = null;
                }
                if (di != null)
                {
                    di = null;
                }
            }
        }


        public void ftn_file_instant_delete(List<string> phat)
        {
            List<string> lstArchivosDelete = new List<string>();
            lstArchivosDelete = phat;

            for (int i = 0; i < lstArchivosDelete.Count; i++)
            {
                //fun.DeleteOldFile(lstArchivos[i], days_deleted);
                if (File.Exists(lstArchivosDelete[i]))
                {
                    try
                    {
                        //Con true, indicamos borrar recursivamente todo lo que cuelgue del directorio
                        Directory.Delete(Directory.GetParent(lstArchivosDelete[i]).ToString(), true);
                        ImprimeConsola("Borra carpeta generada en el servidor " + lstArchivosDelete[i]);
                    }
                    catch (Exception ex_fi)
                    {
                        ImprimeConsola("Error al borrar archivos " + ex_fi);
                    }
                }
                else
                {
                    ImprimeConsola("El archivo no existe ");
                }
            }
        }

        public void ImprimeConsola(string texto)
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


        //public string DataTableToExcel(DataTable dtInfo, string fileName)
        public string DataTableToExcel(DataTable dtInfo, string fileName, List<string> nameSheets)
        {
            int iRow = 0;
            int iColumn = 0;
            SLDocument wbook = null;
            SLStyle styleHeaders = null;

            try
            {
                //WorkBookStyles:
                wbook = new SLDocument();
                styleHeaders = wbook.CreateStyle();
                styleHeaders.Font.FontName = "Arial";
                styleHeaders.Font.FontSize = 8;
                styleHeaders.Font.Bold = true;
                styleHeaders.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                styleHeaders.SetVerticalAlignment(VerticalAlignmentValues.Center);
                styleHeaders.Font.FontColor = System.Drawing.Color.White;
                //styleHeaders.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.White, System.Drawing.Color.Black);
                styleHeaders.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Dark1Color, SLThemeColorIndexValues.Light1Color);


                SLStyle styleRows = wbook.CreateStyle();
                styleRows.Font.FontName = "Arial";
                styleRows.Font.FontSize = 8;
                //styleRows.Font.Bold = true;
                styleRows.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                styleRows.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top);
                styleRows.FormatCode = "0";


                //Headers:
                iRow = 1;
                iColumn = 1;
                foreach (DataColumn dc in dtInfo.Columns)
                {
                    wbook.SetCellValue(iRow, iColumn, " " + dc.ColumnName + " ");
                    wbook.SetCellStyle(iRow, iColumn, styleHeaders);
                    iColumn++;
                }
                wbook.FreezePanes(1, 0);


                //Content:
                iRow = 2;
                iColumn = 1;
                wbook.ImportDataTable(iRow, iColumn, dtInfo, false);
                wbook.AutoFitColumn(0, dtInfo.Columns.Count + 1);

                //StylesRows
                foreach (DataRow dr in dtInfo.Rows)
                {
                    //wbook.AutoFitRow(1);
                    wbook.SetRowStyle(iColumn + 1, styleRows);
                    iColumn++;
                }



                //Save File:
                try
                {
                    if (!dtInfo.TableName.Equals(string.Empty)) { sResult = dtInfo.TableName; }
                    else { sResult = fileName.ToLower().Replace(".xlsx", string.Empty).Replace(".xls", string.Empty); }
                }
                catch
                {
                    sResult = SLDocument.DefaultFirstSheetName;
                }
                finally
                {

                    //Rename Sheet's
                    for (int k = 0; k < nameSheets.Count; k++)
                    {
                       
                        if (k == 0)
                        {
                            wbook.RenameWorksheet(SLDocument.DefaultFirstSheetName, nameSheets[k]);
                        }
                        else {
                            wbook.AddWorksheet(nameSheets[k]);
                        }
                    }

                    
                    sResult = string.Format("{0}/{1}{2}",
                                            sTempPath,
                                            fileName,
                                            fileName.ToLower().EndsWith(".xlsx") ?
                                                string.Empty :
                                                fileName.ToLower().EndsWith(".xls") ?
                                                string.Empty :
                                                string.Format(".xlsx")
                                                );
                    wbook.SaveAs(sResult);
                }
            }
            catch (Exception ex)
            {
                sResult = false.ToString();
                //WriteExceptionOnLog(ex);
            }
            finally
            {
                if (styleHeaders != null)
                {
                    GC.SuppressFinalize(styleHeaders);
                }
                if (wbook != null)
                {
                    wbook.Dispose();
                    GC.SuppressFinalize(wbook);
                }
                GC.Collect();
            }

            return sResult;
        }

        /// <summary>
        /// Libera los recursos utilizados por el objeto.
        /// </summary>
        public void Dispose()
        {
            if (this.first_path != null)
            {
                this.first_path = string.Empty;
                this.first_path = null;
            }
            if (this.second_path != null)
            {
                this.second_path = string.Empty;
                this.second_path = null;
            }
            if (this.IP_servidor1 != null)
            {
                this.IP_servidor1 = string.Empty;
                this.IP_servidor1 = null;
            }
            if (this.IP_servidor2 != null)
            {
                this.IP_servidor2 = string.Empty;
                this.IP_servidor2 = null;
            }
            if (this.ini_path != null)
            {
                this.ini_path = string.Empty;
                this.ini_path = null;
            }
            GC.Collect();
        }
    }
}