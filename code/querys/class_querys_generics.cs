using ReportServer2022.code.include;
using System;
using System.IO;

namespace ReportServer2022.code.querys
{
    public class class_querys_generics
    {
        class_conexion conexion = new class_conexion();
        class_xfunciones obj_xfunciones = new class_xfunciones();
        LogisFunctions fun = new LogisFunctions(System.Reflection.Assembly.GetExecutingAssembly());

        public class_llena_tabla tbl_reporte_ejecutar_consultar = new class_llena_tabla();
        public class_llena_tabla tbl_parametros_reporte_consultar = new class_llena_tabla();
        public class_llena_tabla tbl_libera_InProgress = new class_llena_tabla();


        public bool bandera;
        public string SQL;
        private int i, j;

        public string reporte_name;
        public long days_deleted;
        public string file_name;
        public int id_Reporte;
        public string Carpeta;
        public string servidor;
        public string Fecha_1;
        public string Fecha_2;
        public string dest_mail;
        public string param_string;
        public string[,] tab_archivos = new string[1, 5];


        //**********************************************QUERYS DE INICIALIZACION**********************************************

        /*
         * ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           '' Verificar los datos del reporte que hablo el report server ''
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         */

        public void sub_reporte_ejecutar_consultar(string rep_id, int reporte_temporal)
        {
            conexion.sub_set_conexion();
            tbl_reporte_ejecutar_consultar.cadena_conexion = conexion.db_cadena;
            tbl_reporte_ejecutar_consultar.sqltext = "SELECT	 REP.ID_REP \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,REP.ID_CRON \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,REP.NAME \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,REP.CONFIRMACION \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,REP.FRECUENCIA \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,REP.CLIENTE \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,CLI.CLISTATUS \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		,CLI.CLICLEF || ' - ' || InitCap(CLI.CLINOM) CLI_NOM \n";
            tbl_reporte_ejecutar_consultar.sqltext += "FROM	 rep_detalle_reporte REP \n";
            tbl_reporte_ejecutar_consultar.sqltext += "	INNER JOIN	eclient CLI \n";
            tbl_reporte_ejecutar_consultar.sqltext += "		ON	CLI.CLICLEF	=	REP.CLIENTE \n";
            tbl_reporte_ejecutar_consultar.sqltext += "WHERE	REP.ID_CRON	=	" + rep_id + " \n";
            tbl_reporte_ejecutar_consultar.sub_llenar_tabla();

            if (tbl_reporte_ejecutar_consultar.tabla_llena.Rows.Count < 1)
            {
                //INVOCA LA FUNCION DE ERROR(POR PROGRAMAR)

            }
            else if (tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["id_rep"].ToString() != "317"
                && tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["clistatus"].ToString() == "1"
                && reporte_temporal != 1
                && Int32.Parse(tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["cliente"].ToString()) != 0
                && Int32.Parse(tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["cliente"].ToString()) < 9900 || Int32.Parse(tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["cliente"].ToString()) > 9999)
            {
                //INVOCA A LA FUNCION DE ENVIO DE CORREO DE ERROR
                //INVOCA LA FUNCION DE ERROR(POR PROGRAMAR)

            }


            /*
             * 'no registro
                ''''''''''''''''''''''''''''''''''''''''''''''''
                ''      Verificar si no es un dia libre       ''
                '' cliente 0 -> dia libre de todo la empresa  ''
                '' solo por los reportes diarios (freq : 0,1) ''
                '' ingresar las fechas en variables en caso   ''
                '' que el reporte no necesita confirmacion    ''
                ''''''''''''''''''''''''''''''''''''''''''''''''
             */

            if (reporte_temporal == 0)
            {
                class_llena_tabla tbl_fecha_confirmacion_consulta = new class_llena_tabla();
                conexion.sub_set_conexion();
                tbl_fecha_confirmacion_consulta.cadena_conexion = conexion.db_cadena;

                tbl_fecha_confirmacion_consulta.sqltext = "SELECT	logis.display_fecha_confirmacion4('5', SYSDATE, SYSDATE, 1)	AS	fecha \n";
                tbl_fecha_confirmacion_consulta.sqltext += "FROM	dual \n";

                tbl_fecha_confirmacion_consulta.sub_llenar_tabla();

                if (tbl_fecha_confirmacion_consulta.tabla_llena.Rows.Count > 0)
                {

                    Fecha_1 = fun.Left(tbl_fecha_confirmacion_consulta.tabla_llena.Rows[0]["fecha"].ToString(), 10);
                    Fecha_2 = fun.Right(tbl_fecha_confirmacion_consulta.tabla_llena.Rows[0]["fecha"].ToString(), 10);


                    //Reprocesar:
                    //MM/dd/yyyy
                    //Fecha_1 = "07/16/2022";
                    //Fecha_2 = "07/16/2022";

                    bandera = true;
                }
                else
                {
                    bandera = false;
                }
            }
            else
            {
                class_llena_tabla tbl_fecha_confirmacion2_consulta = new class_llena_tabla();
                conexion.sub_set_conexion();
                tbl_fecha_confirmacion2_consulta.cadena_conexion = conexion.db_cadena;
                tbl_fecha_confirmacion2_consulta.sqltext = "select to_char(LAST_CONF_DATE_1, 'mm/dd/yyyy')  as fecha_1, to_char(LAST_CONF_DATE_2, 'mm/dd/yyyy')  as fecha_2 \n"
                            + "From rep_detalle_reporte \n"
                            + "where id_cron='" + rep_id + "' \n";
                tbl_fecha_confirmacion2_consulta.sub_llenar_tabla();

                if (tbl_fecha_confirmacion2_consulta.tabla_llena.Rows.Count > 0)
                {
                    Fecha_1 = fun.NVL(tbl_fecha_confirmacion2_consulta.tabla_llena.Rows[0]["fecha_1"].ToString());
                    Fecha_2 = fun.NVL(tbl_fecha_confirmacion2_consulta.tabla_llena.Rows[0]["fecha_2"].ToString());

                    bandera = true;
                }
                else
                {
                    bandera = false;
                }
            }

            
            if (Fecha_1.Equals(Fecha_2))
            {
                class_llena_tabla tbl_dias_libres_consultar = new class_llena_tabla();
                conexion.sub_set_conexion();
                tbl_dias_libres_consultar.cadena_conexion = conexion.db_cadena;
                tbl_dias_libres_consultar.sqltext = "select 1 from rep_dias_libres \n"
                            + "where dia_libre = to_date('" + Fecha_1 + "', 'mm/dd/yyyy') \n"
                            + "and cliente in ('" + tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["cliente"].ToString() + "',0) \n";
                tbl_dias_libres_consultar.sub_llenar_tabla();
                //si hay un registro es que el dia o es libre por el cliente o por toda la empresa

                if (tbl_dias_libres_consultar.tabla_llena.Rows.Count > 0)
                //if (tbl_dias_libres_consultar.tabla_llena.Rows.Count == 0)
                {

                    //decir a la tabla rep_chron que esta generado el reporte : ponemos el campo IN_PROGRESS a 0
                    class_llena_tabla tbl_genera_reporte_update = new class_llena_tabla();
                    conexion.sub_set_conexion();
                    tbl_genera_reporte_update.cadena_conexion = conexion.db_cadena;
                    tbl_genera_reporte_update.sqltext = "update rep_chron set in_progress=0 \n"
                               + "where id_rapport= '" + rep_id + "' \n";
                    tbl_genera_reporte_update.sub_llenar_tabla();

                    bandera = true;
                }
            }
            else
            {
                bandera = false;
            }


            /*
           * '''''''''''''''''''''''''''''''''''''''''''
             '' El reporte necesita una confirmacion  ''
             '' verificamos que llego                 ''
             '''''''''''''''''''''''''''''''''''''''''''
           */

            if (tbl_reporte_ejecutar_consultar.tabla_llena.Rows.Count > 0) //comprueba si tiene datos el DT de parametros del reporte
            {
                if (tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["CONFIRMACION"].Equals("1") && reporte_temporal == 0)
                {
                    //verificar en caso de confirmacion 10 todas las aduanas por un reporte
                    /* (comentando en vb6)
                    '    SQL = " select check_fecha_confirmacion2('" & rs.Fields("FRECUENCIA") & "',conf_date, conf_date_2) as ok"
                    '    SQL = SQL & " , to_char(conf.conf_date, 'mm/dd/yyyy') as fecha_1 " 
                    '    SQL = SQL & " , to_char(conf.conf_date_2, 'mm/dd/yyyy') as fecha_2, conf.param "
                    '    SQL = SQL & " from rep_confirmacion conf"
                    '    SQL = SQL & " where  conf.ID_CONF = '" & rep_id & "' "
                    '    SQL = SQL & " and  check_fecha_confirmacion2('" & rs.Fields("FRECUENCIA") & "',conf_date, conf_date_2) = 'ok' "
                    '    SQL = SQL & " order by ok desc"
                     */

                    class_llena_tabla tbl_fecha_confirmacion3_consulta = new class_llena_tabla();
                    conexion.sub_set_conexion();
                    tbl_fecha_confirmacion3_consulta.cadena_conexion = conexion.db_cadena;
                    tbl_fecha_confirmacion3_consulta.sqltext = "select check_fecha_confirmacion2('" + tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["FRECUENCIA"].ToString() + "',conf_date, conf_date_2) as ok \n"
                                + " , to_char(conf.conf_date, 'mm/dd/yyyy') as fecha_1  \n"
                                + " , to_char(conf.conf_date_2, 'mm/dd/yyyy') as fecha_2, conf.param  \n"
                                + " from rep_confirmacion conf \n"
                                + " where  conf.ID_CONF = '" + rep_id + "'  \n"
                                + " and  check_fecha_confirmacion2('" + tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["FRECUENCIA"].ToString() + "',conf_date, conf_date_2) = 'ok'  \n"
                                + " and trunc(conf_date) + decode('" + tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["FRECUENCIA"].ToString() + "',1,1,0) <= trunc(sysdate) \n";
                    //'asi quitamos la validaciones posteriores
                    tbl_fecha_confirmacion3_consulta.sqltext = tbl_fecha_confirmacion3_consulta.sqltext + " order by conf_date desc \n";

                    tbl_fecha_confirmacion3_consulta.sub_llenar_tabla();

                    if (tbl_fecha_confirmacion3_consulta.tabla_llena.Rows.Count < 1)
                    {

                        class_llena_tabla tbl_fecha_confirmacion4_consulta = new class_llena_tabla();
                        conexion.sub_set_conexion();
                        tbl_fecha_confirmacion4_consulta.cadena_conexion = conexion.db_cadena;
                        tbl_fecha_confirmacion4_consulta.sqltext = "select display_fecha_confirmacion4('" + tbl_reporte_ejecutar_consultar.tabla_llena.Rows[0]["FRECUENCIA"].ToString() + "',conf.CONF_DATE,conf.CONF_DATE_2,decode(conf.CONF_DATE,null,1,0)) as next_fecha \n"
                                + "from rep_confirmacion conf \n"
                                + "where  conf.ID_CONF = '" + rep_id + "'  \n"
                                + "order by to_date(next_fecha, 'mm/dd/yyyy') desc \n";
                        tbl_fecha_confirmacion4_consulta.sub_llenar_tabla();

                        if (tbl_fecha_confirmacion4_consulta.tabla_llena.Rows.Count < 1)
                        {
                            //INVOCA A LA FUNCION DE ENVIO DE CORREO DE ERROR
                            //INVOCA LA FUNCION DE ERROR(POR PROGRAMAR)

                            // mail_error = "Ninguna confirmacion llegada."
                            bandera = false;
                        }
                        else
                        {
                            //'la confirmacion no esta lista, mandar un correo con el error
                            //mail_error = rs2.Fields(0).Value

                            bandera = true;
                        }
                    }
                    else
                    {
                        Fecha_1 = fun.NVL(tbl_fecha_confirmacion3_consulta.tabla_llena.Rows[0]["fecha_1"].ToString());
                        Fecha_2 = fun.NVL(tbl_fecha_confirmacion3_consulta.tabla_llena.Rows[0]["fecha_2"].ToString());

                        bandera = true;

                        //(comentando en vb6)
                        //'param = NVL(rs2.Fields("param"))
                        //'penser a agreger les params ds le cas de plusieurs douanes
                    }

                }
            }
        }



        //Consulta el numero y los parametros del reporte:
        public void sub_numero_parametros_reporte_consultar(string rep_id)
        {

            class_llena_tabla tbl_numero_parametros_reporte_consultar = new class_llena_tabla();
            conexion.sub_set_conexion();

            tbl_numero_parametros_reporte_consultar.cadena_conexion = conexion.db_cadena;
            tbl_numero_parametros_reporte_consultar.sqltext = "SELECT	REPORT.NUM_OF_PARAM \n";
            tbl_numero_parametros_reporte_consultar.sqltext += "FROM	REP_REPORTE REPORT \n";
            tbl_numero_parametros_reporte_consultar.sqltext += "	INNER JOIN	REP_DETALLE_REPORTE REP \n";
            tbl_numero_parametros_reporte_consultar.sqltext += "		ON	REPORT.ID_REP	=	REP.ID_REP \n";
            tbl_numero_parametros_reporte_consultar.sqltext += "WHERE	REP.ID_CRON	=	'" + rep_id + "' \n";

            tbl_numero_parametros_reporte_consultar.sub_llenar_tabla();

            if (tbl_numero_parametros_reporte_consultar.tabla_llena.Rows.Count > 0)
            {
                int num_of_param;
                num_of_param = Int32.Parse(tbl_numero_parametros_reporte_consultar.tabla_llena.Rows[0]["NUM_OF_PARAM"].ToString());

                conexion.sub_set_conexion();
                tbl_parametros_reporte_consultar.cadena_conexion = conexion.db_cadena;

                tbl_parametros_reporte_consultar.sqltext = "SELECT	 REP.NAME ,REP.CLIENTE \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.FILE_NAME ,REP.CARPETA \n";
                tbl_parametros_reporte_consultar.sqltext += "		,CLI.CLINOM \n";
                tbl_parametros_reporte_consultar.sqltext += "		,MAIL.NOMBRE ,MAIL.MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REPORT.COMMAND \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.DAYS_DELETED ,REPORT.NUM_OF_PARAM \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.DEST_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "		,TO_CHAR(REP.LAST_CONF_DATE_1, 'mm/dd/yyyy') LAST_CONF_DATE_1 ,TO_CHAR(REP.LAST_CONF_DATE_2, 'mm/dd/yyyy') LAST_CONF_DATE_2 \n";

                //tbl_parametros_reporte_consultar.sqltext += "		,NVL(REP.PARAM_1,'') PARAM_1, NVL(REP.PARAM_2,'') PARAM_2 \n";
                for (i = 1; i <= num_of_param; i++)
                {
                    tbl_parametros_reporte_consultar.sqltext += ", NVL(REP.PARAM_" + i + ",'') PARAM_" + i + "\n";
                }

                tbl_parametros_reporte_consultar.sqltext += "		,MAIL.CLIENT_NUM ,REPORT.ID_REP ,REPORT.SUBCARPETA \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.CREATED_BY, TERCERO \n";
                tbl_parametros_reporte_consultar.sqltext += "FROM	 REP_DETALLE_REPORTE REP \n";
                tbl_parametros_reporte_consultar.sqltext += "	INNER JOIN	REP_REPORTE REPORT \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	REP.ID_REP		=	REPORT.ID_REP \n";
                tbl_parametros_reporte_consultar.sqltext += "	LEFT JOIN	ECLIENT CLI \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	REP.CLIENTE 	= CLI.CLICLEF \n";
                tbl_parametros_reporte_consultar.sqltext += "	LEFT JOIN	REP_DEST_MAIL DEST \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	REP.MAIL_OK		=	DEST.ID_DEST_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "	LEFT JOIN	REP_MAIL MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	DEST.ID_DEST	=	MAIL.ID_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "WHERE	NVL(MAIL.STATUS, 1) = 1 \n";
                tbl_parametros_reporte_consultar.sqltext += "	AND	REP.ID_CRON	=	'" + rep_id + "' \n";
                tbl_parametros_reporte_consultar.sqltext += "UNION ALL \n";
                tbl_parametros_reporte_consultar.sqltext += "SELECT	 REP.NAME ,REP.CLIENTE \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.FILE_NAME ,REP.CARPETA \n";
                tbl_parametros_reporte_consultar.sqltext += "		,CLI.CLINOM \n";
                tbl_parametros_reporte_consultar.sqltext += "		,MAIL.NOMBRE ,MAIL.MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REPORT.COMMAND \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.DAYS_DELETED ,REPORT.NUM_OF_PARAM \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.DEST_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "		,TO_CHAR(REP.LAST_CONF_DATE_1, 'mm/dd/yyyy') LAST_CONF_DATE_1 ,TO_CHAR(REP.LAST_CONF_DATE_2, 'mm/dd/yyyy') LAST_CONF_DATE_2 \n";

                //tbl_parametros_reporte_consultar.sqltext += "		,REP.PARAM_1 ,REP.PARAM_2 \n";
                for (j = 1; j <= num_of_param; j++)
                {
                    tbl_parametros_reporte_consultar.sqltext += ", NVL(REP.PARAM_" + j + ",'') PARAM_" + j + "\n";
                }

                tbl_parametros_reporte_consultar.sqltext += "		,MAIL.CLIENT_NUM ,REPORT.ID_REP ,REPORT.SUBCARPETA \n";
                tbl_parametros_reporte_consultar.sqltext += "		,REP.CREATED_BY ,TERCERO \n";
                tbl_parametros_reporte_consultar.sqltext += "FROM	 REP_DETALLE_REPORTE REP \n";
                tbl_parametros_reporte_consultar.sqltext += "	INNER JOIN	REP_REPORTE REPORT \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	REP.ID_REP	=	REPORT.ID_REP \n";
                tbl_parametros_reporte_consultar.sqltext += "	LEFT JOIN	ECLIENT CLI \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	REP.CLIENTE = CLI.CLICLEF \n";
                tbl_parametros_reporte_consultar.sqltext += "	LEFT JOIN	REP_DEST_MAIL DEST \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	REP.MAIL_OK		=	DEST.ID_DEST_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "	LEFT JOIN	REP_MAIL MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "		ON	DEST.ID_DEST	=	MAIL.ID_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "WHERE	NVL(MAIL.STATUS, 1) = 1 \n";
                tbl_parametros_reporte_consultar.sqltext += "	AND	REP.MAIL_OK IS NOT NULL \n";
                tbl_parametros_reporte_consultar.sqltext += "	AND	NOT EXISTS(	SELECT	NULL \n";
                tbl_parametros_reporte_consultar.sqltext += "					FROM	REP_DEST_MAIL DESTD \n";
                tbl_parametros_reporte_consultar.sqltext += "						INNER JOIN	REP_MAIL MAILD \n";
                tbl_parametros_reporte_consultar.sqltext += "							ON	DESTD.ID_DEST	=	MAILD.ID_MAIL \n";
                tbl_parametros_reporte_consultar.sqltext += "					WHERE	DESTD.ID_DEST_MAIL	=	REP.MAIL_OK \n";
                tbl_parametros_reporte_consultar.sqltext += "						AND	MAILD.STATUS		=	1 ) \n";
                tbl_parametros_reporte_consultar.sqltext += "	AND	DEST.ID_DEST_MAIL	=	2888 \n";
                tbl_parametros_reporte_consultar.sqltext += "	AND	REP.ID_CRON			=	'" + rep_id + "' \n";
                tbl_parametros_reporte_consultar.sqltext += "ORDER BY	CLIENT_NUM ,TERCERO DESC ,NOMBRE \n";

                tbl_parametros_reporte_consultar.sub_llenar_tabla();


                if (tbl_parametros_reporte_consultar.tabla_llena.Rows.Count > 0)

                {
                    bandera = true;


                    reporte_name = fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["NAME"].ToString());
                    days_deleted = fun.NVL_Number(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["DAYS_DELETED"].ToString());
                    file_name = fun.filter_file_name(fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["FILE_NAME"].ToString()), Fecha_1, Fecha_2);
                    id_Reporte = Int32.Parse(fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["ID_REP"].ToString()));


                    //verificamos cual servidor esta bien en caso que uno falla (VB6):
                    //'If FSO.FolderExists(first_path) Then

                    obj_xfunciones.sub_init_var();

                    Carpeta = obj_xfunciones.first_path + fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["CARPETA"].ToString()) + "\\";
                    if (fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["SUBCARPETA"].ToString()) != "")
                    {
                        Carpeta = Carpeta + fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["SUBCARPETA"].ToString()) + "\\";
                    }
                    else
                    {
                        Carpeta = Carpeta + "";
                    }

                    //servidor = "http://" & Trim(Split(Get_IP(), "-")(0))


                    //
                    //'ElseIf FSO.FolderExists(second_path) Then
                    //'    Carpeta = second_path & NVL(rs.Fields("CARPETA")) & "\" & IIf(NVL(rs.Fields("SUBCARPETA")) <> "", NVL(rs.Fields("SUBCARPETA")) & "\", "")
                    //'    servidor = "http://" & IP_servidor2
                    //'    'intercambiamos los path
                    //'    'asi el first sera siempre el que funcinona
                    //'    temp_path = first_path
                    //'    first_path = second_path
                    //'    second_path = temp_path
                    //'
                    //'Else
                    //'    'ninguno de los 2 servidores funciona
                    //'    Call Err.Raise(-1, , "Los 2 servidores no son accesibles")
                    //'    GoTo Errman
                    //'End If



                    //            dest_mail = NVL(rs.Fields("DEST_MAIL"))
                    //For i = 1 To num_of_param
                    //    param_string = param_string & rs.Fields("PARAM_" & i)
                    //    If i<> num_of_param Then param_string = param_string & "|"
                    //Next

                    //If rs.Fields("DEST_MAIL") <> "" Then
                    //    Fecha_1 = NVL(rs.Fields("LAST_CONF_DATE_1"))
                    //    Fecha_2 = NVL(rs.Fields("LAST_CONF_DATE_2"))
                    //End If

                    //                    'verificamos si la carpeta exista, si no la creamos
                    //If Not FSO.FolderExists(Carpeta) Then
                    //    Call Create_Entire_Path(Carpeta)
                    //End If


                    dest_mail = fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["DEST_MAIL"].ToString());
                    for (i = 1; i < num_of_param; i++)
                    {
                        param_string = param_string + tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["PARAM_" + i].ToString();
                        if (i != num_of_param)
                        {
                            param_string = param_string + "|";
                        }
                    }

                    if (tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["DEST_MAIL"].ToString() != "")
                    {
                        Fecha_1 = fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["LAST_CONF_DATE_1"].ToString());
                        Fecha_2 = fun.NVL(tbl_parametros_reporte_consultar.tabla_llena.Rows[0]["LAST_CONF_DATE_2"].ToString());
                    }


                    string creado;

                    //verificamos que la carpeta exista, si no la creamos:
                    if (Directory.Exists(Carpeta) == false)
                    {
                        obj_xfunciones.sub_Create_Entire_Path(Carpeta);
                        creado = "1";
                    }
                    else
                    {
                        creado = "0";
                    }

                    //Console.WriteLine(creado);


                    //tab_archivos(0,i) > nombre del archivo
                    //tab_archivos(1,i) > nombre del reporte
                    //tab_archivos(2,i) > tamaño del archivo
                    //tab_archivos(3,i) > Hash MD5
                    //tab_archivos(4,i) > 1 o 0 (o si se olivida, vacio) (si se necesita o no un zip)
                    //tab_archivos(5,i) > tamaño del zip

                    tab_archivos[0, 0] = file_name;
                    tab_archivos[0, 1] = reporte_name;
                    tab_archivos[0, 4] = "1";

                }

                else
                {
                    bandera = false;
                }
            }
            else
            {
                bandera = false;
            }
        }

        public void sub_libera_InProgress(string id_rapport)
        {
            conexion.sub_set_conexion();
            tbl_libera_InProgress.cadena_conexion = conexion.db_cadena;
            tbl_libera_InProgress.sqltext = "update rep_chron set in_progress = 0 \n";
            tbl_libera_InProgress.sqltext += "where id_rapport = " + id_rapport + "";
            tbl_libera_InProgress.sub_llenar_tabla();
        }


        /// <summary>
        /// Libera los recursos utilizados por el objeto.
        /// </summary>
        public void Dispose()
        {
            if(this.tbl_reporte_ejecutar_consultar!=null)
            {
                GC.SuppressFinalize(this.tbl_reporte_ejecutar_consultar);
            }
            
            if (this.obj_xfunciones != null)
            {
                this.obj_xfunciones.first_path = string.Empty;
                this.obj_xfunciones.ini_path = string.Empty;
                this.obj_xfunciones.IP_servidor1 = string.Empty;
                this.obj_xfunciones.IP_servidor2 = string.Empty;
                this.obj_xfunciones.second_path = string.Empty;
                this.obj_xfunciones = null;
            }
            if (this.conexion != null)
            {
                this.conexion.db_cadena = string.Empty;
                this.conexion = null;
            }
            if (this.fun != null)
            {
                GC.SuppressFinalize(this.fun);
            }
            if (this.SQL != null)
            {
                this.SQL = string.Empty;
                this.SQL = null;
            }
            if (this.reporte_name != null)
            {
                this.reporte_name = string.Empty;
                this.reporte_name = null;
            }
            if (this.file_name != null)
            {
                this.file_name = string.Empty;
                this.file_name = null;
            }
            if (this.Carpeta != null)
            {
                this.Carpeta = string.Empty;
                this.Carpeta = null;
            }
            if (this.servidor != null)
            {
                this.servidor = string.Empty;
                this.servidor = null;
            }
            if (this.Fecha_1 != null)
            {
                this.Fecha_1 = string.Empty;
                this.Fecha_1 = null;
            }
            if (this.Fecha_2 != null)
            {
                this.Fecha_2 = string.Empty;
                this.Fecha_2 = null;
            }
            if (this.dest_mail != null)
            {
                this.dest_mail = string.Empty;
                this.dest_mail = null;
            }
            if (this.param_string != null)
            {
                this.param_string = string.Empty;
                this.param_string = null;
            }
            if (this.tab_archivos != null)
            {
                this.tab_archivos = null;
            }
            GC.SuppressFinalize(this.bandera);
            GC.SuppressFinalize(this.i);
            GC.SuppressFinalize(this.j);
            GC.SuppressFinalize(this.id_Reporte);
            GC.SuppressFinalize(this.days_deleted);

            GC.Collect();
        }
    }
}