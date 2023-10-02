using ReportServer2022.code.include;
using System;

namespace ReportServer2022.code.querys
{
    public class class_querys
    {
        class_conexion conexion = null;
        class_xfunciones obj_xfunciones = null;
        private int i, j;

        public bool bandera;
        public int days_deleted;
        public int id_Reporte;
        public string SQL;
        public string reporte_name;
        public string file_name;
        public string Carpeta;
        public string servidor;
        public string Fecha_1;
        public string Fecha_2;
        public string dest_mail;
        public string param_string;
        public string[,] tab_archivos = new string[1, 5];

        public class_llena_tabla tbl_guias_disponibles_rpt = null;
        public class_llena_tabla tbl_trading_genera_GSK_rpt = null;
        public class_llena_tabla tbl_anexo24_tetrapack_rpt = null;
        public class_llena_tabla tbl_nuis_nc_rpt = null;

        public class_llena_tabla tbl_consulta_pedtos_Marelli = null;
        public class_llena_tabla tbl_consulta_AA_Marelli = null;
        public class_llena_tabla tbl_insert_archivo_gen_Marelli = null;
        public class_llena_tabla tbl_consulta_xml_cg_Marelli = null;
        public class_llena_tabla tbl_consulta_sql_factura_Marelli = null;
        public class_llena_tabla tbl_consulta_sql_base_iva_Marelli = null;
        public class_llena_tabla tbl_consulta_sql_conceptos_Marelli = null;


        //**********************************************QUERYS DE GENERACION DE REPORTES**********************************************
        public void sub_trading_genera_GSK_consultar()
        {
            try
            {
                // Console.WriteLine("Consultando BD´s...");
                conexion = new class_conexion();
                tbl_trading_genera_GSK_rpt = new class_llena_tabla();

                conexion.sub_set_conexion();
                tbl_trading_genera_GSK_rpt.cadena_conexion = conexion.db_cadena;

                tbl_trading_genera_GSK_rpt.sqltext = "SELECT	 NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA) SHIPMENT_NO \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		, '' CARRIER \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		, '' PLANNED_SHIPDATE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,TO_CHAR(WCD.DATE_CREATED, 'dd/mm/yy') SHIP_DATE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,'' PLANNED_DELIVERY_DATE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,INITCAP(DIS.DISNOM) ORIGIN \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,INITCAP(DIS.DISADRESSE1 || ' ' || ' ' || DIS.DISNUMEXT || '  ' || DIS.DISNUMINT || '  ' || DIS.DISADRESSE2 || DECODE(DIS.DISCODEPOSTAL,NULL,NULL, ' C.P. ' || DIS.DISCODEPOSTAL)) ORIGIN_ADDRESS \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,INITCAP(CIU_ORI.VILNOM || ' ('|| EST_ORI.ESTNOMBRE || ')') ORIGIN_CITY \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,INITCAP(CCL.CCL_NOMBRE || ' ' || NVL(DIE.DIE_A_ATENCION_DE, DIE.DIENOMBRE)) DESTINATION \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,INITCAP( DIE.DIEADRESSE1|| ' ' || ' ' || DIE.DIENUMEXT || '  ' || DIE.DIENUMINT || '  ' || DIE.DIEADRESSE2 || DECODE(DIE.DIECODEPOSTAL,NULL,NULL, ' C.P. ' || DIE.DIECODEPOSTAL)) DESTINATION_ADDRESS \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,INITCAP(CIU_DEST.VILNOM || ' ('|| EST_DEST.ESTNOMBRE || ')') DESTINATION_CITY \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,'Road' MODE_\n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,WCD.WCD_FIRMA SHIPMENT_LINE# \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		,TO_CHAR(WCD.DATE_CREATED, 'dd/mm/yy') CREATION_DATE\n";
                tbl_trading_genera_GSK_rpt.sqltext += "FROM	 WCROSS_DOCK WCD \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	EDISTRIBUTEUR DIS \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	DIS.DISCLEF			=	WCD.WCD_DISCLEF \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	ECIUDADES CIU_ORI \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	CIU_ORI.VILCLEF		=	DIS.DISVILLE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	EESTADOS EST_ORI \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	EST_ORI.ESTESTADO	=	CIU_ORI.VIL_ESTESTADO \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	LEFT JOIN	ETRANS_DETALLE_CROSS_DOCK TDCD \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON		WCD.WCD_TDCDCLAVE	=	TDCD.TDCDCLAVE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "			AND	TDCD.TDCDSTATUS		=	'1' \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	EDIRECCIONES_ENTREGA DIE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	DIE.DIECLAVE	=	NVL(NVL(TDCD.TDCD_DIECLAVE_ENT, TDCD.TDCD_DIECLAVE), WCD.WCD_DIECLAVE_ENTREGA) \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	ECIUDADES CIU_DEST \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	CIU_DEST.VILCLEF	=	DIE.DIEVILLE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	EESTADOS EST_DEST \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	EST_DEST.ESTESTADO	=	CIU_DEST.VIL_ESTESTADO \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	INNER JOIN	ECLIENT_CLIENTE CCL \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	CCL.CCLCLAVE	=	NVL(TDCD.TDCD_CCLCLAVE, WCD.WCD_CCLCLAVE) \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	LEFT JOIN	ETRANSFERENCIA_TRADING TRA \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON		WCD.WCD_TRACLAVE	=	TRA.TRACLAVE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "			AND	TRA.TRASTATUS		=	'1' \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	LEFT JOIN	ETRANS_ENTRADA TAE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "		ON	WCD.WCD_TRACLAVE	=	TAE.TAE_TRACLAVE \n";
                tbl_trading_genera_GSK_rpt.sqltext += "WHERE	TRUNC(WCD.DATE_CREATED) BETWEEN TRUNC(sysdate -1) AND TRUNC(sysdate -1) \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	AND	WCD.WCD_CLICLEF	IN	(20501,20502) \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	AND	NOT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA)	LIKE	'%PRUEBA%' \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	AND	NOT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA)	LIKE	'%SENSORES%' \n";
                tbl_trading_genera_GSK_rpt.sqltext += "	AND	NOT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA)	LIKE	'%TARIMAS%' \n";
                tbl_trading_genera_GSK_rpt.sub_llenar_tabla();
            }
            catch { }
            finally
            {
                if (tbl_trading_genera_GSK_rpt != null)
                {
                    tbl_trading_genera_GSK_rpt.Dispose();
                    GC.SuppressFinalize(tbl_trading_genera_GSK_rpt);
                }
            }
        }
        public void sub_guias_disponibles_consultar()
        {
            try
            {
                //Console.WriteLine("Consultando BD´s...");
                conexion = new class_conexion();
                tbl_guias_disponibles_rpt = new class_llena_tabla();

                conexion.sub_set_conexion();
                tbl_guias_disponibles_rpt.cadena_conexion = conexion.db_cadena;

                tbl_guias_disponibles_rpt.sqltext = "SELECT \n";
                tbl_guias_disponibles_rpt.sqltext += "    T.num_cliente num_cliente \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,cli.clinom as nom_cliente \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,T.TIPO TIPO \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,SUM(T.GUIAS_DISPONIBLES) + SUM(T.GUIAS_OCUPADAS) + SUM(T.GUIAS_CANCELADAS) TOTALES \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,SUM(T.GUIAS_DISPONIBLES) DISPONIBLES \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,SUM(T.GUIAS_OCUPADAS) OCUPADAS \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,SUM(T.GUIAS_CANCELADAS) CANCELADAS \n";
                tbl_guias_disponibles_rpt.sqltext += "FROM ( \n";
                tbl_guias_disponibles_rpt.sqltext += "SELECT DISTINCT \n";
                tbl_guias_disponibles_rpt.sqltext += "    wl.web_cliente num_cliente \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,WL.LOTE \n";
                tbl_guias_disponibles_rpt.sqltext += "    ,WL.TIPO \n";
                tbl_guias_disponibles_rpt.sqltext += "	, CASE WHEN WL.TIPO = 'LTL' THEN \n";
                tbl_guias_disponibles_rpt.sqltext += "			(SELECT COUNT(1) FROM web_tracking_stage wts \n";
                tbl_guias_disponibles_rpt.sqltext += "				INNER JOIN WEB_LTL WEL ON WEL.WELclave = wts.nui \n";
                tbl_guias_disponibles_rpt.sqltext += "				LEFT JOIN web_lots wl_DISP ON wl_DISP.lote = wts.numero_lote \n";
                tbl_guias_disponibles_rpt.sqltext += "			 WHERE wts.numero_lote IN (wl.lote) \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND ( \n";
                tbl_guias_disponibles_rpt.sqltext += "						WEL.WELFACTURA = 'RESERVADO' AND WEL.WELSTATUS IN (1,3) AND TRUNC(WEL.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_guias_disponibles_rpt.sqltext += "						OR WEL.WELFACTURA = 'RESERVADA_STNDBY' AND WEL.WELSTATUS IN (1,3) AND TRUNC(WEL.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') ) \n";
                tbl_guias_disponibles_rpt.sqltext += "					) \n";
                tbl_guias_disponibles_rpt.sqltext += "		ELSE \n";
                tbl_guias_disponibles_rpt.sqltext += "			(SELECT COUNT(1) FROM web_tracking_stage wts \n";
                tbl_guias_disponibles_rpt.sqltext += "				INNER JOIN WCROSS_DOCK WCD ON WCD.WCDclave = wts.nui \n";
                tbl_guias_disponibles_rpt.sqltext += "				LEFT JOIN web_lots wl_DISP ON wl_DISP.lote = wts.numero_lote \n";
                tbl_guias_disponibles_rpt.sqltext += "			 WHERE wts.numero_lote IN (wl.lote) \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND ( \n";
                tbl_guias_disponibles_rpt.sqltext += "						WCD.WCDFACTURA = 'RESERVADO' AND WCD.WCDSTATUS IN (1,3) AND TRUNC(WCD.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_guias_disponibles_rpt.sqltext += "						OR WCD.WCDFACTURA = 'RESERVADA_STNDBY' AND WCD.WCDSTATUS IN (1,3) AND TRUNC(WCD.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') ) \n";
                tbl_guias_disponibles_rpt.sqltext += "					) \n";
                tbl_guias_disponibles_rpt.sqltext += "	 END AS GUIAS_DISPONIBLES \n";
                tbl_guias_disponibles_rpt.sqltext += "	,CASE WHEN WL.TIPO = 'LTL' THEN \n";
                tbl_guias_disponibles_rpt.sqltext += "			(SELECT COUNT(1) FROM web_tracking_stage wts \n";
                tbl_guias_disponibles_rpt.sqltext += "				INNER JOIN WEB_LTL WEL ON WEL.WELclave = wts.nui \n";
                tbl_guias_disponibles_rpt.sqltext += "				LEFT JOIN web_lots wl_OCU ON wl_OCU.lote = wts.numero_lote \n";
                tbl_guias_disponibles_rpt.sqltext += "			 WHERE wts.numero_lote = wl.lote \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND WEL.WELSTATUS NOT IN (0, 3) \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND NOT (WEL.WELSTATUS = 1 AND WEL.WELFACTURA = 'RESERVADO') \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND TRUNC(WEL.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY ') ) \n";
                tbl_guias_disponibles_rpt.sqltext += "		ELSE \n";
                tbl_guias_disponibles_rpt.sqltext += "			(SELECT COUNT(1) FROM web_tracking_stage wts \n";
                tbl_guias_disponibles_rpt.sqltext += "				INNER JOIN wcross_dock wcd ON wcd.wcdclave = wts.nui \n";
                tbl_guias_disponibles_rpt.sqltext += "				LEFT JOIN web_lots wl_OCU ON wl_OCU.lote = wts.numero_lote \n";
                tbl_guias_disponibles_rpt.sqltext += "			 WHERE wts.numero_lote = wl.lote \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND WCD.WCDSTATUS NOT IN (0, 3) \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND NOT (WCD.WCDSTATUS = 1 AND WCD.WCDFACTURA = 'RESERVADO') \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND TRUNC(WCD.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') ) \n";
                tbl_guias_disponibles_rpt.sqltext += "	 END AS GUIAS_OCUPADAS \n";
                tbl_guias_disponibles_rpt.sqltext += "	,CASE WHEN WL.TIPO = 'LTL' THEN \n";
                tbl_guias_disponibles_rpt.sqltext += "			(SELECT COUNT(1) FROM web_tracking_stage wts \n";
                tbl_guias_disponibles_rpt.sqltext += "				INNER JOIN WEB_LTL WEL ON WEL.WELclave = wts.nui \n";
                tbl_guias_disponibles_rpt.sqltext += "				LEFT JOIN web_lots wl_CAN ON wl_CAN.lote = wts.numero_lote \n";
                tbl_guias_disponibles_rpt.sqltext += "			 WHERE WEL.WELSTATUS = 0 AND wts.numero_lote = wl.lote \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND WEL.WELSTATUS = 0 AND TRUNC(WEL.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') ) \n";
                tbl_guias_disponibles_rpt.sqltext += "		ELSE \n";
                tbl_guias_disponibles_rpt.sqltext += "			(SELECT COUNT(1) FROM web_tracking_stage wts \n";
                tbl_guias_disponibles_rpt.sqltext += "				INNER JOIN wcross_dock wcd ON wcd.wcdclave = wts.nui \n";
                tbl_guias_disponibles_rpt.sqltext += "				LEFT JOIN web_lots wl_CAN ON wl_CAN.lote = wts.numero_lote \n";
                tbl_guias_disponibles_rpt.sqltext += "			 WHERE WCD.WCDSTATUS=0 AND wts.numero_lote = wl.lote \n";
                tbl_guias_disponibles_rpt.sqltext += "				AND WCD.WCDSTATUS = 0 AND TRUNC(WCD.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') ) \n";
                tbl_guias_disponibles_rpt.sqltext += "	 END AS GUIAS_CANCELADAS \n";
                tbl_guias_disponibles_rpt.sqltext += "FROM	web_lots wl \n";
                tbl_guias_disponibles_rpt.sqltext += "	LEFT JOIN	WEB_TRACKING_STAGE wts ON WL.LOTE = WTS.NUMERO_LOTE \n";
                tbl_guias_disponibles_rpt.sqltext += "WHERE	1=1 \n";
                tbl_guias_disponibles_rpt.sqltext += "    AND wl.web_cliente <> 20123 \n";
                tbl_guias_disponibles_rpt.sqltext += ") T \n";
                tbl_guias_disponibles_rpt.sqltext += "    INNER JOIN eclient cli ON T.num_cliente = cli.cliclef \n";
                tbl_guias_disponibles_rpt.sqltext += "GROUP BY T.num_cliente,cli.clinom,T.TIPO \n";

                ////tbl_guias_disponibles_rpt.sqltext = "select D.* from rep_chron C, REP_DETALLE_REPORTE D WHERE 1=1 AND  d.id_cron(+) = c.id_rapport order by C.LAST_EXECUTION"; //Pruebas


                tbl_guias_disponibles_rpt.sub_llenar_tabla();
            }
            catch { }
            finally
            {
                if (tbl_guias_disponibles_rpt != null)
                {
                    tbl_guias_disponibles_rpt.Dispose();
                    GC.SuppressFinalize(tbl_guias_disponibles_rpt);
                }
            }
        }

        public void sub_Anexo24_tetrapack_14005(int daysInfo)
        {
            try
            {
                //Console.WriteLine("Consultando BD´s...");
                conexion = new class_conexion();
                tbl_anexo24_tetrapack_rpt = new class_llena_tabla();

                conexion.sub_set_conexion();
                tbl_anexo24_tetrapack_rpt.cadena_conexion = conexion.db_cadena;

                tbl_anexo24_tetrapack_rpt.sqltext = "SELECT PEDIMENTO, FECHA_PAGO, FECHA_PRESENTACION, TIPO_CAMBIO, REGIMEN, CLAVE_PEDIMENTO, PATENTE, ADUANA, SECCION, DESCRIPCION, DTA, FP_DTA, PREV, FP_PREV, AF, IC, PC, IM, ST, SU, FACTURA, FECHA, TIPO_CAMBIO_2, CLAVE_CLIENTE, INCOTERM, MONEDA_FACT, FACTOR_MON_FACT, NO_MATERIAL, CANTIDAD_UM_COMERCIAL, UM_COMERCIAL, CANTIDAD_UM_TARIFA, UM_TARIFA, VALOR_MON_FACT, VALOR_USD, FRACCION, PAIS_DEST, TL, PT, PARTIDA_PEDIMENTO, VALOR_ADUANA, VALOR_COMERCIAL, PEDIMENTO_ANEXO_R1, CLAVE_REGIMEN, TIPO_MERCANCIA, FAC_CONV_UM_AMERICANA, COVE, PLANTA, REFERENCIA, CNT, FORMA_PAGO_CNT, VALOR_MP, VALOR_AGREGADO, GUIA_O_CARTA_PORTE, VACIO_1, VACIO_2, VACIO_3, VACIO_4, VACIO_5, VACIO_6, VACIO_7, NICO \n";
                tbl_anexo24_tetrapack_rpt.sqltext += "FROM LOGIS.VW_A24_TETRAPAK T \n";
                tbl_anexo24_tetrapack_rpt.sqltext += "WHERE 1=1 \n";

                // *** REPROCESAR ***
                //tbl_anexo24_tetrapack_rpt.sqltext += "and to_date(fecha_pago ,'yyyy-mm-dd') between to_date('2022-09-12' ,'yyyy-mm-dd') and to_date('2022-09-18','yyyy-mm-dd')";

                tbl_anexo24_tetrapack_rpt.sqltext += "AND TO_DATE(FECHA_PAGO ,'yyyy-mm-dd')  BETWEEN (SYSDATE + (" + daysInfo + ") ) AND SYSDATE - 1 \n";
                tbl_anexo24_tetrapack_rpt.sub_llenar_tabla();
            }
            catch { }
            finally
            {
                if (tbl_anexo24_tetrapack_rpt != null)
                {
                    tbl_anexo24_tetrapack_rpt.Dispose();
                    GC.SuppressFinalize(tbl_anexo24_tetrapack_rpt);
                }
            }
        }

        //nuis pendientes nc
        public void sub_nuis_pendientes_nc()
        {
            try
            {
                //Console.WriteLine("Consultando BD´s...");
                conexion = new class_conexion();
                tbl_nuis_nc_rpt = new class_llena_tabla();

                conexion.sub_set_conexion();
                tbl_nuis_nc_rpt.cadena_conexion = conexion.db_cadena;

                tbl_nuis_nc_rpt.sqltext = "SELECT DISTINCT \n";
                tbl_nuis_nc_rpt.sqltext += " LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " ,WEB_CLIENTE \n";
                tbl_nuis_nc_rpt.sqltext += " ,CLINOM NOMBRE_CLIENTE \n";
                tbl_nuis_nc_rpt.sqltext += " ,FECHA FECHA_CREACION \n";
                tbl_nuis_nc_rpt.sqltext += " ,TOTAL_NUIS \n";
                tbl_nuis_nc_rpt.sqltext += " ,GUIAS_DISPONIBLES \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " ,GUIAS_OCUPADAS_SIN_ENTRADA GUIAS_OCUPADAS_SIN_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " ,GUIAS_OCUPADAS_CON_ENTRADA GUIAS_OCUPADAS_CON_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " ,GUIAS_CANCELADAS \n";
                tbl_nuis_nc_rpt.sqltext += " ,IMPORTE_POR_NUI \n";
                tbl_nuis_nc_rpt.sqltext += " ,NVL(TO_CHAR(MAX(FOLIO_FACTURA_INICIAL)),(CASE WHEN TOTAL_NUIS = GUIAS_CANCELADAS THEN 'N/A' ELSE '' END)) FOLIO_FACTURA_INICIAL \n";
                tbl_nuis_nc_rpt.sqltext += " ,NVL(TO_CHAR(MAX(FACTURA_INICIAL)),(CASE WHEN TOTAL_NUIS = GUIAS_CANCELADAS THEN 'N/A' ELSE '' END)) FACTURA_INICIAL \n";
                tbl_nuis_nc_rpt.sqltext += " ,NVL(TO_CHAR(IMPORTE_FACTURA),(CASE WHEN TOTAL_NUIS = GUIAS_CANCELADAS THEN 'N/A' ELSE '' END)) IMPORTE_FACTURA \n";
                tbl_nuis_nc_rpt.sqltext += " FROM ( \n";
                tbl_nuis_nc_rpt.sqltext += " SELECT \n";
                tbl_nuis_nc_rpt.sqltext += " DISTINCT \n";
                tbl_nuis_nc_rpt.sqltext += " WL.LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " ,TO_CHAR(WL.FECHA_RESERVACION,'dd/mm/yyyy') FECHA \n";
                tbl_nuis_nc_rpt.sqltext += " ,WL.TIPO \n";
                tbl_nuis_nc_rpt.sqltext += " ,FACTURA_INICIAL.FOLFOLIO FOLIO_FACTURA_INICIAL \n";
                tbl_nuis_nc_rpt.sqltext += " ,BASE.FACTURA_CUMPLE FACTURA_INICIAL \n";
                tbl_nuis_nc_rpt.sqltext += " ,BASE.SUBTOTAL IMPORTE_FACTURA \n";
                tbl_nuis_nc_rpt.sqltext += " ,WTS.PRECIO IMPORTE_POR_NUI \n";
                tbl_nuis_nc_rpt.sqltext += " ,WL.CANT_NUIS TOTAL_NUIS \n";
                tbl_nuis_nc_rpt.sqltext += " ,CASE WHEN WL.TIPO = 'LTL' THEN \n";
                tbl_nuis_nc_rpt.sqltext += " ( SELECT COUNT(1) \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_DIS \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WEB_LTL WEL_DIS ON WEL_DIS.WELclave = wts_DIS.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_DIS ON wl_DIS.lote = wts_DIS.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_DIS.numero_lote IN (wl_DIS.lote) \n";
                tbl_nuis_nc_rpt.sqltext += " AND ( \n";
                tbl_nuis_nc_rpt.sqltext += " (WEL_DIS.WELFACTURA = 'RESERVADO' AND WEL_DIS.WELSTATUS IN (1,3) AND TRUNC(WEL_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " OR (WEL_DIS.WELFACTURA = 'RESERVADA_STNDBY' AND WEL_DIS.WELSTATUS IN (1,3) AND TRUNC(WEL_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " AND wl_DIS.lote = wl.lote \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " ELSE \n";
                tbl_nuis_nc_rpt.sqltext += " ( SELECT COUNT(1) \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_DIS \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WCROSS_DOCK WCD_DIS ON WCD_DIS.WCDclave = wts_DIS.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_DIS ON wl_DIS.lote = wts_DIS.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_DIS.numero_lote IN (wl_DIS.lote) \n";
                tbl_nuis_nc_rpt.sqltext += " AND ( \n";
                tbl_nuis_nc_rpt.sqltext += " (WCD_DIS.WCDFACTURA = 'RESERVADO' AND WCD_DIS.WCDSTATUS IN (1,3) AND TRUNC(WCD_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " OR (WCD_DIS.WCDFACTURA = 'RESERVADA_STNDBY' AND WCD_DIS.WCDSTATUS IN (1,3) AND TRUNC(WCD_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " AND wl_DIS.lote = wl.lote \n";
                tbl_nuis_nc_rpt.sqltext += " ) END AS GUIAS_DISPONIBLES \n";
                tbl_nuis_nc_rpt.sqltext += "  \n";
                tbl_nuis_nc_rpt.sqltext += " ,CASE WHEN WL.TIPO = 'LTL' THEN \n";
                tbl_nuis_nc_rpt.sqltext += " ( SELECT COUNT(1) \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_OCU \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WEB_LTL WEL_OCU ON WEL_OCU.WELclave = wts_OCU.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_OCU ON wl_OCU.lote = wts_OCU.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += "  \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_OCU.numero_lote = wl_OCU.lote \n";
                tbl_nuis_nc_rpt.sqltext += " AND WEL_OCU.WELSTATUS NOT IN (0, 3) \n";
                tbl_nuis_nc_rpt.sqltext += " AND NOT (WEL_OCU.WELSTATUS = 1 AND WEL_OCU.WELFACTURA = 'RESERVADO') \n";
                tbl_nuis_nc_rpt.sqltext += " AND TRUNC(WEL_OCU.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_nuis_nc_rpt.sqltext += " AND wl_OCU.lote = wl.lote \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " AND (SELECT (CASE WHEN (WEL_ENTRADA.WELSTATUS NOT IN (0,3) AND TRA_ENTRADA.TRA_MEZTCLAVE_DEST IS NOT NULL) THEN '1' ELSE '0' END) TIENE_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " FROM WEB_LTL WEL_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN ETRANSFERENCIA_TRADING TRA_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " ON WEL_ENTRADA.WEL_TRACLAVE = TRA_ENTRADA.TRACLAVE \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE WEL_ENTRADA.WELCLAVE = wts_OCU.NUI) = '1' \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " ELSE \n";
                tbl_nuis_nc_rpt.sqltext += " ( SELECT COUNT(1) \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_OCU \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN wcross_dock wcd_OCU ON wcd_OCU.wcdclave = wts_OCU.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_OCU ON wl_OCU.lote = wts_OCU.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_OCU.numero_lote = wl_OCU.lote \n";
                tbl_nuis_nc_rpt.sqltext += " AND WCD_OCU.WCDSTATUS NOT IN (0, 3) \n";
                tbl_nuis_nc_rpt.sqltext += " AND NOT (WCD_OCU.WCDSTATUS = 1 AND WCD_OCU.WCDFACTURA = 'RESERVADO') \n";
                tbl_nuis_nc_rpt.sqltext += " AND TRUNC(WCD_OCU.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_nuis_nc_rpt.sqltext += " AND wl_OCU.lote = wl.lote \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " AND (SELECT (CASE WHEN (WCD_ENTRADA.WCDSTATUS NOT IN (0,3) AND TRA_ENTRADA.TRA_MEZTCLAVE_DEST IS NOT NULL) THEN '1' ELSE '0' END) TIENE_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " FROM WCROSS_DOCK WCD_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN ETRANSFERENCIA_TRADING TRA_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " ON WCD_ENTRADA.WCD_TRACLAVE = TRA_ENTRADA.TRACLAVE \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE WCD_ENTRADA.WCDCLAVE = wts_OCU.NUI) = '1' \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " ) END AS GUIAS_OCUPADAS_CON_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " ,CASE WHEN WL.TIPO = 'LTL' THEN \n";
                tbl_nuis_nc_rpt.sqltext += " ( SELECT COUNT(1) \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_OCU \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WEB_LTL WEL_OCU ON WEL_OCU.WELclave = wts_OCU.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_OCU ON wl_OCU.lote = wts_OCU.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_OCU.numero_lote = wl_OCU.lote \n";
                tbl_nuis_nc_rpt.sqltext += " AND WEL_OCU.WELSTATUS NOT IN (0, 3) \n";
                tbl_nuis_nc_rpt.sqltext += " AND NOT (WEL_OCU.WELSTATUS = 1 AND WEL_OCU.WELFACTURA = 'RESERVADO') \n";
                tbl_nuis_nc_rpt.sqltext += " AND TRUNC(WEL_OCU.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_nuis_nc_rpt.sqltext += " AND wl_OCU.lote = wl.lote \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " AND (SELECT (CASE WHEN (WEL_ENTRADA.WELSTATUS NOT IN (0,3) AND TRA_ENTRADA.TRA_MEZTCLAVE_DEST IS NOT NULL) THEN '1' ELSE '0' END) TIENE_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " FROM WEB_LTL WEL_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN ETRANSFERENCIA_TRADING TRA_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " ON WEL_ENTRADA.WEL_TRACLAVE = TRA_ENTRADA.TRACLAVE \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE WEL_ENTRADA.WELCLAVE = wts_OCU.NUI) = '0' \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " ELSE \n";
                tbl_nuis_nc_rpt.sqltext += " ( SELECT COUNT(1) \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_OCU \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN wcross_dock wcd_OCU ON wcd_OCU.wcdclave = wts_OCU.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_OCU ON wl_OCU.lote = wts_OCU.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_OCU.numero_lote = wl_OCU.lote \n";
                tbl_nuis_nc_rpt.sqltext += " AND WCD_OCU.WCDSTATUS NOT IN (0, 3) \n";
                tbl_nuis_nc_rpt.sqltext += " AND NOT (WCD_OCU.WCDSTATUS = 1 AND WCD_OCU.WCDFACTURA = 'RESERVADO') \n";
                tbl_nuis_nc_rpt.sqltext += " AND TRUNC(WCD_OCU.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_nuis_nc_rpt.sqltext += " AND wl_OCU.lote = wl.lote \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " AND (SELECT (CASE WHEN (WCD_ENTRADA.WCDSTATUS NOT IN (0,3) AND TRA_ENTRADA.TRA_MEZTCLAVE_DEST IS NOT NULL) THEN '1' ELSE '0' END) TIENE_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " FROM WCROSS_DOCK WCD_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN ETRANSFERENCIA_TRADING TRA_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " ON WCD_ENTRADA.WCD_TRACLAVE = TRA_ENTRADA.TRACLAVE \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE WCD_ENTRADA.WCDCLAVE = wts_OCU.NUI) = '0' \n";
                tbl_nuis_nc_rpt.sqltext += " --------------------------------------------------------------------------------------------------------------------------------------- \n";
                tbl_nuis_nc_rpt.sqltext += " ) END AS GUIAS_OCUPADAS_SIN_ENTRADA \n";
                tbl_nuis_nc_rpt.sqltext += " , GUIAS_CANCELADAS \n";
                tbl_nuis_nc_rpt.sqltext += " ,WL.FECHA_RESERVACION \n";
                tbl_nuis_nc_rpt.sqltext += " ,WL.WEB_CLIENTE \n";
                tbl_nuis_nc_rpt.sqltext += " ,CLI.CLINOM \n";
                tbl_nuis_nc_rpt.sqltext += " FROM WEB_LOTS WL \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN ECLIENT CLI \n";
                tbl_nuis_nc_rpt.sqltext += " ON WL.WEB_CLIENTE = CLI.CLICLEF \n";
                tbl_nuis_nc_rpt.sqltext += " inner join (SELECT COUNT(1) GUIAS_CANCELADAS,wl_CAN.lote \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_CAN \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WEB_LTL WEL_CAN ON WEL_CAN.WELclave = wts_CAN.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_CAN ON wl_CAN.lote = wts_CAN.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE WEL_CAN.WELSTATUS = 0 \n";
                tbl_nuis_nc_rpt.sqltext += " AND wts_CAN.numero_lote = wl_CAN.lote \n";
                tbl_nuis_nc_rpt.sqltext += " AND TRUNC(WEL_CAN.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_nuis_nc_rpt.sqltext += " group by wl_CAN.lote \n";
                tbl_nuis_nc_rpt.sqltext += " UNION \n";
                tbl_nuis_nc_rpt.sqltext += " SELECT COUNT(1),wl_CAN.lote \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_CAN \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN wcross_dock wcd_CAN ON wcd_CAN.wcdclave = wts_CAN.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_CAN ON wl_CAN.lote = wts_CAN.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE WCD_CAN.WCDSTATUS = 0 \n";
                tbl_nuis_nc_rpt.sqltext += " AND wts_CAN.numero_lote = wl_CAN.lote \n";
                tbl_nuis_nc_rpt.sqltext += " AND TRUNC(WCD_CAN.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY') \n";
                tbl_nuis_nc_rpt.sqltext += " group by wl_CAN.lote) WL_CAN on WL.LOTE = WL_CAN.LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN WEB_TRACKING_STAGE WTS \n";
                tbl_nuis_nc_rpt.sqltext += " ON WL.LOTE = WTS.NUMERO_LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN (SELECT TO_NUMBER(fn.nui) AS nui \n";
                tbl_nuis_nc_rpt.sqltext += " ,f.fctclef AS fctclef \n";
                tbl_nuis_nc_rpt.sqltext += " ,f.fecha_timbrado \n";
                tbl_nuis_nc_rpt.sqltext += " ,TO_NUMBER(RTRIM(LTRIM(NVL(f.factura_cumple,0)))) FACTURA_CUMPLE \n";
                tbl_nuis_nc_rpt.sqltext += " ,f.Subtotal SUBTOTAL \n";
                tbl_nuis_nc_rpt.sqltext += " ,f.cliente \n";
                tbl_nuis_nc_rpt.sqltext += " FROM tb_facturas_ccfdi f \n";
                tbl_nuis_nc_rpt.sqltext += " JOIN tb_facturas_nuis_ccfdi fn \n";
                tbl_nuis_nc_rpt.sqltext += " ON f.id_factura_cumple = fn.id_factura_cumple \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE 1=1 \n";
                tbl_nuis_nc_rpt.sqltext += " AND f.fctclef IS NOT NULL \n";
                tbl_nuis_nc_rpt.sqltext += " AND f.cliente NOT IN (9954,9955,9956,9929,9910) \n";
                tbl_nuis_nc_rpt.sqltext += " AND f.TIPO_CFDI = 'Ingreso' \n";
                tbl_nuis_nc_rpt.sqltext += " AND f.STATUS_CFDI = 'T' \n";
                tbl_nuis_nc_rpt.sqltext += " AND f.FECHA_TIMBRADO IS NOT NULL \n";
                tbl_nuis_nc_rpt.sqltext += " AND f.FECHA_CANCELACION IS NULL \n";
                tbl_nuis_nc_rpt.sqltext += " GROUP BY fn.nui,f.fctclef,f.fecha_timbrado,f.factura_cumple,f.Subtotal, f.cliente) BASE \n";
                tbl_nuis_nc_rpt.sqltext += " ON WTS.NUI = BASE.NUI \n";
                tbl_nuis_nc_rpt.sqltext += " AND WL.WEB_CLIENTE = BASE.CLIENTE \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN ( SELECT fol.folfolio \n";
                tbl_nuis_nc_rpt.sqltext += " ,fct.fctnumero \n";
                tbl_nuis_nc_rpt.sqltext += " ,fct.fcttotingreso \n";
                tbl_nuis_nc_rpt.sqltext += " ,fct.fctuuid \n";
                tbl_nuis_nc_rpt.sqltext += " ,fct.fctclef \n";
                tbl_nuis_nc_rpt.sqltext += " FROM efacturas fct \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN efolios fol \n";
                tbl_nuis_nc_rpt.sqltext += " ON fct.fctfolio = fol.folclave) FACTURA_INICIAL \n";
                tbl_nuis_nc_rpt.sqltext += " ON BASE.fctclef = FACTURA_INICIAL.fctclef \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE 1 = 1 \n";
                tbl_nuis_nc_rpt.sqltext += " --AND WL.WEB_CLIENTE = '22529' \n";
                tbl_nuis_nc_rpt.sqltext += " AND WL.LOTE NOT IN (SELECT wl_DIS.LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_DIS \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WEB_LTL WEL_DIS ON WEL_DIS.WELclave = wts_DIS.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_DIS ON wl_DIS.lote = wts_DIS.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_DIS.numero_lote IN (wl_DIS.lote) \n";
                tbl_nuis_nc_rpt.sqltext += " AND ( \n";
                tbl_nuis_nc_rpt.sqltext += " (WEL_DIS.WELFACTURA = 'RESERVADO' AND WEL_DIS.WELSTATUS IN (1,3) AND TRUNC(WEL_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " OR (WEL_DIS.WELFACTURA = 'RESERVADA_STNDBY' AND WEL_DIS.WELSTATUS IN (1,3) AND TRUNC(WEL_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " UNION \n";
                tbl_nuis_nc_rpt.sqltext += " SELECT wl_DIS.LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " FROM web_tracking_stage wts_DIS \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN WCROSS_DOCK WCD_DIS ON WCD_DIS.WCDclave = wts_DIS.nui \n";
                tbl_nuis_nc_rpt.sqltext += " LEFT JOIN web_lots wl_DIS ON wl_DIS.lote = wts_DIS.numero_lote \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE wts_DIS.numero_lote IN (wl_DIS.lote) \n";
                tbl_nuis_nc_rpt.sqltext += " AND ( \n";
                tbl_nuis_nc_rpt.sqltext += " (WCD_DIS.WCDFACTURA = 'RESERVADO' AND WCD_DIS.WCDSTATUS IN (1,3) AND TRUNC(WCD_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " OR (WCD_DIS.WCDFACTURA = 'RESERVADA_STNDBY' AND WCD_DIS.WCDSTATUS IN (1,3) AND TRUNC(WCD_DIS.DATE_CREATED) >= TO_DATE('01/01/2021', 'DD/MM/YYYY')) \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " AND WL.LOTE NOT IN (SELECT DISTINCT WL_NC.LOTE \n";
                tbl_nuis_nc_rpt.sqltext += " FROM WEB_LOTS WL_NC \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN EFOLIOS FOL_NC ON FOL_NC.FOLFOLIO = WL_NC.NCREDITO_FOLIO_PROFORMA \n";
                tbl_nuis_nc_rpt.sqltext += " INNER JOIN EFACTURAS FCT_NC ON FCT_NC.FCTFOLIO = FOL_NC.FOLCLAVE \n";
                tbl_nuis_nc_rpt.sqltext += " WHERE 1=1 \n";
                tbl_nuis_nc_rpt.sqltext += " AND FCT_NC.FCTUUID IS NOT NULL \n";
                tbl_nuis_nc_rpt.sqltext += " AND FCT_NC.FCT_YFACLEF = '3' \n";
                tbl_nuis_nc_rpt.sqltext += " ) \n";
                tbl_nuis_nc_rpt.sqltext += " ) T \n";
                tbl_nuis_nc_rpt.sqltext += " GROUP BY LOTE,FECHA,TIPO,IMPORTE_FACTURA,IMPORTE_POR_NUI,TOTAL_NUIS \n";
                tbl_nuis_nc_rpt.sqltext += " ,GUIAS_OCUPADAS_CON_ENTRADA,GUIAS_OCUPADAS_SIN_ENTRADA,GUIAS_CANCELADAS \n";
                tbl_nuis_nc_rpt.sqltext += " ,FECHA_RESERVACION,WEB_CLIENTE,CLINOM,GUIAS_DISPONIBLES \n";
                tbl_nuis_nc_rpt.sqltext += " ORDER BY T.LOTE DESC  \n";


                tbl_nuis_nc_rpt.sub_llenar_tabla();
            }
            catch { }
            finally
            {
                if (tbl_nuis_nc_rpt != null)
                {
                    tbl_nuis_nc_rpt.Dispose();
                    GC.SuppressFinalize(tbl_nuis_nc_rpt);
                }
            }
        }
        //nuis pendientes nc




        /************************************************************
         *                                                          *
         * Queryes para generar los expedientes adunales de Marelli *
         *                                                          *   
         ************************************************************/

        public void sub_sql_pedtos_marelli(string mi_cliente, string mi_fecha_ini, string mi_fecha_fin)
        {
            //< CHG-DESA-06092021-01(JEMV): Se crea nueva query para obtener la información de los pedimentos creados para el cliente en el rango de fechas establecido

            try
            {
                conexion = new class_conexion();
                tbl_consulta_pedtos_Marelli = new class_llena_tabla();

                conexion.sub_set_conexion();
                tbl_consulta_pedtos_Marelli.cadena_conexion = conexion.db_cadena;

                tbl_consulta_pedtos_Marelli.sqltext = "SELECT SGE.SGECLAVE SGECLAVE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SGE.SGEPEDNUMERO SGEPEDNUMERO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SGE.SGEDOUCLEF SGEDOUCLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , NVL(SGE.SGE_ADUANA_SECCION,'0') SGE_ADUANA_SECCION \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SGE.SGEANIO SGEANIO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SUBSTR(SGE.SGEPEDNUMERO, 1, 4) ||'-'|| SGE.SGEDOUCLEF || SGE.SGE_ADUANA_SECCION ||'-'|| SUBSTR(SGE.SGEPEDNUMERO, 6, 7) NOM_PEDTO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FOL.FOLFOLIO FOLFOLIO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FOL.FOL_SUCCLEF FOL_SUCCLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , TO_CHAR(PED.PEDDATE,'yyyymm') PED_MES \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SUBSTR(TO_CHAR(SGE.SGEANIO), -2) || SGE.SGEDOUCLEF || SGE.SGE_ADUANA_SECCION || REPLACE(SGE.SGEPEDNUMERO,'-') NOM_ARCHIVO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SUBSTR(TO_CHAR(SGE.SGEANIO), -2) || SGE.SGEDOUCLEF || SGE.SGE_ADUANA_SECCION || REPLACE(SGE.SGEPEDNUMERO,'-') PED_CARPETA \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FOL.FOLCLAVE FOLCLAVE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SGE.SGE_REDCLEF REDCLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FCT.FCTCLEF FCTCLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FCT.FCT_YFACLEF FCT_YFACLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FCT.FCTDIVISA FCTDIVISA \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , DECODE(FCT.FCT_YFACLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 , '1', DECODE(FCT.FCTDIVISA,'USD','0','1') \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 , '0') CON_EXPEDIENTE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FCT.FCTDATEFACTURE FCTDATEFACTURE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FCT.FCTNUMERO FCTNUMERO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FOL.FOLCLAVE FOLCLAVE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FCT.FCT_EMPCLAVE FCT_EMPCLAVE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , SGE.SGE_CLICLEF SGE_CLICLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , FOL.FOL_YCXCLEF FOL_YCXCLEF \n";
                //20200601 -- >
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , NVL((SELECT '1'   \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                FROM EDETAILFACTURE DTF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                    ,ECONCEPTOSHOJA CHO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "               WHERE DTF.DTFFACTURE = FCT.FCTCLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 AND FCT.FCT_YFACLEF = '2' \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 AND CHO.CHOCLAVE = DTF.DTF_CHOCLAVE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 AND CHO.CHO_EMPCLAVE = 29 \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 AND CHO.CHONUMERO = 150 \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "                 AND SGE.SGE_REDCLEF = 'R1'),'0') NOTA_RECTIF \n";
                //20200601 < --
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "    FROM ESAAI_M3_GENERAL SGE \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , EPEDIMENTO PED \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , EFOLIOS FOL \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "       , EFACTURAS FCT \n";

                //<reprocesar cliente:
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "   WHERE SGE.SGE_CLICLEF = " + mi_cliente + " \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "   WHERE SGE.SGE_CLICLEF = 22533 \n";  //**********
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "   WHERE 1=1 \n";  //**********
                //fin-Reprocesar>

                //<<< REPROCESO AVC
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGEDOUCLEF = '24' and SGE.SGEPEDNUMERO = '3420-2008576' and SGE.SGEANIO = '2022'  \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGEANIO = '2022' and SGE.SGE_CLICLEF = '21179' \n";
                //REPROCESO AVC >>>

                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND SGE.SGEFECHA_PAGO IS NOT NULL \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND SGE.SGEFIRMA_ELECTRONICA IS NOT NULL \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND PED.PEDNUMERO = SGE.SGEPEDNUMERO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND PED.PEDANIO = SGE.SGEANIO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND PED.PEDDOUANE = SGE.SGEDOUCLEF \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND PED.PEDDATE IS NOT NULL \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND FOL.FOLCLAVE = PED.PEDFOLIO \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND FCT.FCTFOLIO = FOL.FOLCLAVE \n";

                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "     AND FCT.FCTDATEFACTURE BETWEEN TO_DATE('" + mi_fecha_ini + "','mm/dd/yyyy') AND TO_DATE('" + mi_fecha_fin + "','mm/dd/yyyy') + 1 \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "     AND FCT.FCTDATEFACTURE BETWEEN SYSDATE - 180 AND SYSDATE \n"; //REPROCESO (JEMV)
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "     AND FCT.FCTDATEFACTURE BETWEEN SYSDATE - 30 AND SYSDATE \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "     AND FCT.FCTDATEFACTURE BETWEEN SYSDATE - 7 AND SYSDATE \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "     AND FCT.FCTDATEFACTURE BETWEEN SYSDATE - 15 AND SYSDATE \n"; //'**********

                //<Reprocesar fecha:
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND FCT.FCTDATEFACTURE BETWEEN SYSDATE - 16 AND SYSDATE \n"; //'**********
                //fin-Reprocesar>

                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + "     AND FCT.FCTUUID IS NOT NULL \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND FCT.FCT_EMPCLAVE = 29 \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND FCT.FCT_YFACLEF IN ('1','2') \n";
                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + "     AND FCT.FCTDIVISA IN ('MXN','USD') \n";

                //<reprocesar Pedimento
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " AND SGE.SGEDOUCLEF = 65 and SGE.SGEPEDNUMERO in ('3744-2011949') \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " AND SGE.SGEDOUCLEF IN (47) and SGE.SGEPEDNUMERO in ('3420-2008571','3420-2008570') \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGEPEDNUMERO in ('3744-2008981','3744-2006794','3744-2008561','3744-2005287','3744-2007700','3744-2007937','3744-2008073','3744-2008058','3744-2007894','3744-2007908','3744-2007928','3744-2008013','3744-2008096') \n"; //**********
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGEDOUCLEF in ('65') \n"; //**********


                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGEPEDNUMERO in ('3420-2003132') \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGEANIO = '2022' \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " AND SGE.SGEDOUCLEF IN (43) \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and UPPER(FCE.FCEARCHIVO) LIKE 'M%' \n";
                //tbl_pedtos_marelli.sqltext = tbl_pedtos_marelli.sqltext + " and SGE.SGE_CLICLEF = 22533 \n";  //'**********
                //fin reprocesar pedimento>

                tbl_consulta_pedtos_Marelli.sqltext = tbl_consulta_pedtos_Marelli.sqltext + " ORDER BY FCT.FCTDATEFACTURE DESC \n";
                // CHG-DESA-06092021-01(JEMV) >

                tbl_consulta_pedtos_Marelli.sub_llenar_tabla();


            }
            catch
            {
            }
            finally
            {
                if (tbl_consulta_pedtos_Marelli != null)
                {
                    tbl_consulta_pedtos_Marelli.Dispose();
                    GC.SuppressFinalize(tbl_consulta_pedtos_Marelli);
                }
            }
        }

        public void sub_PDF_de_la_cuenta_de_gastos_AA(string mi_fctclef)
        {
            conexion = new class_conexion();
            tbl_consulta_AA_Marelli = new class_llena_tabla();

            conexion.sub_set_conexion();
            tbl_consulta_AA_Marelli.cadena_conexion = conexion.db_cadena;

            try
            {
                //PDF de la cuenta de gastos AA --------------------------------------------------->
                tbl_consulta_AA_Marelli.sqltext = "SELECT FCT.FCTCLEF FCTCLEF \n";
                tbl_consulta_AA_Marelli.sqltext = tbl_consulta_AA_Marelli.sqltext + " , FCT.FCT_EMPCLAVE FCT_EMPCLAVE \n";
                tbl_consulta_AA_Marelli.sqltext = tbl_consulta_AA_Marelli.sqltext + "       , FCT.FCTCLIENT FCTCLIENT \n";
                tbl_consulta_AA_Marelli.sqltext = tbl_consulta_AA_Marelli.sqltext + "       , FCT.FCTDIVISA FCTDIVISA \n";
                tbl_consulta_AA_Marelli.sqltext = tbl_consulta_AA_Marelli.sqltext + "       , FCT.FCT_YFACLEF FCT_YFACLEF \n";
                tbl_consulta_AA_Marelli.sqltext = tbl_consulta_AA_Marelli.sqltext + "    FROM EFACTURAS FCT \n";
                tbl_consulta_AA_Marelli.sqltext = tbl_consulta_AA_Marelli.sqltext + "   WHERE FCT.FCTCLEF = " + mi_fctclef + " \n";


                tbl_consulta_AA_Marelli.sub_llenar_tabla();

            }
            catch
            {
            }

            finally
            {
                if (tbl_consulta_AA_Marelli != null)
                {
                    tbl_consulta_AA_Marelli.Dispose();
                    GC.SuppressFinalize(tbl_consulta_AA_Marelli);
                }
            }
        }


        public void sub_RegistraArchivoGenerado(string Carpeta, string Carpeta1, string NombreArchivo, int days_deleted)
        {
            conexion = new class_conexion();
            tbl_insert_archivo_gen_Marelli = new class_llena_tabla();

            conexion.sub_set_conexion();
            tbl_insert_archivo_gen_Marelli.cadena_conexion = conexion.db_cadena;

            try
            {
                tbl_insert_archivo_gen_Marelli.sqltext = "insert into rep_archivos(ID_REP, CARPETA, NOMBRE, DATE_CREATED, DEST_MAIL, PARAMS, DAYS_DELETED, SUBCARPETA, TIPO_REPORTE, HASH_MD5, FECHA_INICIO, FECHA_FIN) ";
                tbl_insert_archivo_gen_Marelli.sqltext = tbl_insert_archivo_gen_Marelli.sqltext + "values (317, '" + Carpeta + "', '" + NombreArchivo + "', sysdate, null,''," + days_deleted + ",null,null,null,null,null)";
                tbl_insert_archivo_gen_Marelli.sub_llenar_tabla();
            }
            catch { }
            finally
            {
                tbl_insert_archivo_gen_Marelli.Dispose();
                GC.SuppressFinalize(tbl_insert_archivo_gen_Marelli);
            }
        }

        public void sub_consulta_xml_CG(string my_FCTCLEF)
        {
            conexion = new class_conexion();
            tbl_consulta_xml_cg_Marelli = new class_llena_tabla();

            conexion.sub_set_conexion();
            tbl_consulta_xml_cg_Marelli.cadena_conexion = conexion.db_cadena;

            try
            {
                tbl_consulta_xml_cg_Marelli.sqltext = "SELECT COLUMN_VALUE FROM TABLE(SPLIT_FXM_CLOB(" + my_FCTCLEF + ")) ";
                tbl_consulta_xml_cg_Marelli.sub_llenar_tabla();
            }
            catch { }
            finally
            {
                tbl_consulta_xml_cg_Marelli.Dispose();
                GC.SuppressFinalize(tbl_consulta_xml_cg_Marelli);
            }
        }


        public void sub_SQL_FACTURA(string mi_fctclef)
        {
            conexion = new class_conexion();
            tbl_consulta_sql_factura_Marelli = new class_llena_tabla();

            conexion.sub_set_conexion();
            tbl_consulta_sql_factura_Marelli.cadena_conexion = conexion.db_cadena;

            try
            {
                tbl_consulta_sql_factura_Marelli.sqltext = " SELECT FCT.FCTNUMERO FCTNUMERO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTUUID_UP FCTUUID_UP ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",CLI.CLIRFC CLIRFC ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",CLIEMP.CLIRFC CLIRFC_EMP ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",CLIEMP.CLINOM CLINOM ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",TO_CHAR(FCT.FCTFECHA_TIMBRADO,'DD/MM/YYYY') FCTFECHA_TIMBRADO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",DECODE(FOL.FOL_YCXCLEF, 1, 'IMPORTACION', 2, 'EXPORTACION', 'OTROS') TIPO_OP ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",TO_CHAR(ROUND(SUM(NVL(DTF.DTFSOMME,0)), 2),'FM99999999990.00') SUBTOTAL ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",DECODE(NVL(FCTTOTANTICIPO,0), 0, 'N/A', TO_CHAR(ROUND(FCT.FCTTOTANTICIPO, 2),'FM99999999990.00')) FCTTOTANTICIPO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",TO_CHAR(ROUND(FCT.FCTIVA, 2),'FM99999999990.00') FCTIVA ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",TO_CHAR(ROUND(FCT.FCTRETENCIONFLETE, 2),'FM99999999990.00') FCTRETENCIONFLETE ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",TO_CHAR(ROUND(FCT.FCTTOTAL, 2),'FM99999999990.00') FCTTOTAL ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTDIVISA ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",TO_CHAR(FCT.FCTDATEFACTURE,'DD/MM/YYYY') FCTDATEFACTURE ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTCLIENT FCTCLIENT ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "FROM EFACTURAS FCT ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",ECLIENT CLI ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",ECLIENT CLIEMP ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",EFOLIOS FOL ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",EDETAILFACTURE DTF ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "WHERE FCT.FCTCLEF = " + mi_fctclef;
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "AND CLI.CLICLEF = FCT.FCTCLIENT ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "AND CLIEMP.CLICLEF = 9900 + FCT.FCT_EMPCLAVE ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "AND FOL.FOLCLAVE = FCT.FCTFOLIO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "AND DTF.DTFFACTURE = FCT.FCTCLEF ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + "GROUP BY FCT.FCTNUMERO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTUUID_UP ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",CLI.CLIRFC ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",CLIEMP.CLIRFC ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",CLIEMP.CLINOM ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTFECHA_TIMBRADO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FOL.FOL_YCXCLEF ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTTOTANTICIPO ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTIVA ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTRETENCIONFLETE ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTTOTAL ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTDIVISA ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTDATEFACTURE ";
                tbl_consulta_sql_factura_Marelli.sqltext = tbl_consulta_sql_factura_Marelli.sqltext + ",FCT.FCTCLIENT ";

                tbl_consulta_sql_factura_Marelli.sub_llenar_tabla();

            }
            catch
            {
            }
            finally
            {

                tbl_consulta_sql_factura_Marelli.Dispose();
                GC.SuppressFinalize(tbl_consulta_sql_factura_Marelli);
            }
        }




        public void sub_BASE_IVA(string mi_fctclef)
        {
            conexion = new class_conexion();
            tbl_consulta_sql_base_iva_Marelli = new class_llena_tabla();

            conexion.sub_set_conexion();
            tbl_consulta_sql_base_iva_Marelli.cadena_conexion = conexion.db_cadena;

            try
            {
                tbl_consulta_sql_base_iva_Marelli.sqltext = "SELECT SUM(DTFSOMME) BASE_IVA \n";

                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + " FROM (SELECT /*+INDEX(FCT PK_EFACTURAS) */ \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "DISTINCT DTF.DTFSOMME \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + " FROM EFACTURAS FCT, EDETAILFACTURE DTF, ECONCEPTOSHOJA CHO \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + " WHERE --FCTCLIENT = CLIENTE  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + " FCT_YFACLEF IN ('1', '2', '3') \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--AND FCT_EMPCLAVE = EMPCLAVE  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--AND FCTDIVISA = DIVCLEF  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--AND FCTDATEFACTURE BETWEEN  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--TO_DATE(FECHA_1, 'MM/DD/YYYY HH24:MI') -  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--DECODE(TO_CHAR(SYSDATE, 'D'), 2, 3, 1) AND  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--TO_DATE(FECHA_2, 'MM/DD/YYYY') + 1  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND FCTCLEF = " + mi_fctclef + "\n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND DTFFACTURE = FCTCLEF  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND NVL(DTFSOMME, 0) <> 0  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND CHOCLAVE = DTF_CHOCLAVE \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND CHOTIPOIE = 'E' \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND EXISTS (SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM ESORTIEARGENT \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE SOA_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND SOATYPE = 'E12' \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EDETAILDIARIO \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE DDI_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND DDITYPE IN ('D12', 'D13') \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EGCOMPROBAR \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE GCO_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND GCOTYPE = 'G12' \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EPROVDOCDET \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE PDD_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND PDDTYPE IN ('P12', 'P13') \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM ERELACION_DIARIO_FACTURA \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE RDF_DTFCLEF = DTFCLEF) \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT /*+INDEX(FCT PK_EFACTURAS) */ \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "DISTINCT DTF.DTFSOMME \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EFACTURAS FCT, EDETAILFACTURE DTF, ECONCEPTOSHOJA CHO \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE --FCTCLIENT = CLIENTE  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FCT_YFACLEF IN ('1', '2', '3') \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--AND FCT_EMPCLAVE = 29  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--AND FCTDIVISA = 'MXN'  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--TO_DATE(FECHA_1, 'MM/DD/YYYY HH24:MI') -  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--DECODE(TO_CHAR(SYSDATE, 'D'), 2, 3, 1) AND  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "--TO_DATE(FECHA_2, 'MM/DD/YYYY') + 1  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND FCTCLEF = " + mi_fctclef + "\n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND DTFFACTURE = FCTCLEF  \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND NVL(DTFSOMME, 0) <> 0 \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND CHOCLAVE = DTF_CHOCLAVE \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND CHOTIPOIE = 'I' \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND NOT EXISTS (SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM ESORTIEARGENT \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE SOA_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND SOATYPE = 'E12' \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EDETAILDIARIO \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE DDI_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND DDITYPE IN ('D12', 'D13') \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EGCOMPROBAR \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE GCO_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND GCOTYPE = 'G12' \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM EPROVDOCDET \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE PDD_DTFCLEF = DTFCLEF \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "AND PDDTYPE IN ('P12', 'P13') \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "UNION ALL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "SELECT NULL \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "FROM ERELACION_DIARIO_FACTURA \n";
                tbl_consulta_sql_base_iva_Marelli.sqltext = tbl_consulta_sql_base_iva_Marelli.sqltext + "WHERE RDF_DTFCLEF = DTFCLEF)) \n";

                tbl_consulta_sql_base_iva_Marelli.sub_llenar_tabla();

            }
            catch
            { }
            finally
            {
                tbl_consulta_sql_base_iva_Marelli.Dispose();
                GC.SuppressFinalize(tbl_consulta_sql_base_iva_Marelli);
            }
        }



        public void sub_SQL_CONCEPTOS(string mi_fctclef)
        {
            conexion = new class_conexion();
            tbl_consulta_sql_base_iva_Marelli = new class_llena_tabla();

            conexion.sub_set_conexion();
            tbl_consulta_sql_base_iva_Marelli.cadena_conexion = conexion.db_cadena;

            try
            {
            }
            catch
            {
            }
            finally
            { 
            }

        }



                /// <summary>
                /// Libera los recursos utilizados por el objeto.
                /// </summary>
                public void Dispose()
        {
            if (this.conexion != null)
            {
                this.conexion.Dispose();
                GC.SuppressFinalize(this.conexion);
            }
            if (this.obj_xfunciones != null)
            {
                this.obj_xfunciones.Dispose();
                GC.SuppressFinalize(this.obj_xfunciones);
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

            GC.SuppressFinalize(this.i);
            GC.SuppressFinalize(this.j);
            GC.SuppressFinalize(this.bandera);
            GC.SuppressFinalize(this.id_Reporte);
            GC.SuppressFinalize(this.days_deleted);

            GC.Collect();
        }
    }
}