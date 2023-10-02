using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ReportServer2022;


namespace ReportServer2022.code.querys.procesos
{
    public class class_trading_genera_GSK_bk09052022
    {
        class_conexion conexion = new class_conexion();

        private string[,] tab_titulos;
        private String Line_Buffer;
        private int i, j;
        private StreamWriter File_IO;

        private void sub_init_var()
        {
            tab_titulos = new string[15, 2];

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
            tab_titulos[14, 1] = "mm/dd/yyyy hh:mm";
        }

        public class_llena_tabla tbl_trading_genera_GSK_consultar = new class_llena_tabla();
        public void sub_trading_genera_GSK(String Carpeta, String File_Name, String[,] file_tab, String Fecha_1, String Fecha_2, String Empresa, int idCron)
        {
            File_IO = new StreamWriter(Carpeta + File_Name, true);

            sub_init_var();

            conexion.sub_set_conexion();
            tbl_trading_genera_GSK_consultar.cadena_conexion = conexion.db_cadena;
            tbl_trading_genera_GSK_consultar.sqltext = "SELECT \n"
            + " NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA), '', '', to_char(WCD.DATE_CREATED, 'dd/mm/yy'), '', INITCAP(DIS.DISNOM) REMITENTE, InitCap(DISADRESSE1 || ' ' || ' ' || DISNUMEXT || '  ' || DISNUMINT || '  ' ||DISADRESSE2 || DECODE(DISCODEPOSTAL,NULL,NULL, ' C.P. ' || DISCODEPOSTAL)), \n"
            + " INITCAP(CIU_ORI.VILNOM || ' ('|| EST_ORI.ESTNOMBRE || ')'), INITCAP(CCL.CCL_NOMBRE || ' ' || NVL(DIE.DIE_A_ATENCION_DE, DIE.DIENOMBRE)), InitCap( DIEADRESSE1|| ' ' || ' ' || DIENUMEXT || '  ' || DIENUMINT || '  ' ||DIEADRESSE2 || DECODE(DIECODEPOSTAL,NULL,NULL, ' C.P. ' || DIECODEPOSTAL)), \n"
            + " INITCAP(CIU_DEST.VILNOM || ' ('|| EST_DEST.ESTNOMBRE || ')'), 'Road', WCD.WCD_FIRMA, to_char(WCD.DATE_CREATED, 'dd/mm/yy') \n"
            + " FROM \n"
            + " WCROSS_DOCK WCD, EDIRECCIONES_ENTREGA DIE, ECLIENT_CLIENTE CCL, EDISTRIBUTEUR DIS, ECIUDADES CIU_ORI, EESTADOS EST_ORI, ECIUDADES CIU_DEST, EESTADOS EST_DEST, ETRANS_DETALLE_CROSS_DOCK TDCD, ETRANSFERENCIA_TRADING TRA, ETRANS_ENTRADA TAE \n"
            + " WHERE \n"
            + " TRUNC(WCD.DATE_CREATED) BETWEEN TRUNC(sysdate -1) AND TRUNC(sysdate -1) AND WCD_CLICLEF in(20501,20502) AND NOT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA) LIKE '%PRUEBA%' AND NOT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA) LIKE '%SENSORES%' \n"
            + " AND NOT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA) LIKE '%TARIMAS%' AND DISCLEF = WCD.WCD_DISCLEF AND DIECLAVE = NVL(NVL(TDCD_DIECLAVE_ENT, TDCD_DIECLAVE), WCD_DIECLAVE_ENTREGA) AND CCLCLAVE = NVL(TDCD_CCLCLAVE, WCD.WCD_CCLCLAVE) \n"
            + " AND CIU_ORI.VILCLEF = DISVILLE AND EST_ORI.ESTESTADO = CIU_ORI.VIL_ESTESTADO AND CIU_DEST.VILCLEF = DIEVILLE AND EST_DEST.ESTESTADO = CIU_DEST.VIL_ESTESTADO AND TDCDCLAVE(+) = WCD.WCD_TDCDCLAVE AND TDCDSTATUS (+) = '1' \n"
            + " AND TRACLAVE(+) = WCD.WCD_TRACLAVE AND TRASTATUS (+) = '1' AND TAE_TRACLAVE(+) = WCD.WCD_TRACLAVE ";
            tbl_trading_genera_GSK_consultar.sub_llenar_tabla();


            if (tbl_trading_genera_GSK_consultar.tabla_llena.Rows.Count != 0)
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

                file_tab = new string[3, file_tab.GetLength(1) + 1];


                file_tab[0, file_tab.GetLength(1) - 1] = Carpeta + File_Name;
                file_tab[1, file_tab.GetLength(1) - 1] = "Shipments";
                file_tab[2, file_tab.GetLength(1) - 1] = tab_titulos.GetLength(0) + "";
            }
            //= Concentrado
            if (tbl_trading_genera_GSK_consultar.tabla_llena.Rows.Count != 0)
            {
                Line_Buffer = "";
                for (i = 0; i < tbl_trading_genera_GSK_consultar.tabla_llena.Rows.Count; i++)
                {
                    for (j = 0; j < tbl_trading_genera_GSK_consultar.tabla_llena.Columns.Count; j++)
                    {
                        Line_Buffer = Line_Buffer + tbl_trading_genera_GSK_consultar.tabla_llena.Rows[i][j] + "" + "\n";
                    }
                }

                File_IO.WriteLine(Line_Buffer);
            }
            File_IO.Close();
        }
    }
}
