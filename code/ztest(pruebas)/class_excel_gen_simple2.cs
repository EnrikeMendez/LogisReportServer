using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace ReportServer2022.code.include
{
    public class class_excel_gen_simple2
    {
    
        //public void sub_excel_simple2(string[,] tab_file, string[,] tab_titulos)
        //{
        // SLDocument My_Workbook = new SLDocument();
        // int i, j, k, z;

        //    class_principal principal = new class_principal();
        //    principal.obj_querys.sub_trading_genera_GSK_consultar();

        //    //las cordenadas en excel comienzan en 1,1
        //    int n_column, n_row;
        //    n_column = 1;
        //    n_row = 2; //se inicializa en 2 porque la primera fila pertenece a la cabecera, el DT comenzara a escribir informacion en la segunda fila

        //    //En el primer for Crea cabeceras con el array tab_titulos. En el segundo for crea filas con con el DataTAble:
        //    for (i = 0; i < principal.obj_querys.tbl_trading_genera_GSK_consultar.tabla_llena.Columns.Count ; i++)
        //    {
        //        //El primer parametro son las filas, el segundo son las columnas, y el tercero son los datos de la celda:
        //        My_Workbook.SetCellValue(n_row-1, i, tab_titulos[i,0]);

        //        for (j = 0; j < principal.obj_querys.tbl_trading_genera_GSK_consultar.tabla_llena.Rows.Count; j++)
        //        {
        //            My_Workbook.SetCellValue(n_row + j, n_column + i, principal.obj_querys.tbl_trading_genera_GSK_consultar.tabla_llena.Rows[j][i] + "");
        //        }
        //    }
        //    //Guarda el archivo:
        //    string[] file_name_arr = tab_file[0, 0].Split('.');
        //    My_Workbook.SaveAs(file_name_arr[0] + ".xls");
        //}
    }
}

