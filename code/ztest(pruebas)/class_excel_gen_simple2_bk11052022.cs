using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportServer2022.code.include
{
    /*
Sub Excel_simple2(tab_file() As String, Optional freeze As Boolean)
'igual que el excel_simple pero funciona con una tabla de archivos
'tab_file(0, i) -> ubicacion del archivo
'tab_file(1, i) -> nombre de la pestaña excel
'tab_file(2, i) -> numero de columnas
'tab_file(3, i) -> columnas a borrar (separadas con coma) - hay que clasificarlo al reves ahorita [optional]
'tab_file(4, i) -> tamano de la primera columna [optional]


'Abrimos Excel
Set My_excel = New Excel.Application
'Set My_excel2 = New Excel.Application

My_excel.Visible = True
My_excel.WindowState = xlMinimized
'My_excel2.Visible = True
'Set My_Workbook = My_excel.Workbooks.Add
'Set My_Workbook2 = My_excel.Workbooks

For j = 0 To UBound(tab_file, 2)
    If FSO.FileExists(tab_file(0, j)) Then
        exclColNum = tab_file(2, j)
        If j = 0 Then
            Set My_Workbook = My_excel.Workbooks.Open(tab_file(0, j))
            'Evitar que el nombre sobrepase 31 caracteres oque tenga caracteres prohibidos
            My_Workbook.ActiveSheet.Name = Left(Remove_Forbidden_Chars(tab_file(1, j), FORBIDDEN_SHEET_CHARS), 31)
            
            
        Else
        
            'Set My_Workbook2 = My_excel.Workbooks.Add
            Set My_Workbook2 = My_excel.Workbooks.Open(tab_file(0, j))
            'Evitar que el nombre sobrepase 31 caracteres oque tenga caracteres prohibidos
            My_Workbook2.ActiveSheet.Name = Left(Remove_Forbidden_Chars(tab_file(1, j), FORBIDDEN_SHEET_CHARS), 31)
            My_Workbook2.ActiveSheet.Move Before:=My_Workbook.Sheets(1)
            'My_Workbook2.Activate
            'My_Workbook2.Close
            'My_excel.ActiveWorkbook.Close False  'no guarda los cambios
        End If
        
        'My_excel.Workbooks.
        ' ------------  Mise en forme de la feuille
        'en-têtes
        'With My_excel.ActiveSheet.Range(Cells(1, 1), Cells(1, exclColNum))
        '    .BorderAround xlContinuous, xlMedium    'Encadre la sélection
        '    .Interior.ColorIndex = 15               'Colorie le fond de la sélection (en Gris clair)
        '    .Interior.Pattern = xlSolid
        'End With
        
        'formatting headers :
        With My_excel.Range("A1", LettreCol(exclColNum) & "1")
            .Font.size = 8
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.ColorIndex = 2       'put font in white
            .Interior.ColorIndex = 1   'fill cell with black
        End With
        
        'formatting datas :
        For i = 0 To exclColNum
            With My_excel.Range(LettreCol(i) & "2", LettreCol(i) & "2").EntireColumn
                .Font.size = 8
                .Font.Bold = True
                .HorizontalAlignment = xlCenter     'center
                .VerticalAlignment = xlTop          'top
                tab_datas = Split(My_excel.Range(LettreCol(i) & "1", LettreCol(i) & "1").Value, "|")
                'MsgBox UBound(tab_datas)
                If UBound(tab_datas) > 0 Then
                    .NumberFormat = tab_datas(1)
                    My_excel.Range(LettreCol(i) & "1", LettreCol(i) & "1").Value = tab_datas(0)
                Else
                    .NumberFormat = "0" 'before we used @ but for great number it put 9.15767E+14 for 915767000191576
                End If
                    '.NumberFormat = "0"
                    If i = 0 Then
                        If UBound(tab_file, 1) <= 3 Then
                    .EntireColumn.AutoFit
                        Else
                            'la primera columna tiene un tamano definido
                            If tab_file(4, j) <> "" Then
                                .ColumnWidth = CDbl(tab_file(4, j))
                                .HorizontalAlignment = xlLeft
                            End If
                        End If
                    Else
                        .EntireColumn.AutoFit
                    End If
            End With
        Next
        
        'données
        'aplicamos el freeze despues de comenzar a pegar los datos
        If freeze = True Then
            My_Workbook.ActiveSheet.Range("A2").Select
            My_excel.ActiveWindow.FreezePanes = True
        End If
    End If
    
    'eliminamos columnas
    'ahorita la tabla esta clasificada al reves
    If UBound(tab_file, 1) >= 3 Then
        Dim tab_delete() As String
        If tab_file(3, j) <> "" Then
            tab_delete = Split(tab_file(3, j), ",")
            
            For i = 0 To UBound(tab_delete)
                'borramos la columna
                If Trim(tab_delete(i)) <> "" Then My_excel.Range(LettreCol(CInt(tab_delete(i))) & "2", LettreCol(CInt(tab_delete(i))) & "2").EntireColumn.Delete
            Next
        End If
    End If
Next

My_Workbook.SaveAs Left(tab_file(0, 0), Len(tab_file(0, 0)) - 4) & ".xls", xlWorkbookNormal ', xlExcel9795
'guarda el archiva en formato Excel 97-95


My_excel.Quit
Set My_excel = Nothing

For j = 0 To UBound(tab_file, 2)
    If FSO.FileExists(tab_file(0, j)) Then FSO.DeleteFile (tab_file(0, j))
Next


End Sub
     * 
     */

    public class class_excel_gen_simple2_bk11052022
    {

        private Excel.Application My_excel;
        private int i, j;
        private Excel.Workbook My_Workbook;
        private Excel.Workbook My_Workbook2;
        private Excel._Worksheet My_Sheet;
        

        //public void sub_excel_simple2(String[,] tab_file, Boolean freeze = true )
        public void sub_excel_simple2(string[,] tab_file, Boolean freeze = true )
        {
            //Abrimos Excel
            My_excel = new Excel.Application();
            My_excel.Visible = true;
            My_excel.WindowState = Excel.XlWindowState.xlMinimized;

            for (j = 0; j < tab_file.GetLength(1); j++)
            {
                if (File.Exists(tab_file[0, j]) == true)
                {
                    string exclColNum = tab_file[2, j];

                    if (j == 0)
                    {


                        //My_Workbook = (My_excel.Workbooks.Open(tab_file[0, j]));
                        My_Workbook = My_excel.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory + tab_file[0, j]);
                        My_Sheet = (Excel._Worksheet)My_Workbook.ActiveSheet;
                        My_Sheet.Name = tab_file[1, j];
                    }
                    else
                    {
                        My_Workbook = (My_excel.Workbooks.Add(tab_file[0, j]));
                        My_Sheet = (Excel._Worksheet)My_Workbook.ActiveSheet;
                        My_Sheet.Name = tab_file[1, j];
                        My_Sheet.Move(My_Workbook.Sheets[1]);
                    }

                    
                        //formatting headers
                        My_Sheet.Range["A1", int.Parse(exclColNum) + 65 + "1"].Font.Bold = true;
                        My_Sheet.Range["A1", exclColNum + 65 + "1"].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        My_Sheet.Range["A1", exclColNum + 65 + "1"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; ;
                        My_Sheet.Range["A1", exclColNum + 65 + "1"].Font.ColorIndex = 2;
                        My_Sheet.Range["A1", exclColNum + 65 + "1"].Interior.ColorIndex = 1;
                    

                }
            }

        }
    }
}
