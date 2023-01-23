Attribute VB_Name = "Disponible"
Sub Disponbile()
Attribute Disponbile.VB_ProcData.VB_Invoke_Func = "D\n14"

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.CutCopyMode = False

'Range("A2:M2", Range("A2:M2").End(xlDown)).Delete

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Columns("A:Z").Select
Selection.ClearContents

Range("A1").Formula = "Foto"
Range("B1").Formula = "Item"
Range("C1").Formula = "Genero"
Range("D1").Formula = "Categoria"
Range("E1").Formula = "Calificacion"
Range("F1").Formula = "Existencia total"
Range("G1").Formula = "Inv Actual B044"
Range("H1").Formula = "Disponible B034"
Range("I1").Formula = "S"
Range("J1").Formula = "M"
Range("K1").Formula = "L"
Range("L1").Formula = "Inv Transito"
Range("M1").Formula = "Nuevo"
Range("N1").Formula = "Lista negra"
Range("O1").Formula = "Validaciones"
Range("P1").Formula = "Inv en B044"
Range("Q1").Formula = "S a enviar"
Range("R1").Formula = "M a enviar"
Range("S1").Formula = "L a enviar"
Range("T1").Formula = "Items a enviar total"
Range("U1").Formula = "Unidades en B005"
Range("V1").Formula = "Unidades en B001"
Range("W1").Formula = "Unidades en tienda"
Range("x1").Formula = "Unidades DAFITI"

Sheets("B034").Activate
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Range("A:H").AutoFilter Field:=8, Criteria1:=">=0"
Range("B2:D2", Range("B2:D2").End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Copy
Sheets("Disponible").Range("B2").PasteSpecial xlPasteValues
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Sheets("Disponible").Activate
Columns("B:D").RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
Range("A2").FormulaR1C1 = "=+RC2"
Range("B2").End(xlDown).Offset(0, -1).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Selection.Copy
Selection.PasteSpecial xlPasteValues

Range("E2").FormulaR1C1 = _
        "=+IFERROR(VLOOKUP(RC2,ruta,13,FALSE),"""")"
Range("F2").FormulaR1C1 = "=RC[1]+RC[2]+RC[6]"
Workbooks.OpenText (ruta)
Workbooks("Importación SALE hombre").Activate
        Range("G2").FormulaR1C1 = "=+SUMIFS('[B044.txt]B044'!C6, '[B044.txt]B044'!C2, RC2)"
Range("H2").FormulaR1C1 = "=+SUM(RC[1]:RC[3])"
Range("I2").FormulaR1C1 = "=+SUMIFS(B034!C6, B034!C2, RC2, B034!C5, ""S"")"
Range("J2").FormulaR1C1 = "=+SUMIFS(B034!C6, B034!C2, RC2, B034!C5, ""M"")"
Range("K2").FormulaR1C1 = "=+SUMIFS(B034!C6, B034!C2, RC2, B034!C5, ""L"")"
Workbooks.OpenText (ruta)
Workbooks("Importación SALE hombre").Activate
Range("L2").FormulaR1C1 = "=+SUMIFS(ruta, ruta, RC2)"
Workbooks.OpenText (ruta)
Workbooks("Importación SALE hombre").Activate
Range("M2").FormulaR1C1 = _
        "=IF(OR(COUNTIF(ruta,RC[-12])>0,RC[-1]>0),""NO"",""SI"")"
Range("D2").End(xlDown).Offset(0, 1).Select
Selection.Resize(1, 9).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Columns("I:M").Copy
Columns("I:M").PasteSpecial xlPasteValues
Columns("G").Copy
Columns("G").PasteSpecial xlPasteValues
Workbooks("Fecha importaciones SALE").Close
Workbooks("Transito SALE").Close
Workbooks("B044").Close

'Pintar Viejos
Fin = Range("A1", Range("A2").End(xlDown)).Rows.Count
    Dim i As Variant
    For i = 2 To Fin
    If Cells(i, 13) = "NO" Then
        Cells(i, 13).Interior.Color = RGB(246, 247, 178)
    End If
    Next i

'Tamaño y letra
Range("A2:M2", Range("A2:M2").End(xlDown)).Select
     With Selection.Font
        .Size = 9
    End With
    
    Range("A2:M2", Range("A2:M2").End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With

'Ordenar Genero y Categoría
Range("A:M").Sort key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes
Range("A:M").Sort key1:=Range("C1"), Order1:=xlAscending, Header:=xlYes

'Romper Vinculo
ActiveWorkbook.BreakLink Name:=ruta, Type:=xlLinkTypeExcelLinks


 'Formula para ver inventario en B044

ufila = Cells(Rows.Count, "c").End(xlUp).Row - 1
Range("P2").Resize(ufila).Formula = "=RC[-9]+RC[-4]"
Range("G:G,L:L").Select
Selection.EntireColumn.Hidden = True
    
'Máximo 250 unidades por ítem

Range("Q2").Resize(ufila).Formula = "=IF(RC8>3,IF(RC16<250,IF(RC6<=250,RC[-8],ROUND((RC[-8]/(RC6-RC16))*(250-RC16),0)),0),RC[-8])"
Range("R2").Resize(ufila).Formula = "=IF(RC8>3,IF(RC16<250,IF(RC6<=250,RC[-8],ROUND((RC[-8]/(RC6-RC16))*(250-RC16),0)),0),RC[-8])"
Range("S2").Resize(ufila).Formula = "=IF(RC8>3,IF(RC16<250,IF(RC6<=250,RC[-8],ROUND((RC[-8]/(RC6-RC16))*(250-RC16),0)),0),RC[-8])"

'Validaciones
ufila2 = Cells(Rows.Count, "c").End(xlUp).Row - 1
Workbooks.Open ruta
Workbooks.Open ruta
Workbooks.Open ruta
Workbooks.Open ruta

Workbooks("Importación SALE hombre.xlsm").Activate
Range("T2").Resize(ufila2).Formula = "=SUM(RC[-3]:RC[-1])"
Range("N2").Resize(ufila2).Formula = "=IFNA(VLOOKUP(RC[-12],'[Lista negra.xlsx]Lista negra'!C1:C2,2,0),VLOOKUP(RC[-12],'[Lista negra.xlsx]Foto'!C1:C3,3,0))"
Range("U2").Resize(ufila2).Formula = "=SUMIFS(B005.txt!C6,B005.txt!C2,RC[-19])"
Range("V2").Resize(ufila2).Formula = "=SUMIFS(B001.txt!C6,B001.txt!C2,RC[-20])"
Range("W2").Resize(ufila2).Formula = "=SUMIFS(Consulta1[Existencia],Consulta1[item],Disponible!RC[-22])"
Range("X2").Resize(ufila2).Formula = "=SUMIFS('[Consolidado pedido dafiti.xlsx]INVENTARIO'!C4,'[Consolidado pedido dafiti.xlsx]INVENTARIO'!C1,RC[-22])"


Windows("B005.txt").Activate
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

Windows("B001.txt").Activate
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
Windows("Importación SALE hombre.xlsm").Activate
ActiveWorkbook.BreakLink Name:=ruta, Type:=xlExcelLinks
ActiveWorkbook.BreakLink Name:=ruta, Type:=xlExcelLinks
ActiveWorkbook.BreakLink Name:=ruta, Type:=xlExcelLinks
ActiveWorkbook.BreakLink Name:=ruta, Type:=xlExcelLinks

Workbooks("B001.txt").Close
Workbooks("B005.txt").Close
Workbooks("Consolidado pedido dafiti.xlsx").Close
Workbooks("Lista negra.xlsx").Close
Windows("Importación SALE hombre.xlsm").Activate

Range("A1").Select
'Selection.AutoFilter
'ActiveSheet.Range("$A$1:$X$314").AutoFilter Field:=20, Criteria1:="0"
'Range("2:2").Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.SpecialCells(xlCellTypeVisible).Select
'Selection.Delete Shift:=xlUp
'Selection.AutoFilter

End Sub
