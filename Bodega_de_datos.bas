Attribute VB_Name = "Bodega_de_datos"
Sub Bodega()
'
' Macro2 Macro
'
Workbooks.Open ruta
Range("A:BM").Delete
Workbooks.Open ruta
Workbooks.Open ruta
MsgBox ("Seleccione la TCEDI")
Application.Dialogs(xlDialogOpen).Show
Libro = ActiveWorkbook.Name

despacho = ActiveSheet.Name
ufila = Cells(Rows.Count, "a").End(xlUp).Row - 1
Columns("F:F").Insert Shift:=xlToRight
Range("C2").Resize(ufila).Formula = despacho
Range("F2").Resize(ufila).Formula = "=RC[-2]&""-""&RC[-5]&""-""&RC[-3]"
Range("F:F").Copy
Range("F1").PasteSpecial xlPasteValues
Range("AU:AU,BS:BS").Delete


'ALT TB
ufila = ufila + 1
Application.CutCopyMode = False
Range("F1").Select
    
    
ActiveWorkbook.PivotCaches.Create(SourceType:=xlConsolidation, SourceData:="RC6:R" & ufila & "C69", Version:=7).CreatePivotTable _
    TableDestination:="", TableName:="TablaDinámica1", DefaultVersion:=7
ActiveSheet.PivotTables("TablaDinámica1").DataPivotField.PivotItems( _
    "Suma de Valor").Position = 1
    
    
ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
ActiveSheet.Cells(3, 1).Select
ActiveSheet.PivotTables("TablaDinámica1").DataPivotField.PivotItems("Suma de Valor").Position = 1
ActiveSheet.PivotTables("TablaDinámica1").PivotFields("Columna").Orientation = xlHidden
ActiveSheet.PivotTables("TablaDinámica1").PivotFields("Fila").Orientation = xlHidden
Range("A4").Select
Selection.ShowDetail = True

'Ordenar libro
Selection.Copy
Workbooks("Plantilla preparación.xlsm").Activate
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Workbooks(Libro).Close (False)


Workbooks("Plantilla preparación.xlsm").Activate
Columns("B:B").Select
Selection.Replace What:="B.Eco", Replacement:="B. eco", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
ufila = Cells(Rows.Count, "a").End(xlUp).Row - 1
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Selection.Insert Shift:=xlToRight
Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
    TrailingMinusNumbers:=True
Range("F1").Value = "id_cia"
Range("G1").Value = "concatenado"
Range("H1").Value = "row_id_item_ext"
Range("I1").Value = "año_despacho"
Range("J1").Value = "row_id_bodega"
Range("K1").Value = "Fecha_exhibicion"
Range("F2").Resize(ufila).Formula = "=VLOOKUP(RC[-2],'[Distribución tiendas.xlsx]Distribucion'!C2:C15,14,0)"
Range("G2").Resize(ufila).Formula = "=RC[-6]&RC[-5]&RC[-1]"
Range("H2").Resize(ufila).Formula = "=VLOOKUP(RC[-1],[item_extensiones.xlsx]Consulta!C1:C5,5,0)"
Range("I2").Resize(ufila).Formula = "=IF(MID(RC[-6],2,1)="","",YEAR(TODAY())&""_0""&RC[-6],YEAR(TODAY())&""_""&RC[-6])"
Range("J2").Resize(ufila).Formula = "=VLOOKUP(RC[-6],'[Distribución tiendas.xlsx]Distribucion'!C2:C16,15,0)"
Range("K2").Resize(ufila).Formula = "=(TODAY()-WEEKDAY(TODAY(),3))+14"
Range("F:F", "K:K").Copy
Range("F1").PasteSpecial xlPasteValues


Range("A:D,F:G").Delete
Range("B:B").Cut

Columns("B:B").Select
Selection.Cut
Columns("A:A").Select
Selection.Insert Shift:=xlToRight
Columns("E:E").Select
Selection.Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Columns("E:E").Select
Selection.Cut
Columns("C:C").Select
Selection.Insert Shift:=xlToRight
Columns("E:E").Select
Selection.Cut
Columns("D:D").Select
Selection.Insert Shift:=xlToRight
Selection.Insert Shift:=xlToRight
Range("D1").Select
ActiveCell.FormulaR1C1 = "notas"
Range("E1").Select

Workbooks("item_extensiones.xlsx").Close (False)
Workbooks("Distribución tiendas.xlsx").Close (False)
Workbooks.Open ruta
Workbooks.Open (Libro)
Workbooks("Plantilla preparación.xlsm").Activate
End Sub

