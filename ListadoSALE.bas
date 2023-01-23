Attribute VB_Name = "Listado"
Sub Listado()

Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.DisplayAlerts = False

'Generar Listado SALE
Workbooks.Add

ActiveWorkbook.SaveAs Filename:= _
    ruta, FileFormat _
    :=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
ActiveSheet.Name = "SALE"
Windows("Importación SALE hombre").Activate
Sheets("Disponible").Activate
Range("A:D").Copy
Windows("LIST hombre").Activate
Range("A1").PasteSpecial xlPasteValues
Windows("Importación SALE hombre").Activate
Range("Q:S").Copy
Windows("LIST hombre").Activate
Range("F1").PasteSpecial xlPasteValues
ufila = Cells(Rows.Count, "A").End(xlUp).Row - 1
Cells(1, 1) = "FOTO"
Cells(1, 2) = "ITEM"
Cells(1, 3) = "GENERO"
Cells(1, 4) = "CATEGORÍA"
Cells(1, 5) = "INV TOTAL"
Cells(1, 6) = "S"
Cells(1, 7) = "M"
Cells(1, 8) = "L"
Range("E2").Select
ActiveCell.Formula = "=SUM(RC[1]:RC[3])"
Range("E2").Resize(ufila).Formula = "=SUM(RC[1]:RC[3])"
Range("E:E").Copy
Range("E1").PasteSpecial xlPasteValues

'Cambiar Formato
 Range("A1:H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True

Range("A1:H1").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
       With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        ActiveWindow.DisplayGridlines = False
      
Range("D1").Sort key1:=Range("D2"), Order1:=xlAscending, Header:=xlYes
Range("C1").Sort key1:=Range("C2"), Order1:=xlAscending, Header:=xlYes


Application.Dialogs(xlDialogSaveAs).Show


'
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.CutCopyMode = False
' Macro1 Macro
On Error Resume Next
Workbooks.Open ruta
Workbooks.Open ruta
Workbooks.Open ruta
Workbooks("LIST hombre.xlsx").Activate
ufila = Cells(Rows.Count, "a").End(xlUp).Row - 1
ufila2 = ufila + 1

'Añadir precios
    Range("I2").Resize(ufila).Formula = "=IFNA(VLOOKUP(RC[-7],'[Importación SALE hombre.xlsm]Precios'!C1:C9,9,0),VLOOKUP(RC[-7],'[Fecha importaciones SALE.xlsx]Fecha importación SALE'!C1:C4,4,0))"
    Range("J2").Resize(ufila).Formula = "=RC[-1]&""-""&RC[-8]"
    
    Range("H1").Copy
    Range("I1:J1").PasteSpecial Paste:=xlPasteFormats
    Range("I1").FormulaR1C1 = "PRECIO"
    Range("J1").FormulaR1C1 = "FINAL"
    Columns("I:J").Copy
    Range("I1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Selection.Cut
    Columns("F:F").Insert Shift:=xlToRight
    Range("F1").Select
    
'ALT TB
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlConsolidation, SourceData:= _
        "RC7:R" & ufila2 & "C10", Version:=7).CreatePivotTable _
        TableDestination:="", TableName:="TablaDinámica3", DefaultVersion:=7
    ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
    ActiveSheet.Cells(3, 1).Select
    ActiveSheet.PivotTables("TablaDinámica3").DataPivotField.PivotItems( _
        "Suma de Valor").Position = 1
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Columna").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Fila").Orientation = _
        xlHidden
    Range("A4").Select
    Selection.ShowDetail = True

    Application.CutCopyMode = False
    ChDir "K:\Comercial\Inteligencia\Stocks\ECOMMERCE\SALE"

    Sheets("SALE").Select
    Range("A1").Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "FINAL"
    Range("A1:F1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("A1").FormulaR1C1 = "FOTO"
    Range("B1").FormulaR1C1 = "DESPACHO"
    Range("C1").FormulaR1C1 = "ITEM"
    Range("D1").FormulaR1C1 = "TALLA"
    Range("E1").FormulaR1C1 = "CANT. SIESA"
    Range("F1").FormulaR1C1 = "PRECIO UNITARIO"

    Sheets.Add After:=ActiveSheet
    Sheets("Hoja3").Select
    Columns("A:C").Copy
    Sheets("Hoja5").Range("A1").PasteSpecial xlPasteValues
    Sheets("Hoja5").Select
    Rows("1:1").Delete Shift:=xlUp
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("FINAL").Range("F2").PasteSpecial Paste:=xlPasteValues
    Sheets("Hoja5").Select
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("FINAL").Range("A2").PasteSpecial xlPasteValues
    Sheets("FINAL").Range("C2").PasteSpecial xlPasteValues
    Range("C1").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("FINAL").Range("D2").PasteSpecial xlPasteValues
    Range("D1").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("FINAL").Range("E2").PasteSpecial xlPasteValues
    Sheets("Hoja3").Delete
    Sheets("Hoja2").Delete
    Sheets("Hoja5").Delete
    Sheets("SALE").Delete
    Sheets("FINAL").Name = "SALE"

End Sub




