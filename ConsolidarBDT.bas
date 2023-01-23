Attribute VB_Name = "Módulo1"

Sub Consolidar_BDT()
Attribute Consolidar_BDT.VB_ProcData.VB_Invoke_Func = " \n14"

        
    MsgBox ("Seleccione el listado T.Cedi")
    Application.Dialogs(xlDialogOpen).Show
    TCEDI = ActiveWorkbook.Name
    Workbooks.OpenText Filename:= _
        ruta
        xlWindows , StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, OtherChar:="-", FieldInfo:= _
        Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7 _
        , 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array _
        (14, 1), Array(15, 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:=True
    Windows(TCEDI).Activate
    
    despacho = ActiveSheet.Name
    ufila = Cells(Rows.Count, "a").End(xlUp).Row - 1
    Columns("F:F").Insert Shift:=xlToRight
    Range("C2").Resize(ufila).Formula = despacho
    Range("F2").Resize(ufila).Formula = "=RC[-2]&""-""&RC[-5]&""-""&RC[-3]"
    Range("F:F").Copy
    Range("F1").PasteSpecial xlPasteValues
    Range("AU:AU,BS:BS").Delete
    
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
        Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("A:A").Select
    ufila = Cells(Rows.Count, "c").End(xlUp).Row - 1
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("A2").Resize(ufila).FormulaR1C1 = "=YEAR(TODAY())"
    Range("B2").Resize(ufila).FormulaR1C1 = "=UPPER(TEXT(LEFT(RC[3],FIND("","",RC[3])-1),""MMMM""))"
        Columns("F:F").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("G:G").Cut
    Columns("D:D").Insert Shift:=xlToRight
    Columns("D:D").Insert Shift:=xlToRight

    Range("D2").Resize(ufila).FormulaR1C1 = "=XLOOKUP(RC[2],'Maestro de ítems.txt'!C1,'Maestro de ítems.txt'!C2)"
    Range("I2").Resize(ufila).FormulaR1C1 = "=XLOOKUP(RC[-3],'Maestro de ítems.txt'!C1,'Maestro de ítems.txt'!C3)"
    Range("J2").Resize(ufila).FormulaR1C1 = "=XLOOKUP(RC[-4],'Maestro de ítems.txt'!C1,'Maestro de ítems.txt'!C4)"
    
    
    
        Workbooks.Open Filename:= _
        ruta
    
    Range("A1").End(xlDown).Offset(1, 0).Select
    Windows(TCEDI).Activate

    
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Columns("D:D").Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Selection.Copy
    Application.WindowState = xlNormal
    Windows("BDT.CEDI.xlsx").Activate
    Selection.PasteSpecial Paste:=xlPasteValues
    
End Sub


