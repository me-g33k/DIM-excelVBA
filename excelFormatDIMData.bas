Attribute VB_Name = "excelFormatDIMData"
'
' Change the file path in this location to one that you use.
'

Public Const fldrTarget As String = "C:\temp\"
Sub fixWeapons()
'
' The Main Routine
'
'

weaponTextToColumns
weaponHideColumn
buildWeaponTable

Range("C2").Select
ActiveWindow.FreezePanes = True
    
ds = mkDate()
fn = ds + "-destinyWeapons.xlsx"

filePath = fldrTarget + fn

ActiveWorkbook.SaveAs Filename:= _
        filePath _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub fixArmor()
'
' The Main Routine
'
'

armorTextToColumns
armorHideColumn
buildArmorTable

Range("C2").Select
ActiveWindow.FreezePanes = True

ds = mkDate()
fn = ds + "-destinyArmor.xlsx"

filePath = fldrTarget + fn

ActiveWorkbook.SaveAs Filename:= _
        filePath _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False


End Sub
Private Sub weaponTextToColumns()
'
' textToColumns Macro
' convertText
'

'
    Range(Selection, Selection.End(xlDown)).Select
    Selection.textToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
        Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
        33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1)), TrailingMinusNumbers _
        :=True
End Sub

Private Sub armorTextToColumns()
'
' textToColumns Macro
' convertText
'

'
    Range(Selection, Selection.End(xlDown)).Select
    Selection.textToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1)), TrailingMinusNumbers _
        :=True
End Sub

Private Sub weaponHideColumn()
Attribute weaponHideColumn.VB_Description = "Hide Columns"
Attribute weaponHideColumn.VB_ProcData.VB_Invoke_Func = " \n14"
'
' hideColumn Macro
' Hide Columns
'

'
    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("T:T").Select
    Selection.EntireColumn.Hidden = True
End Sub

Private Sub armorHideColumn()
'
' hideColumn Macro
' Hide Columns
'

'
    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("S:S").Select
    Selection.EntireColumn.Hidden = True
End Sub

Private Sub buildWeaponTable()
'
' buildTable Macro
' Build Table from Raw Data
'

'
    Cells(1, 1).Activate
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
        "Table1"
    Range("Table1[[#Headers],[ % Leveled]]").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 9
    Range("Table1[[#Headers],[ % Leveled]:[ Equip]]").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Columns("H:S").Select
    Columns("H:S").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Range("A2").Select
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
End Sub

Private Sub buildArmorTable()
'
' buildTable Macro
' Build Table from Raw Data
'

'
    Cells(1, 1).Activate
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
        "Table1"
    Range("Table1[[#Headers],[ % Leveled]]").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 9
    Range("Table1[[#Headers],[ Light]:[ Str]]").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Columns("H:R").Select
    Columns("H:R").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Range("A2").Select
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
End Sub

Private Function mkDate() As String

'
' makeDate Macro
' Make the date SubString
'

'

mkDate = Format(Date, "mmdd")



End Function
Private Sub saveFile()
'
' saveFile Macro
'

'
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\merana.VERTEXINC\Documents\2016-ME\003 - Destiny\Inventory\1116-destinyArmor.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub


