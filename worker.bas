Attribute VB_Name = "Module1"
'copySourceAndModule()
Sub phase1_copySourceAndModule()

'needed: 1. name of file to copy; 2. name of tab to copy; 3.


'some declarations
Dim filePath, sourceFileName, newFileName As String
filePath = ThisWorkbook.path  'the default is curdir or current working directory
sourceFileName = "source.xlsx"
newFileName = "new-" & Date & "-" & Hour(Time) & Minute(Time) & Second(Time) & ".xlsx"

'copy the file (requires absolute path)
FileCopy filePath & "\" & sourceFileName, filePath & "\" & newFileName

'open the copy (requires absolute path)
Workbooks.Open Filename:=filePath & "\" & newFileName

'save it as a workbook object (requires just the name of an open wb)
Dim wb As Workbook
Set wb = Workbooks(newFileName)

Dim ws As Worksheet
Set ws = wb.Worksheets("b2win")


'copy this module

'Const TEMPFILE As String = "C:\Users\Dextercat\Desktop\temp.bas" 'temp textfile
Dim TEMPFILE As String
TEMPFILE = filePath & "\temp.bas" 'temp textfile


'export the module as a textfile
'*** to pull this off: https://stackoverflow.com/questions/25638344/programmatic-access-to-visual-basic-project-is-not-trusted
'*** we need to set trust access to vba project object model
Workbooks("worker.xlsm").VBProject.VBComponents("Module1").Export TEMPFILE

'import the module to the new workbook
wb.VBProject.VBComponents.Import TEMPFILE
  
'kill the textfile
Kill TEMPFILE


wb.Activate
ws.Activate

Range("a1").Select

MsgBox "Phase 1: complete"

'enter an if statement that if something went wrong, escape to kill the program

ThisWorkbook.Save
ThisWorkbook.Close

' once thisworkbook is closed, this macro goes with it, but it will remain in the new file
' and from there, we can run the second part of this algorithm! how exciting

' this is good because the file we copied is an .xlsx file
' keep it that way
' but once you close it, the .bas file is gone, since it needs to be macro-enabled to remember the module


' interesting notes:
' FileCopy source, destination
' why use const?
' what is a .bas file?
' more trust center whatnots

End Sub

Private Sub phase2_extractData()

Dim data As Worksheet
Set data = Worksheets("b2win")
data.Name = "data"

Range("K1").Activate
ActiveCell.End(xlDown).Activate
ActiveCell.Offset(0, -1).Activate
ActiveCell.End(xlUp).Offset(0, 0).Activate

Dim x As String
x = ActiveCell.Address
'MsgBox x

Dim nrows As Long 'if use integer, you get overflow error since > 32767
nrows = Val(Range(x).Row)

'MsgBox nrows

data.Name = "b2win"

'works fine
Range("k1").Value = "order"
Range("k2").Formula = "=IF(AND(ISNUMBER(A2),LEN(A2)=6), A2, K1)"

'works fine
Range("l1").Value = "ar"
Range("l2").Formula = "=IF(LEN(RIGHT(TRIM(B2), 6))=6, RIGHT(TRIM(B2), 6), L1)"

'numbervalue(x) does not exist in 2010
'also, in route number, b2win puts bad data in column d

'new formula
Range("m1").Value = "code"
Range("m2").Formula = "=IF(ISNUMBER(RIGHT(C2, 6)+0), RIGHT(C2, 6), M1)"

'new formula -
Range("n1").Value = "route"
Range("n2").Formula = "=IF(ISNUMBER(RIGHT(D2, 4)+0), RIGHT(D2, 4), N1)"


Range("o1").Value = "ref1"
Range("o2").Formula = "=IF(OR(LEFT(TRIM(E2), 2)=" & Chr(34) & "OR" & Chr(34) & ", LEFT(TRIM(E2), 2)=" & Chr(34) & "PT" & Chr(34) & "), E2, O1)"


Range("p1").Value = "ref2"
Range("P2").Formula = "=IF(LEFT(TRIM(E2), 5)=" & Chr(34) & "Ref.:" & Chr(34) & ", E2, P1)"


Range("q1").Value = "itemname"
Range("q2").Formula = "=E2"

Range("r1").Value = "qty"
Range("r2").Formula = "=f2"

Range("s1").Value = "unit"
Range("s2").Formula = "=h2"

Range("t1").Value = "date"
Range("t2").Formula = "=i2"

Range("u1").Value = "amt"
Range("u2").Formula = "=j2"


Dim s As String
s = "k2:u" & nrows


'autofill
Range("k2:u2").Select
Selection.AutoFill Destination:=Range(s), Type:=xlFillDefault

'Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'MsgBox "Phase 2: Complete"


End Sub


Private Sub phase3_cleanData()

Columns("A:J").Delete

'MsgBox "Phase 3: Complete"

End Sub

Private Sub phase4_deleteLinesWithZeroOrNonNumberQuantity()

'qty is in column h

Dim s As Range
Set s = Range("a1")
s.Activate

Dim n As Long
n = Range(s, s.End(xlDown)).Count

Range("l1").Value = "x"
Range("l2").Formula = "=if(isnumber(h2), abs(h2), 0)"

'autofill
Range("l2").Select
Selection.AutoFill Destination:=Range("l2:l" & n), Type:=xlFillDefault

'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'sort
Columns("A:L").Select
ActiveWorkbook.Worksheets("b2win").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("b2win").Sort.SortFields.Add Key:=Range("L2:L" & n), _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("b2win").Sort
    .SetRange Range("A1:L" & n)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Columns("L:L").Select
Selection.Find(What:="0", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Select
    
Range(Selection, Selection.End(xlDown).End(xlToLeft)).Delete

'get rid of this superfluous column
Columns("L").Delete

End Sub

Private Sub phase5_deleteLinesWithZeroAmts()

Dim s As Range
Set s = Range("a1")

Dim n As Long
n = Range(s, s.End(xlDown)).Count

Range("l1").Value = "x"
Range("l2").Formula = "=if(isnumber(k2), abs(k2), 0)"

'autofill
Range("l2").Select
Selection.AutoFill Destination:=Range("l2:l" & n), Type:=xlFillDefault

'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'sort
Columns("A:L").Select
ActiveWorkbook.Worksheets("b2win").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("b2win").Sort.SortFields.Add Key:=Range("L2:L" & n), _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("b2win").Sort
    .SetRange Range("A1:L" & n)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'find the first zero
Columns("L:L").Select
Selection.Find(What:="0", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Select

'delete them
Range(Selection, Selection.End(xlDown).End(xlToLeft)).Delete

'get rid of this superfluous column
Columns("L").Delete


End Sub

Private Sub phase6_addingDataMarkers()

Dim s As Range
Set s = Range("a1")

Dim n As Long
n = Range(s, s.End(xlDown)).Count

'******* insert orderfamily ********
'*******                    ********

Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

Range("b1").Value = "orderfamily"

Range("B2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1], 2)"

'autofill
Selection.AutoFill Destination:=Range("B2:B" & n)
'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value


'******* insert ref1family ********
'*******                   ********

Columns("g:g").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

Range("g1").Value = "ref1family"
Range("g2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1], 2)"

'autofill
Selection.AutoFill Destination:=Range("g2:g" & n)
'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'******* insert fancy1     ********
'*******                   ********

Columns("M:O").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

Range("m1").Value = "entryAmtIsZero"
Range("n1").Value = "ptAmtAvg"
Range("o1").Value = "ptAmtIsZero"

Range("m2").Formula = "=IF(P2=0, 1, 0)"
Range("n2").Formula = "=SUMIFS(M:M,F:F,F2)/COUNTIFS(F:F,F2)"
Range("o2").Formula = "=IF(N2=1, 1, 0)"

'autofill
Range("m2:o2").Select
Selection.AutoFill Destination:=Range("m2:o" & n)
'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'******* insert fancy2     ********
'*******                   ********

Range("r1").Value = "contains1899"
Range("s1").Value = "balance"
Range("t1").Value = "entryIsBalanced"
Range("u1").Value = "ptEntryBalanceAvg"
Range("v1").Value = "ptIsBalanced"

Range("r2").Formula = "=IF(IFERROR(FIND(" & Chr(34) & "1899" & Chr(34) & ", H2), 0), 1, 0)"
Range("s2").Formula = "=SUMIFS(J:J,F:F,F2,D:D,D2)"
Range("t2").Formula = "=IF(S2<>0, 0, 1)"
Range("u2").Formula = "=SUMIFS(T:T,F:F,F2)/COUNTIFS(F:F,F2)"
Range("v2").Formula = "=IF(U2=1, 1, 0)"

'autofill
Range("r2:v2").Select
Selection.AutoFill Destination:=Range("r2:v" & n)
'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

Columns("Q").Delete


'******* fancy3 insert     ********
'*******                   ********

Range("v1").Value = "leftright"
Range("v2").Formula = "=IF(P2>0, 1, 2)"

'autofill
Range("v2").Select
Selection.AutoFill Destination:=Range("v2:v" & n)
'paste values
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

s.Select

End Sub

Private Sub phase7_createPivotTable()
'
' Macro1 Macro
'

'
    Columns("A:V").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "b2win!R1C1:R1048576C22", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ptIsBalanced")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("orderfamily")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("orderfamily").Orientation _
        = xlHidden
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ref1family")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ref1")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("code")
        .Orientation = xlRowField
        .Position = 4
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "code[All]", xlLabelOnly + _
        xlFirstRow, True
    ActiveSheet.PivotTables("PivotTable1").RowAxisLayout xlTabularRow
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ptIsBalanced")
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ref1family")
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("ref1")
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("code")
        .PivotItems("(blank)").Visible = False
    End With
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").Select
    Columns("D:D").EntireColumn.AutoFit
    Range("D4").Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("ar"), "Count of ar", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("order"), "Count of order", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("qty"), "Count of qty", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("amt"), "Count of amt", xlCount
    ActiveWindow.SmallScroll Down:=0
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("leftright")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("leftright")
        .PivotItems("(blank)").Visible = False
    End With
    Range("E6").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("order").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("orderfamily").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ar").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("code").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("route").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ref1").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ref1family").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ref2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("itemname").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("qty").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("unit").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("date").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("entryAmtIsZero").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ptAmtAvg").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ptAmtIsZero").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("amt").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("contains1899").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("balance").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("entryIsBalanced"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ptEntryBalanceAvg"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("ptIsBalanced").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("leftright").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Range("E6").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of order")
        .Caption = "Sum of order"
        .Function = xlSum
    End With
    Range("F6").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of ar")
        .Caption = "Sum of ar"
        .Function = xlSum
    End With
    Range("G6").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of qty")
        .Caption = "Sum of qty"
        .Function = xlSum
    End With
    Range("H6").Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of amt")
        .Caption = "Sum of amt"
        .Function = xlSum
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("M:N").Select
    Selection.ColumnWidth = 4.89
    Range("I2").Select
   
End Sub


Sub phase2_theRest()

    Call phase2_extractData
    Call phase3_cleanData
    Call phase4_deleteLinesWithZeroOrNonNumberQuantity
    Call phase5_deleteLinesWithZeroAmts
    Call phase6_addingDataMarkers
    Call phase7_createPivotTable
    
End Sub
