Sub copySourceAndModule()

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

Sub countRows()

Dim data As Worksheet
Set data = Worksheets("b2win")
data.Name = "data"

Range("K1").Activate
ActiveCell.End(xlDown).Activate
ActiveCell.Offset(0, -1).Activate
ActiveCell.End(xlUp).Offset(0, 0).Activate

Dim x As String
x = ActiveCell.Address
MsgBox x

Dim nrows As Long 'if use integer, you get overflow error since > 32767
nrows = Val(Range(x).Row)

MsgBox nrows

data.Name = "b2win"

'works fine
Range("k1").Value = "w"
Range("k2").Formula = "=IF(AND(ISNUMBER(A2),LEN(A2)=6), A2, K1)"

'works fine
Range("l1").Value = "x"
Range("l2").Formula = "=IF(LEN(RIGHT(TRIM(B2), 6))=6, RIGHT(TRIM(B2), 6), L1)"

'numbervalue(x) does not exist in 2010
'also, in route number, b2win puts bad data in column d

'new formula
Range("m1").Value = "y"
Range("m2").Formula = "=IF(ISNUMBER(RIGHT(C2, 6)+0), RIGHT(C2, 6), M1)"

'new formula -
Range("n1").Value = "z"
Range("n2").Formula = "=IF(ISNUMBER(RIGHT(D2, 4)+0), RIGHT(D2, 4), N1)"


'autofill
Range("k2:n2").Select
Selection.AutoFill Destination:=Range("k2:n14"), Type:=xlFillDefault


End Sub
