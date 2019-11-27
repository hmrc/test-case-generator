Attribute VB_Name = "CSV_Generation"
Option Explicit
Sub GenerateTestCaseCSVs()
Dim WorkBookName As String
WorkBookName = Application.ActiveWorkbook.FullName
Dim LastRow As Long
Dim I As Integer
Dim strIPF As String
Dim strINT As String
Dim stropf As String

Call Create_Directories

strIPF = ThisWorkbook.Path & "/input/"
strINT = ThisWorkbook.Path & "/interim/"
stropf = ThisWorkbook.Path & "/output/"

Sheets("Test case index").Select

LastRow = Range("Last_row").Value
LastRow = LastRow - 1

For I = 1 To LastRow
Call Run_Test_Case(WorkBookName, I)
If Not Application.OperatingSystem Like "*Mac*" Then
    Call Save_Input_File(I, "C:\taxcalc\input\" & I)
    Call Save_Interim_File(I, "C:\taxcalc\interim\" & I)
    Call Save_Output_File(I, "C:\taxcalc\output\" & I)
Else
    Call Save_Input_File(I, strIPF & I)
    Call Save_Interim_File(I, strINT & I)
    Call Save_Output_File(I, stropf & I)
    
End If
Next I
End Sub

Sub Create_Directories()

If Not Application.OperatingSystem Like "*Mac*" Then
    Call Create_Win_Dir
Else
    Call Create_Mac_Dir
End If

End Sub
Sub Create_Win_Dir()
If Len(Dir("C:\taxcalc", vbDirectory)) = 0 Then
    MkDir "C:\taxcalc"
End If

If Len(Dir("C:\taxcalc\input", vbDirectory)) = 0 Then
    MkDir "C:\taxcalc\input"
End If

If Len(Dir("C:\taxcalc\interim", vbDirectory)) = 0 Then
    MkDir "C:\taxcalc\interim"
End If

If Len(Dir("C:\taxcalc\output", vbDirectory)) = 0 Then
    MkDir "C:\taxcalc\output"
    End If
End Sub
Sub Create_Mac_Dir()
        
    MakeFolderIfNotExist (ThisWorkbook.Path & Application.PathSeparator & "input")

    MakeFolderIfNotExist (ThisWorkbook.Path & Application.PathSeparator & "interim")

    MakeFolderIfNotExist (ThisWorkbook.Path & Application.PathSeparator & "output")

End Sub

Sub Generate_Input_Files(LastRow)

Dim WorkBookName As String
Dim I As Integer
WorkBookName = Application.ActiveWorkbook.FullName

For I = 1 To LastRow
Call Run_Test_Case(WorkBookName, I)
If Not Application.OperatingSystem Like "*Mac*" Then
Call Save_Input_File(I, "C:\taxcalc\input\" & I)
Else
    Call Save_Input_File(I, strIPF & I)
End If

Next I
End Sub

Sub Generate_Interim_Files(LastRow)

Dim WorkBookName As String
Dim I As Integer
WorkBookName = Application.ActiveWorkbook.FullName

For I = 1 To LastRow
If Not Application.OperatingSystem Like "*Mac*" Then
Call Save_Interim_File(I, "C:\taxcalc\interim\" & I)
Else
    Call Save_Interim_File(I, strINT & I)
End If

Next I
End Sub

Sub Generate_Output_Files(LastRow)

Dim WorkBookName As String
Dim I As Integer
WorkBookName = Application.ActiveWorkbook.FullName

For I = 1 To LastRow
If Not Application.OperatingSystem Like "*Mac*" Then
Call Save_Output_File(I, "C:\taxcalc\output\" & I)
Else
    Call Save_Output_File(I, stropf & I)
End If

Next I
End Sub

Sub Run_Test_Case(WorkBookName, TestCaseNumber)
Sheets("Test Case index").Select
Range("C3").Select
ActiveCell.FormulaR1C1 = TestCaseNumber
Range("D3").Select
Call HzReplay_Test_case
End Sub

Sub Save_Input_File(TestCaseNumber, FileName)
'   Input Test Case
Call Save_File(TestCaseNumber, "TaxCalc_Input_JUNIT", FileName)
End Sub

Sub Save_Interim_File(TestCaseNumber, FileName)
'   Interim Test Case
Call Save_File(TestCaseNumber, "TaxCalc_Interim_JUNIT", FileName)
End Sub

Sub Save_Output_File(TestCaseNumber, FileName)
'   Output Test Case
Call Save_File(TestCaseNumber, "TaxCalc_FinalOutput_JUNIT", FileName)
End Sub

Sub Save_File(TestCaseNumber, Sheet, FileName)
Sheets(Sheet).Select
Application.CutCopyMode = False
Sheets(Sheet).Copy
Sheets(Sheet).Select
Sheets(Sheet).Name = TestCaseNumber

Application.DisplayAlerts = False

If Not Application.OperatingSystem Like "*Mac*" Then

    ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=6
    ActiveWindow.Close
Else
    ActiveWorkbook.SaveAs FileName:=FileName & ".csv", FileFormat:=6
    ActiveWindow.Close
End If

Application.DisplayAlerts = True

End Sub

Sub HzReplay_Test_case()
Dim intSelected_case As Integer
Dim intLast_case As Integer

intSelected_case = Worksheets("Test Case index").Range("C3")
intLast_case = Worksheets("Library").Range("Last_row")

If intSelected_case > 0 And intSelected_case < intLast_case Then
  reset_test_selection
  Component_string = "Components: "
  Reload_settings
  test_this_case
  Worksheets("SA302").Range("D5") = Component_string
  Worksheets("SA302").Range("G4") = "Test case " & intSelected_case
Else
  MsgBox "This is not a valid test case number"
End If
End Sub

Function MakeFolderIfNotExist(Folderstring As String)
'Ron de Bruin, 22-June-2015
' http://www.rondebruin.nl/mac/mac010.htm
    Dim ScriptToMakeFolder As String
    Dim str As String
    If Val(Application.Version) < 15 Then
        ScriptToMakeFolder = "tell application " & Chr(34) & _
                             "Finder" & Chr(34) & Chr(13)
        ScriptToMakeFolder = ScriptToMakeFolder & _
                "do shell script ""mkdir -p "" & quoted form of posix path of (" & _
                        Chr(34) & Folderstring & Chr(34) & ")" & Chr(13)
        ScriptToMakeFolder = ScriptToMakeFolder & "end tell"
        On Error Resume Next
        MacScript (ScriptToMakeFolder)
        On Error GoTo 0

    Else
        str = MacScript("return POSIX path of (" & _
                        Chr(34) & Folderstring & Chr(34) & ")")
        MkDir str
    End If
End Function
   







