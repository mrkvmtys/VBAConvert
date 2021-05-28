
Option Explicit

Sub ConvertToCsv()
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim myPath As String
    Dim myFile As String
    Dim myExt As String
    Dim NewWBName As String
    Dim ChooseFolder As FileDialog
    
    'Optimize
      Application.ScreenUpdating = False
      Application.EnableEvents = False
      Application.Calculation = xlCalculationManual
    
    'Retrieve Target Folder Path From User
    Set ChooseFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    ChooseFolder.Title = "Select Target Path"
    ChooseFolder.AllowMultiSelect = False
            
    If ChooseFolder.Show <> -1 Then GoTo NextCode
        myPath = ChooseFolder.SelectedItems(1) & "\"
    
    'Cancel
NextCode:
    myPath = myPath
    If myPath = "" Then Exit Sub
    
    'File Ext to Change
    myExt = "*.xls*"
    
    'Target Path with Ending Extention
    myFile = Dir(myPath & myExt)
    
    'Loop through each Excel file in folder
    Do While myFile <> ""
        'Set variable equal to opened workbook
        Set wb = Workbooks.Open(Filename:=myPath & myFile)
        'NewWBName = myPath & Left(myFile, InStr(1, myFile, ".") - 1) & ".csv"
        'disable scientific notation for column G
        Columns("G:G").Select
        Selection.NumberFormat = "0"
        
        Columns("H:J").Select
        Selection.Delete Shift:=xlToLeft
        
        NewWBName = myPath & Left(myFile, InStrRev(myFile, ".") - 1) & ".csv"
        ActiveWorkbook.SaveAs Filename:=NewWBName, FileFormat:=xlCSV, Local:=True
        ActiveWorkbook.Close savechanges:=True
        'Get next file name
        myFile = Dir
    Loop
    
    'Reset Macro Optimization Settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
