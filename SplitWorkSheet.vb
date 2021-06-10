Sub SplitWorkSheet()

    'Made It Easy by ExcelExciting.com | Faraz Shaikh | Updated : 20200801
    '******************************************************
    'Validation Check active workbook file path missing
    If ActiveWorkbook.Path = vbNullString Then
        MsgBox "Please save the workbook to run the code", vbCritical, "Excel Exciting"
        Exit Sub
    End If
    '******************************************************
    Dim FilePath As String
    FilePath = Application.ActiveWorkbook.Path 'Extracting the current workbook File Path
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '******************************************************
    'Check the OS to add the path separator accordingly
        CheckOS = Application.OperatingSystem
        If CheckOS Like "*Mac*" Then
            OS_Separator = "/"  'for Mac
        Else
            OS_Separator = "\"  'for Windows
        End If
    '******************************************************
  
    'Loop Starting here
    For Each ws In ThisWorkbook.Sheets
    
        ws.Copy
        Application.ActiveWorkbook.SaveAs FileName:=FilePath & OS_Separator & ws.Name & ".xlsx"
        Application.ActiveWorkbook.Close False
    
    Next
  
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
  
End Sub
