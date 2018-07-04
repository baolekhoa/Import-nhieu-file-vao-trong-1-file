Option Explicit
Sub import_data()
    Dim ws_main As Worksheet, sh As Worksheet
    Dim wk As Workbook
    Dim strFolderPath As String
    Dim selectedFiles As Variant
    Dim iFileNum As Integer
    Dim strFileName As String
    Dim CR As Range 'Current Range
    Dim iLR As Long, iLRn As Long
    Dim Cate As String
    
    Application.ScreenUpdating = False
'    On Error GoTo ErrHandler

    
    Set ws_main = ActiveWorkbook.Sheets("sheet1")
    
    strFolderPath = ActiveWorkbook.Path
    
    ChDrive strFolderPath
    ChDir strFolderPath
    
    'Lay ten file vao mot mang, cho phep doc nhieu file
    selectedFiles = Application.GetOpenFilename(MultiSelect:=True, Title:="Select Files to Open")
    
    If TypeName(selectedFiles) = "Boolean" Then
        MsgBox "No files were selected"
        GoTo ExitHandler
    End If
    
    'Mo lan luot cac workbook
    For iFileNum = LBound(selectedFiles) To UBound(selectedFiles)
        
        'Lay ten 1 workbook tu mang ten cac files selectFiles
        strFileName = selectedFiles(iFileNum)
        
        'Mo file da chon
        Set wk = Workbooks.Open(strFileName)
        
        'Lam viec voi tung sheet trong tung workbook dang mo
        For Each sh In wk.Sheets
    
            'Tim Dong & Cot cuoi cung
             iLR = ws_main.Range("A" & Rows.Count).End(xlUp).Row

            'Copy range with CurrentRegion without header
            Set CR = sh.Range("A1").CurrentRegion
            Set CR = CR.Offset(1, 0).Resize(CR.Rows.Count - 1)
            CR.Copy ws_main.Cells(iLR + 1, 1)
            
            'Add Category
            iLRn = ws_main.Range("A" & Rows.Count).End(xlUp).Row
            ws_main.Range("AD" & iLR + 1 & ":" & "AD" & iLRn) = ExtractName(sh.Name)
        Next sh
        wk.Close
    Next
'ErrHandler:
'    MsgBox Err.Description
'    GoTo ExitHandler

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
End Sub

Function ExtractName(str As String) As String
    Dim txt As String
    
    If Not str Like "*-*" Then
        ExtractName = ""
    Else
        txt = Right(str, Len(str) - InStrRev(str, "-", -1, vbBinaryCompare))
        ExtractName = txt
    End If
End Function
