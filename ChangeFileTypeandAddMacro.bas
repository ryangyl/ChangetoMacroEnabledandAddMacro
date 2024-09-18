Attribute VB_Name = "Module2"
Sub ConvertXlsxToXlsmAndAddMacro()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim macroCode As String
    Dim xlsxFilePath As String
    Dim xlsmFilePath As String
    Dim folders As Variant
    Dim folder As Variant
    folders = Array("\Sep 24\", "\Oct 24\", "\Nov 24\", "\Dec 24\")
    For Each folder In folders
    folderPath = "\\siwdsntv002\SG_PSC_SG1_PL_08_Control_WHse\Daily Tank Reading\Tanker reading year 2024" & folder
    
     
    ' Check if the folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Define the macro code for Workbook_BeforeClose
    macroCode = _
    "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
    "    Workbooks.Open ""\\siwdsntv002\SG_PSC_SG1_PL_08_Control_WHse\Daily Tank Reading\Solvent Tracking Macro.xlsm""" & vbCrLf & _
    "    ' Custom logic before closing" & vbCrLf & _
    "    ThisWorkbook.Close SaveChanges:=True" & vbCrLf & _
    "End Sub"
    
    ' Loop through each .xlsx file in the folder
    fileName = Dir(folderPath) ' Only .xlsx files are picked up
    
    Do While fileName <> ""
        ' Full path to the .xlsx file
        xlsxFilePath = folderPath & fileName
        
        ' Debugging print to see if the file name is recognized
        Debug.Print "Processing file: " & xlsxFilePath
        
        ' Open the .xlsx file (Error handling in case the file is not recognized)
        On Error Resume Next
        Set wb = Workbooks.Open(xlsxFilePath)
        If wb Is Nothing Then
            MsgBox "Error: Could not open " & fileName, vbExclamation
            Exit Sub
        End If
        On Error GoTo 0
        
        ' Add the macro to the ThisWorkbook object
        With wb.VBProject.VBComponents("ThisWorkbook").CodeModule
            .DeleteLines 1, .CountOfLines ' Remove any existing code in ThisWorkbook
            .AddFromString macroCode      ' Add the Workbook_BeforeClose macro
        End With
        
        ' Create the new .xlsm file path
        xlsmFilePath = folderPath & Replace(fileName, ".xlsx", ".xlsm")
        
        ' Save the workbook as .xlsm
        wb.SaveAs fileName:=xlsmFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
        ' Close the workbook
        wb.Close SaveChanges:=True
        
        ' Delete the original .xlsx file
        Kill xlsxFilePath
        
        ' Get the next file
        fileName = Dir
    Loop
Next folder
    MsgBox "All files converted, macros added, and original .xlsx files deleted!"
End Sub

