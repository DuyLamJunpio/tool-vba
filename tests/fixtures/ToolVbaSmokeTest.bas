Attribute VB_Name = "ToolVbaSmokeTest"
Option Explicit

Public Sub ToolVbaSmokeTest()
    If ThisWorkbook.FileFormat <> 52 Then
        MsgBox "Test failed: workbook is not .xlsm.", vbExclamation, "Tool VBA"
        Exit Sub
    End If

    MsgBox "ToolVbaSmokeTest is loaded in an .xlsm workbook.", vbInformation, "Tool VBA"
End Sub
