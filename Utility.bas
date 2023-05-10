Attribute VB_Name = "Utility"
Option Explicit

Sub FormatWorkSheet()
'Format the worksheet in default setting
Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveWindow.DisplayGridlines = False
    Columns("A:A").Select
    Selection.ColumnWidth = 1.5
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.499984741
        .ThemeFont = xlThemeFontNone
    End With
    Rows("1:1").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.499984740745262
        .ThemeFont = xlThemeFontNone
    End With
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Title Placeholder"
    Range("B2").Select
End Sub

Sub CreateNewWorksheet()
Attribute CreateNewWorksheet.VB_Description = "Create a new worksheet with the default formatting"
Attribute CreateNewWorksheet.VB_ProcData.VB_Invoke_Func = "I\n14"
' Create a new worksheet with the default formatting
' Keyboard Shortcut: Ctrl+Shift+I
'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.CustomProperties.Add _
    Name:="WorksheetFormat", Value:="Default"
    Call FormatWorkSheet
    End Sub
    
Sub GoToTOC():
' Go to the table of contents worksheet
    Sheets("TOC").Select
    Range("A1").Select
End Sub



Sub CreateTableofContents()
' Create table of contents worksheet
    Dim StartCell As Range
    Dim TOCSht As Worksheet
    Dim Sht As Worksheet
    Dim ShtName As String
    Dim MsgConfirm As VBA.VbMsgBoxResult
    
    
    On Error Resume Next
    Set TOCSht = Sheets("Sheet1")
    TOCSht.Name = "TOC"
    Call FormatWorkSheet
    
    Set StartCell = Application.InputBox( _
    prompt:="Where do you want to insert the table?" & vbNewLine & "Please select the cells:", _
    Title:="Insert Table of Contents", _
    Type:=8)
    
    If Err.Number = 424 Then Exit Sub
    
    MsgConfirm = VBA.MsgBox( _
    prompt:="Overwrite the existing table of contents?", _
    Buttons:=vbOKCancel + vbDefaultButton2)

    Set StartCell = StartCell.Cells(1, 1)
        
    For Each Sht In Worksheets
        If Sht.Visible = xlSheetVisible And Sht.Name <> "TOC" Then
            ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:="'" & Sht.Name & "'!A1", TextToDisplay:=Sht.Name
            StartCell.Select
                With Selection.Font
                    .Underline = xlUnderlineStyleNone
                End With
            Set StartCell = StartCell.Offset(1, 0)
        End If
    Next Sht
End Sub
