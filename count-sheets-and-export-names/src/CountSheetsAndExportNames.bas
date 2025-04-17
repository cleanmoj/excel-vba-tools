Attribute VB_Name = "CountSheetsAndExportNames"
' ---------------------------------------------
' Excel VBA Tool - Search Sheet Names
' Author      : Mojtaba Pakrah
' GitHub      : https://github.com/cleanmoj
' Created     : April 2025
' Description : This script adds search functionality for sheet names in Excel.
'
' Copyright (c) 2025 Mojtaba Pakrah
' Permission is granted to use, modify, and share this code for personal or commercial purposes.
' This code is provided "as is", without any warranty or guarantee of any kind.
' ---------------------------------------------
Sub CountSheetsAndExportTheirNames()
    Dim ws As Worksheet
    Dim sheetNames As String
    Dim sheetCount As Integer
    Dim response As VbMsgBoxResult
    Dim filePath As String
    Dim fileFormat As String
    Dim fso As Object
    Dim outFile As Object

    sheetCount = ThisWorkbook.Sheets.Count


    For Each ws In ThisWorkbook.Sheets
        sheetNames = sheetNames & ws.Name & vbCrLf
    Next

    response = MsgBox("Total sheets: " & sheetCount & vbCrLf & vbCrLf & _
                      "Sheet names:" & vbCrLf & sheetNames & vbCrLf & _
                      "Do you want to export these names?", vbYesNo + vbQuestion, "Sheet Info")

    If response = vbYes Then
        filePath = Application.GetSaveAsFilename( _
                    InitialFileName:="SheetNames", _
                    FileFilter:="Text Files (*.txt), *.txt," & _
                                "CSV Files (*.csv), *.csv," & _
                                "Excel Workbook (*.xlsx), *.xlsx", _
                    Title:="Save Sheet Names")

        If filePath = "False" Then Exit Sub

        Select Case LCase(Right(filePath, Len(filePath) - InStrRev(filePath, ".")))
            Case "txt", "csv"
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set outFile = fso.CreateTextFile(filePath, True)
                outFile.WriteLine "Total sheets: " & sheetCount
                outFile.WriteLine "Sheet names:"
                outFile.WriteLine sheetNames
                outFile.Close
                MsgBox "Saved successfully!", vbInformation
            Case "xlsx"
                Dim newWB As Workbook
                Set newWB = Workbooks.Add
                newWB.Sheets(1).Range("A1").Value = "Total sheets: " & sheetCount
                newWB.Sheets(1).Range("A2").Value = "Sheet Names"
                Dim i As Integer: i = 3
                For Each ws In ThisWorkbook.Sheets
                    newWB.Sheets(1).Cells(i, 1).Value = ws.Name
                    i = i + 1
                Next
                newWB.SaveAs filePath
                newWB.Close
                MsgBox "Saved successfully in Excel format!", vbInformation
            Case Else
                MsgBox "Unsupported file type!", vbCritical
        End Select
    End If
End Sub


