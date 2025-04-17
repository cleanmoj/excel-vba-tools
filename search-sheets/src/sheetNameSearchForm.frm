VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sheetNameSearchForm 
   Caption         =   "Search in sheets name"
   ClientHeight    =   2535
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "sheetNameSearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sheetNameSearchform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim matchedSheets As Collection
Dim currentIndex As Integer

Private Sub btnSearch_Click()
    Dim ws As Worksheet
    Dim searchText As String
    
    searchText = txtSearch.Value
    Set matchedSheets = New Collection
    currentIndex = 0
    
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, searchText, vbTextCompare) > 0 Then
            matchedSheets.Add ws.Name
        End If
    Next ws
    
    If matchedSheets.Count = 0 Then
        lblResult.Caption = "No matches found!"
    Else
        currentIndex = 1
        ShowSheetByIndex
    End If
End Sub

Private Sub btnNext_Click()
    If matchedSheets Is Nothing Then Exit Sub
    If currentIndex < matchedSheets.Count Then
        currentIndex = currentIndex + 1
        ShowSheetByIndex
    End If
End Sub

Private Sub btnPrev_Click()
    If matchedSheets Is Nothing Then Exit Sub
    If currentIndex > 1 Then
        currentIndex = currentIndex - 1
        ShowSheetByIndex
    End If
End Sub

Private Sub ShowSheetByIndex()
    If matchedSheets.Count = 0 Then Exit Sub
    Dim sheetName As String
    sheetName = matchedSheets(currentIndex)
    ThisWorkbook.Sheets(sheetName).Select
    lblResult.Caption = "Sheet: " & sheetName & " (" & currentIndex & " of " & matchedSheets.Count & ")"
End Sub

Private Sub UserForm_Click()

End Sub
