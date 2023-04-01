Attribute VB_Name = "Module3"

Sub Clearcells()

Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

    ws.Range("O2", "O4").ClearContents
    ws.Range("P1", "P4").ClearContents
    ws.Range("Q1", "Q4").ClearContents

Next ws


End Sub
