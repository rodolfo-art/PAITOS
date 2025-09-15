Attribute VB_Name = "modTT_Report"
Option Explicit

' Builds a nicely formatted report sheet using the provided headers and data.
Public Sub BuildPrettyReport(ByVal reportSheet As Worksheet, _
                             ByVal headers As Variant, _
                             Optional ByVal rows As Variant)
    Dim colIdx As Long
    Dim columnCount As Long
    Dim dataCol As Long
    Dim dataRow As Long
    Dim destinationRow As Long
    Dim headerLower As Long
    Dim headerUpper As Long
    Dim lastWrittenRow As Long
    Dim rowLowerCol As Long
    Dim rowLowerRow As Long
    Dim rowUpperCol As Long
    Dim rowUpperRow As Long
    Dim hasRowDimension As Boolean
    Dim hasColumnDimension As Boolean

    If reportSheet Is Nothing Then Exit Sub

    reportSheet.Cells.Clear

    If IsEmpty(headers) Then Exit Sub
    If Not IsArray(headers) Then headers = Array(headers)

    headerLower = LBound(headers)
    headerUpper = UBound(headers)
    columnCount = headerUpper - headerLower + 1

    For colIdx = headerLower To headerUpper
        reportSheet.Cells(1, colIdx - headerLower + 1).Value = headers(colIdx)
        reportSheet.Cells(1, colIdx - headerLower + 1).Font.Bold = True
    Next colIdx

    lastWrittenRow = 1
    If Not IsMissing(rows) Then
        If IsArray(rows) Then
            hasRowDimension = TryGetArrayBounds(rows, 1, rowLowerRow, rowUpperRow)
            hasColumnDimension = TryGetArrayBounds(rows, 2, rowLowerCol, rowUpperCol)

            If hasRowDimension And hasColumnDimension Then
                For dataRow = rowLowerRow To rowUpperRow
                    destinationRow = dataRow - rowLowerRow + 2
                    lastWrittenRow = destinationRow

                    For colIdx = headerLower To headerUpper
                        dataCol = colIdx - headerLower + rowLowerCol
                        If dataCol >= rowLowerCol And dataCol <= rowUpperCol Then
                            reportSheet.Cells(destinationRow, colIdx - headerLower + 1).Value = rows(dataRow, dataCol)
                        Else
                            reportSheet.Cells(destinationRow, colIdx - headerLower + 1).Value = vbNullString
                        End If
                    Next colIdx
                Next dataRow
            ElseIf hasRowDimension Then
                destinationRow = 2
                lastWrittenRow = destinationRow

                For colIdx = headerLower To headerUpper
                    dataCol = colIdx - headerLower + rowLowerRow
                    If dataCol >= rowLowerRow And dataCol <= rowUpperRow Then
                        reportSheet.Cells(destinationRow, colIdx - headerLower + 1).Value = rows(dataCol)
                    Else
                        reportSheet.Cells(destinationRow, colIdx - headerLower + 1).Value = vbNullString
                    End If
                Next colIdx
            End If
        End If
    End If

    ApplyReportStyling reportSheet, columnCount, lastWrittenRow
End Sub

' Applies simple formatting to make the report easier to read.
Private Sub ApplyReportStyling(ByVal reportSheet As Worksheet, _
                               ByVal columnCount As Long, _
                               ByVal lastRow As Long)
    Dim colIdx As Long

    If reportSheet Is Nothing Then Exit Sub
    If columnCount < 1 Then Exit Sub

    With reportSheet.rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(230, 243, 255)
    End With

    If lastRow < 2 Then lastRow = 2

    For colIdx = 1 To columnCount
        reportSheet.Columns(colIdx).AutoFit
    Next colIdx

    reportSheet.Range("A1").Select
End Sub

Private Function TryGetArrayBounds(ByRef source As Variant, _
                                   ByVal dimension As Long, _
                                   ByRef lowerBound As Long, _
                                   ByRef upperBound As Long) As Boolean
    On Error GoTo HandleError

    lowerBound = LBound(source, dimension)
    upperBound = UBound(source, dimension)
    TryGetArrayBounds = True
    Exit Function

HandleError:
    TryGetArrayBounds = False
    Err.Clear
End Function
