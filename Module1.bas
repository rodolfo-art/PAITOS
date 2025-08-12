Attribute VB_Name = "Module1"
Option Explicit

' ===== CONFIGURACIÓN =====
Private Const DATA_SHEET As String = "InstantReport"  ' hoja con tus datos

' Encabezados esperados (ajústalos si cambian)
Private Const H_LPN As String = "LPN"
Private Const H_ITEM As String = "Item"
Private Const H_QTY As String = "LPN Qty"
Private Const H_PALLET As String = "Pallet"
Private Const H_CURRLOC As String = "Current Loc"
Private Const H_SOURCE As String = "Source"          ' Source
Private Const H_LASTUPD As String = "Last Update"
Private Const H_TASKID As String = "Task ID"
Private Const H_TASKWAVE As String = "Task Wave"

' ===== Lanzador (asigna este macro al botón del Ribbon) =====
Public Sub TaskTraceFilter()
    On Error Resume Next
    frmTaskTrace.Show
End Sub

' ===== Motor principal =====
Public Sub RunTaskTrace(ByVal method As String, ByVal rawValues As String, Optional ByVal rawPallets As String = "")
    Dim ws As Worksheet: Set ws = SheetExists(DATA_SHEET)
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & DATA_SHEET & "'.", vbExclamation: Exit Sub
    End If

    Dim ur As Range: Set ur = GetTidyRange(ws)
    If ur Is Nothing Then
        MsgBox "No se detectó rango de datos con encabezados.", vbExclamation: Exit Sub
    End If

    ' Columna clave según método
    Dim keyName As String, colIdx As Long
    Select Case UCase$(method)
        Case "WAVE":   keyName = H_TASKWAVE
        Case "TASK":   keyName = H_TASKID
        Case "SOURCE": keyName = H_SOURCE
        Case "LPN":    keyName = H_LPN
        Case "ITEM":   keyName = H_ITEM
        Case Else:      MsgBox "Método no reconocido.", vbExclamation: Exit Sub
    End Select
    colIdx = FindCol(ur, keyName)
    If colIdx = 0 Then
        MsgBox "No se encontró la columna '" & keyName & "'.", vbExclamation: Exit Sub
    End If

    ' Valores a filtrar (coma/espacio/saltos de línea/;)
    Dim setVals As Object: Set setVals = ParseValues(rawValues)
    Dim useFilter As Boolean: useFilter = (setVals.Count > 0)

    ' Pallet adicional
    Dim setPal As Object: Set setPal = ParseValues(rawPallets)
    Dim usePal As Boolean: usePal = (setPal.Count > 0)
    Dim colPalIdx As Long
    If usePal Then
        colPalIdx = FindCol(ur, H_PALLET)
        If colPalIdx = 0 Then
            MsgBox "No se encontró la columna '" & H_PALLET & "'.", vbExclamation: Exit Sub
        End If
    End If

    ' Normalización especial para Source
    If UCase$(method) = "SOURCE" And useFilter Then
        Dim norm As Object, k As Variant
        Set norm = CreateObject("Scripting.Dictionary")
        For Each k In setVals.Keys
            norm(CleanSource(CStr(k))) = 1
        Next k
        Set setVals = norm
    End If

    ' Crear salida
    Dim wsOut As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("TaskTrace_Result").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = Worksheets.Add(After:=ws)
    wsOut.Name = "TaskTrace_Result"

    ' Copiar encabezados
    Dim j As Long, colCount As Long: colCount = ur.Columns.Count
    For j = 1 To colCount
        wsOut.Cells(1, j).Value = ur.Cells(1, j).Value
    Next j

    ' Filtrar y copiar
    Dim r As Long, outR As Long: outR = 2
    Dim v As String
    For r = 2 To ur.Rows.Count
        v = CStr(ur.Cells(r, colIdx).Value)
        If ShouldKeep(v, setVals, useFilter, method) Then
            If Not usePal Or ShouldKeep(ur.Cells(r, colPalIdx).Value, setPal, True, "PALLET") Then
                For j = 1 To colCount
                    wsOut.Cells(outR, j).Value = ur.Cells(r, j).Value
                Next j
                outR = outR + 1
            End If
        End If
    Next r

    ' Ordenar: por clave y luego por Last Update
    Dim colKeyOut As Long, colLU As Long
    colKeyOut = FindCol(wsOut.UsedRange, keyName)
    colLU = FindCol(wsOut.UsedRange, H_LASTUPD)

    If outR > 2 And colKeyOut > 0 Then
        With wsOut.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsOut.Range(wsOut.Cells(1, colKeyOut), wsOut.Cells(outR - 1, colKeyOut)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            If colLU > 0 Then
                .SortFields.Add Key:=wsOut.Range(wsOut.Cells(1, colLU), wsOut.Cells(outR - 1, colLU)), _
                                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            End If
            .SetRange wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outR - 1, colCount))
            .Header = xlYes
            .Apply
        End With
    End If

    ' Estética del reporte
    FormatReport wsOut, colCount, outR - 1

    Dim msg As String: msg = "TaskTrace:"
    If useFilter Then msg = msg & " filtro por " & keyName
    If usePal Then msg = msg & IIf(useFilter, " y ", " filtro por ") & H_PALLET
    If Not useFilter And Not usePal Then msg = msg & " sin filtro (todo)"
    MsgBox msg, vbInformation
End Sub

' ===== Estética / UI =====
Private Sub FormatReport(ByVal ws As Worksheet, ByVal colCount As Long, ByVal lastRow As Long)
    If lastRow < 1 Then lastRow = 1

    ws.Activate
    ActiveWindow.DisplayGridlines = False               ' oculta cuadricula
    ws.Cells.Interior.Color = RGB(246, 248, 252)        ' fondo suave

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
        .Font.Bold = True
        .Interior.Color = RGB(230, 243, 255)
    End With

    Dim rr As Long
    For rr = 2 To lastRow
        If (rr Mod 2) = 0 Then
            ws.Range(ws.Cells(rr, 1), ws.Cells(rr, colCount)).Interior.Color = RGB(252, 254, 255)
        End If
    Next rr

    ws.Columns("A:Z").AutoFit
    ws.Range("A2").Select
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

' ===== Helpers =====
Private Function GetTidyRange(ws As Worksheet) As Range
    If ws.UsedRange Is Nothing Then Exit Function
    Dim ur As Range: Set ur = ws.UsedRange
    If Application.WorksheetFunction.CountA(ur) = 0 Then Exit Function

    Dim f As Long, l As Long, fc As Long, lc As Long
    f = ur.Row: l = ur.Rows(ur.Rows.Count).Row
    fc = ur.Column: lc = ur.Columns(ur.Columns.Count).Column

    Do While Application.WorksheetFunction.CountA(ws.Rows(f)) = 0 And f < l: f = f + 1: Loop
    Do While Application.WorksheetFunction.CountA(ws.Rows(l)) = 0 And l > f: l = l - 1: Loop
    Do While Application.WorksheetFunction.CountA(ws.Columns(fc)) = 0 And fc < lc: fc = fc + 1: Loop
    Do While Application.WorksheetFunction.CountA(ws.Columns(lc)) = 0 And lc > fc: lc = lc - 1: Loop

    On Error Resume Next
    Set GetTidyRange = ws.Range(ws.Cells(f, fc), ws.Cells(l, lc))
    On Error GoTo 0
End Function

Private Function FindCol(ur As Range, headerName As String) As Long
    Dim c As Range
    For Each c In ur.Rows(1).Cells
        If StrComp(Trim$(CStr(c.Value)), headerName, vbTextCompare) = 0 Then
            FindCol = c.Column - ur.Column + 1
            Exit Function
        End If
    Next c
End Function

Private Function ParseValues(ByVal s As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim cleaned As String
    cleaned = Replace(Replace(Replace(Replace(CStr(s), vbCrLf, ","), vbLf, ","), ";", ","), " ", ",")
    cleaned = Replace(cleaned, ",,", ",")
    Dim a() As String, i As Long, t As String
    a = Split(cleaned, ",")
    For i = LBound(a) To UBound(a)
        t = Trim$(a(i))
        If Len(t) > 0 Then d(UCase$(t)) = 1
    Next i
    Set ParseValues = d
End Function

Private Function ShouldKeep(ByVal cellVal As Variant, ByVal d As Object, _
                            ByVal useFilter As Boolean, ByVal method As String) As Boolean
    If Not useFilter Then
        ShouldKeep = True: Exit Function
    End If
    Dim probe As String
    probe = UCase$(Trim$(CStr(cellVal)))
    If UCase$(method) = "SOURCE" Then probe = CleanSource(probe)
    ShouldKeep = d.Exists(probe)
End Function

Private Function CleanSource(ByVal s As String) As String
    ' Mantiene solo A-Z/0-9 en mayúsculas (quita espacios, guiones, etc.)
    Dim i As Long, ch As String, outp As String
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Z0-9]" Then outp = outp & ch
    Next i
    CleanSource = outp
End Function

Private Function SheetExists(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetExists = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function
