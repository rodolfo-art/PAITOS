VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTaskTrace 
   Caption         =   "TaskTrace"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optWave
      Caption         =   "WAVE ID"
      Left            =   180
      Top             =   180
      Width           =   1200
      Height          =   240
   End
   Begin VB.OptionButton optTask
      Caption         =   "TASK ID"
      Left            =   180
      Top             =   420
      Width           =   1200
      Height          =   240
   End
   Begin VB.OptionButton optSource
      Caption         =   "SOURCE"
      Left            =   180
      Top             =   660
      Width           =   1200
      Height          =   240
   End
   Begin VB.OptionButton optLPN
      Caption         =   "LPN"
      Left            =   180
      Top             =   900
      Width           =   1200
      Height          =   240
   End
   Begin VB.OptionButton optItem
      Caption         =   "ITEM"
      Left            =   180
      Top             =   1140
      Width           =   1200
      Height          =   240
   End
   Begin VB.TextBox txtValues
      MultiLine       =   -1  'True
      ScrollBars      =   2   'Vertical
      Tag             =   "VALUES"
      Left            =   1620
      Top             =   180
      Width           =   3600
      Height          =   1200
   End
   Begin VB.TextBox txtPallet
      MultiLine       =   -1  'True
      ScrollBars      =   2   'Vertical
      Tag             =   "PALLET"
      Left            =   1620
      Top             =   1560
      Width           =   3600
      Height          =   1200
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Left            =   3000
      Top             =   3000
      Width           =   1000
      Height          =   300
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Cancel          =   -1  'True
      Left            =   4080
      Top             =   3000
      Width           =   1000
      Height          =   300
   End
End
Attribute VB_Name = "frmTaskTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' Marca "TASK ID" por defecto, usando el Caption del OptionButton
    Dim c As Object
    For Each c In Me.Controls
        If TypeName(c) = "OptionButton" Then
            If UCase$(Trim$(c.Caption)) = "TASK ID" Then
                c.Value = True
                Exit For
            End If
        End If
    Next c
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim m As String: m = GetSelectedMethod()
    If Len(m) = 0 Then
        MsgBox "Selecciona un método.", vbExclamation
        Exit Sub
    End If

    Dim s As String: s = Trim$(GetTextBoxValue("VALUES"))
    Dim p As String: p = Trim$(GetTextBoxValue("PALLET"))
    If Len(s) = 0 And Len(p) = 0 Then
        If MsgBox("No pegaste valores. ¿Continuar para ver todo?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    RunTaskTrace m, s, p
    Unload Me
End Sub

' ---------- helpers del formulario ----------
Private Function GetSelectedMethod() As String
    Dim c As Object, cc As Object
    ' 1) OptionButtons de primer nivel
    For Each c In Me.Controls
        If TypeName(c) = "OptionButton" Then
            If c.Value Then
                GetSelectedMethod = MapCaptionToMethod(c.Caption)
                If Len(GetSelectedMethod) > 0 Then Exit Function
            End If
        ' 2) También busca dentro de Frames / MultiPage si los usaste
        ElseIf TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            For Each cc In c.Controls
                If TypeName(cc) = "OptionButton" Then
                    If cc.Value Then
                        GetSelectedMethod = MapCaptionToMethod(cc.Caption)
                        If Len(GetSelectedMethod) > 0 Then Exit Function
                    End If
                End If
            Next cc
        End If
    Next c
End Function

Private Function MapCaptionToMethod(ByVal captionText As String) As String
    Select Case UCase$(Trim$(captionText))
        Case "WAVE ID": MapCaptionToMethod = "WAVE"
        Case "TASK ID": MapCaptionToMethod = "TASK"
        Case "SOURCE": MapCaptionToMethod = "SOURCE"
        Case "LPN":      MapCaptionToMethod = "LPN"
        Case "ITEM":     MapCaptionToMethod = "ITEM"
        Case Else:        MapCaptionToMethod = ""
    End Select
End Function

Private Function GetTextBoxValue(Optional ByVal boxTag As String = "") As String
    ' Devuelve el texto del TextBox según su Tag (si se especifica)
    Dim c As Object, cc As Object
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            If boxTag = "" Or UCase$(CStr(c.Tag)) = UCase$(boxTag) Then
                GetTextBoxValue = CStr(c.Text): Exit Function
            End If
        ElseIf TypeName(c) = "Frame" Or TypeName(c) = "MultiPage" Then
            For Each cc In c.Controls
                If TypeName(cc) = "TextBox" Then
                    If boxTag = "" Or UCase$(CStr(cc.Tag)) = UCase$(boxTag) Then
                        GetTextBoxValue = CStr(cc.Text): Exit Function
                    End If
                End If
            Next cc
        End If
    Next c
    GetTextBoxValue = ""
End Function
