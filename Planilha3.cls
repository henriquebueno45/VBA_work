VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

Dim senha As String
Dim Resp As String

If Planilha4.Range("A1") = "SIM" Then
Exit Sub
End If

Planilha3.Visible = False
senha = "123"
Resp = InputBox("Informe sua Senha!", "SENHA")
If Resp = senha Then
Planilha3.Visible = True
Planilha4.Range("A1") = "SIM"
Planilha3.Select

Else
MsgBox "Senha errada!", vbCritical, "SENHA"
Planilha3.Visible = False
End If

End Sub

Private Sub Worksheet_Deactivate()
Planilha4.Range("A1") = ""
End Sub
