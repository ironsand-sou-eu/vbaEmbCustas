VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a10SisifoIntimacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    'Se conseguiu salvar e as planilhas de configura��o est�o abertas, pergunta e oculta.
    If Success = True And ThisWorkbook.IsAddin = False Then
        Dim bolConfirma As Boolean
        bolConfirma = MsgBox("As altera��es foram salvas. As planilhas de configura��o ser�o ocultas novamente. ", _
                vbInformation + vbOKCancel, "S�sifo - Altera��es salvas")
        If bolConfirma = True Then
            RestringirEdicaoRibbon ThisWorkbook
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Se as planilhas de configura��o estiverem abertas para edi��o, esconde-as de novo.
    RestringirEdicaoRibbon ThisWorkbook
End Sub
