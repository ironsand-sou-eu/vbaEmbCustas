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
    'Se conseguiu salvar e as planilhas de configuração estão abertas, pergunta e oculta.
    If Success = True And ThisWorkbook.IsAddin = False Then
        Dim bolConfirma As Boolean
        bolConfirma = MsgBox("As alterações foram salvas. As planilhas de configuração serão ocultas novamente. ", _
                vbInformation + vbOKCancel, "Sísifo - Alterações salvas")
        If bolConfirma = True Then
            RestringirEdicaoRibbon ThisWorkbook
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Se as planilhas de configuração estiverem abertas para edição, esconde-as de novo.
    RestringirEdicaoRibbon ThisWorkbook
End Sub
