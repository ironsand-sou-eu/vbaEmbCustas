VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRelatorioAtos 
   Caption         =   "Sísifo - Insira os dados do processo"
   ClientHeight    =   7695
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6015
   OleObjectBlob   =   "frmRelatorioAtos.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmRelatorioAtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chbNaoGerarValorDaCausa_Change()
    ' Desabilita a caixa de valor da causa
    If chbNaoGerarValorDaCausa.value = True Then
        txtValor.Enabled = False
        Label7.Enabled = False
    Else
        txtValor.Enabled = True
        Label7.Enabled = True
    End If

End Sub

Private Sub chbSentencaEmbargos_Change()
    ' Marca ou desmarca a caixa de não gerar DAJE sobre o valor da causa.
    If chbSentencaEmbargos.value = True Then chbNaoGerarValorDaCausa.Visible = True
    If chbSentencaEmbargos.value = False Then chbNaoGerarValorDaCausa.value = False
        
End Sub

Private Sub cmdIr_Click()
    chbDeveGerar.value = True
    Me.Hide
End Sub

Private Sub Label11_Click()
    txtComEletronicas.SetFocus
End Sub

Private Sub Label19_Click()
    txtComPostais.SetFocus
End Sub

Private Sub Label20_Click()
    txtComMandados.SetFocus
End Sub

Private Sub Label21_Click()
    txtDigitalizacoes.SetFocus
End Sub

Private Sub Label22_Click()
    txtLitisconsortes.SetFocus
End Sub

Private Sub Label23_Click()
    txtCalculos.SetFocus
End Sub

Private Sub Label24_Click()
    txtPenhoras.SetFocus
End Sub

Private Sub Label25_Click()
    txtPrecatorias.SetFocus
End Sub

Private Sub Label7_Click()
    txtValor.SetFocus
End Sub

Private Sub txtCalculos_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtCalculos.text = Replace(txtCalculos.text, " ", "")
    txtCalculos.text = Replace(txtCalculos.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtCalculos.text = "" Then txtCalculos.text = "0"
    If Not IsNumeric(txtCalculos.text) Then GoTo ErroNoValor
        
    Label23.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de cálculos inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label23.BackColor = 8421631
    
End Sub

Private Sub txtComEletronicas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtComEletronicas.text = Replace(txtComEletronicas.text, " ", "")
    txtComEletronicas.text = Replace(txtComEletronicas.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtComEletronicas.text = "" Then txtComEletronicas.text = "0"
    If Not IsNumeric(txtComEletronicas.text) Then GoTo ErroNoValor
        
    Label11.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de comunicações eletrônicas inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label11.BackColor = 8421631
    
End Sub

Private Sub txtComMandados_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtComMandados.text = Replace(txtComMandados.text, " ", "")
    txtComMandados.text = Replace(txtComMandados.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtComMandados.text = "" Then txtComMandados.text = "0"
    If Not IsNumeric(txtComMandados.text) Then GoTo ErroNoValor
    
    Label20.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de comunicações por mandado inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label20.BackColor = 8421631
    
End Sub

Private Sub txtComPostais_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtComPostais.text = Replace(txtComPostais.text, " ", "")
    txtComPostais.text = Replace(txtComPostais.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtComPostais.text = "" Then txtComPostais.text = "0"
    If Not IsNumeric(txtComPostais.text) Then GoTo ErroNoValor
    
    Label19.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de comunicações postais inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label19.BackColor = 8421631
    
End Sub

Private Sub txtDigitalizacoes_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtDigitalizacoes.text = Replace(txtDigitalizacoes.text, " ", "")
    txtDigitalizacoes.text = Replace(txtDigitalizacoes.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtDigitalizacoes.text = "" Then txtDigitalizacoes.text = "0"
    If Not IsNumeric(txtDigitalizacoes.text) Then GoTo ErroNoValor
    
    Label21.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de digitalizações inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label21.BackColor = 8421631
    
End Sub

Private Sub txtLitisconsortes_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtLitisconsortes.text = Replace(txtLitisconsortes.text, " ", "")
    txtLitisconsortes.text = Replace(txtLitisconsortes.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtLitisconsortes.text = "" Then txtLitisconsortes.text = "0"
    If Not IsNumeric(txtLitisconsortes.text) Then GoTo ErroNoValor
    
    Label22.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de litisconsortes adicionais inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label22.BackColor = 8421631
    
End Sub

Private Sub txtPenhoras_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtPenhoras.text = Replace(txtPenhoras.text, " ", "")
    txtPenhoras.text = Replace(txtPenhoras.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtPenhoras.text = "" Then txtPenhoras.text = "0"
    If Not IsNumeric(txtPenhoras.text) Then GoTo ErroNoValor
    
    Label24.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de penhoras inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label24.BackColor = 8421631
    
End Sub

Private Sub txtPrecatorias_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtPrecatorias.text = Replace(txtPrecatorias.text, " ", "")
    txtPrecatorias.text = Replace(txtPrecatorias.text, ".", "")
    
    On Error GoTo ErroNoValor
    If txtPrecatorias.text = "" Then txtPrecatorias.text = "0"
    If Not IsNumeric(txtPrecatorias.text) Then GoTo ErroNoValor
    
    Label25.BackColor = -2147483643
    
    Exit Sub
    
ErroNoValor:
    MsgBox DeterminarTratamento & ", a quantidade de precatórias inserida parece não ser um número. Digite um número para continuar!", _
                vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
    Label25.BackColor = 8421631
    
End Sub

Private Sub txtValor_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If txtValor.text = "" Or Not IsNumeric(txtValor.text) Then
        MsgBox DeterminarTratamento & ", o valor da causa/condenação inserido parece não ser um número. Digite um número para continuar!", _
                    vbCritical + vbOKOnly, "Sísifo - Erro na quantidade de atos"
        Label7.BackColor = 8421631
    End If
    
End Sub

Private Sub txtValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Dim strCont As String
    
    ' Só permite digitar os números de 0 a 9 ou vírgula, que adiciona cinco zeros.
    If KeyAscii <> 44 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
    
    strCont = Replace(txtValor.text, ".", "")
    strCont = Replace(strCont, ",", "")
    strCont = IIf(KeyAscii <> 44, strCont & Chr(KeyAscii), strCont & "00000")
    strCont = "00" & strCont
    strCont = Left(strCont, Len(strCont) - 2) & "," & Right(strCont, 2)
    
    txtValor.text = Format(strCont, "#,##0.00")
    Label7.BackColor = -2147483643
    KeyAscii = 0
    
End Sub

Private Sub UserForm_Initialize()

    txtValor.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
    chbDeveGerar.value = False
    Me.Hide
End Sub


