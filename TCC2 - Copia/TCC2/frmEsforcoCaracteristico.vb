Public Class frmEsforcoCaracteristico
    Public formFatorComb As FormFatorComb = New FormFatorComb
    Public Shared f As Double = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.momentoCargaPermanenteX = PropriedadesResistencia.verificaVazio(txtMomentoCargaPermanenteX.Text)
        Form1.momentoCargaPermanenteY = PropriedadesResistencia.verificaVazio(txtMomentoCargaPermanenteY.Text)
        Form1.momentoCargaVariavelX = PropriedadesResistencia.verificaVazio(txtMomentoCargaVariavelX.Text)
        Form1.momentoCargaVariavelY = PropriedadesResistencia.verificaVazio(txtMomentoCargaVariavelY.Text)
        Form1.normalCargaPermanente = PropriedadesResistencia.verificaVazio(txtNormalCargaPermanente.Text)
        Form1.normalCargaVariavel = PropriedadesResistencia.verificaVazio(txtNormalCargaVariavel.Text)
        Form1.f1 = PropriedadesResistencia.verificaVazio(txtF1.Text)
        Form1.f2 = PropriedadesResistencia.verificaVazio(txtF2.Text)

        Form1.voltou = True
        Close()
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        MessageBox.Show("Para peças esbeltas sujeitas à compressão, os esforços normais e
                        momentos fletores, devem ser inseridos em seu valor característico
                        devidos às cargas permanentes e variáveis, ou seja, sem as suas 
                        devidas combinações normativas")
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        FormFatorComb.Show()
    End Sub

    Private Sub txtF1_TextChanged(sender As Object, e As EventArgs) Handles txtF1.TextChanged
        txtF1F2.Text = verificaVazio(txtF1.Text) + verificaVazio(txtF2.Text)
    End Sub

    Private Sub txtF2_TextChanged(sender As Object, e As EventArgs) Handles txtF2.TextChanged
        txtF1F2.Text = verificaVazio(txtF1.Text) + verificaVazio(txtF2.Text)
    End Sub

    Private Sub txtF1_Validated(sender As Object, e As EventArgs) Handles txtF1.Validated
        If Validacao.ValidarDados(txtF1, Validacao.TipoValidacao.RealF1) Or
            f = Convert.ToDouble(txtF1.Text) Then
        End If
    End Sub

    Private Sub txtF2_Validated(sender As Object, e As EventArgs) Handles txtF2.Validated
        If Validacao.ValidarDados(txtF2, Validacao.TipoValidacao.RealF2) Then
            f = Convert.ToDouble(txtF2.Text)
        End If
    End Sub

End Class