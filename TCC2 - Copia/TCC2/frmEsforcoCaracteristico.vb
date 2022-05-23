Public Class frmEsforcoCaracteristico

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.momentoCargaPermanenteX = PropriedadesResistencia.verificaVazio(txtMomentoCargaPermanenteX.Text)
        Dim frm1 As Form1 = New Form1()
        frm1.Retorno()
        Close()
    End Sub
End Class