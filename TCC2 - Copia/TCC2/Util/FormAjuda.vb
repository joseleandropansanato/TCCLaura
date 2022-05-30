Public Class FormAjuda
    Private Sub btnInicio_Click(sender As Object, e As EventArgs) Handles btnInicio.Click
        gbxInicio.Visible = True
        gbxProprMad.Visible = False
        gbxResistCalc.Visible = False
        gbxDInici.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        gbxInicio.Visible = False
        gbxProprMad.Visible = True
        gbxResistCalc.Visible = False
        gbxDInici.Visible = False
        lblbemVindo.Visible = False
        lblao.Visible = False
        lblname.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        gbxInicio.Visible = False
        gbxProprMad.Visible = False
        gbxResistCalc.Visible = True
        gbxDInici.Visible = False
        lblbemVindo.Visible = False
        lblao.Visible = False
        lblname.Visible = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        gbxInicio.Visible = False
        gbxProprMad.Visible = False
        gbxResistCalc.Visible = False
        gbxDInici.Visible = True
        lblbemVindo.Visible = False
        lblao.Visible = False
        lblname.Visible = False

    End Sub

End Class