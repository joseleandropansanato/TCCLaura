Public Class Cisalhamento

	Public Shared tensaoCisalhanteX As Double = 0
	Public Shared tensaoCisalhanteY As Double = 0


    Public Function CalcTensaoCisalhamento(forcaCortanteX As Double, forcaCortanteY As Double, diametro As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Double
        Dim cisalhamento = New Cisalhamento()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tensaoCisalhanteX = ((2 / 3) * (forcaCortanteX / PropriedadesGeometricas.area) / 10 ^ 6)
                tensaoCisalhanteY = ((2 / 3) * (forcaCortanteY / PropriedadesGeometricas.area) / 10 ^ 6)

            Case Madeira.TipoSecao.Circular
                tensaoCisalhanteX = ((4 * forcaCortanteX) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6
                tensaoCisalhanteY = ((4 * forcaCortanteY) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                tensaoCisalhanteX = forcaCortanteX * PropriedadesGeometricas.Qx / PropriedadesGeometricas.eixoXmi
                tensaoCisalhanteY = forcaCortanteY * PropriedadesGeometricas.Qy / PropriedadesGeometricas.eixoYmi

            Case Madeira.TipoSecao.SecaoI
                tensaoCisalhanteX = forcaCortanteX * PropriedadesGeometricas.Qx / PropriedadesGeometricas.eixoXmi
                tensaoCisalhanteY = forcaCortanteY * PropriedadesGeometricas.Qy / PropriedadesGeometricas.eixoYmi

            Case Madeira.TipoSecao.Caixao
                tensaoCisalhanteX = forcaCortanteX * PropriedadesGeometricas.Qx / PropriedadesGeometricas.eixoXmi
                tensaoCisalhanteY = forcaCortanteY * PropriedadesGeometricas.Qy / PropriedadesGeometricas.eixoYmi

            Case Madeira.TipoSecao.ElementosJustaposto2
                tensaoCisalhanteX = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                tensaoCisalhanteX = 0
        End Select

        Return CalcTensaoCisalhamento
    End Function

End Class
