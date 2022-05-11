Public Class Cisalhamento

	Public Shared tensaoCisalhanteX As Double = 0
	Public Shared tensaoCisalhanteY As Double = 0


    Public Function TensaoCisalhamento(forcaCortanteX As Double, forcaCortanteY As Double, diametro As Double, q As Qxy, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Cisalhamento
        Dim cisalhamento = New Cisalhamento()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Cisalhamento.tensaoCisalhanteX = ((2 / 3) * (forcaCortanteX / propriedadesGeometricas.Area) / 10 ^ 6)
                Cisalhamento.tensaoCisalhanteY = ((2 / 3) * (forcaCortanteY / propriedadesGeometricas.Area) / 10 ^ 6)

            Case Madeira.TipoSecao.Circular
                Cisalhamento.tensaoCisalhanteX = ((4 * forcaCortanteX) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6
                Cisalhamento.tensaoCisalhanteY = ((4 * forcaCortanteY) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Cisalhamento.tensaoCisalhanteX = forcaCortanteX * q.Qx / propriedadesGeometricas.EixoXmi
                Cisalhamento.tensaoCisalhanteY = forcaCortanteY * q.Qy / propriedadesGeometricas.EixoYmi

            Case Madeira.TipoSecao.SecaoI
                Cisalhamento.tensaoCisalhanteX = forcaCortanteX * q.Qx / propriedadesGeometricas.EixoXmi
                Cisalhamento.tensaoCisalhanteY = forcaCortanteY * q.Qy / propriedadesGeometricas.EixoYmi

            Case Madeira.TipoSecao.Caixao
                Cisalhamento.tensaoCisalhanteX = forcaCortanteX * q.Qx / propriedadesGeometricas.EixoXmi
                Cisalhamento.tensaoCisalhanteY = forcaCortanteY * q.Qy / propriedadesGeometricas.EixoYmi

            Case Madeira.TipoSecao.ElementosJustaposto2
                Cisalhamento.tensaoCisalhanteX = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                Cisalhamento.tensaoCisalhanteX = 0
        End Select

        Return cisalhamento
    End Function

End Class
