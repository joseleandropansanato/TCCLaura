Public Class Tracao

    Public Shared tensaoTracao As Double = 0
    Public Shared esbeltezTracaoX As Double = 0
    Public Shared esbeltezTracaoY As Double = 0


    Public Function CalcTensaoT(normalTracao As Double, baseX As Double, baseY As Double, c As Double, diametro As Double, d As Double, bw As Double, bf1 As Double, altura As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String) As Double
        'Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tensaoTracao = ((normalTracao) / (PropriedadesGeometricas.area)) 'OK
                esbeltezTracaoX = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area))) 'OK
                esbeltezTracaoY = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area))) 'OK

            Case Madeira.TipoSecao.Circular
                tensaoTracao = normalTracao / PropriedadesGeometricas.area
                esbeltezTracaoX = (diametro / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                esbeltezTracaoY = (altura / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoT
                tensaoTracao = normalTracao / PropriedadesGeometricas.area
                esbeltezTracaoX = (bw / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoI
                tensaoTracao = normalTracao / PropriedadesGeometricas.area
                esbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.Caixao
                tensaoTracao = normalTracao / PropriedadesGeometricas.area
                esbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'botar o espaçador aqui tb
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then

                End If
        End Select

        Return CalcTensaoT
    End Function

End Class
