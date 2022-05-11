Public Class Tracao

	Public Shared tensaoTracao As Double = 0
	Public Shared esbeltezTracaoX As Double = 0
	Public Shared esbeltezTracaoY As Double = 0
	Public Shared eixoXPecaComposta As Double = 0
	Public Shared eixoYPecaComposta As Double = 0

    Public Function Tensao(normalTracao As Double, baseX As Double, baseY As Double, c As Double, diametro As Double, d As Double, bw As Double, bf1 As Double, altura As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String, lvinculado As Double, momentoFletorX As Double, momentoFletorY As Double, h1 As Double) As Tracao
        Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Tracao.tensaoTracao = ((normalTracao) / (PropriedadesGeometricas.area)) 'OK
                Tracao.esbeltezTracaoX = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area))) 'OK
                Tracao.esbeltezTracaoY = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area))) 'OK

            Case Madeira.TipoSecao.Circular
                Tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (diametro / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (altura / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoT
                Tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (bw / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoI
                Tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.Caixao
                Tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'botar o espaçador aqui tb
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    Tracao.esbeltezTracaoX = ((lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                    Tracao.esbeltezTracaoY = ((lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))
                    'eixo X::
                    Tracao.eixoXPecaComposta = (normalTracao / PropriedadesGeometricas.area) + (momentoFletorX / PropriedadesGeometricas.eixoXmi) * (h1 / 2)
                    'eixo y::
                    Tracao.eixoYPecaComposta = (normalTracao / PropriedadesGeometricas.area) + ((momentoFletorY * PropriedadesGeometricas.eixoYmi1) / PropriedadesGeometricas.eixoYie * PropriedadesGeometricas.eixoYmr) + (momentoFletorY / (2 * PropriedadesGeometricas.area * PropriedadesGeometricas.areaA1)) * (1 - 2 * (PropriedadesGeometricas.eixoYmi1 / PropriedadesGeometricas.eixoYie))
                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    Tracao.esbeltezTracaoX = (lvinculado) / PropriedadesGeometricas.eixoXrg
                    Tracao.esbeltezTracaoY = (lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoYie / PropriedadesGeometricas.area))
                    'eixo X::
                    Tracao.eixoXPecaComposta = (normalTracao / PropriedadesGeometricas.area) + (momentoFletorX / PropriedadesGeometricas.eixoXmi) * (h1 / 2)
                    'eixo y::
                    Tracao.eixoYPecaComposta = (normalTracao / PropriedadesGeometricas.area) + ((momentoFletorY * PropriedadesGeometricas.eixoYmi1) / PropriedadesGeometricas.eixoYie * PropriedadesGeometricas.eixoYmr) + (momentoFletorY / (3 * PropriedadesGeometricas.area * PropriedadesGeometricas.areaA1)) * (1 - 2 * (PropriedadesGeometricas.eixoYmi1 / PropriedadesGeometricas.eixoYie))
                End If
        End Select

        Return tracao
    End Function

End Class
