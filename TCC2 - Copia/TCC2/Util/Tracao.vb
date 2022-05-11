Public Class Tracao

	Public Shared tensaoTracao As Double = 0
	Public Shared esbeltezTracaoX As Double = 0
	Public Shared esbeltezTracaoY As Double = 0
	Public Shared eixoXPecaComposta As Double = 0
	Public Shared eixoYPecaComposta As Double = 0

    Public Function TensaoTracao(normalTracao As Double, baseX As Double, baseY As Double, c As Double, diametro As Double, d As Double, bw As Double, bf1 As Double, altura As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String, lvinculado As Double, momentoFletorX As Double, momentoFletorY As Double, h1 As Double) As Tracao
        Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tracao.tensaoTracao = ((normalTracao) / (propriedadesGeometricas.Area)) 'OK
                Tracao.esbeltezTracaoX = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area))) 'OK
                Tracao.esbeltezTracaoY = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area))) 'OK

            Case Madeira.TipoSecao.Circular
                tracao.tensaoTracao = normalTracao / propriedadesGeometricas.Area
                Tracao.esbeltezTracaoX = (diametro / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area)))
                Tracao.esbeltezTracaoY = (altura / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.SecaoT
                tracao.tensaoTracao = normalTracao / propriedadesGeometricas.Area
                Tracao.esbeltezTracaoX = (bw / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.SecaoI
                tracao.tensaoTracao = normalTracao / propriedadesGeometricas.Area
                Tracao.esbeltezTracaoX = (bf1 / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.Caixao
                tracao.tensaoTracao = normalTracao / propriedadesGeometricas.Area
                Tracao.esbeltezTracaoX = (bf1 / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'botar o espaçador aqui tb
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    Tracao.esbeltezTracaoX = ((lvinculado) / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area)))
                    Tracao.esbeltezTracaoY = ((lvinculado) / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area)))
                    'eixo X::
                    Tracao.eixoXPecaComposta = (normalTracao / propriedadesGeometricas.Area) + (momentoFletorX / propriedadesGeometricas.EixoXmi) * (h1 / 2)
                    'eixo y::
                    Tracao.eixoYPecaComposta = (normalTracao / propriedadesGeometricas.Area) + ((momentoFletorY * propriedadesGeometricas.EixoYmi1) / propriedadesGeometricas.EixoYie * propriedadesGeometricas.EixoYmr) + (momentoFletorY / (2 * propriedadesGeometricas.Area * propriedadesGeometricas.AreaA1)) * (1 - 2 * (propriedadesGeometricas.EixoYmi1 / propriedadesGeometricas.EixoYie))
                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    Tracao.esbeltezTracaoX = (lvinculado) / propriedadesGeometricas.EixoXrg
                    Tracao.esbeltezTracaoY = (lvinculado) / (Math.Sqrt(propriedadesGeometricas.EixoYie / propriedadesGeometricas.Area))
                    'eixo X::
                    Tracao.eixoXPecaComposta = (normalTracao / propriedadesGeometricas.Area) + (momentoFletorX / propriedadesGeometricas.EixoXmi) * (h1 / 2)
                    'eixo y::
                    Tracao.eixoYPecaComposta = (normalTracao / propriedadesGeometricas.Area) + ((momentoFletorY * propriedadesGeometricas.EixoYmi1) / propriedadesGeometricas.EixoYie * propriedadesGeometricas.EixoYmr) + (momentoFletorY / (3 * propriedadesGeometricas.Area * propriedadesGeometricas.AreaA1)) * (1 - 2 * (propriedadesGeometricas.EixoYmi1 / propriedadesGeometricas.EixoYie))
                End If
        End Select

        Return tracao
    End Function

End Class
