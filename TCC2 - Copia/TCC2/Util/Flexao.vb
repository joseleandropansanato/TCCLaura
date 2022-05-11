Public Class Flexao

    Public Shared tensaoFlexaoX As Double = 0
	Public Shared tensaoFlexaoY As Double = 0


    Public Function CalcTensaoFlexaoSimples(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Double
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2

            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return CalcTensaoFlexaoSimples
    End Function

    Public Function CalcTensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Double
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2


            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return CalcTensaoFlexaoObliqua
    End Function


    Public Function CalcTensaoFlexotracao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Double
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                tensaoFlexaoX = 0
                tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                tensaoFlexaoX = 0
                tensaoFlexaoY = 0

        End Select

        Return CalcTensaoFlexotracao
    End Function

    Public Function CalcTensaoFlexocompressao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Double
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                tensaoFlexaoX = 0
                tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3

                tensaoFlexaoX = 0
                tensaoFlexaoY = 0
        End Select

        Return CalcTensaoFlexocompressao
    End Function



End Class
