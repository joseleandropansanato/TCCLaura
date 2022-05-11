Public Class Flexao


	Public Shared tensaoFlexaoX As Double = 0
	Public Shared tensaoFlexaoY As Double = 0


    Public Function TensaoFlexaoSimples(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2

            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flexao
    End Function

    Public Function TensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2


            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flexao
    End Function


    Public Function TensaoFlexotracao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0

        End Select

        Return flexao
    End Function

    Public Function TensaoFlexocompressao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3

                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0
        End Select

        Return flexao
    End Function



End Class
