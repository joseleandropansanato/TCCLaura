Public Class Flexao

    'Public Shared tensaoFlexaoX As Double = 0
    'Public Shared tensaoFlexaoY As Double = 0
    Private _tensaoFlexaoX As Double = 0
    Private _tensaoFlexaoY As Double = 0

    Public Sub New()
    End Sub

    Public Property tensaoFlexaoX() As Double
        Get
            Return _tensaoFlexaoX
        End Get
        Set(value As Double)
            _tensaoFlexaoX = value
        End Set
    End Property

    Public Property tensaoFlexaoY() As Double
        Get
            Return _tensaoFlexaoY
        End Get
        Set(value As Double)
            _tensaoFlexaoY = value
        End Set
    End Property

    Public Shared Function CalcTensaoFlexaoSimples(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2

            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flex
    End Function

    Public Shared Function CalcTensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2


            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flex
    End Function


    Public Shared Function CalcTensaoFlexotracao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                flex.tensaoFlexaoX = 0
                flex.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                flex.tensaoFlexaoX = 0
                flex.tensaoFlexaoY = 0

        End Select

        Return flex
    End Function

    Public Shared Function CalcTensaoFlexocompressao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = (momentoFletorX / propriedadesGeometricas.EixoXmr) / 10 ^ 6
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                flex.tensaoFlexaoX = 0
                flex.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3

                flex.tensaoFlexaoX = 0
                flex.tensaoFlexaoY = 0
        End Select

        Return flex
    End Function

End Class
