Public Class Flexao

    'Public Shared tensaoFlexaoX As Double = 0
    'Public Shared tensaoFlexaoY As Double = 0
    Private _tensaoFlexaoX As Double = 0
    Private _tensaoFlexaoY As Double = 0
    Private _tensaoCX As Double = 0
    Private _tensaoCY As Double = 0
    Private _tensaoTX As Double = 0
    Private _tensaoTY As Double = 0

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
    Public Property tensaoCX() As Double
        Get
            Return _tensaoCX
        End Get
        Set(value As Double)
            _tensaoCX = value
        End Set
    End Property

    Public Property tensaoCY() As Double
        Get
            Return _tensaoCY
        End Get
        Set(value As Double)
            _tensaoCY = value
        End Set
    End Property
    Public Property tensaoTX() As Double
        Get
            Return _tensaoTX
        End Get
        Set(value As Double)
            _tensaoTX = value
        End Set
    End Property

    Public Property tensaoTY() As Double
        Get
            Return _tensaoTY
        End Get
        Set(value As Double)
            _tensaoTY = value
        End Set
    End Property

    Public Shared Function CalcTensaoFlexaoSimples(momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, hd As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoCX = ((momentoFletorX * 100) / (b / 2))
                flex.tensaoCY = ((momentoFletorY * 100) / (h / 2))
                flex.tensaoTX = ((momentoFletorX * 100) / (b / 2))
                flex.tensaoTY = ((momentoFletorY * 100) / (h / 2))

            Case Madeira.TipoSecao.Circular
                flex.tensaoCX = ((momentoFletorX * 100) / propriedadesGeometricas.EixoXmr)
                flex.tensaoCY = ((momentoFletorY * 100) / propriedadesGeometricas.EixoYmr)
                flex.tensaoTX = ((momentoFletorX * 100) / propriedadesGeometricas.EixoXmr)
                flex.tensaoTY = ((momentoFletorY * 100) / propriedadesGeometricas.EixoYmr)

            Case Madeira.TipoSecao.SecaoT
                ' flex.tensaoCX = ((momentoFletorX * 100) / x_CGi
                'flex.tensaoCY = ((momentoFletorY * 100) / (hd - y_CGi)
                'flex.tensaoTX = ((momentoFletorX * 100) / x_CGi
               ' flex.tensaoTY = ((momentoFletorY * 100) / y_CGi

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoCX = ((momentoFletorX * 100) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoCY = ((momentoFletorY * 100) / propriedadesGeometricas.EixoYmr) / 100
                flex.tensaoTX = ((momentoFletorX * 100) / 1)
                flex.tensaoTY = ((momentoFletorY * 100) / 1)

            Case Madeira.TipoSecao.Caixao
                flex.tensaoCX = ((momentoFletorX * 100) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoCY = ((momentoFletorY * 100) / propriedadesGeometricas.EixoYmr) / 100
                flex.tensaoTX = ((momentoFletorX * 100) / 1)
                flex.tensaoTY = ((momentoFletorY * 100) / 1)

            Case Madeira.TipoSecao.ElementosJustaposto2

            Case Madeira.TipoSecao.ElementosJustaposto3
        End Select

        Return flex

    End Function

    Public Shared Function CalcTensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.ElementosJustaposto2


            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flex
    End Function


    Public Shared Function CalcTensaoFlexotracao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = (momentoFletorY / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = ((momentoFletorX / 10 ^ 4) / propriedadesGeometricas.EixoXmr) / 100
                flex.tensaoFlexaoY = ((momentoFletorY / 10 ^ 4) / propriedadesGeometricas.EixoYmr) / 100

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
