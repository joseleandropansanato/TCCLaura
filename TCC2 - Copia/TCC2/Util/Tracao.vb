Public Class Tracao

    '    Public Shared tensaoTracao As Double = 0
    '    Public Shared esbeltezTracaoX As Double = 0
    '    Public Shared esbeltezTracaoY As Double = 0

    Private _tensaoTracao As Double = 0
    Private _esbeltezTracaoX As Double = 0
    Private _esbeltezTracaoY As Double = 0

    Public Sub New()

    End Sub

    Public Property TensaoTracao() As Double
        Get
            Return _tensaoTracao
        End Get
        Set(value As Double)
            _tensaoTracao = value
        End Set
    End Property

    Public Property EsbeltezTracaoX() As Double
        Get
            Return _esbeltezTracaoX
        End Get
        Set(value As Double)
            _esbeltezTracaoX = value
        End Set
    End Property

    Public Property EsbeltezTracaoY() As Double
        Get
            Return _esbeltezTracaoY
        End Get
        Set(value As Double)
            _esbeltezTracaoY = value
        End Set
    End Property

    Public Shared Function CalcTensaoT(normalTracao As Double, c As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String) As Tracao
        Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tracao.TensaoTracao = ((normalTracao) / (propriedadesGeometricas.Area)) 'OK
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area))) 'OK
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area))) 'OK

            Case Madeira.TipoSecao.Circular
                tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                tracao.EsbeltezTracaoX = (l * 100 / (Math.Sqrt(propriedadesGeometricas.EixoXmi / propriedadesGeometricas.Area)))
                tracao.EsbeltezTracaoY = (l * 100 / (Math.Sqrt(propriedadesGeometricas.EixoYmi / propriedadesGeometricas.Area)))


            Case Madeira.TipoSecao.SecaoT
                tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoXie / propriedadesGeometricas.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoYie / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.SecaoI
                tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoXie / propriedadesGeometricas.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoYie / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.Caixao
                tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoXie / propriedadesGeometricas.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(propriedadesGeometricas.EixoYie / propriedadesGeometricas.Area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'botar o espaçador aqui tb
                'acho que nao precisa disso, pq na vdd o pilar nao sofra forças axiais de tração??
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                End If
        End Select

        Return tracao
    End Function

End Class
