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

    Public Shared Function CalcTensaoT(normalTracao As Double, baseX As Double, baseY As Double, c As Double, diametro As Double, d As Double, bw As Double, bf1 As Double, altura As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String) As Tracao
        Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tracao.TensaoTracao = ((normalTracao) / (PropriedadesGeometricas.area)) 'OK
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area))) 'OK
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area))) 'OK

            Case Madeira.TipoSecao.Circular
                tracao.TensaoTracao = normalTracao / PropriedadesGeometricas.area
                tracao.EsbeltezTracaoX = (diametro / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                tracao.EsbeltezTracaoY = (altura / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoT
                tracao.TensaoTracao = normalTracao / PropriedadesGeometricas.area
                tracao.EsbeltezTracaoX = (bw / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                tracao.EsbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoI
                tracao.TensaoTracao = normalTracao / PropriedadesGeometricas.area
                tracao.EsbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                tracao.EsbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.Caixao
                tracao.TensaoTracao = normalTracao / PropriedadesGeometricas.area
                tracao.EsbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                tracao.EsbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'botar o espaçador aqui tb
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then

                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then

                End If
        End Select

        Return tracao
    End Function

End Class
