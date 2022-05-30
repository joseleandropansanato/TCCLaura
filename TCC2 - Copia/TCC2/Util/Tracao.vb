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

    Public Shared Function CalcTensaoT(normalTracao As Double, c As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, b1 As Double, a As Double, elementoFixacaoSelecionado As String) As Tracao
        Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tracao.TensaoTracao = (((normalTracao) / (proprGeo.Area)) * 10)
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area)))

            Case Madeira.TipoSecao.Circular
                tracao.TensaoTracao = ((normalTracao / proprGeo.Area) * 10)
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area)))


            Case Madeira.TipoSecao.SecaoT
                tracao.TensaoTracao = ((normalTracao / proprGeo.Area) * 10)
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(proprGeo.EixoXie / proprGeo.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(proprGeo.EixoYie / proprGeo.Area)))

            Case Madeira.TipoSecao.SecaoI
                tracao.TensaoTracao = ((normalTracao / proprGeo.Area) * 10)
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(proprGeo.EixoXie / proprGeo.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(proprGeo.EixoYie / proprGeo.Area)))

            Case Madeira.TipoSecao.Caixao
                tracao.TensaoTracao = ((normalTracao / proprGeo.Area) * 10)
                tracao.EsbeltezTracaoX = (c * 100 / (Math.Sqrt(proprGeo.EixoXie / proprGeo.Area)))
                tracao.EsbeltezTracaoY = (c * 100 / (Math.Sqrt(proprGeo.EixoYie / proprGeo.Area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'acho que nao precisa disso, pq na vdd o pilar nao sofra forças axiais de tração??
                If ElementoJustaposto(c, b1, a, elementoFixacaoSelecionado) Then
                    tracao.TensaoTracao = ((normalTracao / proprGeo.Area) * 10)
                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(c, b1, a, elementoFixacaoSelecionado) Then
                    tracao.TensaoTracao = ((normalTracao / proprGeo.Area) * 10)
                End If
        End Select

        Return tracao
    End Function

End Class
