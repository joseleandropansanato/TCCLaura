Public Class Cisalhamento

    'Public Shared tensaoCisalhanteX As Double = 0
    'Public Shared tensaoCisalhanteY As Double = 0
    Private _tensaoCisalhanteX As Double = 0
    Private _tensaoCisalhanteY As Double = 0

    Public Sub New()
    End Sub

    Public Property TensaoCisalhanteX() As Double
        Get
            Return _tensaoCisalhanteX
        End Get
        Set(value As Double)
            _tensaoCisalhanteX = value
        End Set
    End Property

    Public Property TensaoCisalhanteY() As Double
        Get
            Return _tensaoCisalhanteY
        End Get
        Set(value As Double)
            _tensaoCisalhanteY = value
        End Set
    End Property

    Public Shared Function CalcTensaoCisalhamento(forcaCortanteX As Double, forcaCortanteY As Double, diametro As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Cisalhamento
        Dim cisalh = New Cisalhamento()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                cisalh.tensaoCisalhanteX = ((2 / 3) * (forcaCortanteX / propriedadesGeometricas.Area) / 10 ^ 6)
                cisalh.tensaoCisalhanteY = ((2 / 3) * (forcaCortanteY / propriedadesGeometricas.Area) / 10 ^ 6)

            Case Madeira.TipoSecao.Circular
                cisalh.tensaoCisalhanteX = ((4 * forcaCortanteX) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6
                cisalh.tensaoCisalhanteY = ((4 * forcaCortanteY) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                cisalh.tensaoCisalhanteX = forcaCortanteX * propriedadesGeometricas.Qx / propriedadesGeometricas.EixoXmi
                cisalh.tensaoCisalhanteY = forcaCortanteY * propriedadesGeometricas.Qy / propriedadesGeometricas.EixoYmi

            Case Madeira.TipoSecao.SecaoI
                cisalh.tensaoCisalhanteX = forcaCortanteX * propriedadesGeometricas.Qx / propriedadesGeometricas.EixoXmi
                cisalh.tensaoCisalhanteY = forcaCortanteY * propriedadesGeometricas.Qy / propriedadesGeometricas.EixoYmi

            Case Madeira.TipoSecao.Caixao
                cisalh.tensaoCisalhanteX = forcaCortanteX * propriedadesGeometricas.Qx / propriedadesGeometricas.EixoXmi
                cisalh.tensaoCisalhanteY = forcaCortanteY * propriedadesGeometricas.Qy / propriedadesGeometricas.EixoYmi

            Case Madeira.TipoSecao.ElementosJustaposto2
                cisalh.tensaoCisalhanteX = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                cisalh.tensaoCisalhanteX = 0
        End Select

        Return cisalh
    End Function

End Class
