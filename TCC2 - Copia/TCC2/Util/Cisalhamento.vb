Public Class Cisalhamento

    'Public Shared tensaoCisalhanteX As Double = 0
    'Public Shared tensaoCisalhanteY As Double = 0
    Private _tensaoCisalhanteX As Double = 0
    Private _tensaoCisalhanteY As Double = 0
    Private _vd As Double = 0 'CISALHAMENTO NO ESPAÇADOR

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

    Public Property Vd() As Double
        Get
            Return _vd
        End Get
        Set(value As Double)
            _vd = value
        End Set
    End Property

    Public Shared Function CalcTensaoCisalhamento(forcaCortanteX As Double, forcaCortanteY As Double, diametro As Double, a1 As Double, L1 As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Cisalhamento
        Dim cisalh = New Cisalhamento()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                cisalh.TensaoCisalhanteX = ((2 / 3) * (forcaCortanteX / propriedadesGeometricas.Area))
                cisalh.TensaoCisalhanteY = ((2 / 3) * (forcaCortanteY / propriedadesGeometricas.Area))

            Case Madeira.TipoSecao.Circular
                cisalh.TensaoCisalhanteX = ((4 * forcaCortanteX) / (3 * System.Math.PI * (diametro / 2)))
                cisalh.TensaoCisalhanteY = ((4 * forcaCortanteY) / (3 * System.Math.PI * (diametro / 2)))

            Case Madeira.TipoSecao.SecaoT
                cisalh.TensaoCisalhanteX = ((forcaCortanteX * propriedadesGeometricas.Qx) / propriedadesGeometricas.EixoXmi)
                cisalh.TensaoCisalhanteY = ((forcaCortanteY * propriedadesGeometricas.Qy) / propriedadesGeometricas.EixoYmi)

            Case Madeira.TipoSecao.SecaoI
                cisalh.TensaoCisalhanteX = ((forcaCortanteX * propriedadesGeometricas.Qx) / propriedadesGeometricas.EixoXmi)
                cisalh.TensaoCisalhanteY = ((forcaCortanteY * propriedadesGeometricas.Qy) / propriedadesGeometricas.EixoYmi)

            Case Madeira.TipoSecao.Caixao
                cisalh.TensaoCisalhanteX = ((forcaCortanteX * propriedadesGeometricas.Qx) / propriedadesGeometricas.EixoXmi)
                cisalh.TensaoCisalhanteY = ((forcaCortanteY * propriedadesGeometricas.Qy) / propriedadesGeometricas.EixoYmi)

                'PARA SEÇÃO COMPOSTA POR ESPAÇADORES A GNT SÓ CONFERE O CISALHAMENTO FRENTE DO ESPAÇADOR
            Case Madeira.TipoSecao.ElementosJustaposto2
                'cisalh.Vd = (propriedadesGeometricas.AreaA1 * PropriedadesResistencia.resistCalAoCisalhamento * (L1 * a1))/100

            Case Madeira.TipoSecao.ElementosJustaposto3
                'cisalh.Vd = (propriedadesGeometricas.AreaA1 * PropriedadesResistencia.resistCalAoCisalhamento * (L1 * a1))/100
        End Select

        Return cisalh
    End Function

End Class
