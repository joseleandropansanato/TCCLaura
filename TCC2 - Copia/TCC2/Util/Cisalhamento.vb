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

    Public Shared Function CalcTensaoCisalhamento(forcaCortanteX As Double, forcaCortanteY As Double, diametro As Double, a1 As Double, L1 As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Cisalhamento
        Dim cisalh = New Cisalhamento()

        'PELA REGRA DA MÃO DIREITA, O MOMNETO EM X CAUSA FLEXÃO NO EIXO Y E O MOMENTO EM Y NO EIXO X
        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                cisalh.TensaoCisalhanteX = (((3 / 2) * (forcaCortanteX / proprGeo.Area)) * 10)
                cisalh.TensaoCisalhanteY = (((3 / 2) * (forcaCortanteY / proprGeo.Area)) * 10)

            Case Madeira.TipoSecao.Circular
                cisalh.TensaoCisalhanteX = (((4 * forcaCortanteX) / (3 * System.Math.PI * (diametro / 2))) * 10)
                cisalh.TensaoCisalhanteY = (((4 * forcaCortanteY) / (3 * System.Math.PI * (diametro / 2))) * 10)


            Case Madeira.TipoSecao.SecaoT
                cisalh.TensaoCisalhanteX = (((forcaCortanteX * proprGeo.Qx) / proprGeo.EixoXmi) * 10)
                cisalh.TensaoCisalhanteY = (((forcaCortanteY * proprGeo.Qy) / proprGeo.EixoYmi) * 10)

            Case Madeira.TipoSecao.SecaoI
                cisalh.TensaoCisalhanteX = (((forcaCortanteX * proprGeo.Qx) / proprGeo.EixoXmi) * 10)
                cisalh.TensaoCisalhanteY = (((forcaCortanteY * proprGeo.Qy) / proprGeo.EixoYmi) * 10)

            Case Madeira.TipoSecao.Caixao
                cisalh.TensaoCisalhanteX = (((forcaCortanteX * proprGeo.Qx) / proprGeo.EixoXmi) * 10)
                cisalh.TensaoCisalhanteY = (((forcaCortanteY * proprGeo.Qy) / proprGeo.EixoYmi) * 10)

                'PARA SEÇÃO COMPOSTA POR ESPAÇADORES A GNT SÓ CONFERE O CISALHAMENTO FRENTE DO ESPAÇADOR
            Case Madeira.TipoSecao.ElementosJustaposto2

                'cisalh.Vd = (propriedadesGeometricas.AreaA1 * PropriedadesResistencia.resistCalAoCisalhamento * (L1 * a1))/100

            Case Madeira.TipoSecao.ElementosJustaposto3
                'cisalh.Vd = (propriedadesGeometricas.AreaA1 * PropriedadesResistencia.resistCalAoCisalhamento * (L1 * a1))/100
        End Select

        Return cisalh
    End Function

End Class
