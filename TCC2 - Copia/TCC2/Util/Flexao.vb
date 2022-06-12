Public Class Flexao

    'Public Shared tensaoFlexaoX As Double = 0
    'Public Shared tensaoFlexaoY As Double = 0
    Private _tensaoFlexaoX As Double = 0
    Private _tensaoFlexaoY As Double = 0
    Private _tensaoCX As Double = 0
    Private _tensaoCY As Double = 0
    Private _tensaoTX As Double = 0
    Private _tensaoTY As Double = 0
    Private _apoioX As Double = 0
    Private _apoioY As Double = 0
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
    Public Property ApoioX() As Double
        Get
            Return _apoioX
        End Get
        Set(value As Double)
            _apoioX = value
        End Set
    End Property
    Public Property ApoioY() As Double
        Get
            Return _apoioY
        End Get
        Set(value As Double)
            _apoioY = value
        End Set
    End Property

    Public Shared Function CalcTensaoFlexaoSimples(momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, fx As Double, fy As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoCX = (momentoFletorY * 100) / proprGeo.EixoXmr
                flex.tensaoTX = (momentoFletorY * 100) / proprGeo.EixoXmr

                flex.tensaoCY = (momentoFletorX * 100) / proprGeo.EixoXmr
                flex.tensaoTY = (momentoFletorX * 100) / proprGeo.EixoXmr

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area

            Case Madeira.TipoSecao.Circular
                flex.tensaoCX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                flex.tensaoTX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi

                flex.tensaoCY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                flex.tensaoTY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area


            Case Madeira.TipoSecao.SecaoT
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoCX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

                    flex.tensaoCY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                End If


            Case Madeira.TipoSecao.SecaoI

                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoCX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

                    flex.tensaoCY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                End If


            Case Madeira.TipoSecao.Caixao

                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoCX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

                    flex.tensaoCY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                End If


                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area

            Case Madeira.TipoSecao.ElementosJustaposto2

            Case Madeira.TipoSecao.ElementosJustaposto3
        End Select

        Return flex

    End Function

    Public Shared Function CalcTensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, fx As Double, fy As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = (momentoFletorX * 100) / proprGeo.EixoXmr
                flex.tensaoFlexaoY = (momentoFletorY * 100) / proprGeo.EixoYmr

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area


            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

                flex.ApoioX = fx / proprGeo.Area
                flex.ApoioY = fy / proprGeo.Area

            Case Madeira.TipoSecao.ElementosJustaposto2


            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flex
    End Function


    Public Shared Function CalcTensaoFlexotracao(momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * (b / 2)) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * (h / 2)) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi


            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.ElementosJustaposto2
                flex.tensaoFlexaoX = 0
                flex.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                flex.tensaoFlexaoX = 0
                flex.tensaoFlexaoY = 0

        End Select

        Return flex
    End Function

    Public Shared Function CalcTensaoFlexocompressao(momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * (b / 2)) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * (h / 2)) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.Circular
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.SecaoT
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.SecaoI
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

            Case Madeira.TipoSecao.Caixao
                flex.tensaoFlexaoX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoXmi
                flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoX) / proprGeo.EixoYmi

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
