Public Class Flexao

    'Public Shared tensaoFlexaoX As Double = 0
    'Public Shared tensaoFlexaoY As Double = 0
    Private _tensaoFlexaoX As Double = 0
    Private _tensaoFlexaoY As Double = 0
    Private _tensaoFlexaoXT As Double = 0 'para flexotração
    Private _tensaoFlexaoYT As Double = 0 'para flexotração
    Private _tensaoFlexaoXC As Double = 0 'para flexotração
    Private _tensaoFlexaoYC As Double = 0 'para flexotração
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
    Public Property tensaoFlexaoXT() As Double
        Get
            Return _tensaoFlexaoXT
        End Get
        Set(value As Double)
            _tensaoFlexaoXT = value
        End Set
    End Property
    Public Property tensaoFlexaoYT() As Double
        Get
            Return _tensaoFlexaoYT
        End Get
        Set(value As Double)
            _tensaoFlexaoYT = value
        End Set
    End Property
    Public Property tensaoFlexaoXC() As Double
        Get
            Return _tensaoFlexaoXC
        End Get
        Set(value As Double)
            _tensaoFlexaoXC = value
        End Set
    End Property
    Public Property tensaoFlexaoYC() As Double
        Get
            Return _tensaoFlexaoYC
        End Get
        Set(value As Double)
            _tensaoFlexaoYC = value
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
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoCX = (momentoFletorY * 100) / proprGeo.EixoXmr
                    flex.tensaoTX = (momentoFletorY * 100) / proprGeo.EixoXmr

                    flex.tensaoCY = (momentoFletorX * 100) / proprGeo.EixoXmr
                    flex.tensaoTY = (momentoFletorX * 100) / proprGeo.EixoXmr

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoCX = -(momentoFletorY * 100) / proprGeo.EixoXmr
                    flex.tensaoTX = -(momentoFletorY * 100) / proprGeo.EixoXmr

                    flex.tensaoCY = -(momentoFletorX * 100) / proprGeo.EixoXmr
                    flex.tensaoTY = -(momentoFletorX * 100) / proprGeo.EixoXmr
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Circular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoCX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi

                    flex.tensaoCY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                    flex.tensaoTY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoCX = -((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoTX = -((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi

                    flex.tensaoCY = -((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                    flex.tensaoTY = -((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10


            Case Madeira.TipoSecao.SecaoT
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoCX = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                    flex.tensaoCY = ((momentoFletorX * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoTX = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoCX = -((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                    flex.tensaoTY = -((momentoFletorX * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoCY = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.SecaoI
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoCX = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                    flex.tensaoCY = ((momentoFletorX * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoTX = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoCX = -((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                    flex.tensaoTY = -((momentoFletorX * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoCY = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Caixao
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoCX = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                    flex.tensaoCY = ((momentoFletorX * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoTX = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoCX = -((momentoFletorY * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmi

                    flex.tensaoTY = -((momentoFletorX * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoCY = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYmie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto2
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoCX = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoYie

                    flex.tensaoCY = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoCX = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoTX = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoYie

                    flex.tensaoCY = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoXmi
                    flex.tensaoTY = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto3
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoCX = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoTX = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoYie

                    flex.tensaoCY = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoXmi
                    flex.tensaoTY = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoCX = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoTX = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoYie

                    flex.tensaoCY = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoXmi
                    flex.tensaoTY = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10
        End Select

        Return flex
    End Function

    Public Shared Function CalcTensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, fx As Double, fy As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoX = (momentoFletorX * 100) / proprGeo.EixoXmr
                    flex.tensaoFlexaoY = (momentoFletorY * 100) / proprGeo.EixoYmr

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoX = -(momentoFletorX * 100) / proprGeo.EixoXmr
                    flex.tensaoFlexaoY = -(momentoFletorY * 100) / proprGeo.EixoYmr
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Circular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoX = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoX = -((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = -((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.SecaoT
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoX = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoY = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoX = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.SecaoI
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoX = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoY = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoX = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Caixao
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoX = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoY = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoX = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto2
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoX = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoX = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto3
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoX = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoX = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoY = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10
        End Select

        Return flex
    End Function


    Public Shared Function CalcTensaoFlexotracao(flexaoSimplesOrObiqua As Flexao, momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, fx As Double, fy As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()
        flex = flexaoSimplesOrObiqua

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * (b / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * (h / 2)) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoXT = -((momentoFletorY * 100) * (b / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = -((momentoFletorX * 100) * (h / 2)) / proprGeo.EixoYmi

                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Circular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoXT = -((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = -((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10


            Case Madeira.TipoSecao.SecaoT
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoYT = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoXT = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.SecaoI
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoYT = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoXT = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Caixao
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoYT = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoXT = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto2
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoXT = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto3
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXT = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoXT = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYT = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

        End Select
        Return flex
    End Function

    Public Shared Function CalcTensaoFlexocompressao(flexaoSimplesOrObiqua As Flexao, momentoFletorX As Double, momentoFletorY As Double, b As Double, h As Double, d As Double, fx As Double, fy As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flex = New Flexao()
        flex = flexaoSimplesOrObiqua

        Select Case tipoSecao

            Case Madeira.TipoSecao.Retangular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * (b / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * (h / 2)) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoXC = -((momentoFletorY * 100) * (b / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = -((momentoFletorX * 100) * (h / 2)) / proprGeo.EixoYmi

                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Circular
                If momentoFletorX > 0 Or momentoFletorY > 0 Then

                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoXC = -((momentoFletorY * 100) * (d / 2)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = -((momentoFletorX * 100) * (d / 2)) / proprGeo.EixoYmi
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.SecaoT
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoYC = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoXC = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.SecaoI
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoYC = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1 + PropriedadesGeometricas.h2 + PropriedadesGeometricas.h3) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoXC = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.Caixao
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then

                    flex.tensaoFlexaoYC = -((momentoFletorY * 100) * ((PropriedadesGeometricas.h1c + PropriedadesGeometricas.h2c + PropriedadesGeometricas.h4c) - PropriedadesGeometricas.centroGeometricoY)) / proprGeo.EixoXmi
                    flex.tensaoFlexaoXC = -((momentoFletorX * 100) * PropriedadesGeometricas.centroGeometricoY) / proprGeo.EixoYie
                End If

                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto2
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoXC = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10

            Case Madeira.TipoSecao.ElementosJustaposto3
                If momentoFletorX > 0 Or momentoFletorY > 0 Then
                    flex.tensaoFlexaoXC = ((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = ((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie

                ElseIf momentoFletorX < 0 Or momentoFletorY < 0 Then
                    flex.tensaoFlexaoXC = -((momentoFletorY * 100) * (PropriedadesGeometricas.a1 + (PropriedadesGeometricas.b1 / 2))) / proprGeo.EixoXmi
                    flex.tensaoFlexaoYC = -((momentoFletorX * 100) * (PropriedadesGeometricas.h1 / 2)) / proprGeo.EixoYie
                End If
                flex.ApoioX = (fx / proprGeo.Area) * 10
                flex.ApoioY = (fy / proprGeo.Area) * 10
        End Select

        Return flex
    End Function

End Class
