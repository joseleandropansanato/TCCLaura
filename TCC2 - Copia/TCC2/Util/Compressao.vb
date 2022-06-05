Imports TCC2.frmEsforcoCaracteristico

Public Class Compressao
    'Public Shared tensaoCompressao As Double = 0
    'Public Shared tensaoCompressaoCurtaX As Double = 0
    'Public Shared tensaoCompressaoCurtaY As Double = 0
    'Public Shared tensaoCompressaoMedEsbeltaX As Double = 0
    'Public Shared tensaoCompressaoMedEsbeltaY As Double = 0
    'Public Shared tensaoCompressaoEsbeltaX As Double = 0
    'Public Shared tensaoCompressaoEsbeltaY As Double = 0
    'Public Shared tensaoMxd As Double = 0
    'Public Shared tensaoMyd As Double = 0
    'Public Shared esbeltezCompressaoX As Double = 0
    'Public Shared esbeltezCompressaoY As Double = 0
    'Public Shared forcaElasticaX As Double = 0
    'Public Shared forcaElasticaY As Double = 0
    'Public Shared lvinculado As Double = 0
    'Public Shared ei As Double = 0 'excentricidade inicial
    'Public Shared ea As Double = 0 'excentricidade acidental
    'Public Shared edx As Double = 0 'excentricidade de projeto em x 
    'Public Shared edy As Double = 0 'excentricidade de projeto em y
    'Public Shared ec As Double = 0 'excentricidade suplementar de primeira ordem
    'Public Shared e1ef As Double = 0 'excentricidade de primeira ordem
    'Public Shared eig As Double = 0 'excentricidade de primeira ordem devido as cargas permanentes
    'Public Shared eixoXPecaComposta As Double = 0
    'Public Shared eixoYPecaComposta As Double = 0

    Private _tensaoCompressaoX As Double = 0
    Private _tensaoCompressaoY As Double = 0

    Private _tensaoMxd As Double = 0
    Private _tensaoMyd As Double = 0

    Private _esbeltezCompressaoX As Double = 0
    Private _esbeltezCompressaoY As Double = 0

    Private _forcaElasticaX As Double = 0
    Private _forcaElasticaY As Double = 0

    Private _lvinculado As Double = 0

    Private _eix As Double = 0 'excentricidade inicial em x 
    Private _eiy As Double = 0 'excentricidade inicial em y


    Private _eaX As Double = 0 'excentricidade acidental
    Private _eaY As Double = 0 'excentricidade acidental

    Private _edx As Double = 0 'excentricidade de projeto em x 
    Private _edy As Double = 0 'excentricidade de projeto em y

    Private _ecx As Double = 0 'excentricidade suplementar de primeira ordem
    Private _ecy As Double = 0 'excentricidade suplementar de primeira ordem

    Private _e1ef As Double = 0 'excentricidade de primeira ordem
    Private _eig As Double = 0 'excentricidade de primeira ordem devido as cargas permanentes

    Private _eixoYY As Double = 0
    Private _eixoXX As Double = 0
    Public Sub New()

    End Sub
    Public Property TensaoCompressaoX() As Double
        Get
            Return _tensaoCompressaoX
        End Get
        Set(value As Double)
            _tensaoCompressaoX = value
        End Set
    End Property
    Public Property TensaoCompressaoY() As Double
        Get
            Return _tensaoCompressaoY
        End Get
        Set(value As Double)
            _tensaoCompressaoY = value
        End Set
    End Property
    Public Property TensaoMxd() As Double
        Get
            Return _tensaoMxd
        End Get
        Set(value As Double)
            _tensaoMxd = value
        End Set
    End Property
    Public Property TensaoMyd() As Double
        Get
            Return _tensaoMyd
        End Get
        Set(value As Double)
            _tensaoMyd = value
        End Set
    End Property
    Public Property EsbeltezCompressaoX() As Double
        Get
            Return _esbeltezCompressaoX
        End Get
        Set(value As Double)
            _esbeltezCompressaoX = value
        End Set
    End Property
    Public Property EsbeltezCompressaoY() As Double
        Get
            Return _esbeltezCompressaoY
        End Get
        Set(value As Double)
            _esbeltezCompressaoY = value
        End Set
    End Property
    Public Property ForcaElasticaX() As Double
        Get
            Return _forcaElasticaX
        End Get
        Set(value As Double)
            _forcaElasticaX = value
        End Set
    End Property
    Public Property ForcaElasticaY() As Double
        Get
            Return _forcaElasticaY
        End Get
        Set(value As Double)
            _forcaElasticaY = value
        End Set
    End Property
    Public Property Lvinculado() As Double
        Get
            Return _lvinculado
        End Get
        Set(value As Double)
            _lvinculado = value
        End Set
    End Property
    Public Property EiX() As Double
        Get
            Return _eix
        End Get
        Set(value As Double)
            _eix = value
        End Set
    End Property
    Public Property EiY() As Double
        Get
            Return _eiy
        End Get
        Set(value As Double)
            _eiy = value
        End Set
    End Property
    Public Property EaX() As Double
        Get
            Return _eaX
        End Get
        Set(value As Double)
            _eaX = value
        End Set
    End Property
    Public Property EaY() As Double
        Get
            Return _eaY
        End Get
        Set(value As Double)
            _eaY = value
        End Set
    End Property
    Public Property Edx() As Double
        Get
            Return _edx
        End Get
        Set(value As Double)
            _edx = value
        End Set
    End Property
    Public Property Edy() As Double
        Get
            Return _edy
        End Get
        Set(value As Double)
            _edy = value
        End Set
    End Property
    Public Property E1ef() As Double
        Get
            Return _e1ef
        End Get
        Set(value As Double)
            _e1ef = value
        End Set
    End Property
    Public Property Ecx() As Double
        Get
            Return _ecx
        End Get
        Set(value As Double)
            _ecx = value
        End Set
    End Property
    Public Property Ecy() As Double
        Get
            Return _ecy
        End Get
        Set(value As Double)
            _ecy = value
        End Set
    End Property
    Public Property Eig() As Double
        Get
            Return _eig
        End Get
        Set(value As Double)
            _eig = value
        End Set
    End Property
    Public Property EixoYY() As Double
        Get
            Return _eixoYY
        End Get
        Set(value As Double)
            _eixoYY = value
        End Set
    End Property
    Public Property EixoXX() As Double
        Get
            Return _eixoXX
        End Get
        Set(value As Double)
            _eixoXX = value
        End Set
    End Property

    Public Shared Function CalcTensaoC(normalCompressao As Double, momentoFletorX As Double, momentoFletorY As Double, lvinculado As Double, b As Double, h As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao,
                                     MomentoCargaPermanenteX As Double, MomentoCargaPermanenteY As Double, MomentoCargaVariavelX As Double, MomentoCargaVariavelY As Double, NormalCargaPermanente As Double, NormalCargaVariavel As Double,
                                     coeficiente As Double, f1 As Double, f2 As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String, l1 As Double, a1 As Double) As Compressao

        Dim compr = New Compressao()
        Dim mdX As Double = 0
        Dim mdY As Double = 0

        Form1.esbelta = False
        compr.Lvinculado = lvinculado

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                'CALCULO DA ESBELTEZ
                compr.EsbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area))) 'OK

                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

            Case Madeira.TipoSecao.Circular
                'CALCULO DA ESBELTEZ
                compr.EsbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area))) 'OK

                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

            Case Madeira.TipoSecao.SecaoT
                'CALCULO DA ESBELTEZ
                compr.EsbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area))) 'OK

                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

            Case Madeira.TipoSecao.SecaoI
                'CALCULO DA ESBELTEZ
                compr.EsbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area))) 'OK

                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

            Case Madeira.TipoSecao.Caixao
                'CALCULO DA ESBELTEZ
                compr.EsbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area))) 'OK

                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

            Case Madeira.TipoSecao.ElementosJustaposto2
                compr.EsbeltezCompressaoX = (lvinculado / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado / (Math.Sqrt(proprGeo.EixoYie / proprGeo.Area))) 'OK
                'segurança no eixo YY
                compr.EixoYY = ((normalCompressao / a) + ((mdY * proprGeo.EixoYmi1) / (proprGeo.EixoYie * proprGeo.W2)) + ((mdY / (2 * a1 * proprGeo.AreaA1)) * (1 - (3 * (proprGeo.EixoYmi1 / proprGeo.EixoYmi1)))) * 10)
                'segurança no eixo XX
                compr.EixoXX = ((normalCompressao / a) + ((mdY / proprGeo.EixoYmi1) * (h / 2)))

                'peças interpostas
                If (a <= 3 * b1) And (9 * b1 <= l1 <= 18 * b1) And (elementoFixacaoSelecionado = "Espacador Interposto") Then
                    MessageBox.Show("De acordo com a NBR 7190:1997, a verificação da estabilidade local dos trechos de comprimento L1 pode ser dispensada quando forem atendidas as seguintes condições:
                                    9b1 ≤ L1 ≤ 18b1,
	                                a ≤ 3b1")
                    'chapas
                ElseIf (a <= 6 * b1) And (9 * b1 <= l1 <= 18 * b1) And (elementoFixacaoSelecionado = "Chapas Laterais de Fixação") Then
                    MessageBox.Show("A NBR 7190:1997 dipões que a verificação da estabilidade local dos trechos de comprimento L1 pode ser dispensada quando forem atendidas as seguintes condições:
                                    9b1 ≤ L1 ≤ 18b1 e
                                    a ≤ 6b1: para chapas laterais de fixação")
                Else
                    MessageBox.Show("A verificação da estabilidade global deve ser feita nas direções x e y.")
                End If

                'ESTABILIDADE LOCAL
                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa
                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

            Case Madeira.TipoSecao.ElementosJustaposto3
                compr.EsbeltezCompressaoX = (lvinculado / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado / (Math.Sqrt(proprGeo.EixoYie / proprGeo.Area))) 'OK

                'peças interpostas
                If (a <= 3 * b1) And (9 * b1 <= l1 <= 18 * b1) And (elementoFixacaoSelecionado = "Espacador Interposto") Then
                    MessageBox.Show("De acordo com a NBR 7190:1997, a verificação da estabilidade local dos trechos de comprimento L1 pode ser dispensada quando forem atendidas as seguintes condições:
                                    9b1 ≤ L1 ≤ 18b1,
	                                a ≤ 3b1")
                    'chapas
                ElseIf (a <= 6 * b1) And (9 * b1 <= l1 <= 18 * b1) And (elementoFixacaoSelecionado = "Chapas Laterais de Fixação") Then
                    MessageBox.Show("A NBR 7190:1997 dipões que a verificação da estabilidade local dos trechos de comprimento L1 pode ser dispensada quando forem atendidas as seguintes condições:
                                    9b1 ≤ L1 ≤ 18b1 e
                                    a ≤ 6b1: para chapas laterais de fixação")
                Else
                    MessageBox.Show("A verificação da estabilidade global deve ser feita nas direções x e y.")
                End If

                'ESTABILIDADE LOCAL
                'PEÇA CURTA EM X ==============
                If compr.EsbeltezCompressaoX <= 40 Then
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM X ==============
                ElseIf 40 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 80 Then
                    compr.ForcaElasticaX = ((System.Math.PI ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaX = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.EiX = h / 30
                    Else
                        compr.EiX = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.EiX + compr.EaX) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa


                    'ELEMENTO ESBELTO EM X ==============
                ElseIf 80 < compr.EsbeltezCompressaoX And compr.EsbeltezCompressaoX <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaX = (((System.Math.PI) ^ 2) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (h / 30) Then
                        compr.EaX = h / 30
                    Else
                        compr.EaX = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiX = (((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteX * 100) / normalCompressao
                    compr.Ecx = (compr.EaX + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaX * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaX + compr.EiX + compr.Ecx
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressaoX = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMxd = ((mdX / proprGeo.EixoXmr) * 10) 'multiplicando por 10 para sair em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a Seção transversal da peça.")
                End If

                'PEÇA CURTA EM Y ==============
                If compr.EsbeltezCompressaoY <= 40 Then
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)

                    'ELEMENTO MEDIANAMENTE ESBELTO EM Y ==============
                ElseIf 40 < compr.EsbeltezCompressaoY And compr.EsbeltezCompressaoY <= 80 Then
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    compr.EaY = (lvinculado * 100) / 300
                    'Definindo o valor da excentricidade inicial
                    If ((momentoFletorY / normalCompressao) * 100) < (b / 30) Then
                        compr.EiY = b / 30
                    Else
                        compr.EiY = (momentoFletorY / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.EiY + compr.EaY) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10)
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10)

                    'ELEMENTO ESBELTO EM Y ==============
                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then
                    Form1.esbelta = True
                    compr.ForcaElasticaY = ((System.Math.PI) ^ 2 * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade / 10)) / ((lvinculado * 100) ^ 2)
                    'Definindo o valor da excentricidade acidental
                    If ((lvinculado * 100) / 300) < (b / 30) Then
                        compr.EaY = b / 30
                    Else
                        compr.EaY = (lvinculado * 100) / 300
                    End If
                    'Definindo o valor da excentricidade inicial
                    compr.EiY = (((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao)
                    compr.Eig = (MomentoCargaPermanenteY * 100) / normalCompressao
                    compr.Ecy = (compr.EaY + compr.Eig) * (System.Math.Exp(((coeficiente * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (compr.ForcaElasticaY - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)))) - 1)
                    compr.E1ef = compr.EaY + compr.EiY + compr.Ecy
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressaoY = ((normalCompressao / proprGeo.Area) * 10) 'multiplicando por 10 para sair em MPa
                    compr.TensaoMyd = ((mdY / proprGeo.EixoYmr) * 10) 'multiplicando por 10 para sair em MPa
                ElseIf compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                    Insira novos valores para a seção transversal da peça.")
                End If

        End Select
        Return compr
    End Function
End Class
