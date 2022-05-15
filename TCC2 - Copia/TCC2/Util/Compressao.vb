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

    Private _tensaoCompressao As Double = 0
    Private _tensaoCompressaoCurtaX As Double = 0
    Private _tensaoCompressaoCurtaY As Double = 0
    Private _tensaoCompressaoMedEsbeltaX As Double = 0
    Private _tensaoCompressaoMedEsbeltaY As Double = 0
    Private _tensaoCompressaoEsbeltaX As Double = 0
    Private _tensaoCompressaoEsbeltaY As Double = 0
    Private _tensaoMxd As Double = 0
    Private _tensaoMyd As Double = 0
    Private _esbeltezCompressaoX As Double = 0
    Private _esbeltezCompressaoY As Double = 0
    Private _forcaElasticaX As Double = 0
    Private _forcaElasticaY As Double = 0
    Private _lvinculado As Double = 0
    Private _ei As Double = 0 'excentricidade inicial
    Private _ea As Double = 0 'excentricidade acidental
    Private _edx As Double = 0 'excentricidade de projeto em x 
    Private _edy As Double = 0 'excentricidade de projeto em y
    Private _ec As Double = 0 'excentricidade suplementar de primeira ordem
    Private _e1ef As Double = 0 'excentricidade de primeira ordem
    Private _eig As Double = 0 'excentricidade de primeira ordem devido as cargas permanentes


    Public Sub New()

    End Sub

    Public Property TensaoCompressao() As Double
        Get
            Return _tensaoCompressao
        End Get
        Set(value As Double)
            _tensaoCompressao = value
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
    Public Property Ei() As Double
        Get
            Return _ei
        End Get
        Set(value As Double)
            _ei = value
        End Set
    End Property
    Public Property Ea() As Double
        Get
            Return _ea
        End Get
        Set(value As Double)
            _ea = value
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
    Public Property Ec() As Double
        Get
            Return _ec
        End Get
        Set(value As Double)
            _ec = value
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
    Public Property TensaoCompressaoCurtaX() As Double
        Get
            Return _tensaoCompressaoCurtaX
        End Get
        Set(value As Double)
            _tensaoCompressaoCurtaX = value
        End Set
    End Property
    Public Property TensaoCompressaoCurtaY() As Double
        Get
            Return _tensaoCompressaoCurtaY
        End Get
        Set(value As Double)
            _tensaoCompressaoCurtaY = value
        End Set
    End Property
    Public Property TensaoCompressaoMedEsbeltaX() As Double
        Get
            Return _tensaoCompressaoMedEsbeltaX
        End Get
        Set(value As Double)
            _tensaoCompressaoMedEsbeltaX = value
        End Set
    End Property
    Public Property TensaoCompressaoMedEsbeltaY() As Double
        Get
            Return _tensaoCompressaoMedEsbeltaY
        End Get
        Set(value As Double)
            _tensaoCompressaoMedEsbeltaY = value
        End Set
    End Property
    Public Property TensaoCompressaoEsbeltaX() As Double
        Get
            Return _tensaoCompressaoEsbeltaX
        End Get
        Set(value As Double)
            _tensaoCompressaoEsbeltaX = value
        End Set
    End Property
    Public Property TensaoCompressaoEsbeltaY() As Double
        Get
            Return _tensaoCompressaoEsbeltaY
        End Get
        Set(value As Double)
            _tensaoCompressaoEsbeltaY = value
        End Set
    End Property

    Public Shared Function CalcTensaoC(normalCompressao As Double, momentoFletorX As Double, momentoFletorY As Double, lvinculado As Double, b As Double, h As Double, proprGeo As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao,
                                     MomentoCargaPermanenteX As Double, MomentoCargaPermanenteY As Double, MomentoCargaVariavelX As Double, MomentoCargaVariavelY As Double, NormalCargaPermanente As Double, NormalCargaVariavel As Double,
                                     coefInfluencia As Double, f1 As Double, f2 As Double, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String, h1 As Double) As Compressao

        Dim compr = New Compressao()
        Dim mdX As Double = 0
        Dim mdY As Double = 0

        Select Case tipoSecao

            Case Madeira.TipoSecao.Retangular
                compr.EsbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoXmi / proprGeo.Area))) 'OK
                compr.EsbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(proprGeo.EixoYmi / proprGeo.Area))) 'OK

                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    compr.TensaoCompressao = (normalCompressao / proprGeo.Area)

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    compr.TensaoCompressao = (normalCompressao / proprGeo.Area)

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade * 100)) / ((lvinculado ^ 2) * 100)
                    compr.Ea = (lvinculado * 100) / 300
                    'DEFININDO O VALOR DA EXCENTRICIDADE INICIAL
                    If ((momentoFletorX / normalCompressao) * 100) < (b / 30) Then
                        compr.Ei = b / 30
                    Else
                        compr.Ei = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edx = (compr.Ei + compr.Ea) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressao = (normalCompressao / proprGeo.Area)
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr)

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade * 100)) / ((lvinculado ^ 2) * 100)
                    compr.Ea = (lvinculado * 100) / 300
                    'DEFININDO O VALOR DA EXCENTRICIDADE INICIAL
                    If ((momentoFletorX / normalCompressao) * 100) < (h / 30) Then
                        compr.Ei = h / 30
                    Else
                        compr.Ei = (momentoFletorX / normalCompressao) * 100
                    End If
                    compr.Edy = (compr.Ei + compr.Ea) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoYmr)


                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * (ResistenciaCalculo.moduloElasticidade * 100)) / ((lvinculado ^ 2) * 100)
                    compr.Ea = (lvinculado * 100) / 300
                    'DEFININDO O VALOR DA EXCENTRICIDADE INICIAL
                    If ((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao < (b / 30) Then
                        compr.Ei = b / 30
                    Else
                        compr.Ei = ((MomentoCargaPermanenteX + MomentoCargaVariavelX) * 100) / normalCompressao
                    End If
                    compr.Eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr)


                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * (ResistenciaCalculo.moduloElasticidade * 100)) / ((lvinculado ^ 2) * 100)
                    compr.Ea = (lvinculado * 100) / 300
                    'DEFININDO O VALOR DA EXCENTRICIDADE INICIAL
                    If ((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao < (b / 30) Then
                        compr.Ei = b / 30
                    Else
                        compr.Ei = ((MomentoCargaPermanenteY + MomentoCargaVariavelY) * 100) / normalCompressao
                    End If
                    compr.Eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoXmr)

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")
                End If


                ''CONFERIR COM O PROFESSOR AS FORMULAR ANTERIORES

            Case Madeira.TipoSecao.Circular
                compr.EsbeltezCompressaoX = lvinculado / proprGeo.EixoXrg
                compr.EsbeltezCompressaoY = lvinculado / proprGeo.EixoYrg

                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edx = (compr.Ei + compr.Ea) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edy = (compr.Ei + compr.Ea) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.SecaoT
                compr.EsbeltezCompressaoX = 0
                compr.EsbeltezCompressaoY = 0
                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edx = (compr.Ei + compr.Ea) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edy = (compr.Ei + compr.Ea) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                     Insira novamente a seção transversal da peça para prosseguir com o dimensionamento.")


                End If

            Case Madeira.TipoSecao.SecaoI
                compr.EsbeltezCompressaoX = 0
                compr.EsbeltezCompressaoY = 0

                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edx = (compr.Ei + compr.Ea) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edy = (compr.Ei + compr.Ea) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.Caixao
                compr.EsbeltezCompressaoX = 0
                compr.EsbeltezCompressaoY = 0

                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edx = (compr.Ei + compr.Ea) * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    mdX = normalCompressao * compr.Edx
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    compr.Ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Edy = (compr.Ei + compr.Ea) * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    mdY = normalCompressao * compr.Edy
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaX = ((System.Math.PI) * proprGeo.EixoXmi * proprGeo.EixoXmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdX = normalCompressao * compr.E1ef * (compr.ForcaElasticaX / (compr.ForcaElasticaX - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMxd = (mdX / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    compr.ForcaElasticaY = ((System.Math.PI) * proprGeo.EixoYmi * proprGeo.EixoYmr) / (lvinculado ^ 2)
                    compr.Ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    compr.Ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    compr.Eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    compr.Ec = (compr.Ea + compr.Eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    compr.E1ef = compr.Ea + compr.Ei + compr.Ec
                    mdY = normalCompressao * compr.E1ef * (compr.ForcaElasticaY / (compr.ForcaElasticaY - normalCompressao))
                    compr.TensaoCompressao = normalCompressao / proprGeo.Area
                    compr.TensaoMyd = (mdY / proprGeo.EixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.ElementosJustaposto2

                compr.EsbeltezCompressaoX = 0
                compr.EsbeltezCompressaoY = 0

                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y

                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.ElementosJustaposto3
                compr.EsbeltezCompressaoX = 0
                compr.EsbeltezCompressaoY = 0

                If compr.EsbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X

                ElseIf compr.EsbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y

                ElseIf 40 < compr.EsbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X

                ElseIf 40 < compr.EsbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y

                ElseIf 80 < compr.EsbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf 80 < compr.EsbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf compr.EsbeltezCompressaoX > 140 And compr.EsbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

        End Select

        Return compr
    End Function

End Class
