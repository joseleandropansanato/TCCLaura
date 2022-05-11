Public Class Compressao
	Public Shared tensaoCompressao As Double = 0
	Public Shared tensaoCompressaoCurtaX As Double = 0
	Public Shared tensaoCompressaoCurtaY As Double = 0
	Public Shared tensaoCompressaoMedEsbeltaX As Double = 0
	Public Shared tensaoCompressaoMedEsbeltaY As Double = 0
	Public Shared tensaoCompressaoEsbeltaX As Double = 0
	Public Shared tensaoCompressaoEsbeltaY As Double = 0
	Public Shared tensaoMxd As Double = 0
	Public Shared tensaoMyd As Double = 0
	Public Shared esbeltezCompressaoX As Double = 0
	Public Shared esbeltezCompressaoY As Double = 0
	Public Shared forcaElasticaX As Double = 0
	Public Shared forcaElasticaY As Double = 0
	Public Shared lvinculado As Double = 0
	Public Shared ei As Double = 0 'excentricidade inicial
	Public Shared ea As Double = 0 'excentricidade acidental
	Public Shared edx As Double = 0 'excentricidade de projeto em x 
	Public Shared edy As Double = 0 'excentricidade de projeto em y
	Public Shared ec As Double = 0 'excentricidade suplementar de primeira ordem
	Public Shared e1ef As Double = 0 'excentricidade de primeira ordem
    Public Shared eig As Double = 0 'excentricidade de primeira ordem devido as cargas permanentes
    Public Shared eixoXPecaComposta As Double = 0
    Public Shared eixoYPecaComposta As Double = 0


    Public Function CalcTensaoC(normalCompressao As Double, momentoFletorX As Double, momentoFletorY As Double, lvinculado As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao,
                                     MomentoCargaPermanenteX As Double, MomentoCargaPermanenteY As Double, MomentoCargaVariavelX As Double, MomentoCargaVariavelY As Double, NormalCargaPermanente As Double, NormalCargaVariavel As Double,
                                     coefInfluencia As Double, f1 As Double, f2 As Double, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String, h1 As Double) As Double

        'Dim compressao = New Compressao()
        Dim mdX As Double = 0
        Dim mdY As Double = 0

        Select Case tipoSecao

            Case Madeira.TipoSecao.Retangular

                esbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area))) 'OK
                esbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area))) 'OK

                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edy = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * Compressao.edy
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdX = normalCompressao * e1ef * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdY = normalCompressao * e1ef * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf esbeltezCompressaoX > 140 And esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.Circular
                esbeltezCompressaoX = lvinculado / PropriedadesGeometricas.eixoXrg
                esbeltezCompressaoY = lvinculado / PropriedadesGeometricas.eixoYrg

                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edx = (ei + ea) * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * edx
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edy = (ei + ea) * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * edy
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdX = normalCompressao * e1ef * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdY = normalCompressao * e1ef * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf esbeltezCompressaoX > 140 And esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.SecaoT

                esbeltezCompressaoX = 0
                esbeltezCompressaoY = 0
                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edx = (ei + ea) * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * edx
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edy = (ei + ea) * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * edy
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdX = normalCompressao * e1ef * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdY = normalCompressao * e1ef * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf esbeltezCompressaoX > 140 And esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                     Insira novamente a seção transversal da peça para prosseguir com o dimensionamento.")


                End If

            Case Madeira.TipoSecao.SecaoI
                esbeltezCompressaoX = 0
                esbeltezCompressaoY = 0

                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edy = (ei + ea) * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * edy
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdX = normalCompressao * e1ef * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdY = normalCompressao * e1ef * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf esbeltezCompressaoX > 140 And esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.Caixao
                esbeltezCompressaoX = 0
                esbeltezCompressaoY = 0

                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edx = (ei + ea) * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * edx
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    edy = (ei + ea) * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * edy
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + ec
                    mdX = normalCompressao * e1ef * (forcaElasticaX / (forcaElasticaX - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    ec = (ea + eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    e1ef = ea + ei + Compressao.ec
                    mdY = normalCompressao * e1ef * (forcaElasticaY / (forcaElasticaY - normalCompressao))
                    tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.ElementosJustaposto2
                esbeltezCompressaoX = 0
                esbeltezCompressaoY = 0

                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y

                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf 80 < esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.ElementosJustaposto3


                If esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X

                ElseIf esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y

                ElseIf 40 < esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X

                ElseIf 40 < esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y

                ElseIf 80 < esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If


                'não é assim, arrumar
                esbeltezCompressaoX = 0
                esbeltezCompressaoY = 0

                'botar o espaçador aqui tb
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    esbeltezCompressaoX = ((lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                    esbeltezCompressaoY = ((lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))
                    'eixo X::
                    eixoXPecaComposta = (normalCompressao / PropriedadesGeometricas.area) + (momentoFletorX / PropriedadesGeometricas.eixoXmi) * (h1 / 2)
                    'eixo y::
                    eixoYPecaComposta = (normalCompressao / PropriedadesGeometricas.area) + ((momentoFletorY * PropriedadesGeometricas.eixoYmi1) / PropriedadesGeometricas.eixoYie * PropriedadesGeometricas.eixoYmr) + (momentoFletorY / (2 * PropriedadesGeometricas.area * PropriedadesGeometricas.areaA1)) * (1 - 2 * (PropriedadesGeometricas.eixoYmi1 / PropriedadesGeometricas.eixoYie))
                End If

                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    esbeltezCompressaoX = (lvinculado) / PropriedadesGeometricas.eixoXrg
                    esbeltezCompressaoY = (lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoYie / PropriedadesGeometricas.area))
                    'eixo X::
                    eixoXPecaComposta = (normalCompressao / PropriedadesGeometricas.area) + (momentoFletorX / PropriedadesGeometricas.eixoXmi) * (h1 / 2)
                    'eixo y::
                    eixoYPecaComposta = (normalCompressao / PropriedadesGeometricas.area) + ((momentoFletorY * PropriedadesGeometricas.eixoYmi1) / PropriedadesGeometricas.eixoYie * PropriedadesGeometricas.eixoYmr) + (momentoFletorY / (3 * PropriedadesGeometricas.area * PropriedadesGeometricas.areaA1)) * (1 - 2 * (PropriedadesGeometricas.eixoYmi1 / PropriedadesGeometricas.eixoYie))
                End If




        End Select

        Return CalcTensaoC
    End Function

End Class
