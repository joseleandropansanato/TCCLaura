
Public Class PropriedadesGeometricas

    'serve para receber valor e mandar quando pedir 
    Public Shared area As Double = 0
    Public Shared areaA1 As Double = 0
    Public Shared eixoXmi As Double = 0 'EIXO X MOMENTO DE INERCIA
    Public Shared eixoXmi1 As Double = 0 'EIXO X MOMENTO DE INERCIA PRINCIPAL EM X
    Public Shared eixoXrg As Double = 0 'EIXO X RAIO DE GIRAÇÃO
    Public Shared eixoXmr As Double = 0 'EIXO X 
    Public Shared eixoYmi As Double = 0 'EIXO Y MOMENTO DE INERCIA
    Public Shared eixoYrg As Double = 0 'EIXO Y RAIO DE GIRAÇÃO
    Public Shared eixoYmr As Double = 0 'EIXO Y MODULO DE RESISTENCIA
    Public Shared eixoXie As Double = 0 'EIXO X INERCIA EFETIVA
    Public Shared eixoYie As Double = 0 'EIXO Y INERCIA EFETIVA
    Public Shared eixoYmie As Double = 0 'EIXO Y MOMENTO DE INERCIA EFETIVA EM Y
    Public Shared eixoYmi1 As Double = 0 'EIXO Y MOMENTO DE INERCIA PRINCIPAL EM Y
    Public Shared coefBy As Double = 0 'COEFICIENTE DO ESPAÇADOR
    Public Shared Qx As Double = 0 'MOMENTO ESTÁTICO DE ÁREA EM X
    Public Shared Qy As Double = 0 'MOMENTO ESTÁTICO DE ÁREA EM Y


    Enum Pilar
        Espaçador
        Chapa
    End Enum


    Public Function MomentoEstaticoDeArea(dadosIniciais As DadosIniciais, tipoSecao As Madeira.TipoSecao) As Double
        Dim bases As Double() = {0, 0, 0}
        Dim alturas As Double() = {0, 0, 0}
        Dim xCG As Double = 0
        Dim x_CGi As Double() = {0, 0, 0}
        Dim y_CGi As Double() = {0, 0, 0}

        'CÁLCULO DE Q:
        'Private Sub Pic_TmultS_Click(sender As Object, e As EventArgs) Handles Pic_TmultS.Click
        Select Case tipoSecao
            Case Madeira.TipoSecao.SecaoT, Madeira.TipoSecao.SecaoI
                'SEÇÃO I OU T ==========================================
                'Dim bases As Double() = {10.0, 5.0, 40.0}
                'Dim alturas As Double() = {4, 10.0, 2.0}

                bases = {dadosIniciais.b1, dadosIniciais.b2, dadosIniciais.b3}
                alturas = {dadosIniciais.h1, dadosIniciais.h2, dadosIniciais.h3}

                'calcula as coordenadas x, y do CG de cada figura
                xCG = bases.Max / 2
                x_CGi = {xCG, xCG, xCG}
                y_CGi = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas.Sum - alturas(2) / 2}
        'FIM SEÇÃO I OU T ======================================

            Case Madeira.TipoSecao.Caixao
                ''SEÇÃO CAIXAO ------------------------------------------
                ' Dim bases As Double() = {20, 4, 4, 20}
                'Dim alturas As Double() = {10, 40, 40, 10}
                bases = {dadosIniciais.b1c, dadosIniciais.b2c, dadosIniciais.b3c, dadosIniciais.b4c}
                alturas = {dadosIniciais.h1c, dadosIniciais.h2c, dadosIniciais.h3c, dadosIniciais.h4c}

                ''calcula as coordenadas x,y do CG de cada figura
                xCG = bases.Max / 2
                Dim base As Double = Min(bases(0), bases(3))
                x_CGi = {xCG, xCG - base / 2 + bases(1) / 2, xCG + base / 2 - bases(1) / 2, xCG}
                y_CGi = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) + alturas(3) / 2}
                ''FIM SEÇÃO CAIXÃO ----------------------------------------

        End Select

        Dim area As Double = Calcular_Area(bases, alturas)

        Dim xCG_perfil = Calcular_CG(bases, alturas, x_CGi)
        Dim yCG_perfil = Calcular_CG(bases, alturas, y_CGi)

        Dim InerciaX = Inercia(bases, alturas, y_CGi)
        Dim InerciaY = Inercia(alturas, bases, x_CGi)

        Dim qxy = New Qxy()
        qxy.Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, tipoSecao)
        qxy.Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, tipoSecao)

        'Dim Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, "I")
        'Dim Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, "I")
        'Dim Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, "caixao")
        'Dim Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, "caixao")

        MessageBox.Show("Area: " + area.ToString + vbNewLine +
                        "xCG: " + xCG_perfil.ToString + vbNewLine +
                        "yCG: " + yCG_perfil.ToString + vbNewLine +
                        "Ix: " + InerciaX.ToString + vbNewLine +
                        "Iy: " + InerciaY.ToString + vbNewLine +
                        "Qx: " + qxy.Qx.ToString + vbNewLine +
                        "Qy: " + qxy.Qy.ToString + vbNewLine)

        Return q
    End Function

    Private Function Calcular_Area(ByVal bases() As Double, ByVal alturas() As Double) As Double
        Dim area As Double = 0
        For i As Integer = 0 To bases.Length - 1
            area += bases(i) * alturas(i)
        Next
        Return area
    End Function

    Private Function Calcular_CG(ByVal bases() As Double, ByVal alturas() As Double, ByVal CG_i() As Double) As Double
        Dim CG As Double = 0
        For i As Integer = 0 To bases.Length - 1
            CG += bases(i) * alturas(i) * CG_i(i)
        Next
        Return CG / Calcular_Area(bases, alturas)
    End Function

    Private Function Inercia(ByVal bases() As Double, ByVal alturas() As Double, ByVal yCG_i() As Double) As Double
        Dim yCG As Double = Calcular_CG(bases, alturas, yCG_i)

        'Rotina de cálculo da inércia ao longo do eixo x (eixo horizontal que cruza a secao transversal).
        'Para cálculo ao longo do eixo y (vertical), trocar bases() por alturas() no input da funcao
        'e trocar yCGi por xCGi.
        Inercia = 0
        For i As Integer = 0 To bases.Length - 1
            Inercia += bases(i) * alturas(i) ^ 3 / 12 + bases(i) * alturas(i) * (yCG - yCG_i(i)) ^ 2
        Next
        Return Inercia
    End Function

    Private Function Momento_Estatico_Qx(ByVal bases() As Double, ByVal alturas() As Double, ByVal yCG_i() As Double, section As Madeira.TipoSecao) As Double

        Dim y_barra As Double, A_corte As Double, Q As Double
        Dim yCG As Double = Calcular_CG(bases, alturas, yCG_i)

        Select Case section
            Case Madeira.TipoSecao.SecaoT, Madeira.TipoSecao.SecaoI
                If yCG <= alturas(0) Then
                    'CG da área cisalhada está contido na mesa inferior -> y_barra = yCG/2
                    y_barra = yCG / 2
                    A_corte = bases(0) * yCG

                ElseIf yCG <= alturas(0) + alturas(1) Then
                    'CG da área cisalhada está contido na alma -> y_barra deve ser calculado

                    Dim altura_alma As Double = yCG - alturas(0)
                    y_barra = yCG - Calcular_CG({bases(0), bases(1)}, {alturas(0), altura_alma}, {yCG_i(0), alturas(0) + altura_alma / 2})
                    A_corte = bases(0) * alturas(0) + bases(1) * altura_alma

                Else
                    'CG da área cisalhada está contido na mesa superior -> y_barra = [\sum(alturas) - yCG]/2
                    Dim altura_mesa As Double = alturas.Sum - yCG
                    y_barra = altura_mesa / 2
                    A_corte = bases(2) * altura_mesa
                End If

            Case Madeira.TipoSecao.Caixao
                If yCG <= alturas(0) Then
                    'CG da área cisalhada está contido na mesa inferior -> y_barra = yCG/2
                    y_barra = yCG / 2
                    A_corte = bases(0) * yCG

                ElseIf yCG <= alturas(0) + alturas(1) Then
                    Dim altura_alma As Double = yCG - alturas(0)
                    y_barra = yCG - Calcular_CG({bases(0), bases(1), bases(2)}, {alturas(0), altura_alma, altura_alma}, {yCG_i(0), alturas(0) + altura_alma / 2, alturas(0) + altura_alma / 2})
                    A_corte = bases(0) * alturas(0) + bases(1) * altura_alma * 2

                Else
                    'CG da área cisalhada está contido na mesa superior -> y_barra = [\sum(alturas) - yCG]/2
                    Dim altura_mesa As Double = alturas(0) + alturas(1) + alturas(3) - yCG
                    y_barra = altura_mesa / 2
                    A_corte = bases(2) * altura_mesa
                End If
        End Select

        Q = A_corte * y_barra

        Return Q
    End Function

    Private Function Momento_Estatico_Qy(ByVal bases() As Double, ByVal alturas() As Double, ByVal xCG_i() As Double, ByVal section As String) As Double

        Dim x_barra As Double, A_corte As Double
        Dim xCG As Double = Calcular_CG(bases, alturas, xCG_i)

        Select Case section
            Case Madeira.TipoSecao.SecaoT, Madeira.TipoSecao.SecaoI
                Dim bases_aux As Double() = bases.ToArray
                Dim xCG_i_aux As Double() = xCG_i.ToArray
                For i As Integer = 0 To bases_aux.Length - 1
                    bases_aux(i) /= 2
                    xCG_i_aux(i) -= bases_aux(i) / 2
                Next

                'CG da área cisalhada está contido na alma -> x_barra deve ser calculado
                x_barra = xCG - Calcular_CG(bases_aux, alturas, xCG_i_aux)
                A_corte = Calcular_Area(bases, alturas) / 2

            Case Madeira.TipoSecao.Caixao
                Dim bases_aux As Double() = {bases(0), bases(1), bases(3)}
                Dim xCG_i_aux As Double() = {xCG_i(0), xCG_i(1), xCG_i(3)}
                For i As Integer = 0 To 2 Step 2
                    bases_aux(i) /= 2
                    xCG_i_aux(i) -= bases_aux(i) / 2
                Next
                'CG da área cisalhada está contido na alma -> x_barra deve ser calculado
                x_barra = xCG - Calcular_CG(bases_aux, {alturas(0), alturas(1), alturas(3)}, xCG_i_aux)
                A_corte = Calcular_Area(bases, alturas) / 2

        End Select

        Return A_corte * x_barra
    End Function

    Private Function Min(b1 As Double, b2 As Double) As Double
        If b1 < b2 Then
            Return b1
        ElseIf b1 > b2 Then
            Return b2
        End If
        Return b2
    End Function

    'PROPRIEDADES GEOMÉTRICAS DA PEÇA 
    Public Function CalculoRetangular(baseX As Double, baseY As Double, altura As Double) As PropriedadesGeometricas
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        'seção retangular
        PropriedadesGeometricas.area = baseX * baseY 'OK
        PropriedadesGeometricas.eixoXmi = (baseX * baseY ^ 3) / 12 'OK
        PropriedadesGeometricas.eixoXrg = Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoXmr = PropriedadesGeometricas.eixoXmi / (altura / 2)
        PropriedadesGeometricas.eixoYmi = (baseY * baseX ^ 3) / 12 'OK
        PropriedadesGeometricas.eixoYrg = Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoYmr = PropriedadesGeometricas.eixoYmi / (altura / 2)

        Return propriedadesGeometricas
    End Function

    Public Function CalculoCircular(diametro As Double, altura As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        'seção cirucular
        PropriedadesGeometricas.area = System.Math.PI * (diametro / 2) ^ 2 'OK
        PropriedadesGeometricas.eixoXmi = (System.Math.PI * (diametro) ^ 4) / 64 'OK
        PropriedadesGeometricas.eixoXrg = Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoXmr = PropriedadesGeometricas.eixoXmi / (altura / 2)
        PropriedadesGeometricas.eixoYmi = (System.Math.PI * (diametro) ^ 4) / 64 'OK
        PropriedadesGeometricas.eixoYrg = Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoYmr = PropriedadesGeometricas.eixoYmi / (altura / 2)


        Return propriedadesGeometricas
    End Function

    Public Function CalculoCompostoSecaoT(bf As Double, hf As Double, h As Double, bw As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()
        Dim bases As Double() = {0, bw, bf}
        Dim alturas As Double() = {0, h, hf}
        Dim xCG As Double = bf / 2
        Dim x_CGi As Double() = {xCG, xCG, xCG}
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas.Sum - alturas(2) / 2}

        'seção T
        PropriedadesGeometricas.area = Calcular_Area(bases, alturas)
        PropriedadesGeometricas.eixoXmi = Inercia(bases, alturas, y_CGi)
        PropriedadesGeometricas.eixoXie = PropriedadesGeometricas.eixoXmi * 0.95
        PropriedadesGeometricas.eixoXrg = Math.Sqrt(PropriedadesGeometricas.eixoXie / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoXmr = PropriedadesGeometricas.eixoXie / (h / 2)
        PropriedadesGeometricas.eixoYmi = Inercia(alturas, bases, x_CGi)
        PropriedadesGeometricas.eixoYie = PropriedadesGeometricas.eixoYmi * 0.95
        PropriedadesGeometricas.eixoYrg = Math.Sqrt(PropriedadesGeometricas.eixoYie / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoYmr = PropriedadesGeometricas.eixoYie / (h / 2)

        Return propriedadesGeometricas
    End Function
    Public Function CalculoCompostoSecaoI(bf1 As Double, hf1 As Double, bf2 As Double, hf2 As Double, bw As Double, d As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        Dim bases As Double() = {bf1, bf2, bw}
        Dim alturas As Double() = {hf1, hf2, d}
        Dim xCG As Double = 0
        Dim x_CGi As Double() = {xCG, xCG, xCG}
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas.Sum - alturas(2) / 2}

        'seção I
        'propriedadesGeometricas.Area = (bf1 * hf1) + (bf2 * hf2) + (d - bf1 - bf2)
        PropriedadesGeometricas.area = Calcular_Area(bases, alturas)
        'propriedadesGeometricas.EixoXmi = ((bf1 * hf1 ^ 3) / 12) + ((bf2 * hf2 ^ 3) / 12) + ((bw * (d - hf1 - hf2) ^ 3) / 12)
        PropriedadesGeometricas.eixoXmi = Inercia(bases, alturas, y_CGi)
        PropriedadesGeometricas.eixoXie = PropriedadesGeometricas.eixoXmi * 0.85
        PropriedadesGeometricas.eixoXrg = Math.Sqrt(PropriedadesGeometricas.eixoXie / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoXmr = PropriedadesGeometricas.eixoXie / ((d - hf1 - hf2) / 2)
        'propriedadesGeometricas.EixoYmi = ((bf1 ^ 3 * hf1) / 12) + ((bf2 ^ 3 * hf2) / 12) + ((bw ^ 3 * (d - hf1 - hf2)) / 12)
        PropriedadesGeometricas.eixoYmi = Inercia(bases, alturas, x_CGi)
        PropriedadesGeometricas.eixoYie = PropriedadesGeometricas.eixoYmi * 0.85
        PropriedadesGeometricas.eixoYrg = Math.Sqrt(PropriedadesGeometricas.eixoYie / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoYmr = PropriedadesGeometricas.eixoYie / ((d - hf1 - hf2) / 2)

        Return propriedadesGeometricas
    End Function

    Public Function CalculoCompostoSecaoCaixao(d As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, h1 As Double, h2 As Double, h3 As Double, h4 As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        Dim bases As Double() = {b1, b2, b3, b4}
        Dim alturas As Double() = {h1, h2, h3, h4}
        Dim xCG = bases.Max / 2
        Dim base As Double = Min(bases(0), bases(3))
        Dim x_CGi As Double() = {xCG, xCG - base / 2 + bases(1) / 2, xCG + base / 2 - bases(1) / 2, xCG}
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) + alturas(3) / 2}

        'seção Caixão
        'propriedadesGeometricas.Area = (bf2 * d) - (bf1 * h)
        PropriedadesGeometricas.area = Calcular_Area(bases, alturas)
        'propriedadesGeometricas.EixoXmi = ((bf2 * d ^ 3) - (bf1 * h ^ 3)) / 12
        PropriedadesGeometricas.eixoXmi = Inercia(bases, alturas, y_CGi)
        PropriedadesGeometricas.eixoXie = PropriedadesGeometricas.eixoXmi * 0.85
        PropriedadesGeometricas.eixoXrg = Math.Sqrt(PropriedadesGeometricas.eixoXie / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoXmr = PropriedadesGeometricas.eixoXie / (h3 / 2)
        'propriedadesGeometricas.EixoYmi = ((bf2 ^ 3 * d) - (bf1 ^ 3 * h)) / 12
        PropriedadesGeometricas.eixoYmi = Inercia(bases, alturas, x_CGi)
        PropriedadesGeometricas.eixoYie = PropriedadesGeometricas.eixoYmi * 0.85
        PropriedadesGeometricas.eixoYrg = Math.Sqrt(PropriedadesGeometricas.eixoYie / PropriedadesGeometricas.area)
        PropriedadesGeometricas.eixoYmr = PropriedadesGeometricas.eixoYie / (h3 / 2)

        Return propriedadesGeometricas
    End Function


    Public Function Calculo2Elementos(l As Double, l1 As Double, a As Double, a1 As Double, h As Double, h1 As Double, b1 As Double, pilar As PropriedadesGeometricas.Pilar)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        '2 Elementos Justapostos
        PropriedadesGeometricas.areaA1 = b1 * h1
        PropriedadesGeometricas.area = PropriedadesGeometricas.areaA1 * 2
        PropriedadesGeometricas.eixoXmi1 = (b1 * h1 ^ 3) / 12
        PropriedadesGeometricas.eixoXmi = 2 * PropriedadesGeometricas.eixoXmi1
        PropriedadesGeometricas.eixoYmi1 = (b1 ^ 3 * h1) / 12
        PropriedadesGeometricas.eixoYmi = (2 * PropriedadesGeometricas.eixoYmi1) + (2 * PropriedadesGeometricas.areaA1 * (a1 ^ 2))
        Select Case pilar
            Case PropriedadesGeometricas.Pilar.Espaçador
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * PropriedadesGeometricas.eixoYmi))
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * PropriedadesGeometricas.eixoYmi))
            Case PropriedadesGeometricas.Pilar.Chapa
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * PropriedadesGeometricas.eixoYmi))
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * PropriedadesGeometricas.eixoYmi))
        End Select
        PropriedadesGeometricas.eixoYie = PropriedadesGeometricas.coefBy * PropriedadesGeometricas.eixoYmi
        PropriedadesGeometricas.eixoYmr = PropriedadesGeometricas.eixoYmi1 / (b1 / 2)


        Return propriedadesGeometricas
    End Function

    Public Function Calculo3Elementos(l As Double, l1 As Double, a As Double, a1 As Double, h As Double, h1 As Double, b1 As Double, pilar As PropriedadesGeometricas.Pilar)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        '3 Elementos Justapostos
        PropriedadesGeometricas.areaA1 = b1 * h1
        PropriedadesGeometricas.area = PropriedadesGeometricas.areaA1 * 3
        PropriedadesGeometricas.eixoXmi1 = (b1 * h1 ^ 3) / 12
        PropriedadesGeometricas.eixoXmi = 3 * PropriedadesGeometricas.eixoXmi1
        PropriedadesGeometricas.eixoYmi1 = (b1 ^ 3 * h1) / 12
        PropriedadesGeometricas.eixoYmi = (3 * PropriedadesGeometricas.eixoYmi1) + (2 * PropriedadesGeometricas.areaA1 * (a1 ^ 2))
        Select Case pilar
            Case PropriedadesGeometricas.Pilar.Espaçador
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * PropriedadesGeometricas.eixoYmi))
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * PropriedadesGeometricas.eixoYmi))
            Case PropriedadesGeometricas.Pilar.Chapa
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * PropriedadesGeometricas.eixoXmi))
                PropriedadesGeometricas.coefBy = PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2) / ((PropriedadesGeometricas.eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * PropriedadesGeometricas.eixoYmi))
        End Select
        PropriedadesGeometricas.eixoYie = PropriedadesGeometricas.coefBy * PropriedadesGeometricas.eixoYmi
        PropriedadesGeometricas.eixoYmr = (PropriedadesGeometricas.eixoYmi1 / (b1 / 2))


        Return propriedadesGeometricas
    End Function

End Class
