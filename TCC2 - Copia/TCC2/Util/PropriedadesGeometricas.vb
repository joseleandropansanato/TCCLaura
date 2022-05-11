
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

    'SEÇÃO RETANGULAR

    'SEÇÃO CAIXÃO


    'SEÇÃO T OU I
    Public Shared b1 As Double = 0 'bf2
    Public Shared b2 As Double = 0 'bw
    Public Shared b3 As Double = 0 'bf1
    Public Shared h1 As Double = 0 'hf2
    Public Shared h2 As Double = 0 'h
    Public Shared h3 As Double = 0 'hf1

    'caixão
    Public Shared b1c As Double = 0
    Public Shared b2c As Double = 0
    Public Shared b3c As Double = 0
    Public Shared b4c As Double = 0 'bf2-bw caixao
    Public Shared h1c As Double = 0
    Public Shared h2c As Double = 0
    Public Shared h3c As Double = 0
    Public Shared h4c As Double = 0 'hbf1 caixao



    Enum Pilar
        Espaçador
        Chapa
    End Enum

    ''CÁLCULO DO MOMENTO ESTATICO DE ÁREA======================================
    Public Function MomentoEstaticoDeArea(tipoSecao As Madeira.TipoSecao) As Double
        Dim bases As Double() = {0, 0, 0}
        Dim alturas As Double() = {0, 0, 0}
        Dim xCG As Double = bases.Max / 2
        Dim x_CGi As Double() = {0, 0, 0}
        Dim y_CGi As Double() = {0, 0, 0}

        'CÁLCULO DE Q:
        'Private Sub Pic_TmultS_Click(sender As Object, e As EventArgs) Handles Pic_TmultS.Click
        Select Case tipoSecao
            Case Madeira.TipoSecao.SecaoT, Madeira.TipoSecao.SecaoI
                'SEÇÃO I OU T ==========================================
                bases = {b1, b2, b3}
                alturas = {h1, h2, h3}

                'calcula as coordenadas x, y do CG de cada figura
                xCG = bases.Max / 2
                x_CGi = {xCG, xCG, xCG}
                y_CGi = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas.Sum - alturas(2) / 2}

            Case Madeira.TipoSecao.Caixao
                ''SEÇÃO CAIXAO ------------------------------------------
                bases = {b1c, b2c, b3c, b4c}
                alturas = {h1c, h2c, h3c, h4c}

                'calcula as coordenadas x,y do CG de cada figura
                xCG = bases.Max / 2
                Dim base As Double = Math.Min(bases(0), bases(3))
                x_CGi = {xCG, xCG - base / 2 + bases(1) / 2, xCG + base / 2 - bases(1) / 2, xCG}
                y_CGi = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) + alturas(3) / 2}

        End Select

        area = Calcular_Area(bases, alturas)

        Dim xCG_perfil = Calcular_CG(bases, alturas, x_CGi)
        Dim yCG_perfil = Calcular_CG(bases, alturas, y_CGi)

        Dim InerciaX = Inercia(bases, alturas, y_CGi)
        Dim InerciaY = Inercia(alturas, bases, x_CGi)


        Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, tipoSecao)
        Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, tipoSecao)

        'Dim Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, "I")
        'Dim Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, "I")
        'Dim Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, "caixao")
        'Dim Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, "caixao")

        MessageBox.Show("Area: " + area.ToString + vbNewLine +
                        "xCG: " + xCG_perfil.ToString + vbNewLine +
                        "yCG: " + yCG_perfil.ToString + vbNewLine +
                        "Ix: " + InerciaX.ToString + vbNewLine +
                        "Iy: " + InerciaY.ToString + vbNewLine +
                        "Qx: " + Qx.ToString + vbNewLine +
                        "Qy: " + Qy.ToString + vbNewLine)

        Return MomentoEstaticoDeArea
    End Function

    Private Function Calcular_Area(ByVal bases() As Double, ByVal alturas() As Double) As Double
        Dim A As Double = 0
        For i As Integer = 0 To bases.Length - 1
            A += bases(i) * alturas(i)
        Next
        Return A
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

    '' FIM DO CÁLCULO DO MOMENTO ESTATICO DE ÁREA======================================

    'PROPRIEDADES GEOMÉTRICAS DA PEÇA 
    Public Function CalculoRetangular(baseX As Double, baseY As Double, altura As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        'seção retangular
        area = baseX * baseY 'OK
        eixoXmi = (baseX * baseY ^ 3) / 12 'OK
        eixoXrg = Math.Sqrt(eixoXmi / area)
        eixoXmr = eixoXmi / (altura / 2)
        eixoYmi = (baseY * baseX ^ 3) / 12 'OK
        eixoYrg = Math.Sqrt(eixoYmi / area)
        eixoYmr = eixoYmi / (altura / 2)

        Return propriedadesGeometricas
    End Function

    Public Function CalculoCircular(diametro As Double, altura As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        'seção cirucular
        area = System.Math.PI * (diametro / 2) ^ 2 'OK
        eixoXmi = (System.Math.PI * (diametro) ^ 4) / 64 'OK
        eixoXrg = Math.Sqrt(eixoXmi / area)
        eixoXmr = eixoXmi / (altura / 2)
        eixoYmi = (System.Math.PI * (diametro) ^ 4) / 64 'OK
        eixoYrg = Math.Sqrt(eixoYmi / area)
        eixoYmr = eixoYmi / (altura / 2)


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
        area = Calcular_Area(bases, alturas)
        eixoXmi = Inercia(bases, alturas, y_CGi)
        eixoXie = eixoXmi * 0.95
        eixoXrg = Math.Sqrt(eixoXie / area)
        eixoXmr = eixoXie / (h / 2)
        eixoYmi = Inercia(alturas, bases, x_CGi)
        eixoYie = eixoYmi * 0.95
        eixoYrg = Math.Sqrt(eixoYie / area)
        eixoYmr = eixoYie / (h / 2)

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

        area = Calcular_Area(bases, alturas)
        eixoXmi = Inercia(bases, alturas, y_CGi)
        eixoXie = eixoXmi * 0.85
        eixoXrg = Math.Sqrt(eixoXie / area)
        eixoXmr = eixoXie / ((d - hf1 - hf2) / 2)
        eixoYmi = Inercia(bases, alturas, x_CGi)
        eixoYie = eixoYmi * 0.85
        eixoYrg = Math.Sqrt(eixoYie / area)
        eixoYmr = eixoYie / ((d - hf1 - hf2) / 2)

        Return propriedadesGeometricas
    End Function

    Public Function CalculoCompostoSecaoCaixao(d As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, h1 As Double, h2 As Double, h3 As Double, h4 As Double)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        Dim bases As Double() = {b1, b2, b3, b4}
        Dim alturas As Double() = {h1, h2, h3, h4}
        Dim xCG = bases.Max / 2
        Dim base As Double = Math.Min(bases(0), bases(3))
        Dim x_CGi As Double() = {xCG, xCG - base / 2 + bases(1) / 2, xCG + base / 2 - bases(1) / 2, xCG}
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) + alturas(3) / 2}

        'seção Caixão
        area = Calcular_Area(bases, alturas)
        eixoXmi = Inercia(bases, alturas, y_CGi)
        eixoXie = eixoXmi * 0.85
        eixoXrg = Math.Sqrt(eixoXie / area)
        eixoXmr = eixoXie / (h3 / 2)
        eixoYmi = Inercia(bases, alturas, x_CGi)
        eixoYie = eixoYmi * 0.85
        eixoYrg = Math.Sqrt(eixoYie / area)
        eixoYmr = eixoYie / (h3 / 2)

        Return propriedadesGeometricas
    End Function


    Public Function Calculo2Elementos(l As Double, l1 As Double, a As Double, a1 As Double, h As Double, h1 As Double, b1 As Double, pilar As PropriedadesGeometricas.Pilar)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        '2 Elementos Justapostos
        areaA1 = b1 * h1
        area = areaA1 * 2
        eixoXmi1 = (b1 * h1 ^ 3) / 12
        eixoXmi = 2 * eixoXmi1
        eixoYmi1 = (b1 ^ 3 * h1) / 12
        eixoYmi = (2 * eixoYmi1) + (2 * areaA1 * (a1 ^ 2))
        Select Case pilar
            Case Pilar.Espaçador
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * eixoYmi))
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * eixoYmi))
            Case PropriedadesGeometricas.Pilar.Chapa
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * eixoYmi))
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * eixoYmi))
        End Select
        eixoYie = coefBy * eixoYmi
        eixoYmr = eixoYmi1 / (b1 / 2)


        Return propriedadesGeometricas
    End Function

    Public Function Calculo3Elementos(l As Double, l1 As Double, a As Double, a1 As Double, h As Double, h1 As Double, b1 As Double, pilar As PropriedadesGeometricas.Pilar)
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        '3 Elementos Justapostos
        areaA1 = b1 * h1
        area = areaA1 * 3
        eixoXmi1 = (b1 * h1 ^ 3) / 12
        eixoXmi = 3 * eixoXmi1
        eixoYmi1 = (b1 ^ 3 * h1) / 12
        eixoYmi = (3 * eixoYmi1) + (2 * areaA1 * (a1 ^ 2))
        Select Case pilar
            Case PropriedadesGeometricas.Pilar.Espaçador
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * eixoYmi))
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (1.25 * eixoYmi))
            Case PropriedadesGeometricas.Pilar.Chapa
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * eixoXmi))
                coefBy = eixoYmi1 * ((l / l1) ^ 2) / ((eixoYmi1 * ((l / l1) ^ 2)) + (2.25 * eixoYmi))
        End Select
        eixoYie = coefBy * eixoYmi
        eixoYmr = (eixoYmi1 / (b1 / 2))


        Return propriedadesGeometricas
    End Function

End Class
