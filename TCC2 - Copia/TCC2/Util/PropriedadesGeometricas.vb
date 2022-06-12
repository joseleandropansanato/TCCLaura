Public Class PropriedadesGeometricas

    'serve para receber valor e mandar quando pedir 
    'Public Shared area As Double = 0
    'Public Shared areaA1 As Double = 0
    'Public Shared eixoXmi As Double = 0 'EIXO X MOMENTO DE INERCIA
    'Public Shared eixoXmi1 As Double = 0 'EIXO X MOMENTO DE INERCIA PRINCIPAL EM X
    'Public Shared eixoXrg As Double = 0 'EIXO X RAIO DE GIRAÇÃO
    'Public Shared eixoXmr As Double = 0 'EIXO X 
    'Public Shared eixoYmi As Double = 0 'EIXO Y MOMENTO DE INERCIA
    'Public Shared eixoYrg As Double = 0 'EIXO Y RAIO DE GIRAÇÃO
    'Public Shared eixoYmr As Double = 0 'EIXO Y MODULO DE RESISTENCIA
    'Public Shared eixoXie As Double = 0 'EIXO X INERCIA EFETIVA
    'Public Shared eixoYie As Double = 0 'EIXO Y INERCIA EFETIVA
    'Public Shared eixoYmie As Double = 0 'EIXO Y MOMENTO DE INERCIA EFETIVA EM Y
    'Public Shared eixoYmi1 As Double = 0 'EIXO Y MOMENTO DE INERCIA PRINCIPAL EM Y
    'Public Shared coefBy As Double = 0 'COEFICIENTE DO ESPAÇADOR
    'Public Shared Qx As Double = 0 'MOMENTO ESTÁTICO DE ÁREA EM X
    'Public Shared Qy As Double = 0 'MOMENTO ESTÁTICO DE ÁREA EM Y


    Private _area As Double = 0
    Private _areaA1 As Double = 0
    Private _eixoXmi As Double = 0 'EIXO X MOMENTO DE INERCIA
    Private _eixoXmi1 As Double = 0 'EIXO X MOMENTO DE INERCIA PRINCIPAL EM X
    Private _eixoXrg As Double = 0 'EIXO X RAIO DE GIRAÇÃO
    Private _eixoXmr As Double = 0 'EIXO X MODULO DE RESISTENCIA
    Private _eixoYmi As Double = 0 'EIXO Y MOMENTO DE INERCIA
    Private _eixoYrg As Double = 0 'EIXO Y RAIO DE GIRAÇÃO
    Private _eixoYmr As Double = 0 'EIXO Y MODULO DE RESISTENCIA
    Private _eixoXie As Double = 0 'EIXO X INERCIA EFETIVA
    Private _eixoYie As Double = 0 'EIXO Y INERCIA EFETIVA
    Private _eixoYmie As Double = 0 'EIXO Y MOMENTO DE INERCIA EFETIVA EM Y
    Private _eixoYmi1 As Double = 0 'EIXO Y MOMENTO DE INERCIA PRINCIPAL EM Y
    Private _intervalosM As Double = 0 'NUMERO DE INTERVALOS DE COMPRIMENT L1 QUE FICA DIVIDO O COMPRIMENTO TOTAL (L) DA PEÇA
    Private _w2 As Double = 0
    Private _coefBy As Double = 0 'COEFICIENTE DO ESPAÇADOR
    Private _qx As Double = 0 'MOMENTO ESTÁTICO DE ÁREA EM X
    Private _qy As Double = 0 'MOMENTO ESTÁTICO DE ÁREA EM Y

    Public Shared centroGeometricoX As Double = 0
    Public Shared centroGeometricoY As Double = 0

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
    Public Sub New()
    End Sub
    Public Property Area() As Double
        Get
            Return _area
        End Get
        Set(value As Double)
            _area = value
        End Set
    End Property
    Public Property AreaA1() As Double
        Get
            Return _areaA1
        End Get
        Set(value As Double)
            _areaA1 = value
        End Set
    End Property
    Public Property EixoXmi() As Double
        Get
            Return _eixoXmi
        End Get
        Set(value As Double)
            _eixoXmi = value
        End Set
    End Property
    Public Property EixoXmi1() As Double
        Get
            Return _eixoXmi1
        End Get
        Set(value As Double)
            _eixoXmi1 = value
        End Set
    End Property
    Public Property EixoXrg() As Double
        Get
            Return _eixoXrg
        End Get
        Set(value As Double)
            _eixoXrg = value
        End Set
    End Property
    Public Property EixoXmr() As Double
        Get
            Return _eixoXmr
        End Get
        Set(value As Double)
            _eixoXmr = value
        End Set
    End Property
    Public Property EixoYmi() As Double
        Get
            Return _eixoYmi
        End Get
        Set(value As Double)
            _eixoYmi = value
        End Set
    End Property
    Public Property EixoYmi1() As Double
        Get
            Return _eixoYmi1
        End Get
        Set(value As Double)
            _eixoYmi1 = value
        End Set
    End Property
    Public Property EixoYrg() As Double
        Get
            Return _eixoYrg
        End Get
        Set(value As Double)
            _eixoYrg = value
        End Set
    End Property
    Public Property EixoYmr() As Double
        Get
            Return _eixoYmr
        End Get
        Set(value As Double)
            _eixoYmr = value
        End Set
    End Property
    Public Property EixoXie() As Double
        Get
            Return _eixoXie
        End Get
        Set(value As Double)
            _eixoXie = value
        End Set
    End Property
    Public Property EixoYie() As Double
        Get
            Return _eixoYie
        End Get
        Set(value As Double)
            _eixoYie = value
        End Set
    End Property
    Public Property EixoYmie() As Double
        Get
            Return _eixoYmie
        End Get
        Set(value As Double)
            _eixoYmie = value
        End Set
    End Property
    Public Property CoefBy() As Double
        Get
            Return _coefBy
        End Get
        Set(value As Double)
            _coefBy = value
        End Set
    End Property
    Public Property IntervalosM() As Double
        Get
            Return _intervalosM
        End Get
        Set(value As Double)
            _intervalosM = value
        End Set
    End Property
    Public Property W2() As Double
        Get
            Return _w2
        End Get
        Set(value As Double)
            _w2 = value
        End Set
    End Property
    Public Property Qx() As Double
        Get
            Return _qx
        End Get
        Set(value As Double)
            _qx = value
        End Set
    End Property
    Public Property Qy() As Double
        Get
            Return _qy
        End Get
        Set(value As Double)
            _qy = value
        End Set
    End Property

    'ROTINA PARA CALCULAR O MOMENTO ESTATICO DE ÁREA POIS EU NÃO CONSEGUI DA OUTRA FORMA
    Class Qxy
        Dim _Qx As Double = 0
        Dim _Qy As Double = 0
        Dim _X_CGi As Double() = {0, 0, 0}
        Dim _Y_CGi As Double() = {0, 0, 0}
        Public Sub New()

        End Sub
        Public Property Qx() As Double
            Get
                Return _Qx
            End Get
            Set(value As Double)
                _Qx = value
            End Set
        End Property

        Public Property Qy() As Double
            Get
                Return _Qy
            End Get
            Set(value As Double)
                _Qy = value
            End Set
        End Property

        Public Property X_CGi() As Double()
            Get
                Return _X_CGi
            End Get
            Set(value As Double())
                _X_CGi = value
            End Set
        End Property
        Public Property Y_CGi() As Double()
            Get
                Return _Y_CGi
            End Get
            Set(value As Double())
                _Y_CGi = value
            End Set
        End Property
    End Class

    ''CÁLCULO DO MOMENTO ESTATICO DE ÁREA======================================
    Public Shared Function MomentoEstaticoDeArea(tipoSecao As Madeira.TipoSecao) As Qxy
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

        Dim area As Double = Calcular_Area(bases, alturas)

        Dim xCG_perfil = Calcular_CG(bases, alturas, x_CGi)
        Dim yCG_perfil = Calcular_CG(bases, alturas, y_CGi)

        Dim InerciaX = Inercia(bases, alturas, y_CGi)
        Dim InerciaY = Inercia(alturas, bases, x_CGi)

        Dim qxy = New Qxy()
        qxy.Qx = Momento_Estatico_Qx(bases, alturas, y_CGi, tipoSecao)
        qxy.Qy = Momento_Estatico_Qy(bases, alturas, x_CGi, tipoSecao)
        qxy.X_CGi = x_CGi
        qxy.Y_CGi = y_CGi
        'MessageBox.Show("Area: " + area.ToString + vbNewLine +
        '   "xCG: " + xCG_perfil.ToString + vbNewLine +
        '   "yCG: " + yCG_perfil.ToString + vbNewLine +
        '  "Ix: " + InerciaX.ToString + vbNewLine +
        '  "Iy: " + InerciaY.ToString + vbNewLine +
        '  "Qx: " + qxy.Qx.ToString + vbNewLine +
        '   "Qy: " + qxy.Qy.ToString + vbNewLine)

        Return qxy
    End Function

    Private Shared Function Calcular_Area(ByVal bases() As Double, ByVal alturas() As Double) As Double
        Dim A As Double = 0
        For i As Integer = 0 To bases.Length - 1
            A += bases(i) * alturas(i)
        Next
        Return A
    End Function

    Private Shared Function Calcular_CG(ByVal bases() As Double, ByVal alturas() As Double, ByVal CG_i() As Double) As Double
        Dim CG As Double = 0
        For i As Integer = 0 To bases.Length - 1
            CG += bases(i) * alturas(i) * CG_i(i)
        Next
        Return CG / Calcular_Area(bases, alturas)
    End Function

    Private Shared Function Inercia(ByVal bases() As Double, ByVal alturas() As Double, ByVal yCG_i() As Double) As Double
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

    Private Shared Function Momento_Estatico_Qx(ByVal bases() As Double, ByVal alturas() As Double, ByVal yCG_i() As Double, section As Madeira.TipoSecao) As Double

        Dim y_barra As Double, A_corte As Double, Q As Double
        Dim yCG As Double = Calcular_CG(bases, alturas, yCG_i)

        centroGeometricoY = yCG

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

    Private Shared Function Momento_Estatico_Qy(ByVal bases() As Double, ByVal alturas() As Double, ByVal xCG_i() As Double, ByVal section As String) As Double

        Dim x_barra As Double, A_corte As Double
        Dim xCG As Double = Calcular_CG(bases, alturas, xCG_i)
        centroGeometricoX = xCG

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

    'PROPRIEDADES GEOMÉTRICAS DAS SEÇÕES
    Public Function CalculoRetangular(baseX As Double, baseY As Double) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        'SEÇÃO RETANGULAR: PROPRIEDADES VALIDADAS
        proprGeometricas.Area = baseX * baseY
        proprGeometricas.EixoXmi = (baseX * baseY ^ 3) / 12
        proprGeometricas.EixoXrg = Math.Sqrt(proprGeometricas.EixoXmi / proprGeometricas.Area)
        proprGeometricas.EixoXmr = proprGeometricas.EixoXmi / (baseY / 2)
        proprGeometricas.EixoYmi = (baseY * baseX ^ 3) / 12
        proprGeometricas.EixoYrg = Math.Sqrt(proprGeometricas.EixoYmi / proprGeometricas.Area)
        proprGeometricas.EixoYmr = proprGeometricas.EixoYmi / (baseX / 2)

        Return proprGeometricas
    End Function
    Public Function CalculoCircular(diametro As Double) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        'SEÇÃO CIRCULAR:PROPRIEDADES VALIDADAS
        proprGeometricas.Area = System.Math.PI * (diametro / 2) ^ 2
        proprGeometricas.EixoXmi = (System.Math.PI * (diametro) ^ 4) / 64
        proprGeometricas.EixoXrg = Math.Sqrt(proprGeometricas.EixoXmi / proprGeometricas.Area)
        proprGeometricas.EixoXmr = proprGeometricas.EixoXmi / (diametro / 2)
        proprGeometricas.EixoYmi = (System.Math.PI * (diametro) ^ 4) / 64
        proprGeometricas.EixoYrg = Math.Sqrt(proprGeometricas.EixoYmi / proprGeometricas.Area)
        proprGeometricas.EixoYmr = proprGeometricas.EixoYmi / (diametro / 2)

        Return proprGeometricas
    End Function
    Public Function CalculoCompostoSecaoT(bf As Double, hf As Double, h As Double, bw As Double) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        Dim bases As Double() = {0, bw, bf}
        Dim alturas As Double() = {0, h, hf}
        Dim xCG = bases.Max / 2
        Dim x_CGi As Double() = {xCG, xCG, xCG}
        centroGeometricoX = Calcular_CG(bases, alturas, x_CGi)
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas.Sum - alturas(2) / 2}
        centroGeometricoY = Calcular_CG(bases, alturas, y_CGi)
        'SEÇÃO T: PROPRIEDADES VALIDADAS
        proprGeometricas.Area = Calcular_Area(bases, alturas)
        proprGeometricas.EixoXmi = Inercia(bases, alturas, y_CGi)
        proprGeometricas.EixoXie = proprGeometricas.EixoXmi * 0.95
        proprGeometricas.EixoXrg = Math.Sqrt(proprGeometricas.EixoXmi / proprGeometricas.Area)
        proprGeometricas.EixoXmr = proprGeometricas.EixoXmi / (Calcular_CG(bases, alturas, x_CGi))
        proprGeometricas.EixoYmi = Inercia(alturas, bases, x_CGi)
        proprGeometricas.EixoYie = proprGeometricas.EixoYmi * 0.95
        proprGeometricas.EixoYrg = Math.Sqrt(proprGeometricas.EixoYmi / proprGeometricas.Area)
        proprGeometricas.EixoYmr = proprGeometricas.EixoYmi / (Calcular_CG(bases, alturas, y_CGi))

        Return proprGeometricas
    End Function
    Public Function CalculoCompostoSecaoI(bf1 As Double, hf1 As Double, bf2 As Double, hf2 As Double, bw As Double, h As Double) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        Dim bases As Double() = {bf2, bw, bf1}
        Dim alturas As Double() = {hf2, h, hf1}
        Dim xCG As Double = bases.Max / 2
        Dim x_CGi As Double() = {xCG, xCG, xCG}
        centroGeometricoX = Calcular_CG(bases, alturas, x_CGi)
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas.Sum - alturas(2) / 2}
        centroGeometricoY = Calcular_CG(bases, alturas, y_CGi)
        'SEÇÃO I: PROPRIEDADES VALIDADAS
        proprGeometricas.Area = Calcular_Area(bases, alturas)
        proprGeometricas.EixoXmi = Inercia(bases, alturas, y_CGi)
        proprGeometricas.EixoXie = proprGeometricas.EixoXmi * 0.85
        proprGeometricas.EixoXrg = Math.Sqrt(proprGeometricas.EixoXmi / proprGeometricas.Area)
        proprGeometricas.EixoXmr = proprGeometricas.EixoXmi / (Calcular_CG(bases, alturas, x_CGi))
        proprGeometricas.EixoYmi = Inercia(alturas, bases, x_CGi)
        proprGeometricas.EixoYie = proprGeometricas.EixoYmi * 0.85
        proprGeometricas.EixoYrg = Math.Sqrt(proprGeometricas.EixoYmi / proprGeometricas.Area)
        proprGeometricas.EixoYmr = proprGeometricas.EixoYmi / (Calcular_CG(bases, alturas, y_CGi))

        Return proprGeometricas
    End Function

    Public Function CalculoCompostoSecaoCaixao(d As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, h1 As Double, h2 As Double, h3 As Double, h4 As Double) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        Dim bases As Double() = {b1, b2, b3, b4}
        Dim alturas As Double() = {h1, h2, h3, h4}
        Dim xCG = bases.Max / 2
        'Dim xCG = bases.Max 
        Dim base As Double = Math.Min(bases(0), bases(3))
        Dim x_CGi As Double() = {xCG, xCG - base / 2 + bases(1) / 2, xCG + base / 2 - bases(1) / 2, xCG}
        centroGeometricoX = Calcular_CG(bases, alturas, x_CGi)
        Dim y_CGi As Double() = {alturas(0) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) / 2, alturas(0) + alturas(1) + alturas(3) / 2}
        centroGeometricoY = Calcular_CG(bases, alturas, y_CGi)
        'seção Caixão
        proprGeometricas.Area = Calcular_Area(bases, alturas)
        proprGeometricas.EixoXmi = Inercia(bases, alturas, y_CGi)
        proprGeometricas.EixoXie = proprGeometricas.EixoXmi * 0.85
        proprGeometricas.EixoXrg = Math.Sqrt(proprGeometricas.EixoXmi / proprGeometricas.Area)
        proprGeometricas.EixoXmr = proprGeometricas.EixoXmi / (Calcular_CG(bases, alturas, x_CGi))
        proprGeometricas.EixoYmi = Inercia(alturas, bases, x_CGi)
        proprGeometricas.EixoYie = proprGeometricas.EixoYmi * 0.85
        proprGeometricas.EixoYrg = Math.Sqrt(proprGeometricas.EixoYmi / proprGeometricas.Area)
        proprGeometricas.EixoYmr = proprGeometricas.EixoYmi / (Calcular_CG(bases, alturas, y_CGi))

        Return proprGeometricas
    End Function
    Public Function Calculo2Elementos(l As Double, l1 As Double, a1 As Double, h1 As Double, b1 As Double, pilar As PropriedadesGeometricas.Pilar) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        '2 ELEMENTOS JUSTAPOSTOS:PROPRIEDADES VALIDADAS
        proprGeometricas.AreaA1 = b1 * h1
        proprGeometricas.Area = proprGeometricas.AreaA1 * 2
        proprGeometricas.EixoXmi1 = (b1 * h1 ^ 3) / 12
        proprGeometricas.EixoXmi = 2 * proprGeometricas.EixoXmi1
        proprGeometricas.EixoXmr = proprGeometricas.EixoXmi1 / (h1 / 2)
        proprGeometricas.EixoYmi1 = (b1 ^ 3 * h1) / 12
        proprGeometricas.EixoYmi = (2 * proprGeometricas.EixoYmi1) + (2 * proprGeometricas.AreaA1 * (a1 ^ 2))
        Select Case pilar
            Case Pilar.Espaçador
                proprGeometricas.CoefBy = proprGeometricas.EixoYmi1 * ((l / l1) ^ 2) / ((proprGeometricas.EixoYmi1 * ((l / l1) ^ 2)) + (1.25 * proprGeometricas.EixoYmi))
            Case Pilar.Chapa
                proprGeometricas.CoefBy = proprGeometricas.EixoYmi1 * ((l / l1) ^ 2) / ((proprGeometricas.EixoYmi1 * ((l / l1) ^ 2)) + (2.25 * proprGeometricas.EixoYmi))
        End Select
        proprGeometricas.EixoYie = proprGeometricas.CoefBy * proprGeometricas.EixoYmi
        proprGeometricas.EixoYmr = proprGeometricas.EixoYmi1 / (b1 / 2)
        proprGeometricas.IntervalosM = l / l1
        proprGeometricas.W2 = proprGeometricas.EixoYmi1 / (b1 / 2)

        Return proprGeometricas
    End Function
    Public Function Calculo3Elementos(l As Double, l1 As Double, a1 As Double, h1 As Double, b1 As Double, pilar As PropriedadesGeometricas.Pilar) As PropriedadesGeometricas
        Dim proprGeometricas = New PropriedadesGeometricas()
        '3 ELEMENTOS JUSTAPOSTOS:PROPRIEDADES VALIDADAS
        proprGeometricas.AreaA1 = b1 * h1
        proprGeometricas.Area = proprGeometricas.AreaA1 * 3
        proprGeometricas.EixoXmi1 = (b1 * h1 ^ 3) / 12
        proprGeometricas.EixoXmi = 3 * proprGeometricas.EixoXmi1
        proprGeometricas.EixoYmi1 = (b1 ^ 3 * h1) / 12
        proprGeometricas.EixoYmi = (3 * proprGeometricas.EixoYmi1) + (2 * proprGeometricas.AreaA1 * (a1 ^ 2))
        Select Case pilar
            Case Pilar.Espaçador
                proprGeometricas.CoefBy = proprGeometricas.EixoYmi1 * ((l / l1) ^ 2) / ((proprGeometricas.EixoYmi1 * ((l / l1) ^ 2)) + (1.25 * proprGeometricas.EixoYmi))
            Case Pilar.Chapa
                proprGeometricas.CoefBy = proprGeometricas.EixoYmi1 * ((l / l1) ^ 2) / ((proprGeometricas.EixoYmi1 * ((l / l1) ^ 2)) + (2.25 * proprGeometricas.EixoYmi))
        End Select
        proprGeometricas.EixoYie = proprGeometricas.CoefBy * proprGeometricas.EixoYmi
        proprGeometricas.EixoYmr = (proprGeometricas.EixoYmi1 / (b1 / 2))
        proprGeometricas.IntervalosM = l / l1
        proprGeometricas.W2 = proprGeometricas.EixoYmi1 / (b1 / 2)

        Return proprGeometricas
    End Function


End Class
