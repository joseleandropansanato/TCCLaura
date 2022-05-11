Module PropriedadesResistencia
    'definições: lista com cada madeira que será necessario para o cálculo
    Public Function ListaDenominacaoMadeira() As IEnumerable(Of Madeira)
        Return New List(Of Madeira) From
            {
                New Madeira(0, "Angelim araroba", 50.5, 69.2, 3.1, 7.1, 12876, 688),
                New Madeira(1, "Angelim ferro", 79.9, 117.8, 3.7, 11.8, 208727, 1170),
                New Madeira(2, "Angelim pedra", 59.8, 75.5, 3.5, 8.8, 12912, 694),
                New Madeira(3, "Angelim pedra verdadeiro", 76.7, 104.9, 4.8, 11.3, 16694, 1170),
                New Madeira(4, "Branquilho", 48.1, 87.9, 3.2, 9.8, 13481, 803),
                New Madeira(5, "Cafearana", 59.1, 79.7, 3.0, 5.9, 14098, 677),
                New Madeira(6, "Canafístula", 52.0, 84.9, 6.2, 11.1, 14613, 871),
                New Madeira(7, "Casca Grossa", 56.0, 120.2, 4.1, 8.2, 16224, 801),
                New Madeira(8, "Castelo", 54.8, 99.5, 7.5, 12.8, 11105, 759),
                New Madeira(9, "Cedro amargo", 39.0, 58.1, 3.0, 6.1, 9839, 504),
                New Madeira(10, "Cedo doce", 31.5, 71.4, 3.0, 5.6, 8058, 500),
                New Madeira(11, "Champagne", 93.2, 133.5, 2.9, 10.7, 23002, 1090),
                New Madeira(12, "Cupiúna", 54.4, 62.1, 3.3, 10.4, 13627, 838),
                New Madeira(13, "Catiúna", 83.8, 86.2, 3.3, 11.1, 1942, 1221),
                New Madeira(14, "Eucalipto Alba", 47.3, 69.4, 4.6, 9.5, 13409, 705),
                New Madeira(15, "Eucalipto Camaldulensis", 48.0, 78.1, 4.6, 9.0, 13286, 899),
                New Madeira(16, "Eucalipto Citriodora", 62.0, 123.6, 3.9, 10.7, 18421, 999),
                New Madeira(17, "Eucalipto Cloeziana ", 51.8, 90.8, 4.0, 10.5, 13963, 822),
                New Madeira(18, "Eucalipto Dunni", 48.9, 139.2, 6.9, 9.8, 18029, 690),
                New Madeira(19, "Eucalipto Grandis", 40.3, 70.2, 2.6, 7.0, 12813, 640),
                New Madeira(20, "Eucalipto Maculata", 63.5, 115.6, 4.1, 10.6, 18099, 931),
                New Madeira(21, "Eucalipto Maidene", 48.3, 83.7, 4.8, 10.3, 14431, 924),
                New Madeira(22, "Eucalipto Microcorys", 54.9, 118.6, 4.5, 10.3, 16782, 929),
                New Madeira(23, "Eucalipto Paniculata", 72.7, 147.4, 4.7, 12.4, 19881, 1087),
                New Madeira(24, "Eucalipto Propinqua", 51.6, 89.1, 4.7, 9.7, 15561, 952),
                New Madeira(25, "Eucalipto Punctata", 78.5, 125.6, 6.0, 12.9, 19360, 948),
                New Madeira(26, "Eucalipto Saligna", 46.8, 95.5, 4.0, 8.2, 14933, 731),
                New Madeira(27, "Eucalipto Terticornis", 57.7, 115.9, 4.6, 9.7, 17198, 899),
                New Madeira(28, "Eucalipto Triantha", 53.9, 100.9, 2.7, 9.2, 14617, 755),
                New Madeira(29, "Eucalipto Umbra", 42.7, 90.4, 3.0, 9.4, 14577, 889),
                New Madeira(30, "Eucalipto Urophylla", 46.0, 85.1, 4.1, 8.3, 13166, 739),
                New Madeira(31, "Garapa Roraima", 78.4, 108.0, 6.9, 11.9, 18359, 892),
                New Madeira(32, "Guaiçara", 71.4, 115.6, 4.2, 12.5, 14624, 825),
                New Madeira(33, "Guarucaia", 62.4, 70.9, 5.5, 15.5, 17212, 919),
                New Madeira(34, "Ipê", 76.0, 96.8, 3.1, 13.1, 18011, 1068),
                New Madeira(35, "Jatobá", 93.3, 157.5, 3.2, 15.7, 23607, 1074),
                New Madeira(36, "Louro Preto", 56.5, 111.9, 3.3, 9.0, 14185, 684),
                New Madeira(37, "Maçaranduba", 82.9, 138.5, 5.4, 14.9, 22733, 1143),
                New Madeira(38, "Mandioqueira", 71.4, 89.1, 2.7, 10.6, 18971, 856),
                New Madeira(39, "Oiticica Amarela", 69.9, 82.5, 3.9, 10.6, 14719, 756),
                New Madeira(40, "Quarubarana", 37.8, 58.1, 2.6, 5.8, 9067, 544),
                New Madeira(41, "Sucupira", 95.2, 123.4, 3.4, 11.8, 21724, 1106),
                New Madeira(42, "Tatajuba", 79.5, 78.8, 3.9, 12.2, 19583, 940),
                New Madeira(44, "Pinho do Paraná", 40.9, 93.1, 1.6, 8.8, 15225, 580),
                New Madeira(45, "Pinus Caribea", 35.4, 64.8, 3.2, 7.8, 8431, 579),
                New Madeira(46, "Pinus Bahamensis", 32.6, 52.7, 2.4, 6.8, 7110, 537),
                New Madeira(47, "Pinus Hondurensis", 42.3, 50.3, 2.6, 7.8, 9868, 535),
                New Madeira(48, "Pinus Elliotti", 40.4, 66.0, 2.5, 7.4, 11889, 560),
                New Madeira(49, "Pinus Oocarpa", 43.6, 60.9, 2.5, 8.0, 10904, 538),
                New Madeira(50, "Pinus Taeda", 44.4, 82.8, 2.8, 7.7, 13304, 645)
                }
    End Function
    'RESISTÊNCIA DE CÁLCULO:
    Private Function SelectKmod1(tipo As Integer) As Kmod1
        Dim kmod1 As New Kmod1
        Select Case tipo
            Case 0
                Return New Kmod1(0.6, 0.7, 0.85, 1.0, 1.1)
            Case 1
                Return New Kmod1(0.3, 0.45, 0.65, 1.0, 1.1)
            Case Else
                Return kmod1
        End Select
    End Function
    Private Function SelectKmod2(tipo As Integer) As Kmod2
        Dim kmod2 As New Kmod2
        Select Case tipo
            Case 0
                Return New Kmod2(1.0, 0.9, 0.8, 0.7)
            Case 1
                Return New Kmod2(1.0, 0.95, 0.93, 0.9)
            Case Else
                Return kmod2

        End Select
    End Function
    Private Function SelectKmod3(tipo As Integer) As Kmod3
        Dim kmod3 As New Kmod3
        Select Case tipo
            Case 0
                Return New Kmod3(0.9, 0.85, 0.7, 0.75)
            Case 1
                Return New Kmod3(1.0, 0.9, 0.85, 0.8)
            Case Else
                Return kmod3
        End Select
    End Function
    Public Function Kmod1Final(tipo As Integer, kmod1 As Integer) As Double

        Select Case kmod1
            Case 0
                Return SelectKmod1(tipo).Permanente
            Case 1
                Return SelectKmod1(tipo).LongaDuracao
            Case 2
                Return SelectKmod1(tipo).MediaDuracao
            Case 3
                Return SelectKmod1(tipo).CurtaDuracao
            Case 4
                Return SelectKmod1(tipo).Instantanea
        End Select
    End Function

    Public Function Kmod2Final(tipo As Integer, kmod2 As Integer) As Double
        Select Case kmod2
            Case 0
                Return SelectKmod2(tipo).One
            Case 1
                Return SelectKmod2(tipo).Two
            Case 2
                Return SelectKmod2(tipo).Three
            Case 3
                Return SelectKmod2(tipo).Four
        End Select

    End Function
    Public Function Kmod3Final(tipo As Integer, kmod3 As Integer) As Double
        Select Case kmod3
            Case 0
                Return SelectKmod3(tipo).Se
            Case 1
                Return SelectKmod3(tipo).S1
            Case 2
                Return SelectKmod3(tipo).S2
            Case 3
                Return SelectKmod3(tipo).S3
        End Select
    End Function

    Public Function resistCalCompParalela(kmod As Double, resistCompParalela As Double) As Double
        Return (resistCompParalela * kmod) / 1.4
    End Function

    Public Function resistCalTracaoParalela(kmod As Double, resistTracaoParalela As Double) As Double
        Return (resistTracaoParalela * kmod) / 1.8
    End Function
    Public Function resistCalTracaoNormal(kmod As Double, resistTracaoNormal As Double) As Double
        Return (resistTracaoNormal * kmod) / 1.8
    End Function
    Public Function resistCalAoCisalhamento(kmod As Double, resistAoCisalhamento As Double) As Double
        Return (resistAoCisalhamento * kmod) / 1.8
    End Function
    Public Function ModElasticidade(kmod As Double, modDeElasticidade As Double) As Double
        Return (kmod * modDeElasticidade)
    End Function
    Public Function ModElasticidadeTransversal(modDeElasticidade As Double) As Double
        Return (modDeElasticidade / 15)
    End Function

    'PROPRIEDADES GEOMÉTRICAS DA PEÇA 
    Public Function CalculoRetangular(baseX As Double, baseY As Double, altura As Double) As PropriedadesGeometricas
        Dim propriedadesGeometricas = New PropriedadesGeometricas()

        'TODAS AS LINHAS OK TIVERAM AS CONTAS VALIDADAS NO EXCEL DO PROFESSOR OU MANUALMENTE 
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


    Public Function ElementoJustaposto(l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String) As Boolean

        If ((9 * b1 <= l) And (l <= 18 * b1) And (a <= 3 * b1) And (elementoFixacaoSelecionado = "Espacador Interposto")) Then
            MessageBox.Show("A NBR 7190:1997 dipões que a verificação da estabilidade local dos trechos pode ser dispensada quando forem atendidas as seguintes condições:
                            9b1≤L1≤18b1 e
                            a≤3b1: para espaçadores interpostos")
            Return False

        ElseIf ((9 * b1 <= l) And (l <= 18 * b1) And (a <= 6 * b1) And (elementoFixacaoSelecionado = "Chapas Laterais de Fixação")) Then
            MessageBox.Show("A NBR 7190:1997 dipões que a verificação da estabilidade local dos trechos pode ser dispensada quando forem atendidas as seguintes condições:
                            9b1≤L1≤18b1 e
                            a≤6b1: para chapas laterais de fixação")
            Return False

        Else
            Return True
        End If
    End Function

    'Public Function TensaoTracao(normalTracao As Double, baseX As Double, baseY As Double, diametro As Double, d As Double, bw As Double, bf1 As Double, altura As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Tracao
    Public Function TensaoTracao(normalTracao As Double, baseX As Double, baseY As Double, c As Double, diametro As Double, d As Double, bw As Double, bf1 As Double, altura As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao, l As Double, b1 As Double, a As Double, elementoFixacaoSelecionado As String, lvinculado As Double, momentoFletorX As Double, momentoFletorY As Double, h1 As Double) As Tracao
        Dim tracao = New Tracao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                tracao.tensaoTracao = ((normalTracao) / (PropriedadesGeometricas.area)) 'OK
                Tracao.esbeltezTracaoX = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area))) 'OK
                Tracao.esbeltezTracaoY = (c * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area))) 'OK

            Case Madeira.TipoSecao.Circular
                tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (diametro / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (altura / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoT
                tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (bw / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.SecaoI
                tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.Caixao
                tracao.tensaoTracao = normalTracao / PropriedadesGeometricas.area
                Tracao.esbeltezTracaoX = (bf1 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                Tracao.esbeltezTracaoY = (d / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))

            Case Madeira.TipoSecao.ElementosJustaposto2
                'botar o espaçador aqui tb
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    Tracao.esbeltezTracaoX = ((lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area)))
                    Tracao.esbeltezTracaoY = ((lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area)))
                    'eixo X::
                    Tracao.eixoXPecaComposta = (normalTracao / PropriedadesGeometricas.area) + (momentoFletorX / PropriedadesGeometricas.eixoXmi) * (h1 / 2)
                    'eixo y::
                    Tracao.eixoYPecaComposta = (normalTracao / PropriedadesGeometricas.area) + ((momentoFletorY * PropriedadesGeometricas.eixoYmi1) / PropriedadesGeometricas.eixoYie * PropriedadesGeometricas.eixoYmr) + (momentoFletorY / (2 * PropriedadesGeometricas.area * PropriedadesGeometricas.areaA1)) * (1 - 2 * (PropriedadesGeometricas.eixoYmi1 / PropriedadesGeometricas.eixoYie))
                End If
            Case Madeira.TipoSecao.ElementosJustaposto3
                If ElementoJustaposto(l, b1, a, elementoFixacaoSelecionado) Then
                    'tracao.TensaoTracao = normalTracao / propriedadesGeometricas.Area
                    Tracao.esbeltezTracaoX = (lvinculado) / PropriedadesGeometricas.eixoXrg
                    Tracao.esbeltezTracaoY = (lvinculado) / (Math.Sqrt(PropriedadesGeometricas.eixoYie / PropriedadesGeometricas.area))
                    'eixo X::
                    Tracao.eixoXPecaComposta = (normalTracao / PropriedadesGeometricas.area) + (momentoFletorX / PropriedadesGeometricas.eixoXmi) * (h1 / 2)
                    'eixo y::
                    Tracao.eixoYPecaComposta = (normalTracao / PropriedadesGeometricas.area) + ((momentoFletorY * PropriedadesGeometricas.eixoYmi1) / PropriedadesGeometricas.eixoYie * PropriedadesGeometricas.eixoYmr) + (momentoFletorY / (3 * PropriedadesGeometricas.area * PropriedadesGeometricas.areaA1)) * (1 - 2 * (PropriedadesGeometricas.eixoYmi1 / PropriedadesGeometricas.eixoYie))
                End If
        End Select

        Return tracao
    End Function

    Public Function TensaoCompressao(normalCompressao As Double, momentoFletorX As Double, momentoFletorY As Double, lvinculado As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao,
                                     MomentoCargaPermanenteX As Double, MomentoCargaPermanenteY As Double, MomentoCargaVariavelX As Double, MomentoCargaVariavelY As Double, NormalCargaPermanente As Double, NormalCargaVariavel As Double,
                                     coefInfluencia As Double, f1 As Double, f2 As Double) As Compressao

        Dim compressao = New Compressao()
        Dim mdX As Double = 0
        Dim mdY As Double = 0

        Select Case tipoSecao

            Case Madeira.TipoSecao.Retangular

                Compressao.esbeltezCompressaoX = (lvinculado * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoXmi / PropriedadesGeometricas.area))) 'OK
                Compressao.esbeltezCompressaoY = (lvinculado * 100 / (Math.Sqrt(PropriedadesGeometricas.eixoYmi / PropriedadesGeometricas.area))) 'OK

                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edy = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * Compressao.edy
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdX = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdY = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.Circular
                Compressao.esbeltezCompressaoX = lvinculado / PropriedadesGeometricas.eixoXrg
                Compressao.esbeltezCompressaoY = lvinculado / PropriedadesGeometricas.eixoYrg

                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edy = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * Compressao.edy
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdX = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdY = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.SecaoT
                Compressao.esbeltezCompressaoX = 0
                Compressao.esbeltezCompressaoY = 0
                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edy = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * Compressao.edy
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdX = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdY = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.
                                     Insira novamente a seção transversal da peça para prosseguir com o dimensionamento.")


                End If

            Case Madeira.TipoSecao.SecaoI
                Compressao.esbeltezCompressaoX = 0
                Compressao.esbeltezCompressaoY = 0

                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edy = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * Compressao.edy
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdX = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdY = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.Caixao
                Compressao.esbeltezCompressaoX = 0
                Compressao.esbeltezCompressaoY = 0

                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X
                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorX / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edx = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    mdX = normalCompressao * Compressao.edx
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y
                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER MEDIANAMENTE
                    Compressao.ei = momentoFletorY / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.edy = (Compressao.ei + Compressao.ea) * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    mdY = normalCompressao * Compressao.edy
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaX = ((System.Math.PI) * PropriedadesGeometricas.eixoXmi * PropriedadesGeometricas.eixoXmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteX + MomentoCargaVariavelX) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteX / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdX = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaX / (Compressao.forcaElasticaX - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMxd = (mdX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa


                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                    Compressao.forcaElasticaY = ((System.Math.PI) * PropriedadesGeometricas.eixoYmi * PropriedadesGeometricas.eixoYmr) / (lvinculado ^ 2)
                    Compressao.ea = lvinculado / 300
                    'O MOMENTO CARACTERISTICO O USUARIO VAI TER QUE INSERIR QUANDO A PEÇA DER ESBELTA E MOMENTO CARACTERISTICOXAC DEVIDO A PARCELA ACIDENTAL
                    Compressao.ei = (MomentoCargaPermanenteY + MomentoCargaVariavelY) / normalCompressao 'se esse valor for menor que altura/30 usa h/30 fazer um if else aquior 30
                    Compressao.eig = (MomentoCargaPermanenteY / NormalCargaPermanente)
                    Compressao.ec = (Compressao.ea + Compressao.eig) * System.Math.Exp(((coefInfluencia * (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel)) / (NormalCargaPermanente - (NormalCargaPermanente + (f1 + f2) * NormalCargaVariavel))) - 1)
                    Compressao.e1ef = Compressao.ea + Compressao.ei + Compressao.ec
                    mdY = normalCompressao * Compressao.e1ef * (Compressao.forcaElasticaY / (Compressao.forcaElasticaY - normalCompressao))
                    Compressao.tensaoCompressao = normalCompressao / PropriedadesGeometricas.area
                    Compressao.tensaoMyd = (mdY / PropriedadesGeometricas.eixoXmr) / 10 ^ 6 'w é o modulo de resistencia divindo por 10^6 para transformar em MPa

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.ElementosJustaposto2
                Compressao.esbeltezCompressaoX = 0
                Compressao.esbeltezCompressaoY = 0

                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y

                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

            Case Madeira.TipoSecao.ElementosJustaposto3
                Compressao.esbeltezCompressaoX = 0
                Compressao.esbeltezCompressaoY = 0

                If Compressao.esbeltezCompressaoX <= 40 Then 'PEÇA CURTA EM X

                ElseIf Compressao.esbeltezCompressaoY <= 40 Then 'PEÇA CURTA EM y

                ElseIf 40 < Compressao.esbeltezCompressaoX <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM X

                ElseIf 40 < Compressao.esbeltezCompressaoY <= 80 Then 'ELEMENTO MEDIANAMENTE ESBELTO EM Y

                ElseIf 80 < Compressao.esbeltezCompressaoX <= 140 Then 'ELEMENTO ESBELTO EM X
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf 80 < Compressao.esbeltezCompressaoY <= 140 Then 'ELEMENTO ESBELTO EM Y
                    MessageBox.Show("Aviso: o Elemento é esbelto.")

                ElseIf Compressao.esbeltezCompressaoX > 140 And Compressao.esbeltezCompressaoY > 140 Then
                    MessageBox.Show("Valores acima de 140 não devem ser considerados pois violam o estado limite de serviço imposto pela NBR.")

                End If

        End Select

        Return compressao
    End Function

    Public Function TensaoCisalhamento(forcaCortanteX As Double, forcaCortanteY As Double, diametro As Double, q As Qxy, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Cisalhamento
        Dim cisalhamento = New Cisalhamento()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Cisalhamento.tensaoCisalhanteX = ((2 / 3) * (forcaCortanteX / PropriedadesGeometricas.area) / 10 ^ 6)
                Cisalhamento.tensaoCisalhanteY = ((2 / 3) * (forcaCortanteY / PropriedadesGeometricas.area) / 10 ^ 6)

            Case Madeira.TipoSecao.Circular
                Cisalhamento.tensaoCisalhanteX = ((4 * forcaCortanteX) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6
                Cisalhamento.tensaoCisalhanteY = ((4 * forcaCortanteY) / (3 * System.Math.PI * (diametro / 2))) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Cisalhamento.tensaoCisalhanteX = forcaCortanteX * q.Qx / PropriedadesGeometricas.eixoXmi
                Cisalhamento.tensaoCisalhanteY = forcaCortanteY * q.Qy / PropriedadesGeometricas.eixoYmi

            Case Madeira.TipoSecao.SecaoI
                Cisalhamento.tensaoCisalhanteX = forcaCortanteX * q.Qx / PropriedadesGeometricas.eixoXmi
                Cisalhamento.tensaoCisalhanteY = forcaCortanteY * q.Qy / PropriedadesGeometricas.eixoYmi

            Case Madeira.TipoSecao.Caixao
                Cisalhamento.tensaoCisalhanteX = forcaCortanteX * q.Qx / PropriedadesGeometricas.eixoXmi
                Cisalhamento.tensaoCisalhanteY = forcaCortanteY * q.Qy / PropriedadesGeometricas.eixoYmi

            Case Madeira.TipoSecao.ElementosJustaposto2
                Cisalhamento.tensaoCisalhanteX = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                Cisalhamento.tensaoCisalhanteX = 0
        End Select

        Return cisalhamento
    End Function

    Public Function TensaoFlexaoSimples(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2

            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flexao
    End Function

    Public Function TensaoFlexaoObliqua(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2


            Case Madeira.TipoSecao.ElementosJustaposto3

        End Select

        Return flexao
    End Function


    Public Function TensaoFlexotracao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3
                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0

        End Select

        Return flexao
    End Function

    Public Function TensaoFlexocompressao(momentoFletorX As Double, momentoFletorY As Double, propriedadesGeometricas As PropriedadesGeometricas, tipoSecao As Madeira.TipoSecao) As Flexao
        Dim flexao = New Flexao()

        Select Case tipoSecao
            Case Madeira.TipoSecao.Retangular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Circular
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoT
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.SecaoI
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.Caixao
                Flexao.tensaoFlexaoX = (momentoFletorX / PropriedadesGeometricas.eixoXmr) / 10 ^ 6
                Flexao.tensaoFlexaoY = (momentoFletorY / PropriedadesGeometricas.eixoYmr) / 10 ^ 6

            Case Madeira.TipoSecao.ElementosJustaposto2
                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0

            Case Madeira.TipoSecao.ElementosJustaposto3

                Flexao.tensaoFlexaoX = 0
                Flexao.tensaoFlexaoY = 0
        End Select

        Return flexao
    End Function

    Public Function verificaVazio(value As String) As Double
        If value = "" Then
            Return 0
        End If
        Return value
    End Function

    'TRAÇÃO:
    Public Function validarTensaoTracao(tensaoTracao As Double, resistenciaCalTracaoParalela As Double) As Boolean
        Return tensaoTracao <= resistenciaCalTracaoParalela
    End Function

    Public Function validarEsbeltezTracaoX(esbeltezTracaoX As Double)
        Return esbeltezTracaoX <= 173
    End Function

    Public Function validarEsbeltezTracaoY(esbeltezTracaoY As Double)
        Return esbeltezTracaoY <= 173
    End Function

    'COMPRESSÃO
    Public Function validarTensaoCompressaoCurtaX(TensaoCompressaoX As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return TensaoCompressaoX <= resistenciaCalCompressaoParalela
    End Function

    Public Function validarTensaoCompressaoCurtaY(TensaoCompressaoY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return TensaoCompressaoY <= resistenciaCalCompressaoParalela
    End Function

    Public Function validarTensaoCompressaoMedEsbX(TensaoCompressaoX As Double, TensaoMdX As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return (TensaoCompressaoX / resistenciaCalCompressaoParalela) + (TensaoMdX / resistenciaCalCompressaoParalela) <= 1
    End Function

    Public Function validarTensaoCompressaoMedEsbY(TensaoCompressaoY As Double, TensaoMdY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return (TensaoCompressaoY / resistenciaCalCompressaoParalela) + (TensaoMdY / resistenciaCalCompressaoParalela) <= 1
    End Function

    Public Function validarTensaoCompressaoEsbX(TensaoCompressaoX As Double, TensaoMdX As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return (TensaoCompressaoX / resistenciaCalCompressaoParalela) + (TensaoMdX / resistenciaCalCompressaoParalela) <= 1
    End Function

    Public Function validarTensaoCompressaoEsbY(TensaoCompressaoY As Double, TensaoMdY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return (TensaoCompressaoY / resistenciaCalCompressaoParalela) + (TensaoMdY / resistenciaCalCompressaoParalela) <= 1
    End Function


    'CISALHAMENTO
    Public Function validarTensaoCisalhanteX(tensaoCisalhante As Double, resistenciaCalAoCisalhamento As Double) As Boolean
        Return tensaoCisalhante <= resistenciaCalAoCisalhamento
    End Function

    Public Function validarTensaoCisalhanteY(tensaoCisalhante As Double, resistenciaCalAoCisalhamento As Double) As Boolean
        Return tensaoCisalhante <= resistenciaCalAoCisalhamento
    End Function

    'FLEXÃO SIMPLES RETA
    Public Function validarTensaoFlexaoSimplesRetaX(tensaoFlexaoX As Double, resistenciaCalTracaoParalela As Double) As Boolean
        Return tensaoFlexaoX <= resistenciaCalTracaoParalela
    End Function

    Public Function validarTensaoFlexaoSimplesRetaY(tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return tensaoFlexaoY <= resistenciaCalCompressaoParalela
    End Function

    'FLEXÃO SIMPLES obliqua
    'km é 0.5 se a seção for retangular e outras seções 1
    'km é 0 caso a seção nao for obliqua
    Public Function definirKm(secao As Madeira.TipoSecao) As Double
        Select Case secao
            Case Madeira.TipoSecao.Retangular
                Return 0.5
            Case Madeira.TipoSecao.Circular, Madeira.TipoSecao.SecaoT, Madeira.TipoSecao.SecaoI, Madeira.TipoSecao.Caixao
                Return 1
        End Select
    End Function

    Public Function validarTensaoFlexaoSimplesObliquaX(tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalTracaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return tensaoFlexaoX + (definirKm(tipoSecao) * tensaoFlexaoY) <= resistenciaCalTracaoParalela
    End Function

    Public Function validarTensaoFlexaoSimplesObliquaY(tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (definirKm(tipoSecao) * tensaoFlexaoX) + tensaoFlexaoY <= resistenciaCalCompressaoParalela
    End Function

    'FLEXÃO COMPOSTA
    'km é 0.5 se a seção for retangular e se for obliqua outras seções 1
    'km é 0 caso a seção nao for obliqua
    Public Function validarTensaoFlexotracaoX(tensaoTracao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalTracaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoTracao / resistenciaCalTracaoParalela) + (tensaoFlexaoX / resistenciaCalTracaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoY / resistenciaCalTracaoParalela)) <= 1
    End Function
    Public Function validarTensaoFlexotracaoY(tensaoTracao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalTracaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoTracao / resistenciaCalTracaoParalela) + (tensaoFlexaoY / resistenciaCalTracaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoX / resistenciaCalTracaoParalela)) <= 1
    End Function
    Public Function validarTensaoFlexocompressaoX(tensaoCompressao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoCompressao / resistenciaCalCompressaoParalela) + (tensaoFlexaoX / resistenciaCalCompressaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoY / resistenciaCalCompressaoParalela)) <= 1
    End Function

    Public Function validarTensaoFlexocompressaoY(tensaoCompressao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoCompressao / resistenciaCalCompressaoParalela) + (tensaoFlexaoY / resistenciaCalCompressaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoX / resistenciaCalCompressaoParalela)) <= 1
    End Function

    '2 ELEMENTOS JUSTAPOSTOS
    Public Function ElementosJustaposto2EixoX()
        Return 0
    End Function

    Public Function ElementosJustaposto2EixoY()
        Return 0
    End Function

    '3 ELEMENTOS JUSTAPOSTO
    Public Function ElementosJustaposto3EixoX()
        Return 0
    End Function

    Public Function ElementosJustaposto3EixoY()
        Return 0
    End Function

    Public Function MomentoEstaticoDeArea(dadosIniciais As DadosIniciais, tipoSecao As Madeira.TipoSecao) As Qxy
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

        Return qxy
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

    'CALCULO DIRETO DA ALTURA TOTAL DA PEÇA EM T
    Public Function somaT(hf As Double, h As Double) As Double
        Return hf + h
    End Function

    'CALCULO DIRETO DA ALTURA TOTAL DA PEÇA EM I
    Public Function somaI(hf1 As Double, h As Double, hf2 As Double) As Double
        Return hf1 + h + hf2
    End Function

    'CALCULO DIRETO DA ALTURA TOTAL DA PEÇA CAIXÃO
    Public Function somaCaixao(hf1 As Double, h2 As Double, hf3 As Double, h4 As Double) As Double
        Return hf1 + h2 + hf3 + h4
    End Function

End Module


