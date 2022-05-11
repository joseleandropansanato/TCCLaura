Module PropriedadesResistencia
    'definições: lista com cada madeira que será necessario para o cálculo

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


