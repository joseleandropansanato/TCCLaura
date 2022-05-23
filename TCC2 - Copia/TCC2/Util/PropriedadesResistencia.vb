Module PropriedadesResistencia

    'kmod
    Private Function SelectKmod1(tipo As Integer) As Kmod1
        Dim kmod1 As New Kmod1
        Select Case tipo
            Case 0
                Return New Kmod1(0.6, 0.7, 0.8, 0.9, 1.1)
            Case 1
                Return New Kmod1(0.3, 0.45, 0.65, 0.9, 1.1)
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
                Return New Kmod3(0.9, 0.85, 0.8, 0.75)
            Case 1
                Return New Kmod3(1.0, 0.95, 0.9, 0.85)
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

    'RESISTÊNCIA DE CÁLCULO
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


    'ABA DE COMPRESSÃO
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

    'Verifica vazio: Captura o conteudo do textbox, entra no if e verifica se o conteudo é vazio, caso seja true o conteudo vai receber 0, caso seja falso o conteudo será ele mesmo
    Public Function verificaVazio(value As String) As Double
        If value = "" Then
            Return 0
        End If
        Return value
    End Function

    'VALIDAÇÃO TRAÇÃO:
    Public Function validarTensaoTracao(tensaoTracao As Double, resistenciaCalTracaoParalela As Double) As Boolean
        Return tensaoTracao <= resistenciaCalTracaoParalela
    End Function

    Public Function validarEsbeltezTracaoX(esbeltezTracaoX As Double)
        Return esbeltezTracaoX <= 173
    End Function

    Public Function validarEsbeltezTracaoY(esbeltezTracaoY As Double)
        Return esbeltezTracaoY <= 173
    End Function


    'VALIDAÇÃO COMPRESSÃO
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

    'COMPRESSÃO: SEÇÃO COMPOSTA
    'COMPRESSÃO: PEÇAS MACIÇAS:
    '2 ELEMENTOS JUSTAPOSTOS


    'VALIDAÇÃO CISALHAMENTO
    Public Function validarTensaoCisalhanteX(tensaoCisalhante As Double, resistenciaCalAoCisalhamento As Double) As Boolean
        Return tensaoCisalhante <= resistenciaCalAoCisalhamento
    End Function

    Public Function validarTensaoCisalhanteY(tensaoCisalhante As Double, resistenciaCalAoCisalhamento As Double) As Boolean
        Return tensaoCisalhante <= resistenciaCalAoCisalhamento
    End Function

    Public Function validarTensaoCisalhanteEspaçador() As Boolean
        Return 0
    End Function

    'VALIDAÇÃO FLEXÃO SIMPLES RETA
    Public Function validarTensaoFlexaoSimplesRetaX(tensaoFlexaoX As Double, resistenciaCalTracaoParalela As Double) As Boolean
        Return tensaoFlexaoX <= resistenciaCalTracaoParalela
    End Function

    Public Function validarTensaoFlexaoSimplesRetaY(tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return tensaoFlexaoY <= resistenciaCalCompressaoParalela
    End Function

    Public Function validarTensaoCX(tensaoFlexaoX As Double, resistenciaCalTracaoParalela As Double) As Boolean
        Return tensaoFlexaoX <= resistenciaCalTracaoParalela
    End Function

    Public Function validarTensaoCY(tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return tensaoFlexaoY <= resistenciaCalCompressaoParalela
    End Function

    Public Function validarTensaoTX(tensaoFlexaoX As Double, resistenciaCalTracaoParalela As Double) As Boolean
        Return tensaoFlexaoX <= resistenciaCalTracaoParalela
    End Function

    Public Function validarTensaoTY(tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double) As Boolean
        Return tensaoFlexaoY <= resistenciaCalCompressaoParalela
    End Function

    'VALIDAÇÃO FLEXÃO SIMPLES OBLÍQUA
    'Definindo o coeficiente de correção para a flexão:
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

    'VALIDAÇÃO FLEXÃO COMPOSTA
    'FLEXOTRAÇÃO
    Public Function validarTensaoFlexotracaoX(tensaoTracao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalTracaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoTracao / resistenciaCalTracaoParalela) + (tensaoFlexaoX / resistenciaCalTracaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoY / resistenciaCalTracaoParalela)) <= 1
    End Function
    Public Function validarTensaoFlexotracaoY(tensaoTracao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalTracaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoTracao / resistenciaCalTracaoParalela) + (tensaoFlexaoY / resistenciaCalTracaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoX / resistenciaCalTracaoParalela)) <= 1
    End Function

    'FLEXOCOMPRESSÃO
    Public Function validarTensaoFlexocompressaoX(tensaoCompressao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoCompressao / resistenciaCalCompressaoParalela) + (tensaoFlexaoX / resistenciaCalCompressaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoY / resistenciaCalCompressaoParalela)) <= 1
    End Function
    Public Function validarTensaoFlexocompressaoY(tensaoCompressao As Double, tensaoFlexaoX As Double, tensaoFlexaoY As Double, resistenciaCalCompressaoParalela As Double, tipoSecao As Madeira.TipoSecao) As Boolean
        Return (tensaoCompressao / resistenciaCalCompressaoParalela) + (tensaoFlexaoY / resistenciaCalCompressaoParalela) + (definirKm(tipoSecao) * (tensaoFlexaoX / resistenciaCalCompressaoParalela)) <= 1
    End Function





    'ROTINA PARA O CÁLCULO DIRETO DA ALTURA TOTAL DA PEÇA EM T
    Public Function somaT(hf As Double, h As Double) As Double
        Return hf + h
    End Function

    'ROTINA PARA O CÁLCULO DIRETO DA ALTURA TOTAL DA PEÇA EM I
    Public Function somaI(hf1 As Double, h As Double, hf2 As Double) As Double
        Return hf1 + h + hf2
    End Function

    'ROTINA PARA O CÁLCULO DIRETO DA ALTURA TOTAL DA PEÇA CAIXÃO
    Public Function somaCaixaoAltura(h1 As Double, h2 As Double, h4 As Double) As Double
        Return h1 + h2 + h4
    End Function
    Public Function somaCaixaoComprimento(b1 As Double, b2 As Double, b3 As Double) As Double
        Return b1 + b2 + b3
    End Function

End Module


