Imports TCC2.Kmod1

Public Class Form1

    Public madeiraSelecionada As Madeira = New Madeira 'chamei as propriedades da madeira (modelei as propriedades para ser igual a classe madeira)
    Public listMadeira As List(Of Madeira) = Definicoes.ListaDenominacaoMadeira
    Public tipoMadeira As Integer
    Public kmod1, kmod2, kmod3, kmod As Double
    'Public propriedadesGeometricas As PropriedadesGeometricas = New PropriedadesGeometricas 'inicia como zerado, mas pode ser usado em qualquer lugar do form 1. Então ela é auto preenchida qunado precisa dela
    Public form2Kmod1 As Form2Kmod1 = New Form2Kmod1
    Public form2Kmod2 As Form2Kmod2 = New Form2Kmod2
    Public form2Kmod3 As Form2Kmod3 = New Form2Kmod3
    Public form4Vinculacao As Form4Vinculacao = New Form4Vinculacao

    'Private Sub: Cria uma função
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbldata.Text = Date.Now

        'For: Laço de repetição
        'counter: Inteira - começo com 0. Conta a lista inteira da madeira com o "for" preenchendo o ComboBox
        For counter As Integer = 0 To listMadeira.Count - 1
            cboMadeiraUtilizada.Items.Add(listMadeira(counter).Name)
        Next counter

        CaracteristicaMadeira()

    End Sub

    'função para não deixar inserir letra ou valor nulo e negativo:
    Private Sub SairToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SairToolStripMenuItem.Click
        Application.Exit()

    End Sub

    Private Sub CaracteristicaMadeira()

        Dim checked = rbtInserirValores.Checked
        txtResistCompParalela.Enabled = checked
        txtResistTracaoParalela.Enabled = checked
        txtResistTracaoNormal.Enabled = checked
        txtResistAoCisalhamento.Enabled = checked
        txtModElasticidade.Enabled = checked
        txtDensAparente.Enabled = checked
    End Sub

    'Se o button estiver checked ele é visivel
    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles rbtInserirValores.CheckedChanged
        btnConfirmarInsercao.Visible = True
        CaracteristicaMadeira()
    End Sub

    Private Sub rbrValorNBR_CheckedChanged(sender As Object, e As EventArgs) Handles rbrValorNBR.CheckedChanged
        'madeiraSelecionada = listMadeira(madeiraSelecionada.Id)
        PreenchimentoCaracteristicas()
        CaracteristicaMadeira()
    End Sub

    Private Sub PreenchimentoCaracteristicas()

        txtResistCompParalela.Text = madeiraSelecionada.ResistCompParalela
        txtResistTracaoParalela.Text = madeiraSelecionada.ResistTracaoParalela
        txtResistTracaoNormal.Text = madeiraSelecionada.ResistTracaoNormal
        txtResistAoCisalhamento.Text = madeiraSelecionada.ResistAoCisalhamento
        txtModElasticidade.Text = madeiraSelecionada.ModElasticidade
        txtDensAparente.Text = madeiraSelecionada.DensAparente

    End Sub
    'Recebe o valor da lista que foi selecionado no form madeira
    Private Sub cboMadeiraUtilizada_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMadeiraUtilizada.SelectedIndexChanged
        madeiraSelecionada = listMadeira(cboMadeiraUtilizada.SelectedIndex)
        PreenchimentoCaracteristicas()
    End Sub

    Private Sub btnConfirmarInsercao_Click(sender As Object, e As EventArgs) Handles btnConfirmarInsercao.Click

        btnConfirmarInsercao.Visible = False
        madeiraSelecionada.ResistCompParalela = txtResistCompParalela.Text
        madeiraSelecionada.ResistTracaoParalela = txtResistTracaoParalela.Text
        madeiraSelecionada.ResistTracaoNormal = txtResistTracaoNormal.Text
        madeiraSelecionada.ResistAoCisalhamento = txtResistAoCisalhamento.Text
        madeiraSelecionada.ModElasticidade = txtModElasticidade.Text
        madeiraSelecionada.DensAparente = txtDensAparente.Text

        txtResistCompParalela.Enabled = False
        txtResistTracaoParalela.Enabled = False
        txtResistTracaoNormal.Enabled = False
        txtResistAoCisalhamento.Enabled = False
        txtModElasticidade.Enabled = False
        txtDensAparente.Enabled = False



    End Sub
    Private Sub cboKmod1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKmod1.SelectedIndexChanged
        kmod1 = Definicoes.Kmod1Final(tipoMadeira, cboKmod1.SelectedIndex)
        lblKmod1.Text = kmod1.ToString
        lblKmod1.Visible = True
    End Sub
    Private Sub cboKmod2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKmod2.SelectedIndexChanged
        kmod2 = Definicoes.Kmod2Final(tipoMadeira, cboKmod2.SelectedIndex)
        lblKmod2.Text = kmod2.ToString
        lblKmod2.Visible = True
    End Sub

    Private Sub cboKmod3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKmod3.SelectedIndexChanged
        kmod3 = Definicoes.Kmod3Final(tipoMadeira, cboKmod3.SelectedIndex)
        lblKmod3.Text = kmod3.ToString
        lblKmod3.Visible = True
    End Sub
    ' Private Sub txtEntradaRetangularBx_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaRetangularBx.TextChanged
    'lblBase.Text = txtEntradaRetangularBx.Text
    ' lblBase.Visible = True
    '  End Sub

    'Private Sub txtEntradaRetangularBy_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaRetangularBy.TextChanged
    'lblAltura.Text = txtEntradaRetangularBy.Text
    ' lblAltura.Visible = True
    '  End Sub

    Private Sub bntCalcularKmod_Click(sender As Object, e As EventArgs) Handles bntCalcularKmod.Click
        CalcularKmod()
        ResistCalculo()
    End Sub
    Private Sub cboTipoMadeira_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTipoMadeira.SelectedIndexChanged
        tipoMadeira = cboTipoMadeira.SelectedIndex
    End Sub

    Private Sub rbtSecaoRetangular_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoRetangular.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaoretangular
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub

    Private Sub rbtSecaoCircular_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoCircular.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaocircular
        'lblAltura.Visible = False
        ' lblBase.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub

    Private Sub rbtSecaoT_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoT.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaoT
        'lblAltura.Visible = False
        'lblBase.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub

    Private Sub rbtSecaoI_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoI.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaoi
        'lblAltura.Visible = False
        'lblBase.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub

    Private Sub rbtSecaoCaixao_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoCaixao.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaocaixao
        ' lblAltura.Visible = False
        'lblBase.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub

    Private Sub rbt2Elementos_CheckedChanged(sender As Object, e As EventArgs) Handles rbt2Elementos.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secao2elementos
        'lblAltura.Visible = False
        ' lblBase.Visible = False
        gbxResultadosSecao.Visible = False
        gbxResultadosElementos.Visible = True
    End Sub

    Private Sub rbt3Elementos_CheckedChanged(sender As Object, e As EventArgs) Handles rbt3Elementos.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secao3elementos
        ' lblAltura.Visible = False
        ' lblBase.Visible = False
        gbxResultadosSecao.Visible = False
        gbxResultadosElementos.Visible = True
    End Sub

    Private Sub gbxFlexaoSimples_visible(sender As Object, e As EventArgs) Handles gbxFlexaoSimples.VisibleChanged
        txtCortanteX.Text = 0
        txtMx.Text = 0
        txtMy.Text = 0
    End Sub
    Private Sub rbtFlexoTracao_CheckedChanged(sender As Object, e As EventArgs)
        If txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Tração" Then
            gbxFlexoTracao.Visible = True
            gbxFlexoCompressao.Visible = False

        ElseIf txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            gbxFlexoCompressao.Visible = True
            gbxFlexoTracao.Visible = False
        End If
    End Sub

    Private Sub btnCalcularPropriedades_Click(sender As Object, e As EventArgs) Handles btnCalcularPropriedades.Click

        Dim tipoSecao As Madeira.TipoSecao

        If rbtSecaoRetangular.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.CalculoRetangular(txtEntradaRetangularBx.Text, txtEntradaRetangularBy.Text, txtEntradaRetangularL.Text)
            tipoSecao = Madeira.TipoSecao.Retangular

        ElseIf rbtSecaoCircular.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.CalculoCircular(txtEntradaCircularD.Text, txtEntradaCircularL.Text)
            tipoSecao = Madeira.TipoSecao.Circular

        ElseIf rbtSecaoT.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.CalculoCompostoSecaoT(txtEntradaCompostaBf.Text, txtEntradaCompostaHf.Text, txtEntradaCompostaTh.Text, txtEntradaCompostaBw.Text)
            tipoSecao = Madeira.TipoSecao.SecaoT

        ElseIf rbtSecaoI.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.CalculoCompostoSecaoI(txtEntradaIBf1.Text, txtEntradaIHf1.Text, txtEntradaIBf2.Text, txtEntradaIHf2.Text, txtEntradaIBw.Text, txtEntradaID.Text)
            tipoSecao = Madeira.TipoSecao.SecaoI

        ElseIf rbtSecaoCaixao.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.CalculoCompostoSecaoCaixao(txtEntradaCaixaoD.Text, txtEntradaCaixaoB1.Text, txtEntradaCaixaoB2.Text, txtEntradaCaixaoB3.Text, txtEntradaCaixaoB4.Text, txtEntradaCaixaoH1.Text, txtEntradaCaixaoH2.Text, txtEntradaCaixaoH3.Text, txtEntradaCaixaoH4.Text)
            tipoSecao = Madeira.TipoSecao.Caixao

        ElseIf rbt2Elementos.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.Calculo2Elementos(txtEntrada2L.Text, txtEntrada2L1.Text, txtEntrada2a.Text, txtEntrada2a1.Text, txtEntrada2h.Text, txtEntrada2h1.Text, txtEntrada2b1.Text, pilar())
            tipoSecao = Madeira.TipoSecao.ElementosJustaposto2

        ElseIf rbt3Elementos.Checked Then
            madeiraSelecionada.PropriedadesGeometricas = Definicoes.Calculo3Elementos(txtEntrada3L.Text, txtEntrada3L1.Text, txtEntrada3a.Text, txtEntrada3a1.Text, txtEntrada3h.Text, txtEntrada3h1.Text, txtEntrada3b1.Text, pilar())
            tipoSecao = Madeira.TipoSecao.ElementosJustaposto3
        Else
        End If

        getDadosIniciais(tipoSecao)

        If rbt2Elementos.Checked Or rbt3Elementos.Checked Then
            ResultadoPropriedadesGeometricas23Elemento(madeiraSelecionada.PropriedadesGeometricas)
        Else
            ResultadoPropriedadesGeometricas(madeiraSelecionada.PropriedadesGeometricas)
        End If

        'TRAÇÃO
        If txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Tração" Then
            madeiraSelecionada.Tracao = Definicoes.TensaoTracao(
                Definicoes.verificaVazio(txtNormal.Text),
                Definicoes.verificaVazio(txtEntradaRetangularBx.Text),
                Definicoes.verificaVazio(txtEntradaRetangularBy.Text),
                Definicoes.verificaVazio(txtEntradaRetangularL.Text),
                Definicoes.verificaVazio(txtEntradaCircularD.Text),
                Definicoes.verificaVazio(txtEntradaCompostaH.Text),
                Definicoes.verificaVazio(txtEntradaCompostaBw.Text),
                Definicoes.verificaVazio(txtEntradaIBf1.Text),
                Definicoes.verificaVazio(txtEntradaCircularL.Text),
                madeiraSelecionada.PropriedadesGeometricas,
                tipoSecao,
                Definicoes.verificaVazio(txtEntrada2L.Text),
                Definicoes.verificaVazio(txtEntrada2b1.Text),
                Definicoes.verificaVazio(txtEntrada2a.Text),
                cboElementFixacao.Text,
                vinculacao(comprimento),
                Definicoes.verificaVazio(txtMx.Text),
                Definicoes.verificaVazio(txtMy.Text),
                entradaElementoJustaposto
                )
            'Definicoes.verificaVazio(txtEntrada2h1.Text)
            'Definicoes.verificaVazio(txtEntrada3h1)


            'COMPRESSAO
        ElseIf txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            madeiraSelecionada.Compressao = Definicoes.TensaoCompressao(
                Definicoes.verificaVazio(txtNormal.Text),
                Definicoes.verificaVazio(txtMx.Text),
                Definicoes.verificaVazio(txtMy.Text),
                vinculacao(comprimento),
                madeiraSelecionada.PropriedadesGeometricas,
                tipoSecao,
                Definicoes.verificaVazio(txtMomentoCargaPermanenteX.Text),
                Definicoes.verificaVazio(txtMomentoCargaPermanenteY.Text),
                Definicoes.verificaVazio(txtMomentoCargaVariavelX.Text),
                Definicoes.verificaVazio(txtMomentoCargaVariavelY.Text),
                Definicoes.verificaVazio(txtNormalCargaPermanente.Text),
                Definicoes.verificaVazio(txtNormalCargaVariavel.Text),
                coefInfluencia(),
                Definicoes.verificaVazio(txtF1.Text),
                Definicoes.verificaVazio(txtF2.Text)
                )
        End If

        'CISALHAMENTO
        madeiraSelecionada.Cisalhamento = Definicoes.TensaoCisalhamento(
             Definicoes.verificaVazio(txtCortanteX.Text),
             Definicoes.verificaVazio(txtCortanteY.Text),
             Definicoes.verificaVazio(txtEntradaCircularD.Text),
             Definicoes.MomentoEstaticoDeArea(madeiraSelecionada.DadosIniciais, tipoSecao),
             madeiraSelecionada.PropriedadesGeometricas,
             tipoSecao
             )

        'FLEXÃO SIMPLES RETA

        If txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtCortanteX.Text.ToString <> "" And rbtFlexaoSimples.Checked Then
            gbxFlexaoSimples.Visible = True
            gbxFlexaoObliqua.Visible = False
            madeiraSelecionada.Flexao = Definicoes.TensaoFlexaoSimples(
            Definicoes.verificaVazio(txtMx.Text),
            Definicoes.verificaVazio(txtMy.Text),
           madeiraSelecionada.PropriedadesGeometricas,
           tipoSecao
           )
            'FLEXÃO OBLIQUA
        ElseIf txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtCortanteX.Text.ToString <> "" And rbtFlexaoObliqua.Checked Then
            gbxFlexaoSimples.Visible = False
            gbxFlexaoObliqua.Visible = True
            madeiraSelecionada.Flexao = Definicoes.TensaoFlexaoObliqua(
            Definicoes.verificaVazio(txtMx.Text),
            Definicoes.verificaVazio(txtMy.Text),
            madeiraSelecionada.PropriedadesGeometricas,
            tipoSecao
            )
        End If

        'FLEXAO COMPOSTA
        '------------------------FLEXOTRAÇÃO------------------------
        If txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Tração" Then
            gbxFlexoTracao.Visible = True
            gbxFlexoCompressao.Visible = False
            madeiraSelecionada.Flexao = Definicoes.TensaoFlexotracao(
            Definicoes.verificaVazio(txtMx.Text),
            Definicoes.verificaVazio(txtMy.Text),
            madeiraSelecionada.PropriedadesGeometricas,
            tipoSecao
            )

            '------------------------FLEXOCOMPRESSAO------------------------
        ElseIf txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            gbxFlexoCompressao.Visible = True
            gbxFlexoTracao.Visible = False
            madeiraSelecionada.Flexao = Definicoes.TensaoFlexocompressao(
            Definicoes.verificaVazio(txtMx.Text),
            Definicoes.verificaVazio(txtMy.Text),
            madeiraSelecionada.PropriedadesGeometricas,
            tipoSecao
            )
        End If
        tracao()
        compressao()
        cisalhamento()
        flexaoSimples(tipoSecao)
        flexaoComposta(tipoSecao)


    End Sub

    'TRAÇÃO: CONTAS CORRETAS
    Private Sub tracao()

        txtTensaoTracao.Text = madeiraSelecionada.Tracao.TensaoTracao
        txtEsbeltezTracaoX.Text = madeiraSelecionada.Tracao.EsbeltezTracaoX
        txtEsbeltezTracaoY.Text = madeiraSelecionada.Tracao.EsbeltezTracaoY

        'VALIDAÇÃO TENSÃO TRAÇÃO:
        If Definicoes.validarTensaoTracao(madeiraSelecionada.Tracao.TensaoTracao, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela) Then
            lblValidacaoTensaoTracao.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoTensaoTracao.Visible = True
            lblValidacaoTensaoTracao.ForeColor = Color.Green
        Else
            lblValidacaoTensaoTracao.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoTensaoTracao.Visible = True
            lblValidacaoTensaoTracao.ForeColor = Color.Red
            TabDadosIniciais.Select()
        End If

        'VALIDAÇÃO ESBELTEZ X:
        If Definicoes.validarEsbeltezTracaoX(madeiraSelecionada.Tracao.EsbeltezTracaoX) Then
            lblValidacaoEsbeltezXTracao.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezXTracao.Visible = True
            lblValidacaoEsbeltezXTracao.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltezXTracao.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezXTracao.Visible = True
            lblValidacaoEsbeltezXTracao.ForeColor = Color.Red
        End If

        'VALIDAÇÃO ESBELTEZ Y:
        If Definicoes.validarEsbeltezTracaoY(madeiraSelecionada.Tracao.EsbeltezTracaoY) Then
            lblValidacaoEsbeltezYTracao.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezYTracao.Visible = True
            lblValidacaoEsbeltezYTracao.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltezYTracao.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezYTracao.Visible = True
            lblValidacaoEsbeltezYTracao.ForeColor = Color.Red
        End If
    End Sub

    'COMPRESSÃO
    Private Sub compressao()
        txtEsbeltezCompressaoX.Text = madeiraSelecionada.Compressao.EsbeltezCompressaoX
        txtEsbeltezCompressaoY.Text = madeiraSelecionada.Compressao.EsbeltezCompressaoY
        txtTensaoCompressaoCurtaX.Text = madeiraSelecionada.Compressao.TensaoCompressaoCurtaX
        txtTensaoCompressaoCurtaY.Text = madeiraSelecionada.Compressao.TensaoCompressaoCurtaY
        txtEsforçoNormalMedEsbX.Text = madeiraSelecionada.Compressao.TensaoCompressaoMedEsbeltaX
        txtEsforçoFlexaoMedEsbX.Text = madeiraSelecionada.Compressao.TensaoMxd
        txtEsforçoNormalMedEsbY.Text = madeiraSelecionada.Compressao.TensaoCompressaoMedEsbeltaY
        txtEsforçoFlexaoMedEsbY.Text = madeiraSelecionada.Compressao.TensaoMyd
        txtEsforçoNormalEsbX.Text = madeiraSelecionada.Compressao.TensaoCompressaoEsbeltaX
        txtEsforçoFlexaoEsbX.Text = madeiraSelecionada.Compressao.TensaoMxd
        txtEsforçoNormalEsbY.Text = madeiraSelecionada.Compressao.TensaoCompressaoEsbeltaY
        txtEsforçoFlexaoEsbY.Text = madeiraSelecionada.Compressao.TensaoMyd
        txtForçaElasticaX.Text = madeiraSelecionada.Compressao.ForcaElasticaX
        txtForçaElasticaY.Text = madeiraSelecionada.Compressao.ForcaElasticaY
        txtEaX.Text = madeiraSelecionada.Compressao.Ea
        txtEaY.Text = madeiraSelecionada.Compressao.Ea
        txtEdX.Text = madeiraSelecionada.Compressao.Edx
        txtEdY.Text = madeiraSelecionada.Compressao.Edy
        txtEiX.Text = madeiraSelecionada.Compressao.Ei
        txtEiY.Text = madeiraSelecionada.Compressao.Ei
        txtE1efX.Text = madeiraSelecionada.Compressao.E1ef
        txtE1efY.Text = madeiraSelecionada.Compressao.E1ef
        txtEcX.Text = madeiraSelecionada.Compressao.Ec
        txtEcY.Text = madeiraSelecionada.Compressao.Ec

        'VALIDAÇÃO COMPRESSÂO:


        'VALIDAÇÃO COMPRESSÃO CURTAX:
        If Definicoes.validarTensaoCompressaoCurtaX(madeiraSelecionada.Compressao.TensaoCompressaoCurtaX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            ' gbxEixoXCurta.Visible = True
            'lblClassificaçãoEsbeltezX.Text = "PEÇA CURTA"
            lblValidacaoCurtaXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCurtaXComp.Visible = True
            lblValidacaoCurtaXComp.ForeColor = Color.Green
        Else
            lblValidacaoCurtaXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCurtaXComp.Visible = True
            lblValidacaoCurtaXComp.ForeColor = Color.Green
        End If


        'VALIDAÇÃO COMPRESSÃO CURTAY
        If Definicoes.validarTensaoCompressaoCurtaY(madeiraSelecionada.Compressao.TensaoCompressaoCurtaY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            'gbxEixoYCurta.Visible = True
            'lblClassificaçãoEsbeltezY.Text = "PEÇA CURTA"
            lblValidacaoCurtaYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCurtaYComp.Visible = True
            lblValidacaoCurtaYComp.ForeColor = Color.Green
        Else
            lblValidacaoCurtaYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCurtaYComp.Visible = True
            lblValidacaoCurtaYComp.ForeColor = Color.Green
        End If

        'VALIDAÇÃO COMPRESSAO MEDIANAMENTE ESBELTAX

        If Definicoes.validarTensaoCompressaoMedEsbX(madeiraSelecionada.Compressao.TensaoCompressaoMedEsbeltaX, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            'gbxEixoXMediEsbelta.Visible = True
            'lblClassificaçãoEsbeltezX.Text = "PEÇA MEDIANAMENTE ESBELTA"
            lblValidacaoMedEsbeltXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoMedEsbeltXComp.Visible = True
            lblValidacaoMedEsbeltXComp.ForeColor = Color.Green
        Else
            lblValidacaoMedEsbeltXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoMedEsbeltXComp.Visible = True
            lblValidacaoMedEsbeltXComp.ForeColor = Color.Green
        End If

        'VALIDAÇÃO COMPRESSAO MEDIANAMENTE ESBELTAY
        If Definicoes.validarTensaoCompressaoMedEsbY(madeiraSelecionada.Compressao.TensaoCompressaoMedEsbeltaX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            'gbxEixoYMediEsbelta.Visible = True
            'lblClassificaçãoEsbeltezY.Text = "PEÇA MEDIANAMENTE ESBELTA"
            lblValidacaoMedEsbeltYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoMedEsbeltYComp.Visible = True
            lblValidacaoMedEsbeltYComp.ForeColor = Color.Green
        Else
            lblValidacaoMedEsbeltYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoMedEsbeltYComp.Visible = True
            lblValidacaoMedEsbeltYComp.ForeColor = Color.Green
        End If


        'VALIDAÇÃO COMPRESSÃO ESBELTAX
        If Definicoes.validarTensaoCompressaoEsbX(madeiraSelecionada.Compressao.TensaoCompressaoEsbeltaY, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            'lblClassificaçãoEsbeltezX.Text = "PEÇA ESBELTA"
            lblValidacaoEsbeltXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltXComp.Visible = True
            lblValidacaoEsbeltXComp.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltXComp.Visible = True
            lblValidacaoEsbeltXComp.ForeColor = Color.Green
        End If


        'VALIDAÇÃO COMPRESSÃO ESBELTAY
        If Definicoes.validarTensaoCompressaoEsbY(madeiraSelecionada.Compressao.TensaoCompressaoEsbeltaY, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            'lblClassificaçãoEsbeltezY.Text = "PEÇA ESBELTA"
            lblValidacaoEsbeltYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltYComp.Visible = True
            lblValidacaoEsbeltYComp.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltYComp.Visible = True
            lblValidacaoEsbeltYComp.ForeColor = Color.Green
        End If

    End Sub

    'CISALHAMENTO
    Private Sub cisalhamento()
        txtTensaoCisalhanteX.Text = madeiraSelecionada.Cisalhamento.TensaoCisalhanteX
        txtTensaoCisalhanteY.Text = madeiraSelecionada.Cisalhamento.TensaoCisalhanteY

        'VALIDAÇÃO TENSÃO CISALHANTE X:
        If Definicoes.validarTensaoCisalhanteX(madeiraSelecionada.Cisalhamento.TensaoCisalhanteX, madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento) Then
            lblValidacaoCisalhanteX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteX.Visible = True
            lblValidacaoCisalhanteX.ForeColor = Color.Green
        Else
            lblValidacaoCisalhanteX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteX.Visible = True
            lblValidacaoCisalhanteX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO TENSÃO CISALHANTE Y:
        If Definicoes.validarTensaoCisalhanteY(madeiraSelecionada.Cisalhamento.TensaoCisalhanteX, madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento) Then
            lblValidacaoCisalhanteY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteY.Visible = True
            lblValidacaoCisalhanteY.ForeColor = Color.Green
        Else
            lblValidacaoCisalhanteY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteY.Visible = True
            lblValidacaoCisalhanteY.ForeColor = Color.Red
        End If


    End Sub

    'FLEXÃO SIMPLES
    Private Sub flexaoSimples(tipoSecao As Madeira.TipoSecao)
        'reta
        txtTensaoTx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtTensaoCy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY
        'obliqua
        txtTensaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtTensaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY

        'VALIDAÇÃO FLEXÃO SIMPLES RETA X:
        If Definicoes.validarTensaoFlexaoSimplesRetaX(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela) Then
            lblValidacaoFlexaoSimplesX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesX.Visible = True
            lblValidacaoFlexaoSimplesX.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoSimplesX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesX.Visible = True
            lblValidacaoFlexaoSimplesX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXÃO SIMPLES RETA Y:
        If Definicoes.validarTensaoFlexaoSimplesRetaY(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblValidacaoFlexaoSimplesY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesY.Visible = True
            lblValidacaoFlexaoSimplesY.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoSimplesY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesY.Visible = True
            lblValidacaoFlexaoSimplesY.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXÃO SIMPLES OBLIQUA X:
        If Definicoes.validarTensaoFlexaoSimplesObliquaX(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexaoObliquaX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoObliquaX.Visible = True
            lblValidacaoFlexaoObliquaX.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoObliquaX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoObliquaX.Visible = True
            lblValidacaoFlexaoObliquaX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXÃO SIMPLES OBLIQUA Y:
        If Definicoes.validarTensaoFlexaoSimplesObliquaY(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, tipoSecao) Then
            lblValidacaoFlexaoObliquaY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoObliquaY.Visible = True
            lblValidacaoFlexaoObliquaY.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoObliquaY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoObliquaY.Visible = True
            lblValidacaoFlexaoObliquaY.ForeColor = Color.Red
        End If

    End Sub

    'FLEXÃO COMPOSTA
    Private Sub flexaoComposta(tipoSecao As Madeira.TipoSecao)
        txtFlexoTracao.Text = madeiraSelecionada.Tracao.TensaoTracao
        txtFlexoTracaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtFlexoTracaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY
        txtFlexoCompressao.Text = madeiraSelecionada.Compressao.TensaoCompressao
        txtFlexoCompressaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtFlexoCompressaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY

        'VALIDAÇÃO FLEXOTRAÇÃO X:
        If Definicoes.validarTensaoFlexotracaoX(madeiraSelecionada.Tracao.TensaoTracao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoTracaoX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoX.Visible = True
            lblValidacaoFlexoTracaoX.ForeColor = Color.Green
        Else
            lblValidacaoFlexoTracaoX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoX.Visible = True
            lblValidacaoFlexoTracaoX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOTRAÇÃO Y:
        If Definicoes.validarTensaoFlexotracaoY(madeiraSelecionada.Tracao.TensaoTracao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, tipoSecao) Then
            lblValidacaoFlexoTracaoY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoY.Visible = True
            lblValidacaoFlexoTracaoY.ForeColor = Color.Green
        Else
            lblValidacaoFlexoTracaoY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoY.Visible = True
            lblValidacaoFlexoTracaoY.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOCOMPRESSAO X:
        If Definicoes.validarTensaoFlexocompressaoX(madeiraSelecionada.Compressao.TensaoCompressao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoCompressaoX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoX.Visible = True
            lblValidacaoFlexoCompressaoX.ForeColor = Color.Green
        Else
            lblValidacaoFlexoCompressaoX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoX.Visible = True
            lblValidacaoFlexoCompressaoX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOCOMPRESSAO Y:
        If Definicoes.validarTensaoFlexocompressaoY(madeiraSelecionada.Compressao.TensaoCompressao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoCompressaoY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoY.Visible = True
            lblValidacaoFlexoCompressaoY.ForeColor = Color.Green
        Else
            lblValidacaoFlexoCompressaoY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoY.Visible = True
            lblValidacaoFlexoCompressaoY.ForeColor = Color.Red
        End If
    End Sub
    Private Sub CalcularKmod()
        kmod = kmod1 * kmod2 * kmod3
        txtKmod.Text = kmod
    End Sub

    Private Sub ResistCalculo()
        madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela = Definicoes.resistCalCompParalela(kmod, madeiraSelecionada.ResistCompParalela)
        madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela = Definicoes.resistCalTracaoParalela(kmod, madeiraSelecionada.ResistTracaoParalela)
        madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaNormal = Definicoes.resistCalTracaoNormal(kmod, madeiraSelecionada.ResistTracaoNormal)
        madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento = Definicoes.resistCalAoCisalhamento(kmod, madeiraSelecionada.ResistAoCisalhamento)
        Dim modElasticidade = Definicoes.ModElasticidade(kmod, madeiraSelecionada.ModElasticidade)
        madeiraSelecionada.ResistenciaCalculo.moduloElasticidade = modElasticidade
        madeiraSelecionada.ResistenciaCalculo.moduloElasticidadeTransversal = Definicoes.ModElasticidadeTransversal(modElasticidade)
        madeiraSelecionada.ResistenciaCalculo.densidadeAparente = (madeiraSelecionada.DensAparente)

        txtRcParalela.Text = Format(madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, "#0.0##;-#0.0##")
        txtRtParalelo.Text = Format(madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, "#0.0##;-#0.0##")
        txtRtNormal.Text = Format(madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaNormal, "#0.0##;-#0.0##")
        txtRCisalhamento.Text = Format(madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento, "#0.0##;-#0.0##")
        txtMElasticidade.Text = Format(madeiraSelecionada.ResistenciaCalculo.moduloElasticidade, "#0.0##;-#0.0##")
        txtMeTransversal.Text = Format(madeiraSelecionada.ResistenciaCalculo.moduloElasticidadeTransversal, "#0.0##;-#0.0##")
        txtDAparente.Text = Format(madeiraSelecionada.ResistenciaCalculo.densidadeAparente, "#0.0##;-#0.0##")
    End Sub


    'DADOS INICIAIS QUE SÃO FECHADOS PRO USUARIO NÃO DIGITAR NADA 

    'SEÇÃO T----------------------------------------
    Private Sub txtEntradaCompostaTh_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaTh.TextChanged
        txtEntradaCompostaH.Text = Definicoes.somaT(Definicoes.verificaVazio(txtEntradaCompostaHf.Text), Definicoes.verificaVazio(txtEntradaCompostaTh.Text))
    End Sub

    Private Sub txtEntradaCompostaHf_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaHf.TextChanged
        txtEntradaCompostaH.Text = Definicoes.somaT(Definicoes.verificaVazio(txtEntradaCompostaHf.Text), Definicoes.verificaVazio(txtEntradaCompostaTh.Text))
    End Sub

    'SEÇÃO I----------------------------------------

    Private Sub txtEntradaIHf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIHf1.TextChanged
        txtEntradaIHf2.Text = txtEntradaIHf1.Text
        txtEntradaID.Text = Definicoes.somaI(Definicoes.verificaVazio(txtEntradaIHf1.Text), Definicoes.verificaVazio(txtEntradaIH.Text), Definicoes.verificaVazio(txtEntradaIHf2.Text))
    End Sub

    Private Sub txtEntradaIH_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIH.TextChanged
        txtEntradaID.Text = Definicoes.somaI(Definicoes.verificaVazio(txtEntradaIHf1.Text), Definicoes.verificaVazio(txtEntradaIH.Text), Definicoes.verificaVazio(txtEntradaIHf2.Text))
    End Sub

    Private Sub txtEntradaIHf2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIHf2.TextChanged
        txtEntradaID.Text = Definicoes.somaI(Definicoes.verificaVazio(txtEntradaIHf1.Text), Definicoes.verificaVazio(txtEntradaIH.Text), Definicoes.verificaVazio(txtEntradaIHf2.Text))
    End Sub

    Private Sub txtEntradaIBf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIBf1.TextChanged
        txtEntradaIBf2.Text = txtEntradaIBf1.Text
    End Sub

    'SEÇÃO CAIXÃO----------------------------------------
    Private Sub txtEntradaCaixaoH1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH1.TextChanged
        txtEntradaCaixaoD.Text = Definicoes.somaCaixao(Definicoes.verificaVazio(txtEntradaCaixaoH1.Text), Definicoes.verificaVazio(txtEntradaCaixaoH2.Text), Definicoes.verificaVazio(txtEntradaCaixaoH3.Text), Definicoes.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub

    Private Sub txtEntradaCaixaoH2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH2.TextChanged
        txtEntradaCaixaoD.Text = Definicoes.somaCaixao(Definicoes.verificaVazio(txtEntradaCaixaoH1.Text), Definicoes.verificaVazio(txtEntradaCaixaoH2.Text), Definicoes.verificaVazio(txtEntradaCaixaoH3.Text), Definicoes.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub
    Private Sub txtEntradaCaixaoH3_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH3.TextChanged
        txtEntradaCaixaoD.Text = Definicoes.somaCaixao(Definicoes.verificaVazio(txtEntradaCaixaoH1.Text), Definicoes.verificaVazio(txtEntradaCaixaoH2.Text), Definicoes.verificaVazio(txtEntradaCaixaoH3.Text), Definicoes.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub
    Private Sub txtEntradaCaixaoH4_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH4.TextChanged
        txtEntradaCaixaoD.Text = Definicoes.somaCaixao(Definicoes.verificaVazio(txtEntradaCaixaoH1.Text), Definicoes.verificaVazio(txtEntradaCaixaoH2.Text), Definicoes.verificaVazio(txtEntradaCaixaoH3.Text), Definicoes.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub

    Private Sub txtEntradaCaixaoBf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB1.TextChanged
        txtEntradaCaixaoB2.Text = txtEntradaCaixaoB1.Text
    End Sub

    Private Sub cboTracaoCompressao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTracaoCompressao.SelectedIndexChanged
        TabTracao.Enabled = cboTracaoCompressao.SelectedIndex.CompareTo(1)
        TabCompressao.Enabled = cboTracaoCompressao.SelectedIndex.CompareTo(0)
    End Sub
    Private Sub txtEntrada3a_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3a.TextChanged
        txtEntrada3a1.Text = (txtEntrada3a.Text) / 2
    End Sub
    Private Sub txtEntrada3h_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3h.TextChanged
        txtEntrada3h.Text = txtEntrada3a.Text + txtEntrada3b1.Text
    End Sub

    Private Sub ExibirDadosEntrada()
        Dim selectedComposta = rbtSecaoT.Checked Or rbtSecaoI.Checked Or rbtSecaoCaixao.Checked
        gbxRetangular.Visible = rbtSecaoRetangular.Checked
        gbxCircular.Visible = rbtSecaoCircular.Checked
        gbxCompostaCaixao.Visible = rbtSecaoCaixao.Checked
        gbxCompostaI.Visible = rbtSecaoI.Checked
        gbxCompostaT.Visible = rbtSecaoT.Checked
        gbx2Elementos.Visible = rbt2Elementos.Checked
        gbx3Elementos.Visible = rbt3Elementos.Checked
    End Sub

    'não deu::::
    Private Sub ExibirResultados()
        Dim selectedResultados = rbtSecaoRetangular.Checked Or rbtSecaoCircular.Checked Or rbtSecaoT.Checked Or rbtSecaoI.Checked Or rbtSecaoCaixao.Checked
        gbxResultadosSecao.Visible = rbtSecaoRetangular.Checked Or rbtSecaoCircular.Checked Or rbtSecaoT.Checked Or rbtSecaoI.Checked Or rbtSecaoCaixao.Checked
        gbxResultadosSecao.Visible = rbt2Elementos.Checked Or rbt3Elementos.Checked

    End Sub

    Private Sub ResultadoPropriedadesGeometricas(resultProprGeometricas As PropriedadesGeometricas)
        txtAreaSecao.Text = resultProprGeometricas.Area
        txtEixoXMomentoInercia.Text = resultProprGeometricas.EixoXmi
        txtEixoXRaioGiracao.Text = resultProprGeometricas.EixoXrg
        txtEixoXModuloResistencia.Text = resultProprGeometricas.EixoXmr
        txtEixoXInerciaEfetiva.Text = resultProprGeometricas.EixoXie
        txtEixoYMomentoInercia.Text = resultProprGeometricas.EixoYmi
        txtEixoYRaioGiracao.Text = resultProprGeometricas.EixoYrg
        txtEixoYModuloResistencia.Text = resultProprGeometricas.EixoYmr
        txtEixoYInerciaEfetiva.Text = resultProprGeometricas.EixoYie
    End Sub

    Private Sub ResultadoPropriedadesGeometricas23Elemento(resultProprGeometricas As PropriedadesGeometricas)
        txtArea1.Text = resultProprGeometricas.AreaA1
        txtArea.Text = resultProprGeometricas.Area
        txtInterpoIx1.Text = resultProprGeometricas.EixoXmi1
        txtInterpoIx.Text = resultProprGeometricas.EixoXmi
        txtInterpoIy2.Text = resultProprGeometricas.EixoYmi1
        txtInterpoIy.Text = resultProprGeometricas.EixoYmi
        txtInterpoEixoYCoefB.Text = resultProprGeometricas.CoefBy
        txtInterpoIModuloResist.Text = resultProprGeometricas.EixoYmr
        txtInterpoEixoYInerciaEfetiva.Text = resultProprGeometricas.EixoYie
    End Sub

    'VALIDAÇÕES TXT PROPRIEDADES DA MADEIRA
    Private Sub txtResistCompParalela_Validated(sender As Object, e As EventArgs) Handles txtResistCompParalela.Validated
        If Validacao.ValidarDados(txtResistCompParalela, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtResistCompParalela.Text)
        End If
    End Sub

    Private Sub txtResistTracaoParalela_Validated(sender As Object, e As EventArgs) Handles txtResistTracaoParalela.Validated
        If Validacao.ValidarDados(txtResistTracaoParalela, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtResistTracaoParalela.Text)
        End If
    End Sub

    Private Sub txtResistTracaoNormal_Validated(sender As Object, e As EventArgs) Handles txtResistTracaoNormal.Validated
        If Validacao.ValidarDados(txtResistTracaoNormal, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtResistTracaoNormal.Text)
        End If
    End Sub

    Private Sub txtResistAoCisalhamento_Validated(sender As Object, e As EventArgs) Handles txtResistAoCisalhamento.Validated
        If Validacao.ValidarDados(txtResistAoCisalhamento, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtResistAoCisalhamento.Text)
        End If
    End Sub

    Private Sub txtUmidadeVento_Validated(sender As Object, e As EventArgs) Handles txtUmidadeVento.Validated
        If Validacao.ValidarDados(txtUmidadeVento, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtUmidadeVento.Text)
        End If
    End Sub

    Private Sub txtModElasticidade_Validated(sender As Object, e As EventArgs) Handles txtModElasticidade.Validated
        If Validacao.ValidarDados(txtModElasticidade, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtModElasticidade.Text)
        End If
    End Sub

    Private Sub txtDensAparente_Validated(sender As Object, e As EventArgs) Handles txtDensAparente.Validated
        If Validacao.ValidarDados(txtDensAparente, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtDensAparente.Text)
        End If
    End Sub

    'VALIDAÇÕES TXT DADOS INICIAIS
    'SEÇÃO RETANGULAR
    Private Sub txtEntradaRetangularBx_Validated(sender As Object, e As EventArgs) Handles txtEntradaRetangularBx.Validated
        If Validacao.ValidarDados(txtEntradaRetangularBx, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaRetangularBx.Text)

        End If
    End Sub

    Private Sub txtEntradaRetangularBy_Validated_1(sender As Object, e As EventArgs) Handles txtEntradaRetangularBy.Validated
        If Validacao.ValidarDados(txtEntradaRetangularBy, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaRetangularBy.Text)
        End If
    End Sub

    Private Sub txtEntradaRetangularL_Validated(sender As Object, e As EventArgs) Handles txtEntradaRetangularL.Validated
        If Validacao.ValidarDados(txtEntradaRetangularL, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaRetangularL.Text)
        End If
    End Sub

    'SEÇÃO CIRCULAR 
    Private Sub txtEntradaCircularL_Validated(sender As Object, e As EventArgs) Handles txtEntradaCircularL.Validated
        If Validacao.ValidarDados(txtEntradaCircularL, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCircularL.Text)
        End If
    End Sub

    Private Sub txtEntradaCircularD_Validated(sender As Object, e As EventArgs) Handles txtEntradaCircularD.Validated
        If Validacao.ValidarDados(txtEntradaCircularD, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCircularD.Text)
        End If
    End Sub

    'SEÇÃO T
    Private Sub txtEntradaCompostaH_Validated(sender As Object, e As EventArgs) Handles txtEntradaCompostaH.Validated
        If Validacao.ValidarDados(txtEntradaCompostaH, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCompostaH.Text)
        End If
    End Sub

    Private Sub txtEntradaCompostaHf_Validated(sender As Object, e As EventArgs) Handles txtEntradaCompostaHf.Validated
        If Validacao.ValidarDados(txtEntradaCompostaHf, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCompostaHf.Text)
        End If
    End Sub

    Private Sub txtEntradaCompostaBf_Validated_1(sender As Object, e As EventArgs) Handles txtEntradaCompostaBf.Validated
        If Validacao.ValidarDados(txtEntradaCompostaHf, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCompostaHf.Text)
        End If
    End Sub

    Private Sub txtEntradaCompostaBw_Validated(sender As Object, e As EventArgs) Handles txtEntradaCompostaBw.Validated
        If Validacao.ValidarDados(txtEntradaCompostaBw, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCompostaBw.Text)
        End If
    End Sub

    'SEÇÃO I 
    Private Sub txtEntradaID_Validated(sender As Object, e As EventArgs) Handles txtEntradaID.Validated
        If Validacao.ValidarDados(txtEntradaID, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaID.Text)
        End If
    End Sub

    Private Sub txtEntradaIBw_Validated(sender As Object, e As EventArgs) Handles txtEntradaIBw.Validated
        If Validacao.ValidarDados(txtEntradaIBw, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaIBw.Text)
        End If
    End Sub

    Private Sub txtEntradaIHf1_Validated(sender As Object, e As EventArgs) Handles txtEntradaIHf1.Validated
        If Validacao.ValidarDados(txtEntradaIHf1, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaIHf1.Text)
        End If
    End Sub

    Private Sub txtEntradaIBf1_Validated(sender As Object, e As EventArgs) Handles txtEntradaIBf1.Validated
        If Validacao.ValidarDados(txtEntradaIBf1, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaIBf1.Text)
        End If
    End Sub

    'SEÇÃO CAIXÃO
    Private Sub txtEntradaCaixaoD_Validated(sender As Object, e As EventArgs) Handles txtEntradaCaixaoD.Validated
        If Validacao.ValidarDados(txtEntradaCaixaoD, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCaixaoD.Text)
        End If
    End Sub

    Private Sub txtEntradaCaixaoBf1_Validated(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB1.Validated
        If Validacao.ValidarDados(txtEntradaCaixaoB1, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCaixaoB1.Text)
        End If
    End Sub

    Private Sub txtEntradaCaixaoH_Validated(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH4.Validated
        If Validacao.ValidarDados(txtEntradaCaixaoH4, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtEntradaCaixaoH4.Text)
        End If
    End Sub

    'SOLICITAÇÕES EXTERNAS
    Private Sub txtNormal_Validated(sender As Object, e As EventArgs) Handles txtNormal.Validated
        If Validacao.ValidarDados(txtNormal, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtNormal.Text)
        End If
    End Sub

    Private Sub txtCortante_Validated(sender As Object, e As EventArgs) Handles txtCortanteX.Validated
        If Validacao.ValidarDados(txtCortanteX, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtCortanteX.Text)
        End If
    End Sub
    'kmod1
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click 'janela Geometria
        form2Kmod1.Show()
        'Form2ResistCalculo.Show()
        'PictureBox2.BackgroundImage = My.Resources.help
    End Sub

    'kmod2:
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        form2Kmod2.Show()
    End Sub

    'kmod3
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        form2Kmod3.Show()
    End Sub
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)  'janela Geometria
        form4Vinculacao.Show()
        PictureBox4.BackgroundImage = My.Resources.help
    End Sub

    Private Sub lblEditarKmod_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblEditarKmod.LinkClicked
        txtKmod.Enabled = True
        txtKmod.Focus()
    End Sub

    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles PictureBox2.Click
        MessageBox.Show("As exigências impostas ao dimensionamento de barras comprimidas sujeitas à instabilidade dependem de seu índice de esbeltez, definido para ambas as direções dos eixos principais de inércia da peça:
                        Peças Curta: λ≤40
                        Peças Medianamente esbelta: 40<λ≤80
                                                          Peças Esbelta:80<λ≤140.")
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        MessageBox.Show(" A NBR 7190:1997 proíbe o emprego de peças cujo comprimento teórico de referência exceda 50 vezes a dimensão transversal correspondente a fim de evitar a ocorrência de estados limites de serviço. Tal disposição limita, portanto, o índice de esbeltez de peças simples ou compostas a um valor máximo de aproximadamente 173.")
    End Sub

    Private Sub TabDadosIniciais_Click(sender As Object, e As EventArgs) Handles TabDadosIniciais.Click

    End Sub

    Private Function vinculacao(l As Double) As Double

        Dim lvinculado As Double = 0
        Select Case cboLvinculado.SelectedIndex
            Case 0
                lblKe.Text = 0.65
                lvinculado = 0.65 * l
            Case 1
                lblKe.Text = 0.8
                lvinculado = 0.65 * l
            Case 2
                lblKe.Text = 1.2
                lvinculado = 0.65 * l
            Case 3
                lblKe.Text = 1
                lvinculado = 0.65 * l
            Case 4
                lblKe.Text = 2.1
                lvinculado = 0.65 * l
            Case 5
                lblKe.Text = 2.4
                lvinculado = 0.65 * l
        End Select

        Return lvinculado
    End Function

    Public Function comprimento() As Double

        If rbtSecaoRetangular.Checked Then
            Return Definicoes.verificaVazio(txtEntradaRetangularL.Text)

        ElseIf rbtSecaoCircular.Checked Then
            Return Definicoes.verificaVazio(txtEntradaCircularL.Text)

            'ARRUMAR AQUI
        ElseIf rbtSecaoT.Checked Then
            Return Definicoes.verificaVazio(txtEntradaCompostaHf.Text)

        ElseIf rbtSecaoI.Checked Then
            Return Definicoes.verificaVazio(txtEntradaCompostaH.Text)

        ElseIf rbtSecaoCaixao.Checked Then
            Return Definicoes.verificaVazio(txtEntradaCaixaoH4.Text)

        ElseIf rbt2Elementos.Checked Then
            Return Definicoes.verificaVazio(txtEntrada2L.Text)

        ElseIf rbt3Elementos.Checked Then
            Return Definicoes.verificaVazio(txtEntrada3L.Text)

        Else
            Return 0
        End If

    End Function

    Public Function entradaElementoJustaposto() As Double

        If rbt2Elementos.Checked Then
            Return Definicoes.verificaVazio(txtEntrada2h1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return Definicoes.verificaVazio(txtEntrada3h1.Text)
        Else
            Return 0
        End If

    End Function

    Private Function coefInfluencia() As Double
        Dim coeficiente As Double = 0

        'classe de umidade 1 e 2 e classe de carregamento permanente e longa duração:
        If txtUmidadeVento.Text <= 65 And cboKmod1.Text = "Permanente" And cboKmod1.Text = "Longa Duração" Or txtUmidadeVento.Text > 65 And txtUmidadeVento.Text <= 75 And cboKmod1.Text = "Permanente" And cboKmod1.Text = "Longa Duração" Then
            coeficiente = 0.8

            'classe de umidade 3 e 4 e classe de carregamento permanente e longa duração:
        ElseIf txtUmidadeVento.Text > 75 And txtUmidadeVento.Text <= 85 And cboKmod1.Text = "Permanente" And cboKmod1.Text = "Longa Duração" Or txtUmidadeVento.Text > 85 And cboKmod1.Text = "Permanente" And cboKmod1.Text = "Longa Duração" Then
            coeficiente = 2

            'classe de umidade 1 e 2 e classe de carregamento média duração
        ElseIf txtUmidadeVento.Text <= 65 And cboKmod1.Text = "Média Duração" Or txtUmidadeVento.Text > 65 And txtUmidadeVento.Text <= 75 And cboKmod1.Text = "Média Duração" Then
            coeficiente = 0.3

            'classe de umidade 2 e 3 e classe de carregamento média duração
        ElseIf txtUmidadeVento.Text > 75 And txtUmidadeVento.Text <= 85 And cboKmod1.Text = "Média Duração" Or txtUmidadeVento.Text > 85 And cboKmod1.Text = "Média Duração" Then
            coeficiente = 1

            'classe de umidade 1 e 2 e classe de carregamento Curta Duração
        ElseIf txtUmidadeVento.Text <= 65 And cboKmod1.Text = "Curta Duração" Or txtUmidadeVento.Text > 65 And txtUmidadeVento.Text <= 75 And cboKmod1.Text = "Curta Duração" Then
            coeficiente = 0.1

            'classe de umidade 2 e 3 e classe de carregamento Curta Duração
        ElseIf txtUmidadeVento.Text > 75 And txtUmidadeVento.Text <= 85 And cboKmod1.Text = "Curta Duração" Or txtUmidadeVento.Text > 85 And cboKmod1.Text = "Curta Duração" Then
            coeficiente = 0.5

            'classe de umidade 1 e 2 e classe de carregamento Instantânea
        ElseIf txtUmidadeVento.Text <= 65 And cboKmod1.Text = "Instantânea" Or txtUmidadeVento.Text > 65 And txtUmidadeVento.Text <= 75 And cboKmod1.Text = "Instantânea" Then
            coeficiente = 0

            'classe de umidade 2 e 3 e classe de carregamento Instantânea
        ElseIf txtUmidadeVento.Text > 75 And txtUmidadeVento.Text <= 85 And cboKmod1.Text = "Instantânea" Or txtUmidadeVento.Text > 85 And cboKmod1.Text = "Instantânea" Then
            coeficiente = 0
        End If

        Return coeficiente
    End Function


    Private Sub txtMx_Validated(sender As Object, e As EventArgs) Handles txtMx.Validated
        If Validacao.ValidarDados(txtMx, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtMx.Text)
        End If
    End Sub

    Private Sub txtMy_Validated(sender As Object, e As EventArgs) Handles txtMy.Validated
        If Validacao.ValidarDados(txtMy, Validacao.TipoValidacao.RealDiferenteDeZero) Then
            fcok = Convert.ToDouble(txtMy.Text)
        End If
    End Sub

    Public Function pilar() As PropriedadesGeometricas.Pilar

        Select Case cboElementFixacao.SelectedIndex
            Case 0
                Return PropriedadesGeometricas.Pilar.Espaçador
            Case 1
                Return PropriedadesGeometricas.Pilar.Chapa
            Case Else
                Return PropriedadesGeometricas.Pilar.Espaçador
        End Select

    End Function

    Private Sub getDadosIniciais(tipoSecao As Madeira.TipoSecao)

        Select Case tipoSecao
            Case Madeira.TipoSecao.SecaoT
                madeiraSelecionada.DadosIniciais.b1 = 0 'OK
                madeiraSelecionada.DadosIniciais.b2 = txtEntradaCompostaBw.Text 'OK
                madeiraSelecionada.DadosIniciais.b3 = txtEntradaCompostaBf.Text 'OK
                madeiraSelecionada.DadosIniciais.h1 = 0 'OK
                madeiraSelecionada.DadosIniciais.h2 = txtEntradaCompostaTh.Text 'OK
                madeiraSelecionada.DadosIniciais.h3 = txtEntradaCompostaHf.Text 'OK

            Case Madeira.TipoSecao.SecaoI
                madeiraSelecionada.DadosIniciais.b1 = txtEntradaIBf2.Text 'OK
                madeiraSelecionada.DadosIniciais.b2 = txtEntradaIBw.Text 'OK
                madeiraSelecionada.DadosIniciais.b3 = txtEntradaIBf1.Text 'OK
                madeiraSelecionada.DadosIniciais.h1 = txtEntradaIHf2.Text 'OK
                madeiraSelecionada.DadosIniciais.h2 = txtEntradaIH.Text 'OK
                madeiraSelecionada.DadosIniciais.h3 = txtEntradaIHf1.Text 'OK

            Case Madeira.TipoSecao.Caixao
                madeiraSelecionada.DadosIniciais.b1c = txtEntradaCaixaoB1.Text 'OK
                madeiraSelecionada.DadosIniciais.b2c = txtEntradaCaixaoB2.Text 'OK
                madeiraSelecionada.DadosIniciais.b3c = txtEntradaCaixaoB3.Text 'OK
                madeiraSelecionada.DadosIniciais.b4c = txtEntradaCaixaoB4.Text 'OK
                madeiraSelecionada.DadosIniciais.h1c = txtEntradaCaixaoH1.Text 'OK
                madeiraSelecionada.DadosIniciais.h2c = txtEntradaCaixaoH2.Text 'OK
                madeiraSelecionada.DadosIniciais.h3c = txtEntradaCaixaoH3.Text 'OK
                madeiraSelecionada.DadosIniciais.h4c = txtEntradaCaixaoH4.Text 'OK
        End Select

    End Sub

End Class
