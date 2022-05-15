Imports TCC2.Kmod1
Imports TCC2.Tracao
Imports TCC2.Compressao
Imports TCC2.Cisalhamento
Imports TCC2.Flexao
Imports TCC2.Madeira
Public Class Form1

    Public madeiraSelecionada As Madeira = New Madeira 'chamei as propriedades da madeira (modelei as propriedades para ser igual a classe madeira)
    Public tipoMadeira As Integer
    Public kmod1, kmod2, kmod3, kmod As Double
    Public Shared proprGeometricas As PropriedadesGeometricas = New PropriedadesGeometricas 'inicia como zerado, mas pode ser usado em qualquer lugar do form 1. Então ela é auto preenchida qunado precisa dela
    Public form2Kmod1 As Form2Kmod1 = New Form2Kmod1
    Public form2Kmod2 As Form2Kmod2 = New Form2Kmod2
    Public form2Kmod3 As Form2Kmod3 = New Form2Kmod3
    Public form4Vinculacao As Form4Vinculacao = New Form4Vinculacao

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbldata.Text = Date.Now

        'For: Laço de repetição
        'counter: Inteira - começo com 0. Conta a lista inteira da madeira com o "for" preenchendo o ComboBox
        For counter As Integer = 0 To ListaEspecies.Count - 1
            cboMadeiraUtilizada.Items.Add(ListaEspecies(counter).name)
        Next counter

        CaracteristicaMadeira()
    End Sub

    'SAINDO DO PROGRAMA NO BOTÃO CLICK
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
        'madeiraSelecionada = ListaEspecies(madeiraSelecionada.Id)
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
        madeiraSelecionada = ListaEspecies(cboMadeiraUtilizada.SelectedIndex)
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
        kmod1 = PropriedadesResistencia.Kmod1Final(tipoMadeira, cboKmod1.SelectedIndex)
        lblKmod1.Text = kmod1.ToString
        lblKmod1.Visible = True
    End Sub
    Private Sub cboKmod2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKmod2.SelectedIndexChanged
        kmod2 = PropriedadesResistencia.Kmod2Final(tipoMadeira, cboKmod2.SelectedIndex)
        lblKmod2.Text = kmod2.ToString
        lblKmod2.Visible = True
    End Sub

    Private Sub cboKmod3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboKmod3.SelectedIndexChanged
        kmod3 = PropriedadesResistencia.Kmod3Final(tipoMadeira, cboKmod3.SelectedIndex)
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

    Private Sub btnCalcularPropriedadesGeometricas_Click(sender As Object, e As EventArgs) Handles btnCalcularPropriedadesGeometricas.Click
        Dim tipoSecao As Madeira.TipoSecao

        If rbtSecaoRetangular.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoRetangular(txtEntradaRetangularBx.Text, txtEntradaRetangularBy.Text, txtEntradaRetangularL.Text)
            tipoSecao = Madeira.TipoSecao.Retangular

        ElseIf rbtSecaoCircular.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCircular(txtEntradaCircularD.Text, txtEntradaCircularL.Text)
            tipoSecao = Madeira.TipoSecao.Circular

        ElseIf rbtSecaoT.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCompostoSecaoT(txtEntradaCompostaBf.Text, txtEntradaCompostaHf.Text, txtEntradaCompostaTh.Text, txtEntradaCompostaBw.Text)
            tipoSecao = Madeira.TipoSecao.SecaoT

        ElseIf rbtSecaoI.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCompostoSecaoI(txtEntradaIBf1.Text, txtEntradaIHf1.Text, txtEntradaIBf2.Text, txtEntradaIHf2.Text, txtEntradaIBw.Text, txtEntradaID.Text)
            tipoSecao = Madeira.TipoSecao.SecaoI

        ElseIf rbtSecaoCaixao.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCompostoSecaoCaixao(txtEntradaCaixaoD.Text, txtEntradaCaixaoB1.Text, txtEntradaCaixaoB2.Text, txtEntradaCaixaoB3.Text, txtEntradaCaixaoB4.Text, txtEntradaCaixaoH1.Text, txtEntradaCaixaoH2.Text, txtEntradaCaixaoH3.Text, txtEntradaCaixaoH4.Text)
            tipoSecao = Madeira.TipoSecao.Caixao

        ElseIf rbt2Elementos.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.Calculo2Elementos(txtEntrada2L.Text, txtEntrada2L1.Text, txtEntrada2a.Text, txtEntrada2a1.Text, txtEntrada2h.Text, txtEntrada2h1.Text, txtEntrada2b1.Text, pilar())
            tipoSecao = Madeira.TipoSecao.ElementosJustaposto2

        ElseIf rbt3Elementos.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.Calculo3Elementos(txtEntrada3L.Text, txtEntrada3L1.Text, txtEntrada3a.Text, txtEntrada3a1.Text, txtEntrada3h.Text, txtEntrada3h1.Text, txtEntrada3b1.Text, pilar())
            tipoSecao = Madeira.TipoSecao.ElementosJustaposto3
        Else
        End If
    End Sub


    Private Sub btnCalcularPropriedades_Click(sender As Object, e As EventArgs) Handles btnCalcularPropriedades.Click

        Dim tipoSecao As Madeira.TipoSecao
        'Dim tipoSecao

        If rbtSecaoRetangular.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoRetangular(txtEntradaRetangularBx.Text, txtEntradaRetangularBy.Text, txtEntradaRetangularL.Text)
            tipoSecao = Madeira.TipoSecao.Retangular

        ElseIf rbtSecaoCircular.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCircular(txtEntradaCircularD.Text, txtEntradaCircularL.Text)
            tipoSecao = Madeira.TipoSecao.Circular

        ElseIf rbtSecaoT.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCompostoSecaoT(txtEntradaCompostaBf.Text, txtEntradaCompostaHf.Text, txtEntradaCompostaTh.Text, txtEntradaCompostaBw.Text)
            tipoSecao = Madeira.TipoSecao.SecaoT

        ElseIf rbtSecaoI.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCompostoSecaoI(txtEntradaIBf1.Text, txtEntradaIHf1.Text, txtEntradaIBf2.Text, txtEntradaIHf2.Text, txtEntradaIBw.Text, txtEntradaID.Text)
            tipoSecao = Madeira.TipoSecao.SecaoI

        ElseIf rbtSecaoCaixao.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.CalculoCompostoSecaoCaixao(txtEntradaCaixaoD.Text, txtEntradaCaixaoB1.Text, txtEntradaCaixaoB2.Text, txtEntradaCaixaoB3.Text, txtEntradaCaixaoB4.Text, txtEntradaCaixaoH1.Text, txtEntradaCaixaoH2.Text, txtEntradaCaixaoH3.Text, txtEntradaCaixaoH4.Text)
            tipoSecao = Madeira.TipoSecao.Caixao

        ElseIf rbt2Elementos.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.Calculo2Elementos(txtEntrada2L.Text, txtEntrada2L1.Text, txtEntrada2a.Text, txtEntrada2a1.Text, txtEntrada2h.Text, txtEntrada2h1.Text, txtEntrada2b1.Text, pilar())
            tipoSecao = Madeira.TipoSecao.ElementosJustaposto2

        ElseIf rbt3Elementos.Checked Then
            madeiraSelecionada.propriedadesGeometricas = proprGeometricas.Calculo3Elementos(txtEntrada3L.Text, txtEntrada3L1.Text, txtEntrada3a.Text, txtEntrada3a1.Text, txtEntrada3h.Text, txtEntrada3h1.Text, txtEntrada3b1.Text, pilar())
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
            madeiraSelecionada.tracao = CalcTensaoT(
                PropriedadesResistencia.verificaVazio(txtNormal.Text),
                comprimento(),
                proprGeometricas,
                tipoSecao,
                PropriedadesResistencia.verificaVazio(txtEntrada2L.Text),
                PropriedadesResistencia.verificaVazio(txtEntrada2b1.Text),
                PropriedadesResistencia.verificaVazio(txtEntrada2a.Text),
                cboElementFixacao.Text
                )


            'COMPRESSAO
        ElseIf txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            madeiraSelecionada.compressao = CalcTensaoC(
                PropriedadesResistencia.verificaVazio(txtNormal.Text),
                PropriedadesResistencia.verificaVazio(txtMx.Text),
                PropriedadesResistencia.verificaVazio(txtMy.Text),
                vinculacao(comprimento),
                madeiraSelecionada.propriedadesGeometricas,
                tipoSecao,
                PropriedadesResistencia.verificaVazio(txtMomentoCargaPermanenteX.Text),
                PropriedadesResistencia.verificaVazio(txtMomentoCargaPermanenteY.Text),
                PropriedadesResistencia.verificaVazio(txtMomentoCargaVariavelX.Text),
                PropriedadesResistencia.verificaVazio(txtMomentoCargaVariavelY.Text),
                PropriedadesResistencia.verificaVazio(txtNormalCargaPermanente.Text),
                PropriedadesResistencia.verificaVazio(txtNormalCargaVariavel.Text),
                coefInfluencia(),
                PropriedadesResistencia.verificaVazio(txtF1.Text),
                PropriedadesResistencia.verificaVazio(txtF2.Text),
                vinculacao(comprimento),
                PropriedadesResistencia.verificaVazio(txtEntrada2b1.Text),
                PropriedadesResistencia.verificaVazio(txtEntrada2a.Text),
                cboElementFixacao.Text,
                entradaElementoJustaposto)
        End If

        'CISALHAMENTO
        madeiraSelecionada.cisalhamento = CalcTensaoCisalhamento(
             PropriedadesResistencia.verificaVazio(txtCortanteX.Text),
             PropriedadesResistencia.verificaVazio(txtCortanteY.Text),
             PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text),
             cisalhamento_a1,
             cisalhamento_L1,
             proprGeometricas,
             tipoSecao
             )

        'FLEXÃO SIMPLES RETA

        If txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtCortanteX.Text.ToString <> "" And rbtFlexaoSimples.Checked Then
            gbxFlexaoSimples.Visible = True
            gbxFlexaoObliqua.Visible = False
            madeiraSelecionada.flexao = CalcTensaoFlexaoSimples(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
           madeiraSelecionada.propriedadesGeometricas,
           tipoSecao
           )
            'FLEXÃO OBLIQUA
        ElseIf txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtCortanteX.Text.ToString <> "" And rbtFlexaoObliqua.Checked Then
            gbxFlexaoSimples.Visible = False
            gbxFlexaoObliqua.Visible = True
            madeiraSelecionada.flexao = CalcTensaoFlexaoObliqua(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            madeiraSelecionada.propriedadesGeometricas,
            tipoSecao
            )
        End If

        'FLEXAO COMPOSTA
        '------------------------FLEXOTRAÇÃO------------------------
        If txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Tração" Then
            gbxFlexoTracao.Visible = True
            gbxFlexoCompressao.Visible = False
            madeiraSelecionada.flexao = CalcTensaoFlexotracao(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            madeiraSelecionada.propriedadesGeometricas,
            tipoSecao
            )

            '------------------------FLEXOCOMPRESSAO------------------------
        ElseIf txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            gbxFlexoCompressao.Visible = True
            gbxFlexoTracao.Visible = False
            madeiraSelecionada.flexao = CalcTensaoFlexocompressao(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            madeiraSelecionada.propriedadesGeometricas,
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
        txtTensaoTracao.Text = madeiraSelecionada.tracao.TensaoTracao
        txtEsbeltezTracaoX.Text = madeiraSelecionada.Tracao.esbeltezTracaoX
        txtEsbeltezTracaoY.Text = madeiraSelecionada.Tracao.esbeltezTracaoY

        'VALIDAÇÃO TENSÃO TRAÇÃO:
        If PropriedadesResistencia.validarTensaoTracao(madeiraSelecionada.Tracao.tensaoTracao, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela) Then
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
        If PropriedadesResistencia.validarEsbeltezTracaoX(madeiraSelecionada.Tracao.esbeltezTracaoX) Then
            lblValidacaoEsbeltezXTracao.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezXTracao.Visible = True
            lblValidacaoEsbeltezXTracao.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltezXTracao.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezXTracao.Visible = True
            lblValidacaoEsbeltezXTracao.ForeColor = Color.Red
        End If

        'VALIDAÇÃO ESBELTEZ Y:
        If PropriedadesResistencia.validarEsbeltezTracaoY(madeiraSelecionada.Tracao.esbeltezTracaoY) Then
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
        txtEsbeltezCompressaoX.Text = madeiraSelecionada.Compressao.esbeltezCompressaoX
        txtEsbeltezCompressaoY.Text = madeiraSelecionada.Compressao.esbeltezCompressaoY
        txtTensaoCompressaoCurtaX.Text = madeiraSelecionada.Compressao.tensaoCompressaoCurtaX
        txtTensaoCompressaoCurtaY.Text = madeiraSelecionada.Compressao.tensaoCompressaoCurtaY
        txtEsforçoNormalMedEsbX.Text = madeiraSelecionada.Compressao.tensaoCompressaoMedEsbeltaX
        txtEsforçoFlexaoMedEsbX.Text = madeiraSelecionada.Compressao.tensaoMxd
        txtEsforçoNormalMedEsbY.Text = madeiraSelecionada.Compressao.tensaoCompressaoMedEsbeltaY
        txtEsforçoFlexaoMedEsbY.Text = madeiraSelecionada.Compressao.tensaoMyd
        txtEsforçoNormalEsbX.Text = madeiraSelecionada.Compressao.tensaoCompressaoEsbeltaX
        txtEsforçoFlexaoEsbX.Text = madeiraSelecionada.Compressao.tensaoMxd
        txtEsforçoNormalEsbY.Text = madeiraSelecionada.Compressao.tensaoCompressaoEsbeltaY
        txtEsforçoFlexaoEsbY.Text = madeiraSelecionada.Compressao.tensaoMyd
        txtForçaElasticaX.Text = madeiraSelecionada.Compressao.forcaElasticaX
        txtForçaElasticaY.Text = madeiraSelecionada.Compressao.forcaElasticaY
        txtEaX.Text = madeiraSelecionada.Compressao.ea
        txtEaY.Text = madeiraSelecionada.Compressao.ea
        txtEdX.Text = madeiraSelecionada.Compressao.edx
        txtEdY.Text = madeiraSelecionada.Compressao.edy
        txtEiX.Text = madeiraSelecionada.Compressao.ei
        txtEiY.Text = madeiraSelecionada.Compressao.ei
        txtE1efX.Text = madeiraSelecionada.Compressao.e1ef
        txtE1efY.Text = madeiraSelecionada.Compressao.e1ef
        txtEcX.Text = madeiraSelecionada.Compressao.ec
        txtEcY.Text = madeiraSelecionada.Compressao.ec

        'VALIDAÇÃO COMPRESSÂO:


        'VALIDAÇÃO COMPRESSÃO CURTAX:
        If PropriedadesResistencia.validarTensaoCompressaoCurtaX(madeiraSelecionada.Compressao.tensaoCompressaoCurtaX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
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
        If PropriedadesResistencia.validarTensaoCompressaoCurtaY(madeiraSelecionada.Compressao.tensaoCompressaoCurtaY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
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

        If PropriedadesResistencia.validarTensaoCompressaoMedEsbX(madeiraSelecionada.Compressao.tensaoCompressaoMedEsbeltaX, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
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
        If PropriedadesResistencia.validarTensaoCompressaoMedEsbY(madeiraSelecionada.Compressao.tensaoCompressaoMedEsbeltaX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
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
        If PropriedadesResistencia.validarTensaoCompressaoEsbX(madeiraSelecionada.Compressao.tensaoCompressaoEsbeltaY, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
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
        If PropriedadesResistencia.validarTensaoCompressaoEsbY(madeiraSelecionada.Compressao.tensaoCompressaoEsbeltaY, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
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
        txtTensaoCisalhanteX.Text = madeiraSelecionada.Cisalhamento.tensaoCisalhanteX
        txtTensaoCisalhanteY.Text = madeiraSelecionada.Cisalhamento.tensaoCisalhanteY

        'VALIDAÇÃO TENSÃO CISALHANTE X:
        If PropriedadesResistencia.validarTensaoCisalhanteX(madeiraSelecionada.Cisalhamento.tensaoCisalhanteX, madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento) Then
            lblValidacaoCisalhanteX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteX.Visible = True
            lblValidacaoCisalhanteX.ForeColor = Color.Green
        Else
            lblValidacaoCisalhanteX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteX.Visible = True
            lblValidacaoCisalhanteX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO TENSÃO CISALHANTE Y:
        If PropriedadesResistencia.validarTensaoCisalhanteY(madeiraSelecionada.Cisalhamento.tensaoCisalhanteX, madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento) Then
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
        If PropriedadesResistencia.validarTensaoFlexaoSimplesRetaX(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela) Then
            lblValidacaoFlexaoSimplesX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesX.Visible = True
            lblValidacaoFlexaoSimplesX.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoSimplesX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesX.Visible = True
            lblValidacaoFlexaoSimplesX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXÃO SIMPLES RETA Y:
        If PropriedadesResistencia.validarTensaoFlexaoSimplesRetaY(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblValidacaoFlexaoSimplesY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesY.Visible = True
            lblValidacaoFlexaoSimplesY.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoSimplesY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoSimplesY.Visible = True
            lblValidacaoFlexaoSimplesY.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXÃO SIMPLES OBLIQUA X:
        If PropriedadesResistencia.validarTensaoFlexaoSimplesObliquaX(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexaoObliquaX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoObliquaX.Visible = True
            lblValidacaoFlexaoObliquaX.ForeColor = Color.Green
        Else
            lblValidacaoFlexaoObliquaX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexaoObliquaX.Visible = True
            lblValidacaoFlexaoObliquaX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXÃO SIMPLES OBLIQUA Y:
        If PropriedadesResistencia.validarTensaoFlexaoSimplesObliquaY(madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, tipoSecao) Then
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
        txtFlexoTracao.Text = madeiraSelecionada.Tracao.tensaoTracao
        txtFlexoTracaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtFlexoTracaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY
        txtFlexoCompressao.Text = madeiraSelecionada.Compressao.tensaoCompressao
        txtFlexoCompressaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtFlexoCompressaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY

        'VALIDAÇÃO FLEXOTRAÇÃO X:
        If PropriedadesResistencia.validarTensaoFlexotracaoX(madeiraSelecionada.Tracao.tensaoTracao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoTracaoX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoX.Visible = True
            lblValidacaoFlexoTracaoX.ForeColor = Color.Green
        Else
            lblValidacaoFlexoTracaoX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoX.Visible = True
            lblValidacaoFlexoTracaoX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOTRAÇÃO Y:
        If PropriedadesResistencia.validarTensaoFlexotracaoY(madeiraSelecionada.Tracao.tensaoTracao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, tipoSecao) Then
            lblValidacaoFlexoTracaoY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoY.Visible = True
            lblValidacaoFlexoTracaoY.ForeColor = Color.Green
        Else
            lblValidacaoFlexoTracaoY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoY.Visible = True
            lblValidacaoFlexoTracaoY.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOCOMPRESSAO X:
        If PropriedadesResistencia.validarTensaoFlexocompressaoX(madeiraSelecionada.Compressao.tensaoCompressao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoCompressaoX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoX.Visible = True
            lblValidacaoFlexoCompressaoX.ForeColor = Color.Green
        Else
            lblValidacaoFlexoCompressaoX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoX.Visible = True
            lblValidacaoFlexoCompressaoX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOCOMPRESSAO Y:
        If PropriedadesResistencia.validarTensaoFlexocompressaoY(madeiraSelecionada.Compressao.tensaoCompressao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
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
        madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela = PropriedadesResistencia.resistCalCompParalela(kmod, madeiraSelecionada.ResistCompParalela)
        madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela = PropriedadesResistencia.resistCalTracaoParalela(kmod, madeiraSelecionada.ResistTracaoParalela)
        madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaNormal = PropriedadesResistencia.resistCalTracaoNormal(kmod, madeiraSelecionada.ResistTracaoNormal)
        madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento = PropriedadesResistencia.resistCalAoCisalhamento(kmod, madeiraSelecionada.ResistAoCisalhamento)
        Dim modElasticidade = PropriedadesResistencia.ModElasticidade(kmod, madeiraSelecionada.ModElasticidade)
        madeiraSelecionada.ResistenciaCalculo.moduloElasticidade = modElasticidade
        madeiraSelecionada.ResistenciaCalculo.moduloElasticidadeTransversal = PropriedadesResistencia.ModElasticidadeTransversal(modElasticidade)
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
        txtEntradaCompostaH.Text = PropriedadesResistencia.somaT(PropriedadesResistencia.verificaVazio(txtEntradaCompostaHf.Text), PropriedadesResistencia.verificaVazio(txtEntradaCompostaTh.Text))
    End Sub
    Private Sub txtEntradaCompostaHf_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaHf.TextChanged
        txtEntradaCompostaH.Text = PropriedadesResistencia.somaT(PropriedadesResistencia.verificaVazio(txtEntradaCompostaHf.Text), PropriedadesResistencia.verificaVazio(txtEntradaCompostaTh.Text))
    End Sub

    'SEÇÃO I----------------------------------------
    Private Sub txtEntradaIHf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIHf1.TextChanged
        txtEntradaIHf2.Text = txtEntradaIHf1.Text
        txtEntradaID.Text = PropriedadesResistencia.somaI(PropriedadesResistencia.verificaVazio(txtEntradaIHf1.Text), PropriedadesResistencia.verificaVazio(txtEntradaIH.Text), PropriedadesResistencia.verificaVazio(txtEntradaIHf2.Text))
    End Sub
    Private Sub txtEntradaIH_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIH.TextChanged
        txtEntradaID.Text = PropriedadesResistencia.somaI(PropriedadesResistencia.verificaVazio(txtEntradaIHf1.Text), PropriedadesResistencia.verificaVazio(txtEntradaIH.Text), PropriedadesResistencia.verificaVazio(txtEntradaIHf2.Text))
    End Sub
    Private Sub txtEntradaIHf2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIHf2.TextChanged
        txtEntradaID.Text = PropriedadesResistencia.somaI(PropriedadesResistencia.verificaVazio(txtEntradaIHf1.Text), PropriedadesResistencia.verificaVazio(txtEntradaIH.Text), PropriedadesResistencia.verificaVazio(txtEntradaIHf2.Text))
    End Sub
    Private Sub txtEntradaIBf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIBf1.TextChanged
        txtEntradaIBf2.Text = txtEntradaIBf1.Text
    End Sub

    'SEÇÃO CAIXÃO----------------------------------------
    Private Sub txtEntradaCaixaoH1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH1.TextChanged
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixao(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH3.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub
    Private Sub txtEntradaCaixaoH2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH2.TextChanged
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixao(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH3.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub
    Private Sub txtEntradaCaixaoH3_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH3.TextChanged
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixao(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH3.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub
    Private Sub txtEntradaCaixaoH4_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH4.TextChanged
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixao(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH3.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
    End Sub
    Private Sub txtEntradaCaixaoBf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB1.TextChanged
        txtEntradaCaixaoB2.Text = txtEntradaCaixaoB1.Text
    End Sub

    '3 ELEMENTOS JUSTAPOSTO----------------------------------------
    Private Sub txtEntrada3a_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3a.TextChanged
        txtEntrada3a1.Text = (txtEntrada3a.Text) / 2
    End Sub
    Private Sub txtEntrada3h_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3h.TextChanged
        txtEntrada3h.Text = txtEntrada3a.Text + txtEntrada3b1.Text
    End Sub

    Private Sub cboTracaoCompressao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTracaoCompressao.SelectedIndexChanged
        TabTracao.Enabled = cboTracaoCompressao.SelectedIndex.CompareTo(1)
        TabCompressao.Enabled = cboTracaoCompressao.SelectedIndex.CompareTo(0)
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
        txtAreaSecao.Text = resultProprGeometricas.area
        txtEixoXMomentoInercia.Text = resultProprGeometricas.eixoXmi
        txtEixoXRaioGiracao.Text = resultProprGeometricas.eixoXrg
        txtEixoXModuloResistencia.Text = resultProprGeometricas.eixoXmr
        txtEixoXInerciaEfetiva.Text = resultProprGeometricas.eixoXie
        txtEixoYMomentoInercia.Text = resultProprGeometricas.eixoYmi
        txtEixoYRaioGiracao.Text = resultProprGeometricas.eixoYrg
        txtEixoYModuloResistencia.Text = resultProprGeometricas.eixoYmr
        txtEixoYInerciaEfetiva.Text = resultProprGeometricas.eixoYie
    End Sub

    Private Sub ResultadoPropriedadesGeometricas23Elemento(resultProprGeometricas As PropriedadesGeometricas)
        txtArea1.Text = resultProprGeometricas.areaA1
        txtArea.Text = resultProprGeometricas.area
        txtInterpoIx1.Text = resultProprGeometricas.eixoXmi1
        txtInterpoIx.Text = resultProprGeometricas.eixoXmi
        txtInterpoIy2.Text = resultProprGeometricas.eixoYmi1
        txtInterpoIy.Text = resultProprGeometricas.eixoYmi
        txtInterpoEixoYCoefB.Text = resultProprGeometricas.coefBy
        txtInterpoIModuloResist.Text = resultProprGeometricas.eixoYmr
        txtInterpoEixoYInerciaEfetiva.Text = resultProprGeometricas.eixoYie
    End Sub

    'VALIDAÇÕES TXT PROPRIEDADES DA MADEIRA =========================================================================
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

    'JANELA GEOMETRICA DE AJUDA
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

    'DEFINIÇÃO DA VINCULAÇÃO PARA ESFORÇOS DE COMPRESSÃO =========================
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

    'DEFINIÇÃO DO COEFICIENT DEVIDO A FLUENCIA DO MATERIAL PARA ESFORÇOS DE COMPRESSÃO QUANDO A PEÇA É ESBELTA =================
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

    'CALCULANDO O CISALHAMENTO NOS ESPAÇADORES ===================
    Public Function cisalhamento_a1() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2a1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3a1.Text)
        Else
            Return 0
        End If
    End Function
    Public Function cisalhamento_L1() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2L1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3L1.Text)
        Else
            Return 0
        End If
    End Function

    Public Function comprimento() As Double

        If rbtSecaoRetangular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaRetangularL.Text)

        ElseIf rbtSecaoCircular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaCircularL.Text)

            'ARRUMAR AQUI
        ElseIf rbtSecaoT.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtTComp.Text)

        ElseIf rbtSecaoI.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtIComp.Text)

        ElseIf rbtSecaoCaixao.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtCaixaoComp.Text)

        ElseIf rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2L.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3L.Text)

        Else
            Return 0
        End If

    End Function

    Public Function entradaElementoJustaposto() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2h1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3h1.Text)
        Else
            Return 0
        End If

    End Function

    'DEFINIÇÃO DO ELEMENTO DE FIXAÇÃO DO ESPAÇADOR DO PILAR COMPOSTO
    Public Function pilar() As PropriedadesGeometricas.Pilar

        Select Case cboElementFixacao.SelectedIndex
            Case 0
                Return PropriedadesGeometricas.Pilar.Espaçador
            Case 1
                Return PropriedadesGeometricas.Pilar.Chapa
                'Case Else
                'Return PropriedadesGeometricas.Pilar.Espaçador
        End Select
    End Function

    'DEFINICÇÃO DOS DADOS INICIAIS PARA O CÁCLULO DO MOMENTO DE INERCIA, MOMENTO DE AREA EFETIVO E CG DAS PEÇAS COMPOSTAS (SEÇÃI T, I e CAIXÃO)
    Private Sub getDadosIniciais(tipoSecao As Madeira.TipoSecao)

        Select Case tipoSecao
            Case Madeira.TipoSecao.SecaoT
                madeiraSelecionada.PropriedadesGeometricas.b1 = 0 'OK
                madeiraSelecionada.PropriedadesGeometricas.b2 = txtEntradaCompostaBw.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.b3 = txtEntradaCompostaBf.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h1 = 0 'OK
                madeiraSelecionada.PropriedadesGeometricas.h2 = txtEntradaCompostaTh.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h3 = txtEntradaCompostaHf.Text 'OK

            Case Madeira.TipoSecao.SecaoI
                madeiraSelecionada.PropriedadesGeometricas.b1 = txtEntradaIBf2.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.b2 = txtEntradaIBw.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.b3 = txtEntradaIBf1.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h1 = txtEntradaIHf2.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h2 = txtEntradaIH.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h3 = txtEntradaIHf1.Text 'OK

            Case Madeira.TipoSecao.Caixao
                madeiraSelecionada.PropriedadesGeometricas.b1c = txtEntradaCaixaoB1.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.b2c = txtEntradaCaixaoB2.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.b3c = txtEntradaCaixaoB3.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.b4c = txtEntradaCaixaoB4.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h1c = txtEntradaCaixaoH1.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h2c = txtEntradaCaixaoH2.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h3c = txtEntradaCaixaoH3.Text 'OK
                madeiraSelecionada.PropriedadesGeometricas.h4c = txtEntradaCaixaoH4.Text 'OK
        End Select
    End Sub

End Class
