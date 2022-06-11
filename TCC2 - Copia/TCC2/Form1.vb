Imports TCC2.Kmod1
Imports TCC2.Tracao
Imports TCC2.Compressao
Imports TCC2.Cisalhamento
Imports TCC2.Flexao
Imports TCC2.Madeira
Imports TCC2.FormAjuda
Imports TCC2.frmEsforcoCaracteristico

Imports System.IO
Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class Form1
    'Public Projeto As New CProjeto
    Public madeiraSelecionada As Madeira = New Madeira 'chamei as propriedades da madeira (modelei as propriedades para ser igual a classe madeira)
    Public tipoMadeira As Integer
    Public kmod1, kmod2, kmod3, kmod As Double
    Public Shared proprGeometricas As PropriedadesGeometricas = New PropriedadesGeometricas 'inicia como zerado, mas pode ser usado em qualquer lugar do form 1. Então ela é auto preenchida qunado precisa dela
    Public form2Kmod1 As Form2Kmod1 = New Form2Kmod1
    Public form2Kmod2 As Form2Kmod2 = New Form2Kmod2
    Public form2Kmod3 As Form2Kmod3 = New Form2Kmod3
    Public formAjuda As FormAjuda = New FormAjuda
    Public form4Vinculacao As Form4Vinculacao = New Form4Vinculacao

    Public Shared momentoCargaPermanenteX As Double = 0
    Public Shared momentoCargaPermanenteY As Double = 0
    Public Shared momentoCargaVariavelX As Double = 0
    Public Shared momentoCargaVariavelY As Double = 0
    Public Shared normalCargaPermanente As Double = 0
    Public Shared normalCargaVariavel As Double = 0
    Public Shared f1 As Double = 0
    Public Shared f2 As Double = 0
    Public Shared voltou As Boolean = False
    Public Shared esbelta As Boolean = False

    'ROTINA PARA APARECER A DATA NO CANTINHO DA TELA E SELECINAR A LISTA DE MADEIRA
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbldata.Text = Date.Now

        'For: Laço de repetição
        'counter: Inteira - começo com 0. Conta a lista inteira da madeira com o "for" preenchendo o ComboBox
        For counter As Integer = 0 To ListaEspecies.Count - 1
            cboMadeiraUtilizada.Items.Add(ListaEspecies(counter).name)
        Next counter

        CaracteristicaMadeira()
    End Sub

    'ROTINA PARA SAIR DO PROGRAMA AO APERTAR SAIR
    Private Sub SairToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SairToolStripMenuItem.Click
        Application.Exit()
    End Sub

    ' ROTINA PARA RECEBER OS VALORES DA LISTA QUE FOI SELECIONADO NO FORM MADEIRA
    Private Sub cboMadeiraUtilizada_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMadeiraUtilizada.SelectedIndexChanged
        madeiraSelecionada = ListaEspecies(cboMadeiraUtilizada.SelectedIndex)
        PreenchimentoCaracteristicas()
    End Sub
    'ROTINA PARA ALTERAAR OS VALORES DA RESISTENCIA CARACTERISTICA DA MADEIRA CASO O USUARIO DESEJE
    Private Sub CaracteristicaMadeira()
        Dim checked = rbtInserirValores.Checked

        txtResistCompParalela.Enabled = checked
        txtResistTracaoParalela.Enabled = checked
        txtResistTracaoNormal.Enabled = checked
        txtResistAoCisalhamento.Enabled = checked
        txtModElasticidade.Enabled = checked
        txtDensAparente.Enabled = checked
    End Sub

    'ROTINA PARA USAR VALORES DAS MADEIRAS DIFERENTES DA NBR: Se o button estiver checked ele é visivel
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
        'FORMATANDO A QUANTIDADE DE CASAS DPS DA VIRGULA NOS TXT
        txtResistCompParalela.Text = Format(madeiraSelecionada.ResistCompParalela, "#0.0##;-#0.0##")
        txtResistTracaoParalela.Text = Format(madeiraSelecionada.ResistTracaoParalela, "#0.0##;-#0.0##")
        txtResistTracaoNormal.Text = Format(madeiraSelecionada.ResistTracaoNormal, "#0.0##;-#0.0##")
        txtResistAoCisalhamento.Text = Format(madeiraSelecionada.ResistAoCisalhamento, "#0.0##;-#0.0##")
        txtModElasticidade.Text = Format(madeiraSelecionada.ModElasticidade, "#0.0##;-#0.0###")
        txtDensAparente.Text = Format(madeiraSelecionada.DensAparente, "#0.0##;-#0.0##")

    End Sub

    Private Sub btnConfirmarInsercao_Click(sender As Object, e As EventArgs) Handles btnConfirmarInsercao.Click

        btnConfirmarInsercao.Visible = False
        madeiraSelecionada.ResistCompParalela = txtResistCompParalela.Text
        madeiraSelecionada.ResistTracaoParalela = txtResistTracaoParalela.Text
        madeiraSelecionada.ResistTracaoNormal = txtResistTracaoNormal.Text
        madeiraSelecionada.ResistAoCisalhamento = txtResistAoCisalhamento.Text
        madeiraSelecionada.ModElasticidade = txtModElasticidade.Text
        madeiraSelecionada.DensAparente = txtDensAparente.Text
        'FORMATANDO A QUANTIDADE DE CASAS DPS DA VIRGULA NOS TXT
        madeiraSelecionada.ResistCompParalela = Format(txtResistCompParalela.Text, "#0.0##;-#0.0##")
        madeiraSelecionada.ResistTracaoParalela = Format(txtResistTracaoParalela.Text, "#0.0##;-#0.0##")
        madeiraSelecionada.ResistTracaoNormal = Format(txtResistTracaoNormal.Text, "#0.0##;-#0.0##")
        madeiraSelecionada.ResistAoCisalhamento = Format(txtResistAoCisalhamento.Text, "#0.0##;-#0.0##")
        madeiraSelecionada.ModElasticidade = Format(txtModElasticidade.Text, "#0.0##;-#0.0###")
        madeiraSelecionada.DensAparente = Format(txtDensAparente.Text, "#0.0##;-#0.0##")

        txtResistCompParalela.Enabled = False
        txtResistTracaoParalela.Enabled = False
        txtResistTracaoNormal.Enabled = False
        txtResistAoCisalhamento.Enabled = False
        txtModElasticidade.Enabled = False
        txtDensAparente.Enabled = False

    End Sub

    'ROTINA PARA APARECER OS VALORES DE KMOD1, KMOD2 E KMOD 3 CONFORME A SELÇÃO DO USUARIO
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

    'ROTINA PARA CALCULAR O VALOR DO KMOD
    Private Sub bntCalcularKmod_Click(sender As Object, e As EventArgs) Handles bntCalcularKmod.Click
        CalcularKmod()
        ResistCalculo()
    End Sub

    'ROTINA PARA CALCULAR O KMOD
    Private Sub CalcularKmod()
        kmod = kmod1 * kmod2 * kmod3
        txtKmod.Text = kmod
        txtKmod.Text = Format(kmod, "#0.0##;-#0.0##")
    End Sub

    'ROTINA PARA SELECIONAR O TIPO DA MADEIRA E ASSIM SEUS VALORES CARACTERISTICOS
    Private Sub cboTipoMadeira_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTipoMadeira.SelectedIndexChanged
        tipoMadeira = cboTipoMadeira.SelectedIndex
    End Sub

    'ROTINA PARA APARECER O VALOR DE KE CONFORME A VINCULAÇÃO SELECIONADA PELO USUARIO 
    Private Sub cboLvinculado_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLvinculado.SelectedIndexChanged
        lblKe.Visible = True
    End Sub

    Private Sub txtEntradaRetangularBx_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaRetangularBx.TextChanged
        lblBase.Text = txtEntradaRetangularBx.Text
        lblBase.Visible = True
    End Sub

    Private Sub txtEntradaRetangularBy_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaRetangularBy.TextChanged
        lblAltura.Text = txtEntradaRetangularBy.Text
        lblAltura.Visible = True
    End Sub
    Private Sub txtEntradaRetangularL_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaRetangularL.TextChanged
        lblCompRet.Text = txtEntradaRetangularL.Text
        lblCompRet.Visible = True
    End Sub
    Private Sub txtEntradaCircularD_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCircularD.TextChanged
        lblCircD.Text = txtEntradaCircularD.Text
        lblCircD.Visible = True
    End Sub

    Private Sub txtEntradaCircularL_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCircularL.TextChanged
        lblCircL.Text = txtEntradaCircularL.Text
        lblCircL.Visible = True
    End Sub



    'ROTINA PRA APARECER OS DADOS INICIAIS DE ACORDO COM A SEÇÃO QUE O USUARIO ESCOLHE
    Private Sub rbtSecaoRetangular_ChecimkedChanged(sender As Object, e As EventArgs) Handles rbtSecaoRetangular.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaoretangular
        'seçãocir
        lblCircL.Visible = False
        lblCircD.Visible = False
        'seçãoT
        lblTbf.Visible = False
        lblTbw.Visible = False
        lblTh.Visible = False
        lblTd.Visible = False
        lblThf.Visible = False
        'seçãoI
        lblIbf2.Visible = False
        lblIbw.Visible = False
        lblIbf1.Visible = False
        lblIhf1.Visible = False
        lblIhf2.Visible = False
        lblIH.Visible = False
        lblId.Visible = False
        'secaoCaixao
        lblCaixaob1.Visible = False
        lblCaixaob2.Visible = False
        lblcaixaob3.Visible = False
        lblcaixaob4.Visible = False
        lblcaixaod.Visible = False
        lblCaixaoh1.Visible = False
        lblCaixaoh2.Visible = False
        lblCaixaoh4.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub
    Private Sub rbtSecaoCircular_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoCircular.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaocircular
        'seçãoret
        lblAltura.Visible = False
        lblBase.Visible = False
        lblCompRet.Visible = False
        'seçãoT
        lblTbf.Visible = False
        lblTbw.Visible = False
        lblTh.Visible = False
        lblTd.Visible = False
        lblThf.Visible = False
        'seçãoI
        lblIbf2.Visible = False
        lblIbw.Visible = False
        lblIbf1.Visible = False
        lblIhf1.Visible = False
        lblIhf2.Visible = False
        lblIH.Visible = False
        lblId.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub
    Private Sub rbtSecaoT_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoT.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaoT
        'seçãoret
        lblAltura.Visible = False
        lblBase.Visible = False
        lblCompRet.Visible = False

        'seçãocirc
        lblCircL.Visible = False
        lblCircD.Visible = False
        'seçãoI
        lblIbf2.Visible = False
        lblIbw.Visible = False
        lblIbf1.Visible = False
        lblIhf1.Visible = False
        lblIhf2.Visible = False
        lblIH.Visible = False
        lblId.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub
    Private Sub rbtSecaoI_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoI.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaoi
        'seçãoret
        lblAltura.Visible = False
        lblBase.Visible = False
        lblCompRet.Visible = False
        'seçãocirc
        lblCircL.Visible = False
        lblCircD.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
        'seçãoT
        lblTbf.Visible = False
        lblTbw.Visible = False
        lblTh.Visible = False
        lblTd.Visible = False
        lblThf.Visible = False
    End Sub
    Private Sub rbtSecaoCaixao_CheckedChanged(sender As Object, e As EventArgs) Handles rbtSecaoCaixao.CheckedChanged
        ExibirDadosEntrada()
        imgSecao.Image = My.Resources.secaocaixao
        'seçãoret
        lblAltura.Visible = False
        lblBase.Visible = False
        lblCompRet.Visible = False
        'seçãoT
        lblTbf.Visible = False
        lblTbw.Visible = False
        lblTh.Visible = False
        lblTd.Visible = False
        lblThf.Visible = False
        'seçãocirc
        lblCircL.Visible = False
        lblCircD.Visible = False
        'seçãoI
        lblIbf2.Visible = False
        lblIbw.Visible = False
        lblIbf1.Visible = False
        lblIhf1.Visible = False
        lblIhf2.Visible = False
        lblIH.Visible = False
        lblId.Visible = False
        gbxResultadosSecao.Visible = True
        gbxResultadosElementos.Visible = False
    End Sub
    Private Sub rbt2Elementos_CheckedChanged(sender As Object, e As EventArgs) Handles rbt2Elementos.CheckedChanged
        ExibirDadosEntrada()
        If (cboElementFixacao.Text = "Espacador Interposto") Then
            imgSecao.Image = My.Resources.secao2elementosEspacador
        ElseIf (cboElementFixacao.Text = "Chapas Laterais de Fixação") Then
            imgSecao.Image = My.Resources.secao2elementosChapa
        Else
        End If
        lblAltura.Visible = False
        lblBase.Visible = False
        lblCompRet.Visible = False
        lblCircL.Visible = False
        lblCircD.Visible = False
        gbxResultadosSecao.Visible = False
        gbxResultadosElementos.Visible = True
    End Sub
    Private Sub rbt3Elementos_CheckedChanged(sender As Object, e As EventArgs) Handles rbt3Elementos.CheckedChanged
        ExibirDadosEntrada()
        If (cboElementFixacao.Text = "Espacador Interposto") Then
            imgSecao.Image = My.Resources.secao3elementosEspacador
        ElseIf (cboElementFixacao.Text = "Chapas Laterais de Fixação") Then
            imgSecao.Image = My.Resources.secao3elementosChapa
        Else
        End If
        lblAltura.Visible = False
        lblBase.Visible = False
        lblCompRet.Visible = False
        lblCircL.Visible = False
        lblCircD.Visible = False
        gbxResultadosSecao.Visible = False
        gbxResultadosElementos.Visible = True
    End Sub

    'ROTINA PARA ATUALIZAR AS PROPRIEDADES GEOMÉTRICAS SEM DEPENDER DA MADEIRA SELECIONADA
    Private Sub btnCalcularPropriedadesGeometricas_Click(sender As Object, e As EventArgs) Handles btnCalcularPropriedadesGeometricas.Click
        'inicio - é o que precisa para calcular apenas as propriedades geometricas
        Dim tipoSecao As Madeira.TipoSecao = defineTipoSessao()

        madeiraSelecionada.PropriedadesGeometricas = definePropriedadesGeometricas()

        getDadosIniciais(tipoSecao)

        If rbt2Elementos.Checked Or rbt3Elementos.Checked Then
            ResultadoPropriedadesGeometricas23Elemento(madeiraSelecionada.PropriedadesGeometricas)
        Else
            ResultadoPropriedadesGeometricas(madeiraSelecionada.PropriedadesGeometricas)
        End If
        'fim
    End Sub

    'ROTINA PARA APARECER A JANELA DOS ESFORÇOS CARACTERISTICOS, QND ESBELTO
    Public Function Retorno()
        CalcularPropriedades()
    End Function

    'ROTINA PARA CALCULAR O DIMENSIONAMENTO (DEPENDE DA MADEIRA SELECIONADA):
    Private Sub btnCalcularPropriedades_Click(sender As Object, e As EventArgs) Handles btnCalcularPropriedades.Click
        CalcularPropriedades()
        isEsbelta()
    End Sub

    'ROTINA PARA APARECER OS ESFORÇOS CARACTERISTICOS PARA AS PEÇAS ESBELTAS
    Private Function isEsbelta()
        If esbelta Then
            MessageBox.Show("Aviso: o Elemento é esbelto. 
            Porfavor insira os esforços externos de cálculo da peça.")
            Dim frmEsfCarc As frmEsforcoCaracteristico = New frmEsforcoCaracteristico
            frmEsfCarc.Show()
        End If
    End Function

    Private Function CalcularPropriedades()
        'inicio - é o que precisa para calcular propriedades geometricas e dimensionar toda a peça
        Dim tipoSecao As Madeira.TipoSecao = defineTipoSessao()

        madeiraSelecionada.PropriedadesGeometricas = definePropriedadesGeometricas()

        getDadosIniciais(tipoSecao)

        If rbt2Elementos.Checked Or rbt3Elementos.Checked Then
            ResultadoPropriedadesGeometricas23Elemento(madeiraSelecionada.PropriedadesGeometricas)

        Else
            ResultadoPropriedadesGeometricas(madeiraSelecionada.PropriedadesGeometricas)
        End If
        'fim

        'TRAÇÃO: ACONTEECE QUANDO A NORMAL TRAÇÃO ESTÁ SELECIONADA: OK
        If txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Tração" Then
            madeiraSelecionada.Tracao = CalcTensaoT(
                PropriedadesResistencia.verificaVazio(txtNormal.Text),
                comprimento(),
                madeiraSelecionada.PropriedadesGeometricas,
                tipoSecao,
                b1ElementoJustaposto(),
                aElementoJustaposto(),
                cboElementFixacao.Text
                )

            'COMPRESSAO: ACONTEECE QUANDO A NORMAL TRAÇÃO ESTÁ SELECIONADA: OK
        ElseIf txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            MessageBox.Show("Peças sob esforços de compressão, estão sujeitas aos fenomênos de flambagem.
Por favor, insira a vinculação da peça.")
            'gbxVinculacao.Visible = True
            madeiraSelecionada.Compressao = CalcTensaoC(
                PropriedadesResistencia.verificaVazio(txtNormal.Text),
                PropriedadesResistencia.verificaVazio(txtMx.Text),
                PropriedadesResistencia.verificaVazio(txtMy.Text),
                vinculacao(comprimento),
                base(),
                altura(),
                madeiraSelecionada.PropriedadesGeometricas,
                tipoSecao,
                momentoCargaPermanenteX,
                momentoCargaPermanenteY,
                momentoCargaVariavelX,
                momentoCargaVariavelY,
                normalCargaPermanente,
                normalCargaVariavel,
                coefInfluencia(),
                f1,
                f2,
                b1ElementoJustaposto(),
                aElementoJustaposto(),
                cboElementFixacao.Text,
                l1entradaElementoJustaposto, 'TEM QUE CHAMAR O L1 
                a1ElementoJustaposto()
                )

        End If

        'CISALHAMENTO: ACONTEECE QUANDO Vx ou Vy ESTÁ SELECIONADA: OK 
        If txtCortanteX.Text.ToString <> "" Or txtCortanteY.Text.ToString <> "" Then
            madeiraSelecionada.Cisalhamento = CalcTensaoCisalhamento(
             PropriedadesResistencia.verificaVazio(txtCortanteX.Text),
             PropriedadesResistencia.verificaVazio(txtCortanteY.Text),
             PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text),
             cisalhamento_a1,
             cisalhamento_L1,
             madeiraSelecionada.PropriedadesGeometricas,
             tipoSecao
             )
        End If

        'FLEXÃO SIMPLES RETA: ACONTECE: mX OU mY SELECIONADOS JUNTO A ESFORÇOS DE CISALHAMENTO
        If (txtMx.Text.ToString <> "" And txtMy.Text.ToString = "" And (txtCortanteX.Text.ToString <> "" Or txtCortanteY.Text.ToString <> "")) Or
           (txtMy.Text.ToString <> "" And txtMx.Text.ToString = "" And (txtCortanteX.Text.ToString <> "" Or txtCortanteY.Text.ToString <> "")) Then
            gbxFlexaoSimples.Visible = True
            gbxFlexaoObliqua.Visible = False
            madeiraSelecionada.Flexao = CalcTensaoFlexaoSimples(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBx.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text),
            PropriedadesResistencia.verificaVazio(txtApoioX.Text),
            PropriedadesResistencia.verificaVazio(txtApoioY.Text),
           madeiraSelecionada.PropriedadesGeometricas,
           tipoSecao
           )
            'FLEXÃO OBLIQUA: ACONTECE QUANDO Mx e My ESTÃOO SELECIONADO: OK
        ElseIf txtMx.Text.ToString <> "" And txtMy.Text.ToString <> "" And (txtCortanteX.Text.ToString <> "" Or txtCortanteY.Text.ToString <> "") Then
            gbxFlexaoSimples.Visible = False
            gbxFlexaoObliqua.Visible = True
            madeiraSelecionada.Flexao = CalcTensaoFlexaoObliqua(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBx.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text),
            PropriedadesResistencia.verificaVazio(txtApoioX.Text),
            PropriedadesResistencia.verificaVazio(txtApoioY.Text),
            madeiraSelecionada.PropriedadesGeometricas,
            tipoSecao
            )
        End If

        'FLEXAO COMPOSTA
        '------------------------FLEXOTRAÇÃO------------------------
        If txtMx.Text.ToString <> "" Or txtMy.Text.ToString <> "" And txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Tração" Then
            gbxFlexoTracao.Visible = True
            gbxFlexoCompressao.Visible = False
            madeiraSelecionada.Flexao = CalcTensaoFlexotracao(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBx.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text),
            madeiraSelecionada.PropriedadesGeometricas,
            tipoSecao
            )

            '------------------------FLEXOCOMPRESSAO------------------------
        ElseIf txtMx.Text.ToString <> "" Or txtMy.Text.ToString <> "" And txtNormal.Text.ToString <> "" And cboTracaoCompressao.Text = "Compressão" Then
            gbxFlexoCompressao.Visible = True
            gbxFlexoTracao.Visible = False
            madeiraSelecionada.Flexao = CalcTensaoFlexocompressao(
            PropriedadesResistencia.verificaVazio(txtMx.Text),
            PropriedadesResistencia.verificaVazio(txtMy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBx.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaRetangularBy.Text),
            PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text),
            madeiraSelecionada.PropriedadesGeometricas,
            tipoSecao
            )
        End If

        tracao()
        compressao()
        cisalhamento()
        flexaoSimples(tipoSecao)
        flexaoComposta(tipoSecao)
    End Function

    'ROTINA PARA DEFINIR O TIPO DE SEÇÃO JA QUE OS CÁLCULOS MUDAM PARA CADA UMA :
    Private Function defineTipoSessao() As Madeira.TipoSecao
        If rbtSecaoRetangular.Checked Then
            Return Madeira.TipoSecao.Retangular

        ElseIf rbtSecaoCircular.Checked Then
            Return Madeira.TipoSecao.Circular

        ElseIf rbtSecaoT.Checked Then
            Return Madeira.TipoSecao.SecaoT

        ElseIf rbtSecaoI.Checked Then
            Return Madeira.TipoSecao.SecaoI

        ElseIf rbtSecaoCaixao.Checked Then
            Return Madeira.TipoSecao.Caixao

        ElseIf rbt2Elementos.Checked Then
            Return Madeira.TipoSecao.ElementosJustaposto2

        ElseIf rbt3Elementos.Checked Then
            Return Madeira.TipoSecao.ElementosJustaposto3
        End If
    End Function

    'ROTINA PARA PEGAR OS VALORES QUE O USUARIO INSERIU NOS TXT E CALCULAR AS PROPRIEDADES GEOMETRICAS:
    Private Function definePropriedadesGeometricas() As PropriedadesGeometricas
        If rbtSecaoRetangular.Checked Then
            Return proprGeometricas.CalculoRetangular(txtEntradaRetangularBx.Text, txtEntradaRetangularBy.Text)

        ElseIf rbtSecaoCircular.Checked Then
            Return proprGeometricas.CalculoCircular(txtEntradaCircularD.Text)

        ElseIf rbtSecaoT.Checked Then
            Return proprGeometricas.CalculoCompostoSecaoT(txtEntradaCompostaBf.Text, txtEntradaCompostaHf.Text, txtEntradaCompostaTh.Text, txtEntradaCompostaBw.Text)

        ElseIf rbtSecaoI.Checked Then
            Return proprGeometricas.CalculoCompostoSecaoI(txtEntradaIBf1.Text, txtEntradaIHf1.Text, txtEntradaIBf2.Text, txtEntradaIHf2.Text, txtEntradaIBw.Text, txtEntradaIH.Text)

        ElseIf rbtSecaoCaixao.Checked Then
            Return proprGeometricas.CalculoCompostoSecaoCaixao(txtEntradaCaixaoD.Text, txtEntradaCaixaoB1.Text, txtEntradaCaixaoB2.Text, txtEntradaCaixaoB3.Text, txtEntradaCaixaoB4.Text, txtEntradaCaixaoH1.Text, txtEntradaCaixaoH2.Text, txtEntradaCaixaoH3.Text, txtEntradaCaixaoH4.Text)

        ElseIf rbt2Elementos.Checked Then
            Return proprGeometricas.Calculo2Elementos(txtEntrada2L.Text, txtEntrada2L1.Text, txtEntrada2a1.Text, txtEntrada2h1.Text, txtEntrada2b1.Text, pilar())

        ElseIf rbt3Elementos.Checked Then
            Return proprGeometricas.Calculo3Elementos(txtEntrada3L.Text, txtEntrada3L1.Text, txtEntrada3a1.Text, txtEntrada3h1.Text, txtEntrada3b1.Text, pilar())
        End If
    End Function

    'VALIDAÇÕES APARÊNCIAS: 
    'TRAÇÃO ==============================
    Private Sub tracao()
        txtTensaoTracao.Text = madeiraSelecionada.Tracao.TensaoTracao
        txtEsbeltezTracaoX.Text = madeiraSelecionada.Tracao.EsbeltezTracaoX
        txtEsbeltezTracaoY.Text = madeiraSelecionada.Tracao.EsbeltezTracaoY
        txtFt.Text = madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela

        txtTensaoTracao.Text = Format(madeiraSelecionada.Tracao.TensaoTracao, "#0.0##;-#0.0##")
        txtEsbeltezTracaoX.Text = Format(madeiraSelecionada.Tracao.EsbeltezTracaoX, "#0.0##;-#0.0##")
        txtEsbeltezTracaoY.Text = Format(madeiraSelecionada.Tracao.EsbeltezTracaoY, "#0.0##;-#0.0##")
        txtFt.Text = Format(madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, "#0.0##;-#0.0##")

        'VALIDAÇÃO TENSÃO TRAÇÃO:
        If PropriedadesResistencia.validarTensaoTracao(madeiraSelecionada.Tracao.TensaoTracao, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela) Then
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
        If PropriedadesResistencia.validarEsbeltezTracaoX(madeiraSelecionada.Tracao.EsbeltezTracaoX) Then
            lblValidacaoEsbeltezXTracao.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezXTracao.Visible = True
            lblValidacaoEsbeltezXTracao.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltezXTracao.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezXTracao.Visible = True
            lblValidacaoEsbeltezXTracao.ForeColor = Color.Red
        End If

        'VALIDAÇÃO ESBELTEZ Y:
        If PropriedadesResistencia.validarEsbeltezTracaoY(madeiraSelecionada.Tracao.EsbeltezTracaoY) Then
            lblValidacaoEsbeltezYTracao.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezYTracao.Visible = True
            lblValidacaoEsbeltezYTracao.ForeColor = Color.Green
        Else
            lblValidacaoEsbeltezYTracao.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoEsbeltezYTracao.Visible = True
            lblValidacaoEsbeltezYTracao.ForeColor = Color.Red
        End If
    End Sub

    'COMPRESSÃO ==============================
    Private Sub compressao()
        txtEsbeltezCompressaoX.Text = madeiraSelecionada.Compressao.EsbeltezCompressaoX
        txtEsbeltezCompressaoY.Text = madeiraSelecionada.Compressao.EsbeltezCompressaoY

        txtTensaoCompressaoX.Text = madeiraSelecionada.Compressao.TensaoCompressaoX
        txtTensaoCompressaoY.Text = madeiraSelecionada.Compressao.TensaoCompressaoY

        txtEsforçoFlexaoX.Text = madeiraSelecionada.Compressao.TensaoMxd
        txtEsforçoFlexaoY.Text = madeiraSelecionada.Compressao.TensaoMyd

        txtForçaElasticaX.Text = madeiraSelecionada.Compressao.ForcaElasticaX
        txtForçaElasticaY.Text = madeiraSelecionada.Compressao.ForcaElasticaY

        txtEaX.Text = madeiraSelecionada.Compressao.EaX
        txtEaY.Text = madeiraSelecionada.Compressao.EaY

        txtEdX.Text = madeiraSelecionada.Compressao.Edx
        txtEdY.Text = madeiraSelecionada.Compressao.Edy

        txtEiX.Text = madeiraSelecionada.Compressao.EiX
        txtEiY.Text = madeiraSelecionada.Compressao.EiY

        txtE1efX.Text = madeiraSelecionada.Compressao.E1ef
        txtE1efY.Text = madeiraSelecionada.Compressao.E1ef

        txtEcX.Text = coefInfluencia()
        txtEcY.Text = coefInfluencia()

        txtEsbeltezCompressaoX.Text = Format(madeiraSelecionada.Compressao.EsbeltezCompressaoX, "#0.0##;-#0.0##")
        txtEsbeltezCompressaoY.Text = Format(madeiraSelecionada.Compressao.EsbeltezCompressaoY, "#0.0##;-#0.0##")

        txtTensaoCompressaoX.Text = Format(madeiraSelecionada.Compressao.TensaoCompressaoX, "#0.0##;-#0.0##")
        txtTensaoCompressaoY.Text = Format(madeiraSelecionada.Compressao.TensaoCompressaoY, "#0.0##;-#0.0##")

        txtEsforçoFlexaoX.Text = Format(madeiraSelecionada.Compressao.TensaoMxd, "#0.0##;-#0.0##")
        txtEsforçoFlexaoY.Text = Format(madeiraSelecionada.Compressao.TensaoMyd, "#0.0##;-#0.0##")

        txtForçaElasticaX.Text = Format(madeiraSelecionada.Compressao.ForcaElasticaX, "#0.0##;-#0.0##")
        txtForçaElasticaY.Text = Format(madeiraSelecionada.Compressao.ForcaElasticaY, "#0.0##;-#0.0##")

        txtEaX.Text = Format(madeiraSelecionada.Compressao.EaX, "#0.0##;-#0.0##")
        txtEaY.Text = Format(madeiraSelecionada.Compressao.EaY, "#0.0##;-#0.0##")

        txtEdX.Text = Format(madeiraSelecionada.Compressao.Edx, "#0.0##;-#0.0##")
        txtEdY.Text = Format(madeiraSelecionada.Compressao.Edy, "#0.0##;-#0.0##")

        txtEiX.Text = Format(madeiraSelecionada.Compressao.EiX, "#0.0##;-#0.0##")
        txtEiY.Text = Format(madeiraSelecionada.Compressao.EiY, "#0.0##;-#0.0##")

        txtE1efX.Text = Format(madeiraSelecionada.Compressao.E1ef, "#0.0##;-#0.0##")
        txtE1efY.Text = Format(madeiraSelecionada.Compressao.E1ef, "#0.0##;-#0.0##")

        If (madeiraSelecionada.Compressao.EsbeltezCompressaoX <= 40) Then
            lblClassifEsbeltezX.Text = "PEÇA CURTA"
        ElseIf (40 < madeiraSelecionada.Compressao.EsbeltezCompressaoX And madeiraSelecionada.Compressao.EsbeltezCompressaoX <= 80) Then
            lblClassifEsbeltezX.Text = "PEÇA MEDIANAMENTE ESBELTA"
        ElseIf (80 < madeiraSelecionada.Compressao.EsbeltezCompressaoY And madeiraSelecionada.Compressao.EsbeltezCompressaoY <= 140) Then
            lblClassifEsbeltezX.Text = "PEÇA ESBELTA"
        End If

        If (madeiraSelecionada.Compressao.EsbeltezCompressaoY <= 40) Then
            lblClassifEsbeltezY.Text = "PEÇA CURTA"
        ElseIf (40 < madeiraSelecionada.Compressao.EsbeltezCompressaoY And madeiraSelecionada.Compressao.EsbeltezCompressaoY <= 80) Then
            lblClassifEsbeltezY.Text = "PEÇA MEDIANAMENTE ESBELTA"
        ElseIf (80 < madeiraSelecionada.Compressao.EsbeltezCompressaoY And madeiraSelecionada.Compressao.EsbeltezCompressaoY <= 140) Then
            lblClassifEsbeltezY.Text = "PEÇA ESBELTA"
        End If

        'VALIDAÇÃO COMPRESSÃO CURTA X:
        If PropriedadesResistencia.validarTensaoCompressaoCurtaX(madeiraSelecionada.Compressao.TensaoCompressaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblValidacaoXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoXComp.Visible = True
            lblValidacaoXComp.ForeColor = Color.Green
        Else
            lblValidacaoXComp.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoXComp.Visible = True
            lblValidacaoXComp.ForeColor = Color.Red
        End If

        'VALIDAÇÃO COMPRESSÃO CURTA Y
        If PropriedadesResistencia.validarTensaoCompressaoCurtaY(madeiraSelecionada.Compressao.TensaoCompressaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblValidacaoYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoYComp.Visible = True
            lblValidacaoYComp.ForeColor = Color.Green
        Else
            lblValidacaoYComp.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoYComp.Visible = True
            lblValidacaoYComp.ForeColor = Color.Red
        End If

        'VALIDAÇÃO COMPRESSAO MEDIANAMENTE ESBELTA E ESBELTA X
        If PropriedadesResistencia.validarTensaoCompressaoMedEsbX(madeiraSelecionada.Compressao.TensaoCompressaoX, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblValidacaoXComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoXComp.Visible = True
            lblValidacaoXComp.ForeColor = Color.Green
        Else
            lblValidacaoXComp.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoXComp.Visible = True
            lblValidacaoXComp.ForeColor = Color.Red
        End If

        'VALIDAÇÃO COMPRESSAO MEDIANAMENTE ESBELTA E ESBELTA Y
        If PropriedadesResistencia.validarTensaoCompressaoMedEsbY(madeiraSelecionada.Compressao.TensaoCompressaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblValidacaoYComp.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoYComp.Visible = True
            lblValidacaoYComp.ForeColor = Color.Green
        Else
            lblValidacaoYComp.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoYComp.Visible = True
            lblValidacaoYComp.ForeColor = Color.Red
        End If

        'VALIDAÇÃO SEÇÃO COSTA PEÇAS SOLIDARIZADAS DESCONTÍNUAMENTE 



    End Sub


    'CISALHAMENTO ====================================
    Private Sub cisalhamento()
        txtTensaoCisalhanteX.Text = madeiraSelecionada.Cisalhamento.TensaoCisalhanteX
        txtTensaoCisalhanteY.Text = madeiraSelecionada.Cisalhamento.TensaoCisalhanteY
        txtfv.Text = madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento

        txtTensaoCisalhanteX.Text = Format(madeiraSelecionada.Cisalhamento.TensaoCisalhanteX, "#0.0##;-#0.0##")
        txtTensaoCisalhanteY.Text = Format(madeiraSelecionada.Cisalhamento.TensaoCisalhanteY, "#0.0##;-#0.0##")
        txtfv.Text = Format(madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento, "#0.0##;-#0.0##")

        'VALIDAÇÃO TENSÃO CISALHANTE X:
        If PropriedadesResistencia.validarTensaoCisalhanteX(madeiraSelecionada.Cisalhamento.TensaoCisalhanteX, madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento) Then
            lblValidacaoCisalhanteX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteX.Visible = True
            lblValidacaoCisalhanteX.ForeColor = Color.Green
        Else
            lblValidacaoCisalhanteX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteX.Visible = True
            lblValidacaoCisalhanteX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO TENSÃO CISALHANTE Y:
        If PropriedadesResistencia.validarTensaoCisalhanteY(madeiraSelecionada.Cisalhamento.TensaoCisalhanteY, madeiraSelecionada.ResistenciaCalculo.resistCalculoAoCisalhamento) Then
            lblValidacaoCisalhanteY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteY.Visible = True
            lblValidacaoCisalhanteY.ForeColor = Color.Green
        Else
            lblValidacaoCisalhanteY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoCisalhanteY.Visible = True
            lblValidacaoCisalhanteY.ForeColor = Color.Red
        End If

    End Sub

    'FLEXÃO SIMPLES ====================================
    Private Sub flexaoSimples(tipoSecao As Madeira.TipoSecao)

        txtftFlexo.Text = madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela
        txtfc.Text = madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela

        'apoio
        txtCompApoioX.Text = madeiraSelecionada.Flexao.ApoioX
        txtCompApoioX.Text = madeiraSelecionada.Flexao.ApoioY
        'reta
        txtTensaoTx.Text = madeiraSelecionada.Flexao.tensaoTY
        txtTensaoCy.Text = madeiraSelecionada.Flexao.tensaoCY
        'obliqua
        txtTensaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtTensaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY
        'apoio
        txtApoioX.Text = Format(madeiraSelecionada.Flexao.ApoioX, "#0.0##;-#0.0##")
        txtApoioY.Text = Format(madeiraSelecionada.Flexao.ApoioY, "#0.0##;-#0.0##")
        'reta
        txtTensaoTx.Text = Format(madeiraSelecionada.Flexao.tensaoTY, "#0.0##;-#0.0##")
        txtTensaoCy.Text = Format(madeiraSelecionada.Flexao.tensaoCY, "#0.0##;-#0.0##")
        'obliqua
        txtTensaoMx.Text = Format(madeiraSelecionada.Flexao.tensaoFlexaoX, "#0.0##;-#0.0##")
        txtTensaoMy.Text = Format(madeiraSelecionada.Flexao.tensaoFlexaoY, "#0.0##;-#0.0##")

        'VALIDAÇÃO APOIO X
        If PropriedadesResistencia.validarApoioX(madeiraSelecionada.Flexao.ApoioX, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblApoioX.Text = "PASSOU NA VERIFICAÇÃO"
            lblApoioX.Visible = True
            lblApoioX.ForeColor = Color.Green
        Else
            lblApoioX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblApoioX.Visible = True
            lblApoioX.ForeColor = Color.Red
        End If
        'VALIDAÇÃO APOIO y
        If PropriedadesResistencia.validarApoioY(madeiraSelecionada.Flexao.ApoioY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela) Then
            lblApoioY.Text = "PASSOU NA VERIFICAÇÃO"
            lblApoioY.Visible = True
            lblApoioY.ForeColor = Color.Green
        Else
            lblApoioY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblApoioY.Visible = True
            lblApoioY.ForeColor = Color.Red
        End If

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

    'FLEXÃO COMPOSTA ================================
    Private Sub flexaoComposta(tipoSecao As Madeira.TipoSecao)
        txtFlexoTracao.Text = madeiraSelecionada.Tracao.TensaoTracao
        txtFlexoTracaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtFlexoTracaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY
        txtFlexoCompressao.Text = madeiraSelecionada.Compressao.TensaoCompressaoX 'OU Y POIS SÃO IGUAIS
        txtFlexoCompressaoMx.Text = madeiraSelecionada.Flexao.tensaoFlexaoX
        txtFlexoCompressaoMy.Text = madeiraSelecionada.Flexao.tensaoFlexaoY
        txtftFlexo.Text = madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela
        txtfc.Text = madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela

        txtFlexoTracao.Text = Format(madeiraSelecionada.Tracao.TensaoTracao, "#0.0##;-#0.0##")
        txtFlexoTracaoMx.Text = Format(madeiraSelecionada.Flexao.tensaoFlexaoX, "#0.0##;-#0.0##")
        txtFlexoTracaoMy.Text = Format(madeiraSelecionada.Flexao.tensaoFlexaoY, "#0.0##;-#0.0##")
        txtFlexoCompressao.Text = Format(madeiraSelecionada.Compressao.TensaoCompressaoX, "#0.0##;-#0.0##") 'OU Y POIS SÃO IGUAIS
        txtFlexoCompressaoMx.Text = Format(madeiraSelecionada.Flexao.tensaoFlexaoX, "#0.0##;-#0.0##")
        txtFlexoCompressaoMy.Text = Format(madeiraSelecionada.Flexao.tensaoFlexaoY, "#0.0##;-#0.0##")




        'VALIDAÇÃO FLEXOTRAÇÃO X:
        If PropriedadesResistencia.validarTensaoFlexotracaoX(madeiraSelecionada.Tracao.TensaoTracao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoTracaoX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoX.Visible = True
            lblValidacaoFlexoTracaoX.ForeColor = Color.Green
        Else
            lblValidacaoFlexoTracaoX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoX.Visible = True
            lblValidacaoFlexoTracaoX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOTRAÇÃO Y:
        If PropriedadesResistencia.validarTensaoFlexotracaoY(madeiraSelecionada.Tracao.TensaoTracao, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoTracaoParalela, tipoSecao) Then
            lblValidacaoFlexoTracaoY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoY.Visible = True
            lblValidacaoFlexoTracaoY.ForeColor = Color.Green
        Else
            lblValidacaoFlexoTracaoY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoTracaoY.Visible = True
            lblValidacaoFlexoTracaoY.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOCOMPRESSAO X:
        If PropriedadesResistencia.validarTensaoFlexocompressaoX(madeiraSelecionada.Compressao.TensaoCompressaoX, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, tipoSecao) Then
            lblValidacaoFlexoCompressaoX.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoX.Visible = True
            lblValidacaoFlexoCompressaoX.ForeColor = Color.Green
        Else
            lblValidacaoFlexoCompressaoX.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoX.Visible = True
            lblValidacaoFlexoCompressaoX.ForeColor = Color.Red
        End If

        'VALIDAÇÃO FLEXOCOMPRESSAO Y:
        If PropriedadesResistencia.validarTensaoFlexocompressaoY(madeiraSelecionada.Compressao.TensaoCompressaoY, madeiraSelecionada.Flexao.tensaoFlexaoX, madeiraSelecionada.Flexao.tensaoFlexaoY, madeiraSelecionada.ResistenciaCalculo.resistCalculoCompressaoParalela, tipoSecao) Then
            lblValidacaoFlexoCompressaoY.Text = "PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoY.Visible = True
            lblValidacaoFlexoCompressaoY.ForeColor = Color.Green
        Else
            lblValidacaoFlexoCompressaoY.Text = "NÃO PASSOU NA VERIFICAÇÃO"
            lblValidacaoFlexoCompressaoY.Visible = True
            lblValidacaoFlexoCompressaoY.ForeColor = Color.Red
        End If
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
        txtMElasticidade.Text = Format(madeiraSelecionada.ResistenciaCalculo.moduloElasticidade, "#0.0#;-#0.0#")
        txtMeTransversal.Text = Format(madeiraSelecionada.ResistenciaCalculo.moduloElasticidadeTransversal, "#0.000##;-#0.000##")
        txtDAparente.Text = Format(madeiraSelecionada.ResistenciaCalculo.densidadeAparente, "#0.0##;-#0.0##")
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

        'FORMATANDO A QUANTIDADE DE CASAS DPS DA VIRGULA NOS TXT
        txtAreaSecao.Text = Format(resultProprGeometricas.Area, "#0.0##;-#0.0##")
        txtEixoXMomentoInercia.Text = Format(resultProprGeometricas.EixoXmi, "#0.0##;-#0.0##")
        txtEixoXRaioGiracao.Text = Format(resultProprGeometricas.EixoXrg, "#0.0##;-#0.0##")
        txtEixoXModuloResistencia.Text = Format(resultProprGeometricas.EixoXmr, "#0.0##;-#0.0##")
        txtEixoXInerciaEfetiva.Text = Format(resultProprGeometricas.EixoXie, "#0.0##;-#0.0##")
        txtEixoYMomentoInercia.Text = Format(resultProprGeometricas.EixoYmi, "#0.0##;-#0.0##")
        txtEixoYRaioGiracao.Text = Format(resultProprGeometricas.EixoYrg, "#0.0##;-#0.0##")
        txtEixoYModuloResistencia.Text = Format(resultProprGeometricas.EixoYmr, "#0.0##;-#0.0##")
        txtEixoYInerciaEfetiva.Text = Format(resultProprGeometricas.EixoYie, "#0.0##;-#0.0##")
    End Sub

    'DADOS INICIAIS QUE SÃO FECHADOS PRO USUARIO NÃO DIGITAR NADA 
    'SEÇÃO T----------------------------------------
    Private Sub txtEntradaCompostaTh_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaTh.TextChanged
        txtEntradaCompostaH.Text = PropriedadesResistencia.somaT(PropriedadesResistencia.verificaVazio(txtEntradaCompostaHf.Text), PropriedadesResistencia.verificaVazio(txtEntradaCompostaTh.Text))
        lblTh.Text = txtEntradaCompostaTh.Text
        lblTh.Visible = True
    End Sub
    Private Sub txtEntradaCompostaHf_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaHf.TextChanged
        txtEntradaCompostaH.Text = PropriedadesResistencia.somaT(PropriedadesResistencia.verificaVazio(txtEntradaCompostaHf.Text), PropriedadesResistencia.verificaVazio(txtEntradaCompostaTh.Text))
        lblThf.Text = txtEntradaCompostaHf.Text
        lblThf.Visible = True
    End Sub
    Private Sub txtEntradaCompostaH_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaH.TextChanged
        lblTd.Text = txtEntradaCompostaH.Text
        lblTd.Visible = True
    End Sub

    Private Sub txtEntradaCompostaBw_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaBw.TextChanged
        lblTbw.Text = txtEntradaCompostaBw.Text
        lblTbw.Visible = True
    End Sub

    Private Sub txtEntradaCompostaBf_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCompostaBf.TextChanged
        lblTbf.Text = txtEntradaCompostaBf.Text
        lblTbf.Visible = True
    End Sub


    'SEÇÃO I----------------------------------------
    Private Sub txtEntradaIHf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIHf1.TextChanged
        txtEntradaIHf2.Text = txtEntradaIHf1.Text
        txtEntradaID.Text = PropriedadesResistencia.somaI(PropriedadesResistencia.verificaVazio(txtEntradaIHf1.Text), PropriedadesResistencia.verificaVazio(txtEntradaIH.Text), PropriedadesResistencia.verificaVazio(txtEntradaIHf2.Text))
        lblIhf1.Text = txtEntradaIHf1.Text
        lblIhf1.Visible = True
    End Sub
    Private Sub txtEntradaIH_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIH.TextChanged
        txtEntradaID.Text = PropriedadesResistencia.somaI(PropriedadesResistencia.verificaVazio(txtEntradaIHf1.Text), PropriedadesResistencia.verificaVazio(txtEntradaIH.Text), PropriedadesResistencia.verificaVazio(txtEntradaIHf2.Text))
        lblIH.Text = txtEntradaIH.Text
        lblIH.Visible = True
    End Sub
    Private Sub txtEntradaIHf2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIHf2.TextChanged
        txtEntradaID.Text = PropriedadesResistencia.somaI(PropriedadesResistencia.verificaVazio(txtEntradaIHf1.Text), PropriedadesResistencia.verificaVazio(txtEntradaIH.Text), PropriedadesResistencia.verificaVazio(txtEntradaIHf2.Text))
        lblIhf2.Text = txtEntradaIHf2.Text
        lblIhf2.Visible = True
    End Sub
    Private Sub txtEntradaIBf1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIBf1.TextChanged
        txtEntradaIBf2.Text = txtEntradaIBf1.Text
        lblIbf1.Text = txtEntradaIBf1.Text
        lblIbf1.Visible = True
    End Sub
    Private Sub txtEntradaID_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaID.TextChanged
        lblId.Text = txtEntradaID.Text
        lblId.Visible = True
    End Sub

    Private Sub txtEntradaIBw_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIBw.TextChanged
        lblIbw.Text = txtEntradaIBw.Text
        lblIbw.Visible = True
    End Sub

    Private Sub txtEntradaIBf2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaIBf2.TextChanged
        lblIbf2.Text = txtEntradaIBf2.Text
        lblIbf2.Visible = True
    End Sub

    'SEÇÃO CAIXÃO----------------------------------------
    'ALTURAS:
    Private Sub txtEntradaCaixaoH1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH1.TextChanged
        txtEntradaCaixaoH4.Text = txtEntradaCaixaoH1.Text
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixaoAltura(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
        lblCaixaoh1.Text = txtEntradaCaixaoH1.Text
        lblCaixaoh1.Visible = True
    End Sub
    Private Sub txtEntradaCaixaoH2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH2.TextChanged
        txtEntradaCaixaoH3.Text = txtEntradaCaixaoH2.Text
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixaoAltura(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
        lblCaixaoh2.Text = txtEntradaCaixaoH2.Text
        lblCaixaoh2.Visible = True
    End Sub
    Private Sub txtEntradaCaixaoH4_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoH4.TextChanged
        txtEntradaCaixaoD.Text = PropriedadesResistencia.somaCaixaoAltura(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoH4.Text))
        lblCaixaoh4.Text = txtEntradaCaixaoH4.Text
        lblCaixaoh4.Visible = True
    End Sub
    'LARGURAS:
    Private Sub txtEntradaCaixaoB1_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB1.TextChanged
        txtCaixaoComp.Text = PropriedadesResistencia.somaCaixaoComprimento(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB3.Text))
        txtEntradaCaixaoB4.Text = txtEntradaCaixaoB1.Text
        lblCaixaob1.Text = txtEntradaCaixaoB1.Text
        lblCaixaob1.Visible = True
    End Sub
    Private Sub txtEntradaCaixaoB2_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB2.TextChanged
        txtCaixaoComp.Text = PropriedadesResistencia.somaCaixaoComprimento(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB3.Text))
        txtEntradaCaixaoB3.Text = txtEntradaCaixaoB2.Text
        lblCaixaob2.Text = txtEntradaCaixaoB2.Text
        lblCaixaob2.Visible = True
    End Sub
    Private Sub txtEntradaCaixaoB3_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB3.TextChanged
        txtCaixaoComp.Text = PropriedadesResistencia.somaCaixaoComprimento(PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB1.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB2.Text), PropriedadesResistencia.verificaVazio(txtEntradaCaixaoB3.Text))
        lblcaixaob3.Text = txtEntradaCaixaoB3.Text
        lblcaixaob3.Visible = True
    End Sub
    Private Sub txtEntradaCaixaoD_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoD.TextChanged
        lblcaixaod.Text = txtEntradaCaixaoD.Text
        lblcaixaod.Visible = True
    End Sub
    Private Sub txtEntradaCaixaoB4_TextChanged(sender As Object, e As EventArgs) Handles txtEntradaCaixaoB4.TextChanged
        lblcaixaob4.Text = txtEntradaCaixaoB4.Text
        lblcaixaob4.Visible = True
    End Sub

    'ROTINA QUE TA BLOQUEANDO UM TAB (TRAÇÃO OU COMPRESSÃO) CASO ELE NAO SEJA SELECIONADO
    Private Sub cboTracaoCompressao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTracaoCompressao.SelectedIndexChanged
        TabTracao.Enabled = cboTracaoCompressao.SelectedIndex.CompareTo(1)
        TabCompressao.Enabled = cboTracaoCompressao.SelectedIndex.CompareTo(0)
    End Sub

    'ROTINA PARA APARECER OS GBX RESPECTIVOS DOS DADOS DE ENTRADA SELECIONADOS
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


    'ROTINA PARA EXIBIR OS RESULTADOS DAS SEÇÕES SIMPLES E COMPOSTAS POR SOLIDARIZAÇÃO DESCONTÍNUA
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

        'FORMATANDO A QUANTIDADE DE CASAS DPS DA VIRGULA NOS TXT
        txtArea1.Text = Format(resultProprGeometricas.AreaA1, "#0.0##;-#0.0##")
        txtArea.Text = Format(resultProprGeometricas.Area, "#0.0##;-#0.0##")
        txtInterpoIx1.Text = Format(resultProprGeometricas.EixoXmi1, "#0.0##;-#0.0##")
        txtInterpoIx.Text = Format(resultProprGeometricas.EixoXmi, "#0.0##;-#0.0##")
        txtInterpoIy2.Text = Format(resultProprGeometricas.EixoYmi1, "#0.0##;-#0.0##")
        txtInterpoIy.Text = Format(resultProprGeometricas.EixoYmi, "#0.0##;-#0.0##")
        txtInterpoEixoYCoefB.Text = Format(resultProprGeometricas.CoefBy, "#0.0##;-#0.0##")
        txtInterpoIModuloResist.Text = Format(resultProprGeometricas.EixoYmr, "#0.0##;-#0.0##")
        txtInterpoEixoYInerciaEfetiva.Text = Format(resultProprGeometricas.EixoYie, "#0.0##;-#0.0##")
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

    'JANELA GEOMETRICA DE AJUDA ===================================================================================
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
    'vinculação compressão
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        form4Vinculacao.Show()
        PictureBox4.BackgroundImage = My.Resources.help
    End Sub
    'ROTINA PARA APARECER MANUAL DO SOFTWARES
    Private Sub AjudaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AjudaToolStripMenuItem1.Click
        formAjuda.Show()
    End Sub
    'ROTINA PARA FAZER A ABA PROSSEGUIR DEPOIS QUE APERTAR O NOVO PROJETO
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
    'DEFINIÇÃO DA VINCULAÇÃO PARA ESFORÇOS DE COMPRESSÃO =========================
    Private Function vinculacao(l As Double) As Double
        Dim lvinculado As Double = 0

        Select Case cboLvinculado.SelectedIndex
            Case 0
                lblKe.Text = 0.65.ToString
                lvinculado = 0.65 * l
            Case 1
                lblKe.Text = 0.8.ToString
                lvinculado = 0.8 * l
            Case 2
                lblKe.Text = 1.2.ToString
                lvinculado = 1.2 * l
            Case 3
                lblKe.Text = 1.ToString
                lvinculado = 1 * l
            Case 4
                lblKe.Text = 2.1.ToString
                lvinculado = 2.1 * l
            Case 5
                lblKe.Text = 2.4.ToString
                lvinculado = 2.4 * l
        End Select
        Return lvinculado
    End Function

    'DEFINIÇÃO DO COEFICIENTe DEVIDO A FLUÊNCIA DO MATERIAL PARA ESFORÇOS DE COMPRESSÃO QUANDO A PEÇA É ESBELTA =================
    Private Function coefInfluencia() As Double
        Dim coeficiente As Double = 0

        ' If (madeiraSelecionada.Compressao.EsbeltezCompressaoX <= 40) Then
        '     coeficiente = 0
        ' ElseIf (madeiraSelecionada.Compressao.EsbeltezCompressaoY <= 40) Then
        'coeficiente = 0
        ' ElseIf (40 < madeiraSelecionada.Compressao.EsbeltezCompressaoX And madeiraSelecionada.Compressao.EsbeltezCompressaoX <= 80) And
        'Coeficiente = 0 Then
        ' ElseIf (40 < madeiraSelecionada.Compressao.EsbeltezCompressaoY And madeiraSelecionada.Compressao.EsbeltezCompressaoY <= 80) Then
        'coeficiente = 0
        'classe de umidade 1 e 2 e classe de carregamento permanente e longa duração:
        'ElseIf (cboKmod1.Text = "Permanente" Or cboKmod1.Text = "Longa Duração") And (cboKmod2.Text = "Classe (1)" Or cboKmod2.Text = "Classe (2)") Then
        If (cboKmod1.Text = "Permanente" Or cboKmod1.Text = "Longa Duração") And (cboKmod2.Text = "Classe (1)" Or cboKmod2.Text = "Classe (2)") Then
            coeficiente = 0.8
            'classe de umidade 3 e 4 e classe de carregamento permanente e longa duração:
        ElseIf (cboKmod1.Text = "Permanente" Or cboKmod1.Text = "Longa Duração") And (cboKmod2.Text = "Classe (3)" Or cboKmod2.Text = "Classe (4)") Then
            coeficiente = 2

            'classe de umidade 1 e 2 e classe de carregamento média duração
        ElseIf cboKmod1.Text = "Média Duração" And (cboKmod2.Text = "Classe (1)" Or cboKmod2.Text = "Classe (2)") Then
            coeficiente = 0.3

            'classe de umidade 2 e 3 e classe de carregamento média duração
        ElseIf cboKmod1.Text = "Média Duração" And (cboKmod2.Text = "Classe (3)" Or cboKmod2.Text = "Classe (4)") Then
            coeficiente = 1

            'classe de umidade 1 e 2 e classe de carregamento Curta Duração
        ElseIf cboKmod1.Text = "Curta Duração" And (cboKmod2.Text = "Classe (1)" Or cboKmod2.Text = "Classe (2)") Then
            coeficiente = 0.1

            'classe de umidade 2 e 3 e classe de carregamento Curta Duração
        ElseIf cboKmod1.Text = "Curta Duração" And (cboKmod2.Text = "Classe (3)" Or cboKmod2.Text = "Classe (4)") Then
            coeficiente = 0.5

            'classe de umidade 1 e 2 e classe de carregamento Instantânea
        ElseIf cboKmod1.Text = "Instantânea" Then
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
    'rotina para chamar o comprimento de cada seção
    'nas verificações
    Public Function comprimento() As Double
        If rbtSecaoRetangular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaRetangularL.Text)

        ElseIf rbtSecaoCircular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaCircularL.Text)

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

    'rotina para chamar a base de cada seção nas verificações à compressão
    Public Function base() As Double

        If rbtSecaoRetangular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaRetangularBx.Text)
        ElseIf rbtSecaoCircular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text)

        ElseIf rbtSecaoT.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbtSecaoI.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbtSecaoCaixao.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbt2Elementos.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbt3Elementos.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)
        Else
            'Return 0
        End If
    End Function

    'rotina para chamar a altura de cada seção nas verificações à compressão
    Public Function altura() As Double

        If rbtSecaoRetangular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaRetangularBy.Text)

        ElseIf rbtSecaoCircular.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntradaCircularD.Text)

        ElseIf rbtSecaoT.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbtSecaoI.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbtSecaoCaixao.Checked Then
            'Return PropriedadesResistencia.verificaVazio(txt.Text)

        ElseIf rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2h1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3h1.Text)
        Else
            Return 0
        End If
    End Function

    'rotina para chamar o valor de a nos esforços de compressão
    Public Function aElementoJustaposto() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(2 * txtEntrada2a1.Text) - PropriedadesResistencia.verificaVazio(txtEntrada2b1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3a1.Text) - PropriedadesResistencia.verificaVazio(txtEntrada3b1.Text)
        Else
            Return 0
        End If
    End Function
    'rotina para chamar o valor de b1 nos esforços de compressão
    Public Function b1ElementoJustaposto() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2b1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3b1.Text)
        Else
            Return 0
        End If
    End Function

    'rotina para chamar o valor de a1 nos esforços de compressão
    Public Function a1ElementoJustaposto() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2a1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3a1.Text)
        Else
            Return 0
        End If
    End Function

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If voltou Then
            voltou = False
            Retorno()
        End If
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        MessageBox.Show("Peças submetidas à flexão, devem ter a região dos apoios verificadas devido ao fato de
                        estarem sujeitas a aos esforços de compressão normal às fibras nas regiões dos apoios.")

    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        Dim gerarPdf As GeradorPDF = New GeradorPDF(madeiraSelecionada)
        gerarPdf.GerarRelatorio()
    End Sub

    Private Sub txtEntrada3L_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3L.TextChanged

    End Sub

    Private Sub txtEntrada3L1_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3L1.TextChanged

    End Sub

    Private Sub txtEntrada3a1_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3a1.TextChanged

    End Sub

    Private Sub txtEntrada3h1_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3h1.TextChanged

    End Sub

    Private Sub txtEntrada3b1_TextChanged(sender As Object, e As EventArgs) Handles txtEntrada3b1.TextChanged

    End Sub

    Private Sub SalvarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalvarToolStripMenuItem.Click
        'Projeto.Update()
        ' CBackupFile.SalvarArquivoBin(Projeto, 1)
    End Sub

    'rotina para chamar o valor de L1 nos esforços de compressão
    Public Function l1entradaElementoJustaposto() As Double

        If rbt2Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada2L1.Text)

        ElseIf rbt3Elementos.Checked Then
            Return PropriedadesResistencia.verificaVazio(txtEntrada3L1.Text)
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
