Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports TCC2

Public Class GeradorPDF

    Private _madeira As Madeira
    Private _exibir As Exibicao

    Public Sub New(mMadeira As Madeira, mExibir As Exibicao)
        _madeira = mMadeira
        _exibir = mExibir
    End Sub

    Public Sub GerarRelatorio()
        Dim doc As New Document(iTextSharp.text.PageSize.A4, 20, 20, 20, 20)
        Dim arquivo As FileStream = New FileStream("C:\Users\Public\Documents\RelatTCC.pdf", FileMode.Create) 'Colocar destino do arquivo a ser gerado
        Dim writer As PdfWriter = PdfWriter.GetInstance(doc, arquivo)

        doc.Open()

        Dim fontTitle As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 16, Font.BOLD, BaseColor.BLACK)
        Dim fontSubTitle As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 14, Font.BOLD, BaseColor.BLACK)
        Dim fontSubTitle2 As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 14, Font.NORMAL, BaseColor.BLACK)
        Dim fontForm As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, Font.NORMAL, BaseColor.BLACK)
        Dim fontItem As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, Font.BOLD, BaseColor.BLACK)
        Dim fontDescription As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, Font.NORMAL, BaseColor.BLACK)


        'define imagem
        'Dim arquivoImagem As String = "C:\Users\Usuario\Desktop\TCC1-TRABALHO\CABE"
        'Dim gif As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(New Uri(arquivoImagem))
        'doc.Add(gif)


        Dim titulo As New Paragraph("TCC da depressão", fontTitle)
        titulo.Alignment = Element.ALIGN_CENTER
        doc.Add(titulo)

        'PROPRIEDADES DA MADEIRA
        doc.Add(New Paragraph("1.PROPRIEDADES DA MADEIRA", fontSubTitle))

        doc.Add(New Paragraph("1.1.Resistência Característica da Madeira", fontSubTitle2))

        doc.Add(New Paragraph("Nome da Madeira:", fontItem))
        doc.Add(New Paragraph(_madeira.Name, fontDescription))

        doc.Add(New Paragraph("Resistência à Compressão Paralela às Fibras (fco):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistCompParalela & "MPa", fontDescription))

        doc.Add(New Paragraph("Resistência à Tração Paralela às Fibras (ft0):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistTracaoParalela & "MPa", fontDescription))

        doc.Add(New Paragraph("Resistência à Tração Normal às Fibras (ft90):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistTracaoNormal & "MPa", fontDescription))

        doc.Add(New Paragraph("Resistência ao Cisalhamento (fv):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistAoCisalhamento & "MPa", fontDescription))

        doc.Add(New Paragraph("Modulo de Elasticidade (Ec0):", fontItem))
        doc.Add(New Paragraph(_madeira.ModElasticidade & "MPa", fontDescription))

        doc.Add(New Paragraph("Densidade Aparente:", fontItem))
        doc.Add(New Paragraph(_madeira.DensAparente & "kg/m³", fontDescription))

        doc.Add(New Paragraph("2.RESISTÊNCIA DE CÁLCULO DA MADEIRA", fontTitle))
        doc.Add(New Paragraph("2.1Coeficiente de Modificação:", fontSubTitle))


        doc.Add(New Paragraph("kmod1", fontItem))
        doc.Add(New Paragraph(_madeira.Kmod1, fontDescription))

        doc.Add(New Paragraph("kmod1", fontItem))
        doc.Add(New Paragraph(_madeira.Kmod2, fontDescription))

        doc.Add(New Paragraph("kmod3", fontItem))
        doc.Add(New Paragraph(_madeira.Kmod3, fontDescription))

        doc.Add(New Paragraph("kmod", fontItem))

        doc.Add(New Paragraph("kmod = kmod1 * kmod2 * kmod3", fontForm))
        doc.Add(New Paragraph(_madeira.Kmod, fontDescription))

        doc.Add(New Paragraph("fd = kmod/ gama", fontForm))

        doc.Add(New Paragraph("Resistência de Cálculo à Compressão Paralela às Fibras (fc,0)", fontItem))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.resistCalculoCompressaoParalela & "MPa", fontDescription))

        doc.Add(New Paragraph("Resistência de Cálculo à Tração Paralela às Fibras (ft,0):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.resistCalculoTracaoParalela & "MPa", fontDescription))

        doc.Add(New Paragraph("Resistência de Cálculo à Tração Normal às Fibras (ft,90):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.resistCalculoTracaNormal & "MPa", fontDescription))

        doc.Add(New Paragraph("Resistência de Cálculo ao Cisalhamento (fv):", fontItem))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.resistCalculoAoCisalhamento & "MPa", fontDescription))

        doc.Add(New Paragraph("Rigidez de Cálculo (Eco,ef):", fontItem))
        doc.Add(New Paragraph("Eco,ef = kmod * Eco,m", fontForm))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.moduloElasticidade & "MPa", fontDescription))

        doc.Add(New Paragraph("Módulo de Elasticidade Transversal (Gef):", fontItem))
        doc.Add(New Paragraph("Gef = Eco,ef/15", fontForm))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.moduloElasticidadeTransversal & "MPa", fontDescription))

        doc.Add(New Paragraph("Densidade Aparente", fontItem))
        doc.Add(New Paragraph(_madeira.ResistenciaCalculo.densidadeAparente & "kg/m³", fontDescription))

        doc.Add(New Paragraph("4.DADOS INICIAIS:", fontSubTitle))
        doc.Add(New Paragraph("SEÇÃO TRANSVERSAL:", fontSubTitle))
        'como?
        'seção retangular
        doc.Add(New Paragraph("Base (b):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Altura (h):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento (L):", fontItem))
        'doc.Add(New Paragraph(& "m", fontDescription))

        'seção circular
        doc.Add(New Paragraph("Diametro (D):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento (L):", fontItem))
        'doc.Add(New Paragraph(& "m", fontDescription))

        'seção T

        doc.Add(New Paragraph("(bf):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(bw):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(hf):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(h):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(d):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento (L):", fontItem))
        'doc.Add(New Paragraph(& "m", fontDescription))

        'seção I
        doc.Add(New Paragraph("(bf1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(bf2):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(bw):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(hf1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(hf2):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(h):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(d):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento (L):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        'seção caixão
        doc.Add(New Paragraph("(b1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(b2):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(b3):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(b4):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(h1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(h2):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(h3):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("(h4):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento (L):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        '2 elementos justaposto

        doc.Add(New Paragraph("Comprimento Tota (L):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimentro Entre Eixos dos Espaçadores (L1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento Entre Eixos do Elementos (a1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Altura do Elemento Composto (h1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Base do Elemento Composto (b1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        '3 elementos justaposto

        doc.Add(New Paragraph("Comprimento Tota (L):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimentro Entre Eixos dos Espaçadores (L1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Comprimento Entre Eixos do Elementos (a1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Altura do Elemento Composto (h1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))

        doc.Add(New Paragraph("Base do Elemento Composto (b1):", fontItem))
        'doc.Add(New Paragraph(& "cm", fontDescription))


        doc.Add(New Paragraph("PROPRIEDADES GEOMÉTRICAS:", fontSubTitle))
        'seções simples
        If _exibir.ExibeProprSimples Then
            doc.Add(New Paragraph("Área (A):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.Area & "m²", fontDescription))

            doc.Add(New Paragraph("Eixo X:", fontItem))
            doc.Add(New Paragraph("Momento de Inércia (Ix):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoXmi & "m^4", fontDescription))

            doc.Add(New Paragraph("Raio de Giração (rx):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoXrg & "m²", fontDescription))

            doc.Add(New Paragraph("Módulo de Rêsistência (Wx):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoXmr & "m³", fontDescription))

            doc.Add(New Paragraph("Eixo Y:", fontItem))
            doc.Add(New Paragraph("Momento de Inércia (Iy):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYmi & "m^4", fontDescription))

            doc.Add(New Paragraph("Raio de Giração (ry):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYrg & "m²", fontDescription))

            doc.Add(New Paragraph("Módulo de Rêsistência (Wy):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYmr & "m³", fontDescription))
        End If

        'seções solidarizadas continuamente
        If _exibir.ExibeProprComp Then
            doc.Add(New Paragraph("Elemento componente:", fontItem))
            doc.Add(New Paragraph("Área (A1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.AreaA1 & "m²", fontDescription))

            doc.Add(New Paragraph("Momento de Inércia (I1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoXmi1 & "m^4", fontDescription))

            doc.Add(New Paragraph("Momento de Inércia (I2):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYmi1 & "m^4", fontDescription))

            doc.Add(New Paragraph("Seção Composta:", fontItem))
            doc.Add(New Paragraph("Área (A):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.Area & "m²", fontDescription))

            doc.Add(New Paragraph("Momento de Inércia (Ix):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoXmi & "m^4", fontDescription))

            doc.Add(New Paragraph("Momento de Inércia (Iy):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYmi & "m^4", fontDescription))

            doc.Add(New Paragraph("Coeficiente (By):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.CoefBy, fontDescription))

            doc.Add(New Paragraph("Módulo de Rêsistência (Wy):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYmr & "m³", fontDescription))

            doc.Add(New Paragraph("Inércia Efetiva (Ief):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.EixoYie & "m^4", fontDescription))

            doc.Add(New Paragraph(" Númetos de intervalos de comprimento (L1) em que fica divido o comprimento (L) total da peça (m):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.IntervalosM, fontDescription))

            doc.Add(New Paragraph("W2:", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.W2 & "m³", fontDescription))
        End If


        'TRAÇÃO:
        If _exibir.ExibeTracao Then
            doc.Add(New Paragraph("TRAÇÃO:", fontSubTitle))
            doc.Add(New Paragraph("Tensao solicitante de Projeto à Tracao", fontItem))
            doc.Add(New Paragraph("sigmat = Nt/A", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.TensaoTracao, "MPa", fontDescription))

            doc.Add(New Paragraph("Esbeltez Eixo x:", fontItem))
            doc.Add(New Paragraph("lambidax = L/ix ", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.EsbeltezTracaoX, fontDescription))

            doc.Add(New Paragraph("Esbeltez no Eixo y:", fontItem))
            doc.Add(New Paragraph("lambiday = L/iy", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.EsbeltezTracaoY, fontDescription))
        End If

        'COMPRESSÃO:
        If _exibir.ExibeCompressao Then
            doc.Add(New Paragraph("COMPRESSÃO:", fontSubTitle))

            doc.Add(New Paragraph("Tensao solicitante de Projeto à Compressao no Eixo X:", fontItem))
            doc.Add(New Paragraph("sigmacx = Nc/A", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoCompressaoX, fontDescription))

            doc.Add(New Paragraph("Tensao solicitante de Projeto à Compressao no Eixo Y:", fontItem))
            doc.Add(New Paragraph("sigmacy = Nc/A", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoCompressaoY, fontDescription))

            doc.Add(New Paragraph("Tensao Normal Devida à Flexão no Eixo X:", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoMxd, fontDescription))

            doc.Add(New Paragraph("Tensao Normal Devida à Flexão no Eixo Y:", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoMyd, fontDescription))

            doc.Add(New Paragraph("Esbeltez no Eixo X:", fontItem))
            doc.Add(New Paragraph("lambidax = L/ix ", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EsbeltezCompressaoX, fontDescription))

            doc.Add(New Paragraph("Esbeltez no Eixo Y:", fontItem))
            doc.Add(New Paragraph("lambidaY = L/iY ", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EsbeltezCompressaoY, fontDescription))

            doc.Add(New Paragraph("Comprimento Teórico de Referência", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.Lvinculado, fontDescription))

            doc.Add(New Paragraph("Forca Elastica em X:", fontItem))
            doc.Add(New Paragraph("Fe = pi² * Eco,ef *Ix)/(L0²)", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.ForcaElasticaX, fontDescription))

            doc.Add(New Paragraph("Forca Elastica em Y:", fontItem))
            doc.Add(New Paragraph("Fe = pi² * Eco,ef *Iy)/(L0²)", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.ForcaElasticaY, fontDescription))

            doc.Add(New Paragraph("Excentricidade Inicial no Eixo X (eix):", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EiX, fontDescription))

            doc.Add(New Paragraph("Excentricidade Inicial no Eixo Y (eiY):", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EiY, fontDescription))

            doc.Add(New Paragraph("Excentricidade Acidental no Eixo X (eax)", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EaX, fontDescription))

            doc.Add(New Paragraph("Excentricidade Acidental no Eixo Y (eay)", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EaY, fontDescription))

            doc.Add(New Paragraph("Excentricidade de Projeto no Eixo X (edx)", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.Edx, fontDescription))

            doc.Add(New Paragraph("Excentricidade de Projeto no Eixo Y (edy)", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.Edy, fontDescription))

            doc.Add(New Paragraph("Excentricidade Suplementar de primeira Ordem no Eixo X (ecx):", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.Ecx, fontDescription))

            doc.Add(New Paragraph("Excentricidade Suplementar de primeira Ordem no Eixo Y (ecy)", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.Ecy, fontDescription))

            doc.Add(New Paragraph("Excentricidade Efetiva de Primeira Ordem (e1,ef):", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.E1ef, fontDescription))

            doc.Add(New Paragraph("eig", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.Eig, fontDescription))

            doc.Add(New Paragraph("Condição de Segurança no EixoXX", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EixoXX, fontDescription))

            doc.Add(New Paragraph("Condição de Segurança no EixoYY", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EixoYY, fontDescription))
        End If

        'CISALHAMENTO:
        If _exibir.ExibeCisalhamento Then
            doc.Add(New Paragraph("CISALHAMENTO:", fontSubTitle))

            doc.Add(New Paragraph("Tensao de Cisalhamento Solicitante de Projeto no Eixo X:", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Cisalhamento.TensaoCisalhanteX, fontDescription))

            doc.Add(New Paragraph("Tensao de Cisalhamento Solicitante de Projeto no Eixo Y:", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Cisalhamento.TensaoCisalhanteY, fontDescription))

            doc.Add(New Paragraph("Esforço Solicitante de cisalhamento dos espaçadores (vd):", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Cisalhamento.Vd, fontDescription))
        End If

        'FLEXÃO:
        'COLOCAR OS IF ELSE 
        If _exibir.ExibeFlexaoSimplesReta Then
            doc.Add(New Paragraph("FLEXÃO:", fontSubTitle))
            doc.Add(New Paragraph("FLEXÃO SIMPLES RETA:", fontSubTitle))
            doc.Add(New Paragraph("Tensão de Cálculo na Borda Comprimida (sigma Cx):", fontItem))
            doc.Add(New Paragraph("sigmac,dy = Nsd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoCX, fontDescription))
        End If


        doc.Add(New Paragraph("Tensão de Cálculo na Borda Comprimida (sigma Cy)", fontItem))
        doc.Add(New Paragraph("sigmac,dy = Nsd/A", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.tensaoCY, fontDescription))

        doc.Add(New Paragraph("Tensão de Cálculo na Borda Tracionada (sigma Tx)", fontItem))
        doc.Add(New Paragraph("sigmat,dy = Nsd/A", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.tensaoTX, fontDescription))



        doc.Add(New Paragraph("Tensão de Cálculo na Borda Tracionada (sigma Ty)", fontItem))
        doc.Add(New Paragraph("sigmat,dy = Nsd/A", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.tensaoTY, fontDescription))


        doc.Add(New Paragraph("tensaoFlexaoX", fontItem))
        doc.Add(New Paragraph("", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoX, fontDescription))

        doc.Add(New Paragraph("tensaoFlexaoY", fontItem))
        doc.Add(New Paragraph("", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoY, fontDescription))

        doc.Add(New Paragraph("Compressão Normal no Apoio no Eixo X", fontItem))
        doc.Add(New Paragraph("sigmac,dx = Fd/A", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.ApoioX, fontDescription))

        doc.Add(New Paragraph("Compressão Normal no Apoio no Eixo Y", fontItem))
        doc.Add(New Paragraph("sigmac,dy = Fd/A", fontItem))
        doc.Add(New Paragraph(_madeira.Flexao.ApoioY, fontDescription))



        doc.Add(New Paragraph("FLEXÃO SIMPLES OBLÍQUA:", fontSubTitle))
        doc.Add(New Paragraph("FLEXOTRAÇÃO:", fontSubTitle))
        doc.Add(New Paragraph("FLEXOCOMPRESSÃO:", fontSubTitle))



        doc.Close()
        arquivo.Close()
    End Sub
End Class

