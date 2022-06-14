Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports TCC2


Public Class GeradorPDF

    Private _madeira As Madeira
    Private _exibir As Exibicao
    'MOVE TO PROJECT CLASS ===============================
    Shared CaminhoRelat As String
    Shared pdfpath As String = "relatorio.pdf"
    Shared img As Image

    Public Sub New(mMadeira As Madeira, mExibir As Exibicao)
        _madeira = mMadeira
        _exibir = mExibir
    End Sub


    'PDF size measured in points. A4 size = 595 x 841 points (210 x 297 mm)
    Shared ReadOnly margin_left As Single = 70.8661 '70.8661 points = 2.5 cm
    Shared ReadOnly margin_top As Single = 70.8661
    Shared ReadOnly margin_right As Single = 70.8661
    Shared ReadOnly margin_bottom As Single = 70.8661


    'DOCUMENT FONTS
    Shared ReadOnly Arial As BaseFont = BaseFont.CreateFont("C:\Windows\Fonts\Arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)

    'ReadOnly TituloDocumento As Font = FontFactory.GetFont(FontFactory.HELVETICA, 12, iTextSharp.text.Font.BOLD)
    Shared ReadOnly TituloDocumento As Font = New Font(Arial, 12, iTextSharp.text.Font.NORMAL)
    Shared ReadOnly Titulo1 As Font = New Font(Arial, 10, iTextSharp.text.Font.BOLD)
    Shared ReadOnly Normal As Font = New Font(Arial, 10, iTextSharp.text.Font.NORMAL)
    Shared ReadOnly Normalnegrito As Font = New Font(Arial, 10, iTextSharp.text.Font.BOLD)
    Shared ReadOnly NormalSublinhado As New Font(Arial, 10, iTextSharp.text.Font.UNDERLINE)
    Shared ReadOnly Normal8 As Font = New Font(Arial, 8, iTextSharp.text.Font.NORMAL)
    Shared ReadOnly Normal8Negrito As Font = New Font(Arial, 8, iTextSharp.text.Font.BOLD)

    'CLASS LOCAL VARIABLES
    Shared ReadOnly p As Paragraph
    'Shared img As Image
    Shared table As PdfPTable
    Shared cell As PdfPCell
    Shared ReadOnly solid As ILineDash = New SolidLine()
    Shared ReadOnly dotted As ILineDash = New DottedLine()
    Shared ReadOnly dashed As ILineDash = New DashedLine()

    'CHAPTERS
    Shared ReadOnly topico(10) As Integer
    Public Shared nivel_figura As Integer
    Public Shared nivel_tabela As Integer

    Public Shared Sub CreatePDF(ByVal tipo As Integer, sender As Object, e As EventArgs) 'tipo=0: generate pdf; tipo=1: print
        For i As Integer = 0 To topico.Length - 1
            topico(i) = 0
        Next
        nivel_figura = 0
        nivel_tabela = 0
        '=====================================================================
        '                     RECOVER PROJECT FOLDER
        '=====================================================================
        If tipo = 0 Then
            Dim quebra As String = Chr(13) & Chr(10)
            Dim SFD As New SaveFileDialog

            If CaminhoRelat = vbLf Then
                SFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Else
                SFD.InitialDirectory = CaminhoRelat
            End If

            SFD.FileName = "relatorio.pdf"
            SFD.Filter = "(*pdf)|*.pdf|Todos os Arquivos (*.*)|*.*"

            Try
                If SFD.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    pdfpath = SFD.FileName

                    'Tratamento de string
                    Dim dest As String
                    dest = SFD.FileName
                    Dim w As Integer = dest.Length
                    Try
                        While dest(w - 1).ToString <> "\"
                            w -= 1
                        End While
                        dest = dest.Remove(w, (dest.Length - w))
                    Catch
                    End Try
                    CaminhoRelat = dest
                Else
                    Exit Sub
                End If
            Catch
            End Try
        Else
            pdfpath = System.IO.Path.GetTempPath + "\relatorio.pdf"
        End If

        '=====================================================================
        '                     CREATE PDF
        '=====================================================================
        Cursor.Current = Cursors.WaitCursor
        'pdfpath = System.IO.Path.GetTempPath + "relatorio.pdf"
        Console.WriteLine("path: {0}", pdfpath)
        Try
            Using msReport As FileStream = New FileStream(pdfpath, FileMode.Create)
                Using pdfDoc As Document = New Document(PageSize.A4, margin_left, margin_right, margin_top, margin_bottom)
                    Try
                        Dim writer As PdfWriter = PdfWriter.GetInstance(pdfDoc, msReport)
                        writer.PageEvent = New ITextEvents()
                        pdfDoc.Open()
                        pdfDoc.Add(New Paragraph(vbLf, New Font(Font.FontFamily.HELVETICA, 22)))

                        '=====================================================================
                        '                     CORPO DO RELATÓRIO
                        '=====================================================================


                        'PRINTS PRODUCT INFORMATION ON END PAGE -----------------------------
                        Dim availableSpace As Single = writer.GetVerticalPosition(True) - pdfDoc.BottomMargin()
                        If availableSpace < 56 Then pdfDoc.NewPage()

                        'Separator line
                        Dim canvas As PdfContentByte = writer.DirectContent
                        canvas.SetLineWidth(1)
                        canvas.SetColorStroke(BaseColor.BLACK)
                        canvas.MoveTo(70.8661, 122)
                        canvas.LineTo(pdfDoc.PageSize.Width - 70.8661, 122)
                        canvas.Stroke()

                        'StrucTool logo
                        'EU COMENTEI *****
                        'img = Image.GetInstance(Global.StrucTool_gabriel.My.Resources.Resources.StrucTool4, Drawing.Imaging.ImageFormat.Png)
                        img.ScaleToFit(48, 48)
                        img.SetAbsolutePosition(pdfDoc.PageSize.Width - 70.8661 - 48, 72)
                        img.Alignment = Element.ALIGN_RIGHT
                        canvas.AddImage(img)

                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase("StrucTool 2021. Módulo de Levantamento de Cargas devidas ao Vento em planta retangulares.", Titulo1), 70.8661, 122 - 12, 0)
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase("Software desenvolvido para fins educacionais.", Normal), 70.8661, 122 - 12 - 12, 0)
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase("Desenvolvedor: Alex Sandro da Costa, Engenheiro Civil", Normal), 70.8661, 122 - 12 - 12 - 12, 0)
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase("Supervisor: Prof. Rodolfo Krul Tessari, Doutor em Engenharia de Estruturas", Normal), 70.8661, 122 - 12 - 12 - 12 - 12, 0)
                        '--------------------------------------------------------------------

                        pdfDoc.Close()
                        'FProgressBar.Operacao(1, "Memorial concluído")
                        MessageBox.Show("Arquivo PDF salvo com sucesso.")
                        'FProgressBar.Close()
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString())
                    Finally
                    End Try
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Não foi possível salvar arquivo PDF." + ControlChars.NewLine + "O arquivo está sendo usado por outra pessoa ou programa.")
        End Try
        Cursor.Current = Cursors.Default

        '=====================================================================
        '                     PRINT/OPENS PDF
        '=====================================================================
        Try
            If tipo = 0 Then
                Process.Start(pdfpath)

            ElseIf tipo = 1 Then
                Try
                    Dim pathToExecutable As String = "AcroRd32.exe"
                    Dim sReport = pdfpath 'Complete name/path of PDF file
                    Dim printdlg As New PrintDialog

                    '%%%%%%%%%%% REVER NAO FECHAR TUDO ANTES DE SAIR %%%%%%%%%%%%
                    Process.Start(pdfpath)
                    System.Threading.Thread.Sleep(1000)

                    If printdlg.ShowDialog = Windows.Forms.DialogResult.OK Then

                        Dim SPrinter As String = printdlg.PrinterSettings.PrinterName.ToString
                        Dim objStartInfo As New ProcessStartInfo
                        Dim objProcess As New System.Diagnostics.Process

                        Try
                            objProcess.StartInfo.CreateNoWindow = True
                            objProcess.StartInfo.UseShellExecute = True
                            objProcess.StartInfo.FileName = pdfpath
                            objProcess.StartInfo.Arguments = """" & SPrinter & """"
                            objProcess.StartInfo.Verb = "PrintTo"
                            objProcess.StartInfo.CreateNoWindow = True
                            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            objProcess.Start()

                            objProcess.Close()
                            objProcess.Dispose()
                            objProcess = Nothing

                        Catch ex As Exception
                            objProcess.Kill()
                        End Try

                    End If
                Catch
                End Try
            Else
                Process.Start(pdfpath)
            End If
        Catch em As Exception
            MessageBox.Show("Erro ao gerar o arquivo...")
        End Try
    End Sub
    Public Shared Sub AddFigura(pdfDoc As Document, Picture_Box As PictureBox, legenda As String, Optional Prop As Integer = 100)
        nivel_figura += 1
        pdfDoc.Add(New Paragraph(vbLf, Normal))
        table = New PdfPTable(1) With {.WidthPercentage = Prop, .KeepTogether = True}
        cell = New PdfPCell(New Phrase("Figura " + nivel_figura.ToString + " - " + legenda, Normal)) With {.HorizontalAlignment = Element.ALIGN_CENTER, .VerticalAlignment = Element.ALIGN_MIDDLE, .Border = 0, .CellEvent = New CustomBorder(Nothing, Nothing, Nothing, Nothing)}
        table.AddCell(cell)
        Dim Figura As New Bitmap(Picture_Box.Width, Picture_Box.Height)
        Picture_Box.DrawToBitmap(Figura, Picture_Box.ClientRectangle)
        img = Image.GetInstance(Figura, Imaging.ImageFormat.Png)
        cell = New PdfPCell With {.HorizontalAlignment = Element.ALIGN_CENTER, .VerticalAlignment = Element.ALIGN_MIDDLE, .Border = 0}
        cell.AddElement(img)
        table.AddCell(cell)
        pdfDoc.Add(table)
        pdfDoc.Add(New Paragraph(vbLf, Normal))
    End Sub
    Public Shared Sub AddTopico(pdfDoc As Document, nivel As Integer, texto As String)
        topico(nivel) += 1
        For i As Integer = nivel + 1 To topico.Length - 1
            topico(i) = 0
        Next
        Dim numeracao As String = ""
        For i As Integer = 0 To nivel
            numeracao += topico(i).ToString + "."
        Next
        If nivel = 0 Then
            pdfDoc.Add(New Paragraph(numeracao + " " + texto, Titulo1))
        Else
            pdfDoc.Add(New Paragraph(numeracao + " " + texto, Normal))
        End If
        pdfDoc.Add(New Paragraph(vbLf, Normal))
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


        Dim titulo As New Paragraph("Universidade Tecnológica Federal do Paraná", fontTitle)
        titulo.Alignment = Element.ALIGN_CENTER
        doc.Add(titulo)

        Dim subtitulo As New Paragraph("Curso de Engenharia Civil - Campus Apucarana", fontTitle)
        subtitulo.Alignment = Element.ALIGN_CENTER
        doc.Add(subtitulo)

        Dim subtitulo1 As New Paragraph("Memorial de Cálculo", fontTitle)
        subtitulo1.Alignment = Element.ALIGN_CENTER
        doc.Add(subtitulo1)


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
        If _exibir.ExibeRetangular Then
            doc.Add(New Paragraph("Seção Retangular", fontSubTitle))
            doc.Add(New Paragraph("Base (b):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.base & "cm", fontDescription))

            doc.Add(New Paragraph("Altura (h):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.altura & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimento (L):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "m", fontDescription))
        End If

        'seção circular
        If _exibir.ExibeCircular Then
            doc.Add(New Paragraph("Seção Retangular", fontSubTitle))
            doc.Add(New Paragraph("Diametro (D):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.diametro & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimento (L):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "m", fontDescription))
        End If

        'seção T
        If _exibir.ExibeSecaoT Then
            doc.Add(New Paragraph("(bf):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b3 & "cm", fontDescription))

            doc.Add(New Paragraph("(bw):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b2 & "cm", fontDescription))

            doc.Add(New Paragraph("(hf):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h3 & "cm", fontDescription))

            doc.Add(New Paragraph("(h):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h2 & "cm", fontDescription))

            doc.Add(New Paragraph("(d):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "m", fontDescription))

        End If

        'seção I
        If _exibir.ExibeSecaoI Then
            doc.Add(New Paragraph("(bf1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b3 & "cm", fontDescription))

            doc.Add(New Paragraph("(bf2):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b1 & "cm", fontDescription))

            doc.Add(New Paragraph("(bw):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b2 & "cm", fontDescription))

            doc.Add(New Paragraph("(hf1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h3 & "cm", fontDescription))

            doc.Add(New Paragraph("(hf2):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h1 & "cm", fontDescription))

            doc.Add(New Paragraph("(h):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h2 & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimento (L):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "m", fontDescription))
        End If

        'seção caixão
        If _exibir.ExibeCaixao Then
            doc.Add(New Paragraph("(b1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b1c & "cm", fontDescription))

            doc.Add(New Paragraph("(b2):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b2c & "cm", fontDescription))

            doc.Add(New Paragraph("(b3):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b3c & "cm", fontDescription))

            doc.Add(New Paragraph("(b4):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b4c & "cm", fontDescription))

            doc.Add(New Paragraph("(h1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b1c & "cm", fontDescription))

            doc.Add(New Paragraph("(h2):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b2c & "cm", fontDescription))

            doc.Add(New Paragraph("(h3):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h3c & "cm", fontDescription))

            doc.Add(New Paragraph("(h4):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h4c & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimento (L):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "m", fontDescription))
        End If

        '2 elementos justaposto
        If _exibir.Exibe2Elementos Then
            doc.Add(New Paragraph("Comprimento Tota (L):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimentro Entre Eixos dos Espaçadores (L1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.l1 & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimento Entre Eixos do Elementos (a1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.a1 & "cm", fontDescription))

            doc.Add(New Paragraph("Altura do Elemento Composto (h1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h1 & "cm", fontDescription))

            doc.Add(New Paragraph("Base do Elemento Composto (b1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b1 & "cm", fontDescription))
        End If

        '3 elementos justaposto
        If _exibir.Exibe3Elementos Then
            doc.Add(New Paragraph("Comprimento Tota (L):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.comprimento & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimentro Entre Eixos dos Espaçadores (L1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.l1 & "cm", fontDescription))

            doc.Add(New Paragraph("Comprimento Entre Eixos do Elementos (a1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.a1 & "cm", fontDescription))

            doc.Add(New Paragraph("Altura do Elemento Composto (h1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.h1 & "cm", fontDescription))

            doc.Add(New Paragraph("Base do Elemento Composto (b1):", fontItem))
            doc.Add(New Paragraph(_madeira.PropriedadesGeometricas.b1 & "cm", fontDescription))
        End If

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
            doc.Add(New Paragraph("Tensãoo solicitante de Projeto à Tracao", fontItem))
            doc.Add(New Paragraph("Tensão = Nt/A", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.TensaoTracao, "MPa", fontDescription))

            doc.Add(New Paragraph("Esbeltez Eixo x:", fontItem))
            doc.Add(New Paragraph("Esbeltez = L/ix ", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.EsbeltezTracaoX, fontDescription))

            doc.Add(New Paragraph("Esbeltez no Eixo y:", fontItem))
            doc.Add(New Paragraph("Esbeltez = L/iy", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.EsbeltezTracaoY, fontDescription))
        End If

        'COMPRESSÃO:
        If _exibir.ExibeCompressao Then
            doc.Add(New Paragraph("COMPRESSÃO:", fontSubTitle))

            doc.Add(New Paragraph("Tensão solicitante de Projeto à Compressao no Eixo X:", fontItem))
            doc.Add(New Paragraph("Tensão = Nc/A", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoCompressaoX, fontDescription))

            doc.Add(New Paragraph("Tensao solicitante de Projeto à Compressao no Eixo Y:", fontItem))
            doc.Add(New Paragraph("Tensão = Nc/A", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoCompressaoY, fontDescription))

            doc.Add(New Paragraph("Tensao Normal Devida à Flexão no Eixo X:", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoMxd, fontDescription))

            doc.Add(New Paragraph("Tensao Normal Devida à Flexão no Eixo Y:", fontItem))
            doc.Add(New Paragraph("", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoMyd, fontDescription))

            doc.Add(New Paragraph("Esbeltez no Eixo X:", fontItem))
            doc.Add(New Paragraph("Esbeltez X = L/ix ", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EsbeltezCompressaoX, fontDescription))

            doc.Add(New Paragraph("Esbeltez no Eixo Y:", fontItem))
            doc.Add(New Paragraph("Esbeltez Y = L/iY ", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EsbeltezCompressaoY, fontDescription))

            doc.Add(New Paragraph("Comprimento Teórico de Referência", fontItem))
            doc.Add(New Paragraph("l0 = l * ke", fontItem))
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
            doc.Add(New Paragraph("normalCompressao / Area + mdY * y / Ix * (h1/2)", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EixoXX, fontDescription))

            doc.Add(New Paragraph("Condição de Segurança no EixoYY", fontItem))
            doc.Add(New Paragraph("normalCompressao / Area + mdY * I2 / Iy,ef * W2 + md / 2 * a1 * A1 * (1 - n * I2 / Iy,ef)", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.EixoYY, fontDescription))
        End If

        'CISALHAMENTO:
        If _exibir.ExibeCisalhamento Then
            doc.Add(New Paragraph("CISALHAMENTO:", fontSubTitle))

            doc.Add(New Paragraph("Tensão de Cisalhamento Solicitante de Projeto no Eixo X:", fontItem))
            doc.Add(New Paragraph("Tensão", fontItem))
            doc.Add(New Paragraph(_madeira.Cisalhamento.TensaoCisalhanteX, fontDescription))

            doc.Add(New Paragraph("Tensão de Cisalhamento Solicitante de Projeto no Eixo Y:", fontItem))
            doc.Add(New Paragraph("Tensão", fontItem))
            doc.Add(New Paragraph(_madeira.Cisalhamento.TensaoCisalhanteY, fontDescription))

            doc.Add(New Paragraph("Esforço Solicitante de cisalhamento dos espaçadores (vd):", fontItem))
            doc.Add(New Paragraph("vd", fontItem))
            doc.Add(New Paragraph(_madeira.Cisalhamento.Vd, fontDescription))
        End If

        'FLEXÃO:
        'COLOCAR OS IF ELSE 
        If _exibir.ExibeFlexaoSimplesReta Then
            doc.Add(New Paragraph("FLEXÃO:", fontSubTitle))
            doc.Add(New Paragraph("FLEXÃO SIMPLES RETA:", fontSubTitle))
            doc.Add(New Paragraph("Tensão de Cálculo::", fontItem))
            doc.Add(New Paragraph("Tensão = Nsd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoCX, fontDescription))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoTX, fontDescription))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoCY, fontDescription))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoTY, fontDescription))

            doc.Add(New Paragraph("COMPRESSÃO NORMAL NOS APOIOS:", fontSubTitle))
            doc.Add(New Paragraph("Eixo X:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioX, fontDescription))

            doc.Add(New Paragraph("Eixo Y:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioY, fontDescription))
        End If

        If _exibir.ExibeFlexaoSimplesObliqua Then
            doc.Add(New Paragraph("FLEXÃO SIMPLES OBLÍQUA:", fontSubTitle))
            doc.Add(New Paragraph("Tensão de Cálculo:", fontItem))
            doc.Add(New Paragraph("Tensão = Nsd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoX, fontDescription))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoY, fontDescription))

            doc.Add(New Paragraph("COMPRESSÃO NORMAL NOS APOIOS:", fontSubTitle))
            doc.Add(New Paragraph("Eixo X:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioX, fontDescription))

            doc.Add(New Paragraph("Eixo Y:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioY, fontDescription))
        End If

        If _exibir.ExibeFlexoTracao Then
            doc.Add(New Paragraph("FLEXOTRAÇÃO:", fontSubTitle))
            doc.Add(New Paragraph("Tensão = Nsd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoXT, fontDescription))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoYT, fontDescription))
            doc.Add(New Paragraph("Tensão solicitante de Projeto à Tracao", fontItem))
            doc.Add(New Paragraph("Tensão = Nt/A", fontItem))
            doc.Add(New Paragraph(_madeira.Tracao.TensaoTracao, "MPa", fontDescription))

            doc.Add(New Paragraph("COMPRESSÃO NORMAL NOS APOIOS:", fontSubTitle))
            doc.Add(New Paragraph("Eixo X:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioX, fontDescription))

            doc.Add(New Paragraph("Eixo Y:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioY, fontDescription))

        End If

        If _exibir.ExibeFlexoCompressao Then
            doc.Add(New Paragraph("FLEXOCOMPRESSÃO:", fontSubTitle))
            doc.Add(New Paragraph("Tensão = Nsd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoXC, fontDescription))
            doc.Add(New Paragraph(_madeira.Flexao.tensaoFlexaoYC, fontDescription))
            doc.Add(New Paragraph("Tensão solicitante de Projeto à Compressão", fontItem))
            doc.Add(New Paragraph("Tensão = Nc/A", fontItem))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoCompressaoX & "MPa", fontDescription))
            doc.Add(New Paragraph(_madeira.Compressao.TensaoCompressaoY & "MPa", fontDescription))

            doc.Add(New Paragraph("COMPRESSÃO NORMAL NOS APOIOS:", fontSubTitle))
            doc.Add(New Paragraph("Eixo X:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioX, fontDescription))

            doc.Add(New Paragraph("Eixo Y:", fontItem))
            doc.Add(New Paragraph("Tensão = Fd/A", fontItem))
            doc.Add(New Paragraph(_madeira.Flexao.ApoioY, fontDescription))

        End If


        doc.Close()
        arquivo.Close()
    End Sub
End Class

