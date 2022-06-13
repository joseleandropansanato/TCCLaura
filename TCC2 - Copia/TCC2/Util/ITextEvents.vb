Imports System.IO
Imports System.Drawing.Printing
Imports iTextSharp.text
Imports iTextSharp.text.pdf
'Imports System.Runtime.Serialization.Formatters.Binary
'Imports System.Runtime.Serialization
'Imports System.Data.OleDb
Imports System.Math
Imports Org.BouncyCastle.Crypto.Generators


Public Class ITextEvents
    Inherits PdfPageEventHelper

    Private Cb As PdfContentByte
    Private headerTemplate, footerTemplate As PdfTemplate
    Private bf As BaseFont = Nothing
    Private PrintTime As DateTime = DateTime.Now
    Private _header As String
    Private _footer As String

    Public Property Header As String
        Get
            Return _header
        End Get
        Set(ByVal value As String)
            _header = value
        End Set
    End Property
    Public Property Footer As String
        Get
            Return _footer
        End Get
        Set(ByVal value As String)
            _footer = value
        End Set
    End Property


    Public Overrides Sub OnOpenDocument(ByVal writer As PdfWriter, ByVal document As Document)
        Try
            PrintTime = DateTime.Now
            bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
            Cb = writer.DirectContent
            headerTemplate = Cb.CreateTemplate(100, 100)
            footerTemplate = Cb.CreateTemplate(50, 50)
        Catch de As DocumentException
        Catch ioe As System.IO.IOException
        End Try
    End Sub


    Public Overrides Sub OnEndPage(ByVal writer As pdf.PdfWriter, ByVal document As Document)
        MyBase.OnEndPage(writer, document)

        Dim baseFontSmall As Font = New Font(Font.FontFamily.HELVETICA, 8.0F, Font.NORMAL, BaseColor.BLACK)
        Dim baseFontSmallBold As Font = New Font(Font.FontFamily.HELVETICA, 8.0F, Font.BOLD, BaseColor.BLACK)
        Dim baseFontNormal As Font = New Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)
        Dim baseFontNormalBold As Font = New Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)
        Dim baseFonText_ig As Font = New Font(Font.FontFamily.HELVETICA, 12.0F, Font.NORMAL, BaseColor.BLACK)
        Dim baseFonText_igBold As Font = New Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.BLACK)

        Dim len As Single
        Dim img As Image
        Dim canvas As PdfContentByte = writer.DirectContent

        'HEADER ==============================================
        If writer.PageNumber = 1 Then
            'StrucTool logo
            'img = Image.GetInstance(Global.StrucTool_gabriel.My.Resources.Resources.StrucTool4, Drawing.Imaging.ImageFormat.Png)
            'img.ScaleToFit(70, 70)
            'img.SetAbsolutePosition(70.8661, document.PageSize.GetTop(96))
            'img.Alignment = Element.ALIGN_CENTER
            'canvas.AddImage(img)

            'UTFPR logo
            'img = Image.GetInstance(Global.StrucTool_gabriel.My.Resources.Resources.Render_UTFPR, Drawing.Imaging.ImageFormat.Png)
            'img.ScaleToFit(70, 70)
            'img.SetAbsolutePosition(document.PageSize.GetRight(70.8661 + 70), document.PageSize.GetTop(96))
            'img.Alignment = Element.ALIGN_CENTER
            'canvas.AddImage(img)

            'Plain text
            'ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER, New Phrase("StrucTool - Projeto: " + FDadosProjeto.Text_Projeto.Text, baseFonText_igBold), document.PageSize.Width() / 2, document.PageSize.GetTop(30 + 12), 0)
            ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER, New Phrase("Universidade Tecnológica Federal do Paraná", baseFonText_ig), document.PageSize.Width() / 2, document.PageSize.GetTop(30 + 12 + 6 + 12), 0)
            ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER, New Phrase("Curso de Engenharia Civil - Campus Apucarana", baseFontNormal), document.PageSize.Width() / 2, document.PageSize.GetTop(30 + 12 + 6 + 12 + 4 + 10), 0)
            ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER, New Phrase("Memorial De Cálculo", baseFontNormal), document.PageSize.Width() / 2, document.PageSize.GetTop(30 + 12 + 6 + 12 + 4 + 10 + 3 + 10), 0)

            'Separator lines
            Cb.SetLineWidth(2)
            Cb.SetColorStroke(BaseColor.BLACK)
            Cb.MoveTo(70.8661, document.PageSize.GetTop(98))
            Cb.LineTo(document.PageSize.Width - 70.8661, document.PageSize.GetTop(98))
            Cb.Stroke()
            Cb.SetColorStroke(New BaseColor(245, 197, 0))
            Cb.MoveTo(70.8661, document.PageSize.GetTop(102))
            Cb.LineTo(document.PageSize.Width - 70.8661, document.PageSize.GetTop(102))
            Cb.Stroke()
        Else
            'Plain text
            'ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase("StrucTool - Projeto: " + FDadosProjeto.Text_Projeto.Text + " - Memorial De Cálculo", baseFontSmallBold), 70.8661, document.PageSize.GetTop(68), 0)
            Dim Text_header As String = writer.PageNumber
            len = bf.GetWidthPoint(Text_header, 8)
            ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase(Text_header, baseFontSmall), document.PageSize.GetRight(70.8661 + len), document.PageSize.GetTop(68), 0)
            'Combo_.AddTemplate(headerTemplate, document.PageSize.GetRight(70.8661 + bf.GetWidthPoint("99", 8)), document.PageSize.GetTop(68))

            'Separator line
            Cb.SetLineWidth(1)
            Cb.SetColorStroke(BaseColor.BLACK)
            Cb.MoveTo(70.8661, document.PageSize.GetTop(70))
            Cb.LineTo(document.PageSize.Width - 70.8661, document.PageSize.GetTop(70))
            Cb.Stroke()
        End If



        'FOOTER ==============================================
        'Separator line
        Cb.SetLineWidth(1)
        Cb.SetColorStroke(BaseColor.BLACK)
        Cb.MoveTo(70.8661, 70)
        Cb.LineTo(document.PageSize.Width - 70.8661, 70)
        Cb.Stroke()

        'Date, time and page number
        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase(PrintTime.ToShortDateString() & " " & String.Format("{0:t}", DateTime.Now), baseFontSmallBold), 70.8661, 60, 0)
        'Dim Text_footer As String = "Page " & writer.PageNumber & " of "
        'len = bf.GetWidthPoint(Text_footer + "99", 8)
        'ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, New Phrase("Page " & writer.PageNumber & " of ", baseFontSmall), document.PageSize.GetRight(70.8661 + len), 60, 0)
        'Combo_.AddTemplate(footerTemplate, document.PageSize.GetRight(70.8661 + bf.GetWidthPoint("99", 8)), 60)


    End Sub

    Public Overrides Sub OnCloseDocument(ByVal writer As PdfWriter, ByVal document As Document)
        MyBase.OnCloseDocument(writer, document)
        headerTemplate.BeginText()
        headerTemplate.SetFontAndSize(bf, 8)
        headerTemplate.SetTextMatrix(0, 0)
        headerTemplate.ShowText((writer.PageNumber - 1).ToString())
        headerTemplate.EndText()

        footerTemplate.BeginText()
        footerTemplate.SetFontAndSize(bf, 8)
        footerTemplate.SetTextMatrix(0, 0)
        footerTemplate.ShowText((writer.PageNumber - 1).ToString())
        footerTemplate.EndText()
    End Sub
End Class



'PdfPTable BORDER CUSTOMIZATION COMMANDS
Interface ILineDash
    Sub ApplyLineDash(ByVal canvas As PdfContentByte)
End Interface
Class SolidLine : Implements ILineDash
    Public Sub ApplyLineDash(ByVal canvas As PdfContentByte) Implements ILineDash.ApplyLineDash
    End Sub
End Class

Class DottedLine : Implements ILineDash
    Public Sub ApplyLineDash(ByVal canvas As PdfContentByte) Implements ILineDash.ApplyLineDash
        canvas.SetLineCap(PdfContentByte.LINE_CAP_ROUND)
        canvas.SetLineDash(0, 4, 2)
    End Sub
End Class

Class DashedLine : Implements ILineDash
    Public Sub ApplyLineDash(ByVal canvas As PdfContentByte) Implements ILineDash.ApplyLineDash
        canvas.SetLineDash(3, 3)
    End Sub
End Class

Class CustomBorder : Implements IPdfPCellEvent
    Protected left As ILineDash
    Protected right As ILineDash
    Protected top As ILineDash
    Protected bottom As ILineDash
    Public Sub New(ByVal left As ILineDash, ByVal right As ILineDash, ByVal top As ILineDash, ByVal bottom As ILineDash)
        Me.left = left
        Me.right = right
        Me.top = top
        Me.bottom = bottom
    End Sub
    Public Sub CellLayout(cell As PdfPCell, position As Rectangle, canvases As PdfContentByte()) Implements IPdfPCellEvent.CellLayout
        Dim canvas As PdfContentByte = canvases(PdfPTable.LINECANVAS)

        If top IsNot Nothing Then
            canvas.SaveState()
            top.ApplyLineDash(canvas)
            canvas.MoveTo(position.Right, position.Top)
            canvas.LineTo(position.Left, position.Top)
            canvas.Stroke()
            canvas.RestoreState()
        End If

        If bottom IsNot Nothing Then
            canvas.SaveState()
            bottom.ApplyLineDash(canvas)
            canvas.MoveTo(position.Right, position.Bottom)
            canvas.LineTo(position.Left, position.Bottom)
            canvas.Stroke()
            canvas.RestoreState()
        End If

        If right IsNot Nothing Then
            canvas.SaveState()
            right.ApplyLineDash(canvas)
            canvas.MoveTo(position.Right, position.Top)
            canvas.LineTo(position.Right, position.Bottom)
            canvas.Stroke()
            canvas.RestoreState()
        End If

        If left IsNot Nothing Then
            canvas.SaveState()
            left.ApplyLineDash(canvas)
            canvas.MoveTo(position.Left, position.Top)
            canvas.LineTo(position.Left, position.Bottom)
            canvas.Stroke()
            canvas.RestoreState()
        End If
    End Sub
End Class
