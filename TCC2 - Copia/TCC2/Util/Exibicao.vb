﻿Public Class Exibicao

    Private _exibeTracao As Boolean = True
    Private _exibeCompressao As Boolean = True
    Private _exibeCisalhamento As Boolean = True
    Private _exibeFlexaoSimplesReta As Boolean = True
    Private _exibeFlexaoSimplesObliqua As Boolean = True
    Private _exibeFlexoTracao As Boolean = True
    Private _exibeFlexoCompressao As Boolean = True
    Private _exibeProprSimples As Boolean = True
    Private _exibeProprComp As Boolean = True

    Public Sub New()

    End Sub

    Public Property ExibeTracao() As Boolean
        Get
            Return _exibeTracao
        End Get
        Set(value As Boolean)
            _exibeTracao = value
        End Set
    End Property

    Public Property ExibeCompressao() As Boolean
        Get
            Return _exibeCompressao
        End Get
        Set(value As Boolean)
            _exibeCompressao = value
        End Set
    End Property

    Public Property ExibeCisalhamento() As Boolean
        Get
            Return _exibeCisalhamento
        End Get
        Set(value As Boolean)
            _exibeCisalhamento = value
        End Set
    End Property

    Public Property ExibeFlexaoSimplesReta() As Boolean
        Get
            Return _exibeFlexaoSimplesReta
        End Get
        Set(value As Boolean)
            _exibeFlexaoSimplesReta = value
        End Set
    End Property

    Public Property ExibeFlexaoSimplesObliqua() As Boolean
        Get
            Return _exibeFlexaoSimplesObliqua
        End Get
        Set(value As Boolean)
            _exibeFlexaoSimplesObliqua = value
        End Set
    End Property
    Public Property ExibeFlexoTracao() As Boolean
        Get
            Return _exibeFlexoTracao
        End Get
        Set(value As Boolean)
            _exibeFlexoTracao = value
        End Set
    End Property
    Public Property ExibeFlexoCompressao() As Boolean
        Get
            Return _exibeFlexoCompressao
        End Get
        Set(value As Boolean)
            _exibeFlexoCompressao = value
        End Set
    End Property
    Public Property ExibeProprSimples() As Boolean
        Get
            Return _exibeProprSimples
        End Get
        Set(value As Boolean)
            _exibeProprSimples = value
        End Set
    End Property

    Public Property ExibeProprComp() As Boolean
        Get
            Return _exibeProprComp
        End Get
        Set(value As Boolean)
            _exibeProprComp = value
        End Set
    End Property


End Class
