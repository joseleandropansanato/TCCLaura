Public Class Validacao
    Enum TipoValidacao
        RealQualquer '-1
        RealMaiorOuIgualZero '0
        RealMaiorQueZero '1
        RealDiferenteDeZero '2
        InteiroMaiorOuIgualZero '3
        InteiroMaiorQueZero '4
    End Enum

    Public Shared Function ValidarDados(ByVal TxBox As TextBox, ByVal num As TipoValidacao) As Boolean
        Dim check As Boolean = False
        If TxBox.Text <> Nothing Then
            If Not IsNumeric(TxBox.Text) Then
                MessageBox.Show("Por favor, entre apenas com dados numéricos.")
                TxBox.Text = ""
                TxBox.Focus()
            Else
                TxBox.Text = TxBox.Text.Replace(".", ",")
                Try

                    Select Case num
                        Case TipoValidacao.RealQualquer
                            Convert.ToSingle(TxBox.Text)
                            check = True

                        Case TipoValidacao.RealMaiorOuIgualZero
                            If Convert.ToSingle(TxBox.Text) < 0 Then
                                MessageBox.Show("Por favor, entre com um valor positivo ou igual a 0 (zero) para o campo.")
                                TxBox.Text = ""
                                TxBox.Focus()
                            Else
                                check = True
                            End If

                        Case TipoValidacao.RealMaiorQueZero
                            If Convert.ToSingle(TxBox.Text) <= 0 Then
                                MessageBox.Show("Por favor, entre com um valor positivo e não nulo para o campo.")
                                TxBox.Text = ""
                                TxBox.Focus()
                            Else
                                check = True
                            End If

                        Case TipoValidacao.RealDiferenteDeZero
                            If Convert.ToSingle(TxBox.Text) <= 0 Then
                                MessageBox.Show("Por favor, entre com um valor positivo e não nulo para o campo.")
                                TxBox.Text = ""
                                TxBox.Focus()
                            Else
                                check = True
                            End If

                        Case TipoValidacao.InteiroMaiorOuIgualZero
                            Try
                                Convert.ToInt32(TxBox.Text)
                                If Convert.ToInt32(TxBox.Text) < 0 Then
                                    MessageBox.Show("Por favor, entre com um número inteiro POSITIVO ou igual a 0 (zero) para o campo.")
                                    TxBox.Text = ""
                                    TxBox.Focus()
                                Else
                                    check = True
                                End If
                            Catch
                                MessageBox.Show("Por favor, entre com um número INTEIRO positivo ou igual a 0 (zero) para o campo.")
                                TxBox.Text = ""
                                TxBox.Focus()
                            End Try

                        Case TipoValidacao.InteiroMaiorQueZero
                            Try
                                Convert.ToInt32(TxBox.Text)
                                If Convert.ToInt32(TxBox.Text) <= 0 Then
                                    MessageBox.Show("Por favor, entre com um número inteiro POSITIVO e NÃO NULO para o campo.")
                                    TxBox.Text = ""
                                    TxBox.Focus()
                                Else
                                    check = True
                                End If
                            Catch
                                MessageBox.Show("Por favor, entre com um número INTEIRO positivo e não nulo para o campo.")
                                TxBox.Text = ""
                                TxBox.Focus()
                            End Try

                    End Select

                Catch ex As Exception
                    'DValidar.ShowDialog()
                    TxBox.Focus()
                End Try
            End If
        End If
        Return check
    End Function

End Class
