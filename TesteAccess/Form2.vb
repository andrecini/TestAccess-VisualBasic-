Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb

Public Class Form2
    
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        atualizaSaldo()
        CarregaDados()
    End Sub 'Genérico


    #Region "Variáveis"
    Dim Lista As New List(Of String)
    Dim Conta As New ContaCorrente("André")
    Dim Da As New OleDbDataAdapter
    Dim Dt As DataTable
    Dim Cmd As New OleDbCommand
    #End Region


    ''' <summary>
    ''' Treinamento de criação e manipulação de Strings
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim strInput1 As String = txtInput1.Text
        Dim strInput2 As String = txtInput2.Text

        txtOutput.Text = strInput1 + strInput2
    End Sub 


    ''' <summary>
    ''' Treinamento de criação e manipulação 
    ''' de variáveis numéricas
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    #Region "Numerics"

    Private Sub rb_CheckedChanged(sender As Object, e As EventArgs) Handles rbInteiros.CheckedChanged, rbDoubles.CheckedChanged, rbDecimals.CheckedChanged
        gpInputNum.Enabled = True
        gpOutputNum.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If rbInteiros.Checked = True Then
            try
                Soma(Convert.ToInt32(txtValue1.Text), Convert.ToInt32(txtValue2.Text))
                Multiplicacao(Convert.ToInt32(txtValue1.Text), Convert.ToInt32(txtValue2.Text))
            Catch ex As Exception
                MessageBox.Show("Digite um inteiro!")
            End Try
        
        ElseIf rbDoubles.Checked = True then
            try
                Soma(Convert.ToDouble(txtValue1.Text), Convert.ToDouble(txtValue2.Text))
                Multiplicacao(Convert.ToDouble(txtValue1.Text), Convert.ToDouble(txtValue2.Text))
            Catch ex As Exception
                MessageBox.Show("Digite um Double (Com vírgula)!")
            End Try
        
        ElseIf rbDecimals.Checked = True then
            try
                Soma(Convert.ToDecimal(txtValue1.Text), Convert.ToDecimal(txtValue2.Text))
                Multiplicacao(Convert.ToDecimal(txtValue1.Text), Convert.ToDecimal(txtValue2.Text))
            Catch ex As Exception
                MessageBox.Show("Digite um Decimal (Com vírgula)!")
            End Try
        End If
    End Sub

    #Region "MÉTODOS -> VARIÁVEIS NUMÉRICAS"

    #Region "SOBRECARGAS -> SOMA"
    Sub Soma (v1 as Integer, v2 as Integer)
        txtSoma.Text = (v1 + v2).ToString()
    End Sub

    Sub Soma (v1 as Double, v2 as Double)
        txtSoma.Text = (v1 + v2).ToString()
    End Sub

    Sub Soma (v1 as Decimal, v2 as Decimal)
        txtSoma.Text = (v1 + v2).ToString()
    End Sub
    #End Region
    
    #Region "SOBRECARGAS -> MULTIPLICAÇÃO"
    Sub Multiplicacao (v1 as Integer, v2 as Integer)
        txtMulti.Text = (v1 * v2).ToString()
    End Sub

    Sub Multiplicacao (v1 as Double, v2 as Double)
        txtMulti.Text = (v1 * v2).ToString()
    End Sub

    Sub Multiplicacao (v1 as Decimal, v2 as Decimal)
        txtMulti.Text = (v1 * v2).ToString()
    End Sub
    #End Region

    #End Region

    #End Region


    ''' <summary>
    ''' Treinamento de utilização dos laços de repetição
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>    
    #Region "Repetitions"

    Private Sub rbNum_CheckedChanged(sender As Object, e As EventArgs) Handles rbAleatorio.CheckedChanged, rbDecresc.CheckedChanged, rbCresc.CheckedChanged

        If rbAleatorio.Checked Then
            Sorteia()

        ElseIf rbCresc.Checked Then
            OrdenaCrescente()

        ElseIf rbDecresc.Checked Then
            OrdenaDecrescente()

        End If

    End Sub

    #Region "MÉTODOS -> LAÇOS DE REPETIÇÃO"
        Sub Sorteia()
            lbNumbers.Items.Clear()

            Dim randomGenerator As Random = New Random
            Dim sorteado As Integer

            For index As Integer = 1 To nupQtdd.Value
                sorteado = randomGenerator.Next(nupMin.Value, nupMax.Value)

                lbNumbers.Items.Add(sorteado)
            Next
        End Sub

        Sub OrdenaCrescente()
            lbNumbers.Items.Clear()

            Dim contador As Integer = nupMin.Value

            For auxiliar As Integer = 1 To nupQtdd.Value
                Do While contador <= nupMax.Value
                    lbNumbers.Items.Add(contador.ToString())
                    contador+=1
                Loop

                contador = nupMin.value
            Next
        End Sub

        Sub OrdenaDecrescente() 
             lbNumbers.Items.Clear()

            Dim contador As Integer = nupMax.Value

            For auxiliar As Integer = 1 To nupQtdd.Value
                Do While contador >= nupMin.Value
                    lbNumbers.Items.Add(contador.ToString())
                    contador-=1
                Loop

                contador = nupMax.value
            Next
        End Sub

    #End Region
    #End Region


    ''' <summary>
    ''' Treinamento de criação e utilização de coleções genéricas
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>    
    #Region "Coleções"

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Adiciona()
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        Remove()
    End Sub

    Private Sub btnLength_Click(sender As Object, e As EventArgs) Handles btnLength.Click
        lblMessage.Text = Lista.Count.ToString()
    End Sub

    Private Sub btnSort_Click(sender As Object, e As EventArgs) Handles btnSort.Click
        Ordena()
    End Sub

    #Region "MÉTODOS -> COLEÇÕES"

            Sub Adiciona()
                Try
                    Lista.Add(txtInputWord.Text)
                    lbCollections.Items.Add(txtInputWord.Text)
                    lblMessage.Text = "Sucess!"
                Catch 
                    MessageBox.Show("Preencha todos os campos")
                End Try
            End Sub

            Sub Remove()
                Try
                    Lista.Remove(txtInputWord.Text)
                    lbCollections.Items.Remove(txtInputWord.Text)
                    lblMessage.Text = "Sucess!"
                Catch 
                    MessageBox.Show("Preencha todos os campos e com uma palavra que exista na lista!")
                End Try
            End Sub

            Sub Ordena()
                Lista.Sort()
                lbCollections.Items.Clear()
                lbCollections.Items.Add("Lista:")
       
                For Each item As String In Lista
                    lbCollections.Items.Add(item)
            Next
            End Sub

     #End Region

    #End Region

    ''' <summary>
    ''' Treinamento de implementação de Classes e Objetos
    ''' </summary>
    #Region "POO"
    Private Sub atualizaSaldo()
        lblConta.Text = Conta.Saldo
        
        If Convert.ToInt32(Conta.Saldo) >= 0 Then 
            lblConta.ForeColor = New Color().Green
        Else
            lblConta.ForeColor = New Color().Red
        End If
    End Sub

    Private Sub btnDepositar_Click(sender As Object, e As EventArgs) Handles btnDepositar.Click
        Dim valor As Decimal = 0D

        If String.IsNullOrEmpty(txtDeposito.Text) Then
            MessageBox.Show("Informe um valor para depósito!")
        Else
            try
                valor = Convert.ToDecimal(txtDeposito.Text)
                Conta.Deposito(valor)
                atualizaSaldo()
            Catch
                MessageBox.Show("Digite um valor válido!")
             End Try
        End If
    End Sub

    Private Sub btnSacar_Click(sender As Object, e As EventArgs) Handles btnSacar.Click
        Dim valor As Decimal = 0D

        If String.IsNullOrEmpty(txtSaque.Text) Then
            MessageBox.Show("Informe um valor para Saque!")
        Else
            try
                valor = Convert.ToDecimal(txtSaque.Text)
                Conta.Saque(valor)
                atualizaSaldo()
            Catch
                MessageBox.Show("Digite um valor válido!")
            End Try
        End If
    End Sub
    #End Region
    
    ''' <summary>
    ''' Treinamento de conexão com o banco de dados Access
    ''' </summary>
    #Region "Banco de Dados"
    Private Sub CarregaDados()

        Dim cn As New OleDb.OleDbConnection
        cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=T:\André Cini\TesteAccess\My Project\db50-test.mdb; Jet OLEDB:Database Password=98750"
        cn.Open()

         Try
            With Cmd 'Variável de comandos do OleDB
                .CommandType = CommandType.Text
                .CommandText = "SELECT * from Teste"
                .Connection = cn
            End With 'Realiza o Select do Banco de Dados específicado

            With Da 'Adaptador do OleDB
                .SelectCommand = Cmd
                Dt = New DataTable
                .Fill(Dt)
                dgvAlunos.DataSource = Dt
            End With 'Manda os dados carregados para o DataGridView

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

    End Sub 'Chamado no FormLoad

    Private Sub btnIncluir_Click(sender As System.Object, e As System.EventArgs) Handles btnIncluir.Click
        Dim cn As New OleDb.OleDbConnection
        cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=T:\André Cini\TesteAccess\My Project\db50-test.mdb; Jet OLEDB:Database Password=98750"
        cn.Open()

        Dim arrImagem() As Byte
        Dim strImagem As String
        Dim ms As New IO.MemoryStream

        If txtNome.Text = String.Empty Then
            MsgBox("Informe o nome do aluno")
            txtNome.Focus()
            Return
        End If

        '
        If Not IsNothing(Me.picFoto.Image) Then
            Me.picFoto.Image.Save(ms, Me.picFoto.Image.RawFormat)
            arrImagem = ms.GetBuffer
            strImagem = "?"
        Else
            arrImagem = Nothing
            strImagem = "NULL"
        End If

        Dim myCmd As New OleDb.OleDbCommand
        myCmd.Connection = cn
        myCmd.CommandText = "INSERT INTO Teste(codigo, nome, imagem) " & _
                            " VALUES( '" & Me.txtCodigo.Text & "', '" & Me.txtNome.Text & "'," & strImagem & ")"

        If strImagem = "?" Then
            myCmd.Parameters.Add(strImagem, OleDb.OleDbType.Binary).Value = arrImagem
        End If

        Try
        myCmd.ExecuteNonQuery()
        MsgBox("Dados Salvos com sucesso!")
        Catch
         MsgBox("Código já utilizado!")
        End Try


        cn.Close()
        CarregaDados()
    End Sub

    Private Sub btnProcurar_Click(sender As System.Object, e As System.EventArgs) Handles btnProcurar.Click
        If txtCodigo.Text = String.Empty Then
            MsgBox("Informe o codigo do aluno")
        Else
            Procurar(Me.txtCodigo.Text)
        End If
    End Sub

    Private Sub Procurar(ByVal codigo As Integer)
        Dim cn As New OleDb.OleDbConnection
        cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=T:\André Cini\TesteAccess\My Project\db50-test.mdb; Jet OLEDB:Database Password=98750"
        cn.Open()

        Dim arrImagem() As Byte
        Dim ms As New IO.MemoryStream
        Dim da As New OleDb.OleDbDataAdapter("SELECT * FROM Teste " & _
                                             " WHERE codigo='" & codigo & "'", cn)
        Dim dt As New DataTable
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            Me.txtCodigo.Text = dt.Rows(0).Item("codigo")
            Me.txtNome.Text = dt.Rows(0).Item("nome") & ""
            If Not IsDBNull(dt.Rows(0).Item("imagem")) Then
                arrImagem = dt.Rows(0).Item("imagem")
                For Each ar As Byte In arrImagem
                    ms.WriteByte(ar)
                Next
                '
                Me.picFoto.Image = System.Drawing.Image.FromStream(ms)
            Else
                Me.picFoto.Image = System.Drawing.Image.FromFile(Application.StartupPath & "/semfoto.jpg")
            End If
            Me.btnIncluir.Enabled = False
        Else
            MsgBox("Registro não localizado")
        End If

        cn.Close()

    End Sub

    Private Sub lnkProcurar_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkProcurar.LinkClicked
        If Me.opfImagem.ShowDialog = 1 Then
            Me.picFoto.Image = System.Drawing.Image.FromFile(Me.opfImagem.FileName)
        Else
            Me.picFoto.Image = System.Drawing.Image.FromFile("T:\André Cini\TesteAccess\bin\Debug")
        End If
    End Sub
    
    Private Sub dgvAlunos_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAlunos.CellClick
        Try
            'Dim codigo As Integer = dgvAlunos.Rows(e.RowIndex).Cells(e.ColumnIndex).Value()
            Dim codigo As Integer = dgvAlunos.Rows(e.RowIndex).Cells(0).Value()
            Procurar(codigo)
        Catch ex As Exception
            MsgBox("Seleção Inválida. Clique em uma célula com dados.")
        End Try
    End Sub

    Private Sub btnLimpar_Click(sender As System.Object, e As System.EventArgs) Handles btnLimpar.Click
        Me.txtCodigo.Text = ""
        Me.txtNome.Text = ""
        Me.picFoto.Image = Nothing
        Me.txtCodigo.Focus()
        Me.btnIncluir.Enabled = True
    End Sub

    Private Sub txtNome_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles txtNome.Validating
        If (txtNome.Text.Trim().Length = 0) Then
            erro.SetError(txtNome, "Informe o nome do aluno")
        Else
            erro.SetError(txtNome, "")
        End If
    End Sub

    Private Sub btnDeletar_Click(sender As System.Object, e As System.EventArgs) Handles btnDeletar.Click

        Dim resultado As DialogResult = MessageBox.Show("Confirma a exclusão deste registro ?", _
            "Excluir", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If resultado = vbYes Then
            Dim cn As New OleDb.OleDbConnection
            cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=T:\André Cini\TesteAccess\My Project\db50-test.mdb; Jet OLEDB:Database Password=98750"
            cn.Open()

            If txtCodigo.Text = String.Empty Then
                MsgBox("Informe o codigo do aluno")
                txtNome.Focus()
                Return
            End If

            Dim myCmd As New OleDb.OleDbCommand
            myCmd.Connection = cn
            myCmd.CommandText = "DELETE FROM Teste WHERE codigo = '" & txtCodigo.Text & "'"

            myCmd.ExecuteNonQuery()
            MsgBox("Dados excluídos com sucesso!")

            cn.Close()
            CarregaDados()
        End If
    End Sub

    Private Sub btnAlterar_Click(sender As System.Object, e As System.EventArgs) Handles btnAlterar.Click
        Dim cn As New OleDb.OleDbConnection
        cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=T:\André Cini\TesteAccess\My Project\db50-test.mdb; Jet OLEDB:Database Password=98750"
        cn.Open()

        Dim arrImagem() As Byte
        Dim strImagem As String
        Dim ms As New IO.MemoryStream

        If txtNome.Text = String.Empty Then
            MsgBox("Informe o nome do aluno")
            txtNome.Focus()
            Return
        End If

        '
        If Not IsNothing(Me.picFoto.Image) Then
            Me.picFoto.Image.Save(ms, Me.picFoto.Image.RawFormat)
            arrImagem = ms.GetBuffer
            strImagem = "?"
        Else
            arrImagem = Nothing
            strImagem = "NULL"
        End If

        Dim myCmd As New OleDb.OleDbCommand
        myCmd.Connection = cn
        myCmd.CommandText = "Update Teste SET nome = '" & txtNome.Text & "'," & "imagem = " & strImagem & " WHERE codigo = '" & txtCodigo.Text & "'"

        If strImagem = "?" Then
            myCmd.Parameters.Add(strImagem, OleDb.OleDbType.Binary).Value = arrImagem
        End If

        myCmd.ExecuteNonQuery()
        MsgBox("Dados Alterados com sucesso!")

        cn.Close()
        CarregaDados()
    End Sub
    #End Region
    
End Class