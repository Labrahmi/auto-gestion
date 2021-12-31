Imports System.Data.SqlClient
'
Public Class Form1
    '
    Dim con As SqlConnection = New SqlConnection("Data Source=DESKTOP-70GDM1N\SQLEXPRESS;Initial Catalog=gestion_voitures;Integrated Security=True")
    '
    Private Sub addFct()
        Dim ajQuery As String = "insert into voiture values('" & matrBox.Text & "','" & pajBox.Text & "','" & ndjBox.Text & "','" & marqCombo.SelectedValue.ToString & "')"
        Dim VerQuery As String = "select * from voiture where matricul = '" & matrBox.Text & "' "
        Dim cmdD As SqlCommand
        Dim cmdV As SqlCommand
        cmdD = New SqlCommand(ajQuery, con)
        cmdV = New SqlCommand(VerQuery, con)
        con.Open()
        Dim VerDatread As SqlDataReader
        VerDatread = cmdV.ExecuteReader
        If VerDatread.HasRows Then
            MessageBox.Show("Matricule already exist!")
        Else
            VerDatread.Close()
            cmdD.ExecuteNonQuery()
            MessageBox.Show("Added successfully!")
        End If
        con.Close()
        getCars()
        calcttl()
    End Sub
    '
    Private Sub delFct()
        Dim delQuery As String = "delete from voiture where matricul ='" + matrBox.Text + "' "
        Dim VerQuery As String = "select * from voiture where matricul = '" + matrBox.Text + "' "
        Dim cmdD As SqlCommand
        Dim cmdV As SqlCommand
        cmdD = New SqlCommand(delQuery, con)
        cmdV = New SqlCommand(VerQuery, con)
        con.Open()
        Dim VerDatread As SqlDataReader
        VerDatread = cmdV.ExecuteReader
        If VerDatread.HasRows Then
            VerDatread.Close()
            cmdD.ExecuteNonQuery()
            MessageBox.Show("Deleted successfully!")
        Else
            MessageBox.Show("Not existing Matricule!")
        End If
        con.Close()
        getCars()
        calcttl()
    End Sub
    '
    Private Sub modFct()
        Dim modQuery As String = "UPDATE voiture set prix = '" + pajBox.Text + "', nbrjrs = '" + ndjBox.Text + "', id_marque = '" + marqCombo.SelectedValue.ToString + "' where matricul = '" + matrBox.Text + "'"
        Dim VerQuery As String = "select * from voiture where matricul = '" + matrBox.Text + "' "
        Dim cmdM As SqlCommand
        Dim cmdV As SqlCommand
        cmdM = New SqlCommand(modQuery, con)
        cmdV = New SqlCommand(VerQuery, con)
        con.Open()
        Dim VerDatread As SqlDataReader
        VerDatread = cmdV.ExecuteReader
        If VerDatread.HasRows Then
            VerDatread.Close()
            cmdM.ExecuteNonQuery()
            MessageBox.Show("Modified successfully!")
        Else
            MessageBox.Show("Not existing Matricule!")
        End If
        con.Close()
        getCars()
        calcttl()
    End Sub
    '
    Private Sub rechFct()
        Dim rechQuery As String = "select * from voiture v, marque m where m.id_marque=v.id_marque and matricul = '" + matrBox.Text + "'"
        Dim VerQuery As String = "select * from voiture where matricul = '" + matrBox.Text + "' "
        Dim cmdM As SqlCommand
        Dim cmdV As SqlCommand
        cmdM = New SqlCommand(rechQuery, con)
        cmdV = New SqlCommand(VerQuery, con)
        con.Open()
        Dim VerDatread As SqlDataReader
        VerDatread = cmdV.ExecuteReader
        If VerDatread.HasRows Then
            VerDatread.Close()
            Dim DrechDataread As SqlDataReader
            DrechDataread = cmdM.ExecuteReader
            Dim rechDatatable As New DataTable
            rechDatatable.Load(DrechDataread)
            pajBox.Text = rechDatatable.Rows(0)("prix").ToString
            ndjBox.Text = rechDatatable.Rows(0)("nbrjrs").ToString
            marqCombo.Text = rechDatatable.Rows(0)("libelle").ToString
        Else
            MessageBox.Show("Not existing Matricule!")
        End If
        con.Close()
        getCars()
        calcttl()
    End Sub
    '
    Private Sub ajouButton_Click(sender As Object, e As EventArgs) Handles ajouButton.Click
        If matrBox.Text <> "" Then
            If IsNumeric(pajBox.Text) Then
                If IsNumeric(ndjBox.Text) And Not ndjBox.Text.Contains(".") Then
                    addFct()
                    getCars()
                Else
                    MessageBox.Show("Invalid days form")
                End If
            Else
                MessageBox.Show("Invalid Price per day")
            End If
        Else
            MessageBox.Show("Matricule is necessary")
        End If
    End Sub






    Private Sub suppButton_Click(sender As Object, e As EventArgs) Handles suppButton.Click
        If matrBox.Text <> "" Then
            delFct()
            getCars()
        End If
    End Sub

    Private Sub modButton_Click(sender As Object, e As EventArgs) Handles modButton.Click
        If matrBox.Text <> "" Then
            If IsNumeric(pajBox.Text) Then
                If IsNumeric(ndjBox.Text) And Not ndjBox.Text.Contains(".") Then
                    modFct()
                    getCars()
                Else
                    MessageBox.Show("Invalid days form")
                End If
            Else
                MessageBox.Show("Invalid Price per day")
            End If
        Else
            MessageBox.Show("Matricule is necessary")
        End If
    End Sub

    Private Sub rechButton_Click(sender As Object, e As EventArgs) Handles rechButton.Click
        rechFct()
        getCars()
    End Sub




    Private Sub calcttl()
        con.Open()

        Dim query As String = "select sum(prix*nbrjrs) as SUM_ from voiture"
        Dim cmd As SqlCommand
        cmd = New SqlCommand(query, con)
        totPrixBox.Text = cmd.ExecuteScalar.ToString + " Dh"

        con.Close()
        ''
        'Dim calcDatareader As SqlDataReader
        '
        'calcDatareader = cmd.ExecuteReader
        'Dim calcTable As New DataTable
        'calcTable.Load(calcDatareader)
        'Dim varC As String = calcTable.Rows(0)("SUM_").ToString
        'If varC = "NULL" Then
        ' totPrixBox.Text = "0 Dh"
        ' Else
        ' totPrixBox.Text = varC + " Dh"
        ' End If
        ' 

    End Sub

    '
    Private Sub getCars()
        Dim query As String = "select matricul, prix as 'Prix par jour', nbrjrs as 'Nombre de jours', libelle , (prix*nbrjrs) as Totale from voiture v, marque m where m.id_marque=v.id_marque order by(prix*nbrjrs) DESC"
        Dim cmd As SqlCommand
        cmd = New SqlCommand(query, con)
        Dim getComDatRead As SqlDataReader
        con.Open()
        getComDatRead = cmd.ExecuteReader
        Dim getComDatatable As New DataTable
        getComDatatable.Load(getComDatRead)
        Guna2DataGridView1.DataSource = getComDatatable
        con.Close()
    End Sub
    '
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        calcttl()
        getCars()
        Dim loadComQuery As String = "select * from marque order by libelle"
        Dim cmd As SqlCommand
        cmd = New SqlCommand(loadComQuery, con)
        Dim loadComDatRead As SqlDataReader
        con.Open()
        loadComDatRead = cmd.ExecuteReader
        Dim loadComDatatable As New DataTable
        loadComDatatable.Load(loadComDatRead)
        marqCombo.DataSource = loadComDatatable
        marqCombo.DisplayMember = "libelle"
        marqCombo.ValueMember = "id_marque"
        con.Close()
    End Sub


    '
End Class
