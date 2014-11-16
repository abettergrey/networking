Imports System.IO
Imports System.Data.OleDb

'Module1 is called on startup rather than the standard form.
'This allows us to check the database before the form is loaded.
Module Module1
    Sub Main()
        Dim strConn As String
        Dim strDBName As String = "DB.mdb"
        Dim myForm As New Form1

        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBName

        If Not File.Exists(strDBName) Then
            MakeDataBase(strConn)
        End If
        'ClearDatabase(strConn)

        myForm.ShowDialog()
    End Sub

    'Creates the database if it doesn't exist
    Sub MakeDataBase(strConn As String)
        Dim DBCat As New ADOX.Catalog
        Dim DBConn As OleDbConnection
        Dim DBCmd As OleDbCommand = New OleDbCommand

        DBCat.Create(strConn)
        DBConn = New OleDbConnection(strConn)

        DBConn.Open()

        DBCmd.CommandText = "CREATE TABLE checkers(id INT NOT NULL AUTO_INCREMENT, PRIMARY KEY( id ), turn VARCHAR(10), tile_number_from INT, tile_number_to INT)"
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()
        DBConn.Close()

    End Sub

    'clears all data for a new game
    Sub ClearDatabase(strConn As String)
        Dim DBConn As OleDbConnection
        Dim DBCmd As OleDbCommand = New OleDbCommand

        DBConn = New OleDbConnection(strConn)
        DBConn.Open()
        DBCmd.Connection = DBConn
        DBCmd.CommandText = "DELETE * FROM checkers"
        DBCmd.ExecuteNonQuery()
        DBConn.Close()
    End Sub
End Module
