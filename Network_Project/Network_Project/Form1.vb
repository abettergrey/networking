Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms
Public Class Form1
    Dim strConn As String
    Dim strDBName As String = "DB.mdb"
    Dim DBCmd As OleDbCommand = New OleDbCommand
    Dim DBConn As OleDbConnection
    Sub send_move(tile_from As Integer, tile_to As Integer, myturn As String)
        Dim strSQLCommand As String = "INSERT INTO checkers (id, turn, tile_number_from, tile_number_to) VALUES (NULL, " & myturn _
                                      & ", " & tile_from & ", " & tile_to & ")"

        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBName
        DBConn = New OleDbConnection(strConn)
        DBConn.Open()
        DBCmd.CommandText = strSQLCommand
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()
        DBConn.Close()
    End Sub
End Class
