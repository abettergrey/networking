Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms
Public Class Form1
    Dim strConn As String
    Dim strDBName As String = "DB.mdb"
    Dim DBCmd As OleDbCommand = New OleDbCommand
    Dim DBConn As OleDbConnection
    Dim SelectedLocation As Integer
    Dim pieceLocation As Integer
    Dim pieceButton As Button
    Dim invalid As Boolean = False

    Private Sub send_move(tile_from As Integer, tile_to As Integer, tile_cap As Integer, myturn As String)
        Dim strSQLCommand As String = "INSERT INTO checkers (color, tile_number_from, tile_number_to, tile_jump) VALUES (" & myturn _
                                      & ", " & tile_from & ", " & tile_to & ", " & tile_cap & ")"

        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBName
        DBConn = New OleDbConnection(strConn)
        DBConn.Open()
        DBCmd.CommandText = strSQLCommand
        DBCmd.Connection = DBConn
        DBCmd.ExecuteNonQuery()
        DBConn.Close()
    End Sub

    Private Sub wait_for_opponent()
        For Each c As Button In Controls.OfType(Of Button)()
            c.Enabled = False
        Next
        Dim flag As Boolean = False
        Dim strSQLCommand As String = "SELECT LAST(color, tile_number_from, tile_number_to, tile_jump) FROM checkers"
        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBName
        DBConn = New OleDbConnection(strConn)
        While flag = False
            DBCmd.CommandText = strSQLCommand
            Dim reader As OleDbDataReader = DBCmd.ExecuteReader()
            reader.Read()
            If Not reader(0).ToString = myColor.BackColor.ToString Then
                Dim tmp_color As Color = CType(reader(0), Color)
                Dim tile_from As Integer = CType(reader(1), Integer)
                Dim tile_to As Integer = CType(reader(2), Integer)
                Dim tile_cap As Integer = CType(reader(3), Integer)

                CType(convert_int_to_name(tile_from), Button).BackColor = Color.Silver
                CType(convert_int_to_name(tile_to), Button).BackColor = tmp_color
                If Not tile_cap = 0 Then
                    CType(convert_int_to_name(tile_cap), Button).BackColor = Color.Silver
                End If
                flag = True
            End If
        End While
        For Each c As Button In Controls.OfType(Of Button)()
            c.Enabled = True
        Next
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        For Each c As Button In Controls.OfType(Of Button)()
            AddHandler c.Click, AddressOf Button_Click
        Next
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        If pieceLocation = 0 Then  'if a button has not been clicked yet
            If btn.BackColor = turnColor.BackColor Then 'if it is your turn
                pieceLocation = convert_name_to_int(btn.Name.ToString) 'get button name
                pieceButton = btn
                'disable all buttons except the right ones

            Else 'it is not that colors turn
                MessageBox.Show("You cannot move that piece because it is not yours.")
            End If
        Else 'if it needs input of a second button
            SelectedLocation = convert_name_to_int(btn.Name.ToString)
            If pieceLocation = SelectedLocation Then
                pieceLocation = 0
                SelectedLocation = 0
                unreadyMove()
            Else
                canmove(btn)
            End If
        End If
    End Sub

    Public Sub unreadyMove()
        For Each c As Button In Controls.OfType(Of Button)()
            c.Enabled = True
        Next
    End Sub

    Public Sub readyMove(btn As Button)
        For Each c As Button In Controls.OfType(Of Button)()
            If Not c.Name.ToString.Equals(btn.Name.ToString) Then
                c.Enabled = False
            End If
        Next
        btn.Enabled = True
        If turnColor.BackColor = Color.Black Then
            If btn.Tag = "K" Then
                'do king things
            Else
                Select Case convert_name_to_int(btn.Name.ToString)
                    Case 1, 9, 17, 25, 33, 41, 49, 57
                        CType(convert_int_to_name(convert_name_to_int(btn.Name.ToString) + 9), Button).Enabled = True
                End Select
            End If
        End If
    End Sub



    Public Sub canmove(btn As Button)
        If pieceButton.BackColor = Color.Black Then
            If pieceButton.Tag = "K" Then
                isking(btn)
            Else
                If (pieceLocation < SelectedLocation) Then 'If pieceLocation is less than selectedLocation (moving in the proper direction)
                    Select Case pieceLocation
                        Case 1, 9, 17, 25, 33, 41, 49, 57
                            If (SelectedLocation = (pieceLocation + 9)) Then
                                checkforcapture(btn)
                            End If
                            MsgBox("Invalid Move")
                            InvalidMove()
                        Case 8, 16, 24, 32, 40, 48, 56, 64
                            If (SelectedLocation = (pieceLocation + 7)) Then
                                checkforcapture(btn)
                                'movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                                InvalidMove()
                            End If
                        Case Else
                            If (SelectedLocation = (pieceLocation + 7)) Or (SelectedLocation = (pieceLocation + 9)) Then
                                checkforcapture(btn)
                                'movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                                InvalidMove()
                            End If
                    End Select
                Else
                    MsgBox("Invalid Move")
                    InvalidMove()
                End If
                If Not invalid Then
                    turnColor.BackColor = Color.Red
                Else
                    invalid = False
                End If
            End If

        ElseIf pieceButton.BackColor = Color.Red Then
            If pieceButton.Tag = "K" Then
                isking(btn)
            Else
                If (pieceLocation > SelectedLocation) Then 'If pieceLocation is less than selectedLocation (moving in the proper direction)
                    Select Case pieceLocation
                        Case 1, 9, 17, 25, 33, 41, 49, 57
                            If (SelectedLocation = (pieceLocation - 7)) Then
                                checkforcapture(btn)
                            Else
                                MsgBox("Invalid Move")
                                InvalidMove()
                            End If
                        Case 8, 16, 24, 32, 40, 48, 56, 64
                            If (SelectedLocation = (pieceLocation - 9)) Then
                                checkforcapture(btn)
                            Else
                                MsgBox("Invalid Move")
                                InvalidMove()
                            End If
                        Case Else
                            If (SelectedLocation = (pieceLocation - 7)) Or (SelectedLocation = (pieceLocation - 9)) Then
                                checkforcapture(btn)
                            Else
                                MsgBox("Invalid Move")
                                InvalidMove()
                            End If
                    End Select
                Else
                    MsgBox("Invalid Move")
                    InvalidMove()
                End If
            End If
            If Not invalid Then 'change whos turn it is
                turnColor.BackColor = Color.Black
            Else
                invalid = False
            End If

        End If
    End Sub

    Public Sub checkforcapture(btn As Button)
        If Not btn.BackColor = Color.Silver And Not btn.BackColor = turnColor.BackColor Then
            If turnColor.BackColor = Color.Black Then
                Select Case SelectedLocation
                    Case 1, 9, 17, 25, 33, 41, 49, 57, 8, 16, 24, 32, 40, 48, 56, 64
                        'do not move
                        MsgBox("You can not move here!")
                        InvalidMove()
                        Return
                    Case 58 To 64 'For red make this 1 to 8
                        MsgBox("Can't jump a border piece.")
                        InvalidMove()
                        Return
                    Case Else
                        If (SelectedLocation = pieceLocation + 7) Then
                            If Not (CType(convert_int_to_name(SelectedLocation + 7), Button)).BackColor = Color.Silver Then
                                MsgBox("You cannot jump two pieces")
                                InvalidMove()
                            Else
                                btn.BackColor = Color.Silver
                                btn.Tag = ""
                                SelectedLocation += 7
                                movepiece(btn)
                            End If
                            Return

                        ElseIf (SelectedLocation = pieceLocation + 9) Then
                            If Not (CType(convert_int_to_name(SelectedLocation + 9), Button)).BackColor = Color.Silver Then
                                MsgBox("You cannot jump two pieces")
                                InvalidMove()
                            Else
                                btn.BackColor = Color.Silver
                                btn.Tag = ""
                                SelectedLocation += 9
                                movepiece(btn)
                            End If

                            Return
                        Else
                            MsgBox("Invalid Move")
                            InvalidMove()
                        End If
                End Select
                Return
            End If

            If turnColor.BackColor = Color.Red Then
                Select Case SelectedLocation
                    Case 1, 9, 17, 25, 33, 41, 49, 57, 8, 16, 24, 32, 40, 48, 56, 64
                        'do not move
                        MsgBox("You can not move here!")
                        InvalidMove()
                        Return
                    Case 1 To 8
                        MsgBox("Can't jump a border piece.")
                        InvalidMove()
                        Return
                    Case Else
                        If (SelectedLocation = pieceLocation - 7) Then
                            If Not (CType(convert_int_to_name(SelectedLocation - 7), Button)).BackColor = Color.Silver Then
                                MsgBox("You cannot jump two pieces")
                                InvalidMove()
                            Else
                                btn.BackColor = Color.Silver
                                btn.Tag = ""
                                SelectedLocation -= 7
                                movepiece(btn)
                            End If
                            Return

                        ElseIf (SelectedLocation = pieceLocation - 9) Then
                            If Not (CType(convert_int_to_name(SelectedLocation - 9), Button)).BackColor = Color.Silver Then
                                MsgBox("You cannot jump two pieces")
                                InvalidMove()
                            Else
                                btn.BackColor = Color.Silver
                                btn.Tag = ""
                                SelectedLocation -= 9
                                movepiece(btn)
                            End If

                            Return
                        Else
                            MsgBox("Invalid Move")
                            InvalidMove()
                        End If
                End Select
            Else
                MsgBox("You cannot jump your own piece!")
                InvalidMove()
                Return
            End If
        ElseIf Not btn.BackColor = turnColor.BackColor Then
            movepiece(btn)
        Else
            MessageBox.Show("cannot jump your own piece")
            InvalidMove()
        End If
    End Sub

    Sub movepiece(btn As Button)
        pieceLocation = SelectedLocation
        btn.BackColor = Color.Silver
        CType(convert_int_to_name(pieceLocation), Button).BackColor = turnColor.BackColor
        pieceButton.BackColor = Color.Silver
        'if it needs to be kinged
        If btn.BackColor = Color.Black Then
            If SelectedLocation > 51 Then
                btn.Tag = "K"
                btn.BackgroundImage = System.Drawing.Image.FromFile("King.jpeg")
            End If
        Else
            If SelectedLocation < 9 Then
                btn.Tag = "K"
            End If
        End If
        pieceLocation = 0
    End Sub

    Public Function convert_int_to_name(theInt As Integer)
        For Each c As Button In Controls.OfType(Of Button)()
            If c.Name = "Button_" & theInt.ToString Then
                Return c
            End If
        Next
        Dim error_o As Button
        Return error_o
    End Function

    Public Function convert_name_to_int(theName As String)

        For i As Integer = 1 To 64
            If theName = "Button_" & i.ToString Then
                Return i
            End If
        Next
        Return 0
    End Function


    ''kings cannot capture pieces... 
    Sub isking(btn As Button)
        Select Case pieceLocation
            Case 1, 9, 17, 25, 33, 41, 49, 57
                If SelectedLocation = (pieceLocation + 9) Or SelectedLocation = (pieceLocation - 9) Then
                    checkforcapture(btn)
                End If
            Case 8, 16, 24, 32, 40, 48, 56, 64
                If SelectedLocation = (pieceLocation + 7) Or SelectedLocation = (pieceLocation - 7) Then
                    checkforcapture(btn)
                End If
            Case Else
                If SelectedLocation = (pieceLocation + 7) Or SelectedLocation = (pieceLocation + 9) Or SelectedLocation = (pieceLocation - 7) Or SelectedLocation = (pieceLocation - 9) Then
                    checkforcapture(btn)
                End If
                If Not invalid Then
                    If turnColor.BackColor = Color.Red Then
                        turnColor.BackColor = Color.Black
                    Else
                        turnColor.BackColor = Color.Red
                    End If

                Else
                    invalid = False
                End If
        End Select
    End Sub

    Public Sub InvalidMove()
        pieceLocation = 0
        SelectedLocation = 0
        invalid = True
    End Sub
End Class
