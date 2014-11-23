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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        For Each c As Button In Controls.OfType(Of Button)()
            AddHandler c.Click, AddressOf Button_Click
        Next
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        If pieceLocation = 0 Then
            If btn.BackColor = myColor.BackColor Then
                pieceLocation = convert_name_to_int(btn.Name.ToString)
                pieceButton = btn
                'disable all buttons except the right ones
                'readyMove(btn)
            Else
                MessageBox.Show("You cannot move that piece because it is not yours.")
            End If
        Else
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
        If myColor.BackColor = Color.Black Then
            If btn.Tag = "K" Then
                'do king things
            Else
                Select convert_name_to_int(btn.Name.ToString)
                    Case 1, 9, 17, 25, 33, 41, 49, 57
                        CType(convert_int_to_name(convert_name_to_int(btn.Name.ToString) + 9), Button).Enabled = True
                End Select
            End If
        End If
    End Sub



    Public Sub canmove(btn As Button)
        If pieceButton.BackColor = Color.Black Then
            If pieceButton.Tag = "K" Then
                'is king fucntion needs to be called
            Else
                If (pieceLocation < SelectedLocation) Then 'If pieceLocation is less than selectedLocation (moving in the proper direction)
                    Select Case pieceLocation
                        Case 1, 9, 17, 25, 33, 41, 49, 57
                            If (SelectedLocation = (pieceLocation + 9)) Then
                                checkforcapture(btn)
                                movepiece(btn)
                            End If
                            MsgBox("Invalid Move")
                        Case 8, 16, 24, 32, 40, 48, 56, 64
                            If (SelectedLocation = (pieceLocation + 7)) Then
                                checkforcapture(btn)
                                movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                            End If
                        Case Else
                            If (SelectedLocation = (pieceLocation + 7)) Or (SelectedLocation = (pieceLocation + 9)) Then
                                checkforcapture(btn)
                                movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                            End If
                    End Select
                Else
                    MsgBox("Invalid Move")
                End If
            End If
        ElseIf pieceButton.BackColor = Color.Red Then
            If pieceButton.Tag = "K" Then
                'isking()
            Else
                If (pieceLocation < SelectedLocation) Then 'If pieceLocation is less than selectedLocation (moving in the proper direction)
                    Select Case pieceLocation
                        Case 1, 9, 17, 25, 33, 41, 49, 57
                            If (SelectedLocation = (pieceLocation - 7)) Then
                                checkforcapture(btn)
                                movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                            End If
                        Case 8, 16, 24, 32, 40, 48, 56, 64
                            If (SelectedLocation = (pieceLocation - 9)) Then
                                checkforcapture(btn)
                                movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                            End If
                        Case Else
                            If (SelectedLocation = (pieceLocation - 7)) Or (SelectedLocation = (pieceLocation - 9)) Then
                                checkforcapture(btn)
                                movepiece(btn)
                            Else
                                MsgBox("Invalid Move")
                            End If
                    End Select
                Else
                    MsgBox("Invalid Move")
                End If
            End If
        End If
    End Sub

    Public Sub checkforcapture(btn As Button)
        If Not btn.BackColor = Color.Silver Then
            If myColor.BackColor = Color.Black Then
                Select Case SelectedLocation
                    Case 1, 9, 17, 25, 33, 41, 49, 57, 8, 16, 24, 32, 40, 48, 56, 64
                        'do not move
                        MsgBox("You can not move here!")
                        pieceLocation = 0
                        SelectedLocation = 0
                        Return
                    Case Else
                        If (SelectedLocation = pieceLocation + 7) Then
                            'capturedPiece = SelectedLocation
                            btn.BackColor = Color.Silver
                            btn.Tag = ""
                            SelectedLocation += 7
                            'movepiece()
                            Return

                        ElseIf (SelectedLocation = pieceLocation + 9) Then
                            'capturedPiece = SelectedLocation
                            btn.BackColor = Color.Silver
                            btn.Tag = ""
                            SelectedLocation += 9
                            'movepiece()
                            Return
                        Else
                            MsgBox("Invalid Move")
                        End If
                End Select
            Else
                MsgBox("You cannot jump your own piece!")
                Return
            End If

            If myColor.BackColor = Color.Red Then
                Select Case SelectedLocation
                    Case 1, 9, 17, 25, 33, 41, 49, 57, 8, 16, 24, 32, 40, 48, 56, 64
                        'do not move
                        MsgBox("You can not move here!")
                        pieceLocation = 0
                        SelectedLocation = 0
                        Return
                    Case Else
                        If (SelectedLocation = pieceLocation - 7) Then
                            'capturedPiece = SelectedLocation
                            btn.BackColor = Color.Silver
                            btn.Tag = ""
                            SelectedLocation -= 7
                            'movepiece()

                        ElseIf (SelectedLocation = pieceLocation - 9) Then
                            'capturedPiece = SelectedLocation
                            btn.BackColor = Color.Silver
                            btn.Tag = ""
                            SelectedLocation -= 9
                            'movepiece()
                        Else
                            MsgBox("Invalid Move")
                        End If
                End Select
            Else
                MsgBox("You cannot jump your own piece!")
                Return
            End If
        End If
    End Sub

    Sub movepiece(btn As Button)
        ''''make pi
        'capturedPiece = 0 'move piece and delete captured piece (reset piece location)
        pieceLocation = 0

        pieceLocation = SelectedLocation
        MsgBox("Piece:" & pieceLocation)
        btn.BackColor = Color.Silver
        CType(convert_int_to_name(pieceLocation), Button).BackColor = myColor.BackColor
        pieceButton.BackColor = Color.Silver
        'if it needs to be kinged
        If btn.BackColor = Color.Black Then
            If SelectedLocation > 51 Then
                btn.Tag = "K"
            End If
        Else
            If SelectedLocation < 9 Then
                btn.Tag = "K"
            End If
        End If
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
End Class


