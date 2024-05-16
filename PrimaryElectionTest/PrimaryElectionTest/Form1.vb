Imports IBM.Data.DB2.iSeries
' Clifton Lindsey
' 3.10.23
' This program reads in election data from selected source databases and then send the appropriate data to the AS400
Public Class Form1
    ' database connection setup
    Dim FileNameLocal As String = ""
    Dim iSeriesCn As New iDB2Connection("DATASOURCE=10.1.80.1;USERID=FCCDBACC;PASSWORD=LEXFAY")

    Dim cn As New OleDb.OleDbConnection : Dim da As New OleDb.OleDbDataAdapter : Dim ds As New DataSet : Dim dbcmd As New OleDb.OleDbCommand : Dim cmd As New iDB2Command
    Dim cn2 As New OleDb.OleDbConnection : Dim da2 As New OleDb.OleDbDataAdapter : Dim ds2 As New DataSet : Dim dbcmd2 As New OleDb.OleDbCommand : Dim cmd2 As New iDB2Command

    ' change these strings to change the files being used in the counts
    Dim thisVoteCountDocument As String = "\Users\clifton.lindsey\Desktop\Primary Election Example\Vote Totals 0522.accdb" : Dim thisCandidateNumberDocument As String = "\users\clifton.lindsey\desktop\Primary Election Example\CandidateList.accdb"

    ' change these strings to change the db tables that are read
    Dim thisVoteTable As String = "VoteTotals0522" : Dim thisCandidateTable As String = "Candidates"
    Dim totalCandidates As Integer = 73 - 1 : Dim dbLengthCount As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub
    Private Sub SendMachCnt(ByVal MachCnt As String, ByVal PrecCode As String)
        'Console.WriteLine(PrecCode & " " & MachCnt)
        'My.Computer.FileSystem.WriteAllText("C:\Users\clifton.lindsey\Desktop\testMachCnt2.txt", PrecCode & " " & MachCnt & (vbLf), True)
        'cmd2.CommandText = "Update ELLIB.ELPREC500 Set #Voted500 = @Voted Where PreCode500 = '" & PrecCode & "'"
        'cmd2.Connection = iSeriesCn
        'cmd2.DeriveParameters()
        'cmd2.Parameters("@Voted").Value = MachCnt
        'cmd2.ExecuteNonQuery()
    End Sub
    Private Sub SendDataToISeries(ByVal CandNum As String, ByVal PrecCode As String, ByVal NumVotes As String, ByVal thisCand As String)
        If CandNum <> "023" Then
            Console.WriteLine(CandNum & " " & PrecCode & " " & NumVotes & " " & thisCand)
            My.Computer.FileSystem.WriteAllText("C:\Users\clifton.lindsey\Desktop\testISeries2.txt", CandNum & " " & PrecCode & " " & NumVotes & " " & thisCand & (vbLf), True)
            'cmd.CommandText = "insert into ellib.elctot301 (precode301, cand#301, votes301) values (@preccode, @candnum, @votes)"
            'cmd.Connection = iSeriesCn
            'cmd.DeriveParameters()
            'cmd.Parameters("@preccode").Value = PrecCode
            'cmd.Parameters("@candnum").Value = CandNum
            'cmd.Parameters("@votes").Value = Val(NumVotes)
            'cmd.ExecuteNonQuery()
        End If
    End Sub
    Sub getTotalPrecVotes(ByVal thisPrec As String)
        Dim x As Integer = 0 : Dim precVoteCount As String = vbNullString

        For x = 0 To dbLengthCount
            If thisPrec = ds.Tables(0).Rows(x).Item("#Precinct") Then
                precVoteCount = ds.Tables(0).Rows(x).Item("Ballots Cast")
            End If
        Next
        SendMachCnt(precVoteCount, thisPrec.Substring(0, 4))
    End Sub
    Sub fixNamesForiSeries(ByVal thisPrec As String, ByVal thisCand As String, ByVal thisVote As String, ByVal thisID As String)
        Dim x As Integer = 0 : Dim y As Integer = 0 : Dim candNum As String = vbNullString

        ' this may need to be reworked each election
        ' changes Yes/No options to their proper counterparts
        For y = 0 To dbLengthCount
            If thisPrec = ds.Tables(0).Rows(y).Item("#Precinct") And thisCand = ds.Tables(0).Rows(y).Item("Choice Name") Then
                If thisCand = "YES" And thisID = "2" Then
                    thisCand = "YES1"
                End If
                If thisCand = "YES" And thisID = "1" Then
                    thisCand = "YES2"
                End If
                If thisCand = "NO" And thisID = "1" Then
                    thisCand = "NO1"
                End If
                If thisCand = "NO" And thisID = "2" Then
                    thisCand = "NO2"
                End If
            End If
        Next

        ' sets the correct candidate id for the yes/no options
        For x = 0 To totalCandidates
            If thisCand = ds2.Tables(0).Rows(x).Item("Candidate") Then
                If thisCand = "YES1" And thisID = "2" Then
                    candNum = "113"
                    Exit For
                ElseIf thisCand = "NO1" And thisID = "1" Then
                    candNum = "114"
                    Exit For
                ElseIf thisCand = "Yes2" And thisID = "1" Then
                    candNum = "115"
                    Exit For
                ElseIf thisCand = "NO2" And thisID = "2" Then
                    candNum = "116"
                    Exit For
                Else
                    candNum = ds2.Tables(0).Rows(x).Item("Number")
                    Exit For
                End If
            End If
        Next

        If thisVote > 0 Then
            SendDataToISeries(candNum, thisPrec.Substring(0, 4), thisVote, thisCand)
        End If
    End Sub
    Private Sub PrecinctPartyCount()
        Dim rebPrecCount As Integer = 0
        Dim demPrecCount As Integer = 0
        Dim nrPrecCount As Integer = 0
        Dim PrecID As String = vbNullString
        Dim PrecNum As String = vbNullString
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim yString As String = vbNullString

        For y = 1 To 28
            For x = 0 To dbLengthCount
                PrecID = ds.Tables(0).Rows(x).Item("#Precinct")
                PrecNum = Microsoft.VisualBasic.Right(PrecID, 3)
                PrecNum = Microsoft.VisualBasic.Left(PrecNum, 2)
                PrecID = Microsoft.VisualBasic.Right(PrecID, 1)

                If Len(y.ToString()) = 1 Then
                    yString = "0" & y
                Else
                    yString = y.ToString()
                End If

                If yString = PrecNum Then

                    If PrecID = "D" Then
                        demPrecCount += (ds.Tables(0).Rows(x).Item("Election Day Voting Votes") + ds.Tables(0).Rows(x).Item("Absentee Voting Votes") + ds.Tables(0).Rows(x).Item("Early Voting Votes"))
                    End If
                    If PrecID = "R" Then
                        rebPrecCount += (ds.Tables(0).Rows(x).Item("Election Day Voting Votes") + ds.Tables(0).Rows(x).Item("Absentee Voting Votes") + ds.Tables(0).Rows(x).Item("Early Voting Votes"))
                    End If
                    If PrecID = "N" Then
                        nrPrecCount += (ds.Tables(0).Rows(x).Item("Election Day Voting Votes") + ds.Tables(0).Rows(x).Item("Absentee Voting Votes") + ds.Tables(0).Rows(x).Item("Early Voting Votes"))
                    End If
                End If
                Console.WriteLine("Step " & x)
            Next
            My.Computer.FileSystem.WriteAllText("C:\Users\clifton.lindsey\Desktop\testPrecinctPartyCount.txt", "PrecID: 0" & PrecNum & PrecID & " Reb: " & rebPrecCount & " Dem: " & demPrecCount & " NR: " & nrPrecCount & (vbLf), True)
            rebPrecCount = 0
            demPrecCount = 0
            nrPrecCount = 0
        Next
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        ' initializes and sets up variables
        Dim x As Integer = 0 : Dim y As Integer = 0 : Dim n As Integer = 0 : Dim p As Integer = 0
        Dim voteCount As String = vbNullString : Dim candName As String = vbNullString : Dim candNum As String = vbNullString
        Dim precVoteCount As Integer = 0 : Dim precinctVoteCountButAsAString As String = vbNullString : Dim thisPrec As String = vbNullString : Dim thisVote As String = vbNullString : Dim tVote As Integer = 0
        Dim thisID As String = vbNullString : Dim thisABSVote As Integer = 0 : Dim thisEarlyVote As Integer = 0 : Dim writeCount As Integer = 0 : Dim percentComplete As Integer = 0

        iSeriesCn.Open()

        ' database connection and querie block
        cn.ConnectionString = "provider=microsoft.ace.oledb.12.0;data source=c:" & thisVoteCountDocument : cn.Open()
        dbcmd.CommandText = "select * from " & thisVoteTable : dbcmd.Connection = cn : da.SelectCommand = dbcmd : da.Fill(ds)
        cn2.ConnectionString = "provider=microsoft.ace.oledb.12.0;data source=c:" & thisCandidateNumberDocument : cn2.Open()
        dbcmd2.CommandText = "select * from " & thisCandidateTable : dbcmd2.Connection = cn2 : da2.SelectCommand = dbcmd2 : da2.Fill(ds2)
        dbLengthCount = ds.Tables(0).Rows.Count - 1

        PrecinctPartyCount()

        ' statement that iterates through the table and gets desired information
        With ds.Tables(0)

            ' tracks day-of votes
            For x = 0 To dbLengthCount
                candName = .Rows(x).Item("Choice Name")
                If candName <> "Unassigned write-ins" And candName <> "Rejected write-ins" And candName <> "Republican Party" And candName <> "Democratic Party" Then
                    thisPrec = .Rows(x).Item("#Precinct") : thisVote = .Rows(x).Item("Election Day Voting Votes") : thisID = .Rows(x).Item("Choice ID")

                    ' fixes the Yes/No option to reflect which contest is being tracked
                    'fixNamesForiSeries(thisPrec, candName, thisVote, thisID)
                    If .Rows(x).Item("#Precinct") <> .Rows(x + 1).Item("#Precinct") Then
                        On Error Resume Next
                        'getTotalPrecVotes(thisPrec)
                    End If
                End If
                writeCount += 1
            Next
            writeCount = 0

            ' tracks absentee and early votes
            For n = 0 To totalCandidates
                candName = ds2.Tables(0).Rows(n).Item("Candidate") : candNum = ds2.Tables(0).Rows(n).Item("Number")
                tVote = 0
                For p = 0 To dbLengthCount
                    If candName <> "Unassigned write-ins" And candName <> "Rejected write-ins" And candName <> "Republican Party" And candName <> "Democratic Party" Then
                        If candName = .Rows(p).Item("Choice Name") Then
                            thisID = .Rows(p).Item("Choice ID") : thisABSVote = .Rows(p).Item("Absentee Voting Votes") : thisEarlyVote = .Rows(p).Item("Early Voting Votes")
                            tVote += thisABSVote + thisEarlyVote
                        End If

                        ' the yes/no options may need to be reworked for each election depending on potential amendments that are on the ballot
                        If candName = "YES1" Then
                            thisID = .Rows(p).Item("Choice ID")
                            If .Rows(p).Item("Choice Name") = "YES" And .Rows(p).Item("Contest ID") = "4" Then
                                thisABSVote = .Rows(p).Item("Absentee Voting Votes") : thisEarlyVote = .Rows(p).Item("Early Voting Votes")
                                tVote += thisABSVote + thisEarlyVote
                            End If
                        End If
                        If candName = "YES2" Then
                            thisID = .Rows(p).Item("Choice ID")
                            If .Rows(p).Item("Choice Name") = "YES" And .Rows(p).Item("Contest ID") = "1" Then
                                thisABSVote = .Rows(p).Item("Absentee Voting Votes") : thisEarlyVote = .Rows(p).Item("Early Voting Votes")
                                tVote += thisABSVote + thisEarlyVote
                            End If
                        End If
                        If candName = "NO1" Then
                            thisID = .Rows(p).Item("Choice ID")
                            If .Rows(p).Item("Choice Name") = "NO" And .Rows(p).Item("Contest ID") = "4" Then
                                thisABSVote = .Rows(p).Item("Absentee Voting Votes") : thisEarlyVote = .Rows(p).Item("Early Voting Votes")
                                tVote += thisABSVote + thisEarlyVote
                            End If
                        End If
                        If candName = "NO2" Then
                            thisID = .Rows(p).Item("Choice ID")
                            If .Rows(p).Item("Choice Name") = "NO" And .Rows(p).Item("Contest ID") = "1" Then
                                thisABSVote = .Rows(p).Item("Absentee Voting Votes") : thisEarlyVote = .Rows(p).Item("Early Voting Votes")
                                tVote += thisABSVote + thisEarlyVote
                            End If
                        End If
                    End If
                Next
                If tVote > 0 Then
                    'SendDataToISeries(candNum, "****", tVote, candName)
                End If
                writeCount += 1
            Next
        End With
        iSeriesCn.Close()
    End Sub
End Class
