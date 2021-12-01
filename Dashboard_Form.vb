Imports System.Data.SqlClient

Public Class Dashboard_Form

    Dim DataConnection As New SqlConnection()
    Dim Reader As SqlDataReader 'Creates a new reader for SQL queries that require data to be read from the database
    Dim Writer As SqlDataAdapter 'Creates a writer, used to write data to the database
    Dim command 'By not defining the variable's data type, it can be given the data type of the value that it is set to

    Dim value As String 'Set to the latest value returned from the database, when reading from the database
    Dim valueCount As Integer = 0 'Set to the length of the array of values returned from the database
    Dim Query As String
    Dim Level(1) As String
    Dim UserLevel As String
    Dim AllTheWords(20) As String
    Dim AllLength As Integer
    Dim WordInstances(20) As String
    Dim OutcomeInstances(20) As String
    Dim InstanceLength As Integer
    Dim ThreeTimes As Integer = 0
    Dim v As Integer
    Dim sameCount As Integer
    Dim length As Integer
    Dim threeless(20) As String
    Dim InstancesWord(20) As String
    Dim InstancesOutcome(20) As String
    Dim wrong(10) As String
    Dim InstanceOrder(20) As String
    Dim LastOne As Integer
    Dim InstanceCount As Integer = 0
    Dim Never(10) As String
    Dim InstanceWords(20) As String
    Dim instanceSize As Integer
    Dim same As Boolean
    Dim IncoEmpty As Boolean
    Dim NeverEmpty As Boolean
    Dim NotEmpty As Boolean
    Dim IncorrectPicked As Boolean
    Dim NeverPickedPicked As Boolean
    Dim NotThreePicked As Boolean
    Dim pickValue As Integer
    Dim random As New Random()
    Shared Word As String
    Shared WordLength As Integer
    Dim boundary As Integer
    Dim pickPos As Integer
    Dim pickVal As String
    Dim WordArray() As String
    Dim nextQuery As String
    Dim TheWordLength() As String
    Dim FinalWrong(20) As String
    Dim FinalCorrect(20) As String
    Dim PastWrong(20) As String
    Dim PastCorrect(20) As String
    Dim WrongLength As Integer = 0
    Dim CorrectLength As Integer = 0
    Dim LevelString As String
    Dim FinalIncorrect(20) As String
    Dim FinalNotThree(20) As String
    Dim FinalNeverPicked(20) As String
    Dim Which As String
    Dim IncorrectLength As Integer
    Dim NotThreeLength As Integer
    Dim NeverPickedLength As Integer
    Dim UserLevelArray(1) As String
    Dim getVal As String
    Dim LevelUp As Boolean
    Dim alreadyPlaced As Boolean

    Function establishCon()
        Dim ConnectionData As String = "server=DESKTOP-H3A0KDT\SQLEXPRESS; Database=Logins;Integrated Security=true" 'Local variable that holds the connection data
        Dim Connection As New SqlConnection(ConnectionData) 'establishes a connection in a local varibale, passed then to the global variable DataConnection
        Return Connection
    End Function

    Function ApplyQuery(Query, returnVal)
        Dim valuesArray(20) As String
        valueCount = 0
        DataConnection.Open() 'Opens the connection with the database
        command = New SqlCommand(Query, DataConnection) 'sets the command that will be sent to the database, with the correct query and connection
        If returnVal = False Then
            Writer = command.ExecuteScalar() 'Executes the command when the query requires writing to the database, hence no values being returned is expected
        Else
            Reader = command.ExecuteReader() 'Executes the command when the query requires reading from the database.
            While Reader.Read() 'Reads the data from the database, loops until all values have been read from the database
                value = String.Format("{0}", Reader(0)) 'formats the value so that it is a readable and storable string
                'MessageBox.Show("Value " + value)
                valuesArray(valueCount) = value 'appends to the values array, which is later returned
                valueCount += 1 'Increments the size of the value array
            End While
        End If
        DataConnection.Close() 'Closes the connection, regardless of if the query was a read or a write one
        Return valuesArray
    End Function

    Function AllWords(UserName)
        Query = "SELECT Userlevel FROM dbo.NewUserInfo WHERE UserName = " + "'" + UserName + "'"
        'MessageBox.Show("1")
        Level = ApplyQuery(Query, True)
        UserLevel = Level(0)
        If UserLevel = "1" Then
            Query = "SELECT WordID FROM dbo.FirstWord"
        ElseIf UserLevel = "2" Then
            Query = "SELECT WordID FROM dbo.SecondWord"
        ElseIf UserLevel = "3" Then
            Query = "SELECT WordID FROM dbo.ThirdWord"
        Else
            MessageBox.Show("Error at query")
        End If
        'MessageBox.Show("2")
        AllTheWords = ApplyQuery(Query, True) 'Only need to do this query once, then can just use the all the words array wherever needed
        AllLength = valueCount
        Query = "SELECT WordID FROM dbo.instances WHERE UserName = " + "'" + UserName + "'"
        'MessageBox.Show("3")
        WordInstances = ApplyQuery(Query, True)
        InstanceLength = valueCount
        Query = "SELECT completed FROM dbo.instances WHERE UserName = " + "'" + UserName + "'" 'Split the instances and words into seperate arrays for simplicity
        'MessageBox.Show("4")
        OutcomeInstances = ApplyQuery(Query, True)
        ThreeTimes = 0
        For n = 0 To AllLength - 1
            value = AllTheWords(n)
            v = 0
            sameCount = 0
            While v < InstanceLength
                If WordInstances(v) = value And OutcomeInstances(v) = "1" Then
                    sameCount += 1
                End If
                v += 1
            End While
            If sameCount >= 3 Then
                ThreeTimes += 1
            End If
        Next
        If ThreeTimes = AllLength Then
            UserLevel += 1
            If UserLevel > 3 Then
                UserLevel = 3
            Else
                Query = "INSERT INTO dbo.NewUserInfo(UserLevel) VALUES (" + "'" + UserLevel + "') WHERE UserName = " + "'" + UserName + "'"
                'MessageBox.Show("5")
                ApplyQuery(Query, False)
            End If
            Return True
        Else
            Return False
        End If
    End Function

    Function IsEmpty(ThisArray())
        'MessageBox.Show("Called")
        If ThisArray(0) = "" Then
            'MessageBox.Show("Empty")
            Return True
        Else
            Return False
        End If
    End Function

    Function NotThree()
        length = 0
        If UserLevel = 1 Then
            Query = "SELECT WordID FROM dbo.FirstWord"
        ElseIf UserLevel = 2 Then
            Query = "SELECT WordID FROM dbo.SecondWord"
        ElseIf UserLevel = 3 Then
            Query = "SELECT WordID FROM dbo.ThirdWord"
        End If
        'MessageBox.Show("Username " + LogIn_Form.User)
        Query = "SELECT WordID FROM dbo.instances WHERE UserName = " + "'" + LogIn_Form.User + "' AND WordID LIKE" + "'" + UserLevel + "%'"
        'MessageBox.Show("6")
        InstancesWord = ApplyQuery(Query, True)
        'MessageBox.Show("Instances word " + InstancesWord(0))
        InstanceLength = valueCount
        Query = "SELECT completed FROM dbo.instances WHERE UserName = " + "'" + LogIn_Form.User + "' AND WordID LIKE" + "'" + UserLevel + "%'"
        'MessageBox.Show("7")
        InstancesOutcome = ApplyQuery(Query, True)
        'MessageBox.Show("Instances Outcome " + InstancesOutcome(0))
        ThreeTimes = 0
        For n = 0 To AllLength - 1
            value = AllTheWords(n)
            v = 0
            sameCount = 0
            'MessageBox.Show("Value " + value)
            While v < InstanceLength
                'MessageBox.Show("Instance Word " + InstancesWord(v))
                'MessageBox.Show("value " + value)

                If InstancesWord(v) = value And InstancesOutcome(v) = "1" Then
                    sameCount += 1
                    'MessageBox.Show("Same Count")
                    'MessageBox.Show(sameCount)
                End If
                v += 1
            End While
            If sameCount < 3 Then
                'MessageBox.Show("8")
                'MessageBox.Show("Instance to append " + value)
                threeless(length) = value
                length += 1
            End If
        Next
        length = InstanceLength
        Return threeless
    End Function

    Function Incorrect()
        length = 0
        For n = 0 To AllLength - 1
            Query = "SELECT completed FROM dbo.instances WHERE UserName = " + "'" + LogIn_Form.User + "' AND WordID = '" + AllTheWords(n) + "'"
            'MessageBox.Show("9")
            InstanceOrder = ApplyQuery(Query, True)
            InstanceLength = valueCount
            If IsEmpty(InstanceOrder) = False Then
                For v = 0 To InstanceLength - 1
                    LastOne = -1
                    If InstanceOrder(v) = 0 Then
                        LastOne = v
                    End If
                Next
                If LastOne <> -1 Then
                    InstanceCount = 0
                    For k = LastOne To InstanceLength - 1
                        If InstanceOrder(k) = "1" Then
                            InstanceCount += 1
                        End If
                    Next
                    If InstanceCount < 3 Then
                        wrong(length) = AllTheWords(n)
                        length += 1
                    End If
                End If
            End If
        Next
        For n = 0 To length - 1
            'MessageBox.Show("Incorrect " + wrong(n))
        Next
        Return wrong
    End Function

    Function NeverPicked()
        length = 0
        Query = "SELECT WordID FROM dbo.instances WHERE UserName = " + "'" + LogIn_Form.User + "' AND WordID LIKE" + "'" + UserLevel + "%'"
        'MessageBox.Show("11")
        InstanceWords = ApplyQuery(Query, True)
        instanceSize = valueCount
        'MessageBox.Show("Instance size")
        'MessageBox.Show(instanceSize)
        For n = 0 To AllLength - 1
            same = False
            For v = 0 To instanceSize - 1
                'MessageBox.Show("All words " + AllTheWords(n))
                'MessageBox.Show("Instance " + InstancesWord(v))
                If same = True Then
                    same = True
                Else
                    If AllTheWords(n) = InstanceWords(v) Then
                        'MessageBox.Show("Same")
                        same = True
                    End If
                End If
            Next
            If same = False Then
                Never(length) = AllTheWords(n)
                length += 1
            End If
        Next
        Return Never
    End Function

    Function WhichSearch(Incorrect(), NeverPicked(), NotThree())
        IncoEmpty = IsEmpty(Incorrect)
        NeverEmpty = IsEmpty(NeverPicked)
        NotEmpty = IsEmpty(NotThree)
        IncorrectPicked = False
        NeverPickedPicked = False
        NotThreePicked = False
        For t = 0 To NotThreeLength - 1
            'MessageBox.Show("In function not three " + NotThree(t))
        Next
        For b = 0 To IncorrectLength - 1
            'MessageBox.Show("In function incorrect " + Incorrect(b))
        Next
        For g = 0 To NeverPickedLength - 1
            'MessageBox.Show("In function never picked " + NeverPicked(g))
        Next
        If IncoEmpty = True Or NeverEmpty = True Or NotEmpty = True Then
            If IncoEmpty = True And NeverEmpty = True And NotEmpty = True Then
                MessageBox.Show("All empty")
            Else
                If IncoEmpty = True And NeverEmpty = False And NotEmpty = False Then
                    pickValue = random.Next(1, 3)
                    If pickValue = 1 Then
                        NotThreePicked = True
                    Else
                        NeverPickedPicked = True
                    End If
                ElseIf IncoEmpty = False And NeverEmpty = True And NotEmpty = False Then
                    pickValue = random.Next(1, 3)
                    If pickValue = 1 Then
                        NotThreePicked = True
                    Else
                        IncorrectPicked = True
                    End If
                ElseIf IncoEmpty = False And NeverEmpty = False And NotEmpty = True Then
                    pickValue = random.Next(1, 3)
                    If pickValue = 1 Then
                        NeverPickedPicked = True
                    Else
                        IncorrectPicked = True
                    End If
                ElseIf IncoEmpty = True And NeverEmpty = True And NotEmpty = False Then
                    NotThreePicked = True
                ElseIf IncoEmpty = True And NeverEmpty = True And NotEmpty = True Then
                    NeverPickedPicked = True
                ElseIf IncoEmpty = False And NeverEmpty = True And NotEmpty = True Then
                    IncorrectPicked = True
                End If
            End If
        Else
            pickValue = random.Next(1, 6)
            If pickValue = 1 Or 2 Or 3 Then
                IncorrectPicked = True
            ElseIf pickValue = 4 Or 5 Then
                NeverPickedPicked = True
            ElseIf pickValue = 6 Then
                NotThreePicked = True
            End If
        End If
        If IncorrectPicked = True Then
            Return "1"
        ElseIf NeverPickedPicked = True Then
            Return "2"
        ElseIf NotThreePicked = True Then
            Return "3"
        Else
            Return "4"
        End If
    End Function

    Function PickedWord(PickedArray(), Arraylength)
        WordLength = 0
        pickPos = random.Next(0, Arraylength - 1)
        pickVal = PickedArray(pickPos)
        If UserLevel = 1 Then
            Query = "SELECT Word FROM dbo.FirstWord WHERE WordID = " + "'" + pickVal + "'" ''''''
            nextQuery = "SELECT WordLength FROM dbo.FirstWord WHERE WordID = " + "'" + pickVal + "'"
        ElseIf UserLevel = 2 Then
            Query = "SELECT Word FROM dbo.SecondWord WHERE WordID = " + "'" + pickVal + "'"
            nextQuery = "SELECT WordLength FROM dbo.SecondWord WHERE WordID = " + "'" + pickVal + "'"
        ElseIf UserLevel = 3 Then
            Query = "SELECT Word FROM dbo.ThirdWord WHERE WordID = " + "'" + pickVal + "'"
            nextQuery = "SELECT WordLength FROM dbo.ThirdWord WHERE WordID = " + "'" + pickVal + "'"
        End If
        'MessageBox.Show("13")
        WordArray = ApplyQuery(Query, True)
        Word = WordArray(0)
        'MessageBox.Show("14")
        TheWordLength = ApplyQuery(nextQuery, True)
        WordLength = TheWordLength(0)
        Return pickVal
    End Function

    Sub UserInfo()
        Query = "SELECT WordID FROM dbo.instances WHERE UserName =" + "'" + LogIn_Form.User + "' AND completed = '0'"
        'MessageBox.Show("15")
        PastWrong = ApplyQuery(Query, True)
        WrongLength = valueCount
        Query = "SELECT WordID FROM dbo.instances WHERE UserName = " + "'" + LogIn_Form.User + "' AND completed = '1'"
        'MessageBox.Show("16")
        PastCorrect = ApplyQuery(Query, True)
        CorrectLength = valueCount
        If WrongLength = 0 Then
            Wrong_txt.Text = "No values"
        Else
            If WrongLength > 10 Then
                For n = WrongLength - 11 To WrongLength - 1
                    If UserLevel = "1" Then
                        Query = "SELECT Word FROM dbo.FirstWord WHERE WordID = " + "'" + PastWrong(n) + "'"
                    ElseIf UserLevel = "2" Then
                        Query = "SELECT Word FROM dbo.SecondWord WHERE WordID = " + "'" + PastWrong(n) + "'"
                    ElseIf UserLevel = "3" Then
                        Query = "SELECT Word FROM dbo.ThirdWord WHERE WordID = " + "'" + PastWrong(n) + "'"
                    End If
                    'MessageBox.Show("17")
                    FinalWrong = ApplyQuery(Query, True)
                    Wrong_txt.Text += FinalWrong(0)
                    Wrong_txt.Text += Environment.NewLine
                Next
            Else
                For n = 0 To WrongLength - 1
                    getVal = PastWrong(n)
                    If UserLevel = "1" Then
                        Query = "SELECT Word FROM dbo.FirstWord WHERE WordID = " + "'" + PastWrong(n) + "'"
                    ElseIf UserLevel = "2" Then
                        Query = "SELECT Word FROM dbo.SecondWord WHERE WordID = " + "'" + PastWrong(n) + "'"
                    ElseIf UserLevel = "3" Then
                        Query = "SELECT Word FROM dbo.ThirdWord WHERE WordID = " + "'" + PastWrong(n) + "'"
                    End If
                    '"SELECT Word FROM dbo.ThirdWord WHERE WordID = " + getVal
                    'MessageBox.Show("18")
                    FinalWrong = ApplyQuery(Query, True)
                    Wrong_txt.Text += FinalWrong(0)
                    Wrong_txt.Text += Environment.NewLine
                Next
            End If
        End If
        If CorrectLength = 0 Then
            Right_txt.Text = "No results"
        Else
            If CorrectLength > 10 Then
                For n = CorrectLength - 11 To CorrectLength - 1
                    If UserLevel = "1" Then
                        Query = "SELECT Word FROM dbo.FirstWord WHERE WordID = " + "'" + PastCorrect(v) + "'"
                    ElseIf UserLevel = "2" Then
                        Query = "SELECT Word FROM dbo.SecondWord WHERE WordID = " + "'" + PastCorrect(v) + "'"
                    ElseIf UserLevel = "3" Then
                        Query = "SELECT Word FROM dbo.ThirdWord WHERE WordID = " + "'" + PastCorrect(v) + "'"
                    Else
                        MessageBox.Show("Error")
                    End If
                    'MessageBox.Show("19")
                    FinalCorrect = ApplyQuery(Query, True)
                    Right_txt.Text += FinalCorrect(0)
                    Right_txt.Text += Environment.NewLine
                Next
            Else
                For v = 0 To CorrectLength - 1
                    getVal = PastCorrect(v)
                    If UserLevel = "1" Then
                        Query = "SELECT Word FROM dbo.FirstWord WHERE WordID = " + "'" + PastCorrect(v) + "'"
                    ElseIf UserLevel = "2" Then
                        Query = "SELECT Word FROM dbo.SecondWord WHERE WordID = " + "'" + PastCorrect(v) + "'"
                    ElseIf UserLevel = "3" Then
                        Query = "SELECT Word FROM dbo.ThirdWord WHERE WordID = " + "'" + PastCorrect(v) + "'"
                    Else
                        MessageBox.Show("Error")
                    End If
                    'MessageBox.Show("20")
                    FinalCorrect = ApplyQuery(Query, True)
                    Right_txt.Text += FinalCorrect(0)
                    Right_txt.Text += Environment.NewLine
                Next
            End If
        End If
        Level_txt.Text = UserLevel
    End Sub

    Private Sub Dashboard_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataConnection = establishCon()
        Query = "SELECT UserLevel FROM dbo.NewUserInfo WHERE UserName = " + "'" + LogIn_Form.User + "'"
        UserLevelArray = ApplyQuery(Query, True)
        UserLevel = UserLevelArray(0)
        UserInfo()
        LevelUp = AllWords(LogIn_Form.User)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FinalIncorrect = Incorrect()
        IncorrectLength = length
        FinalNotThree = NotThree()
        NotThreeLength = length
        FinalNeverPicked = NeverPicked()
        NeverPickedLength = length
        Which = WhichSearch(FinalIncorrect, FinalNeverPicked, FinalNotThree)
        If Which = "1" Then
            PickedWord(FinalIncorrect, IncorrectLength)
        ElseIf Which = "2" Then
            PickedWord(FinalNeverPicked, NeverPickedLength)
        ElseIf Which = "3" Then
            PickedWord(FinalNotThree, NotThreeLength)
        ElseIf Which = "4" Then
            MessageBox.Show("Issue")
        End If
        MessageBox.Show("Word ", Word)
        Game_Form.Show()
        Me.Close()
    End Sub
End Class