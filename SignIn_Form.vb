Imports System.Data.SqlClient

Public Class SignIn_Form

    Dim DataConnection As New SqlConnection() 'Will be used to call the establishCon function
    Dim Reader As SqlDataReader 'Creates a new reader for SQL queries that require data to be read from the database
    Dim Writer As SqlDataAdapter 'Creates a writer, used to write data to the database
    Dim command 'By not defining the variable's data type, it can be given the data type of the value that it is set to
    Dim valuesArray(100) 'Set to be 100 long because there's no way to know how many values are going to be returned, but it's almost impossible for it to be longer than 100 items.
    Dim passwordArray(SignUp_InitialPassword_txt.Text.Length()) 'Sets the size of the array to be the same as the password's length, this is as long as the array needs to be and no longer, good for efficiency.
    Dim number As Boolean 'Will be set to true if a number is within the password
    Dim capital As Boolean 'Will be set to true if there's a capital value within the password
    Dim NotEmpty As Boolean 'Will be set to true when TextEntered returns true, which will be when all fields are filled
    Dim Query As String 'Will be used to store the next query for the database
    Dim Users As String 'Will give any users with the same name, should only ever be 1, considering all usernames are meant to be different
    Dim passwordLength As Integer 'stores the length of the password
    Dim Correct As Boolean 'Will store the value returned by validData(), which returns true when the username and password is valid
    Dim insertDetails As String = "" 'Used to call the applyQuery() function when the new username, password and user's level is set as a new record within the database's User's database
    Dim FirstVal As String = ""
    Dim empty As Boolean = False
    Dim value As String 'Set to the latest value returned from the database, when reading from the database
    Dim valueCount As Integer = 0 'Set to the length of the array of values returned from the database


    Function establishCon()
        Dim ConnectionData As String = "server=DESKTOP-H3A0KDT\SQLEXPRESS; Database=Logins;Integrated Security=true" 'Local variable that holds the connection data
        Dim Connection As New SqlConnection(ConnectionData) 'establishes a connection in a local varibale, passed then to the global variable DataConnection
        Return Connection
    End Function

    Function applyQuery(Query, returnVal)
        DataConnection.Open() 'Opens the connection with the database
        command = New SqlCommand(Query, DataConnection) 'sets the command that will be sent to the database, with the correct query and connection
        If returnVal = True Then
            Writer = command.ExecuteScalar() 'Executes the command when the query requires writing to the database, hence no values being returned is expected
        Else
            Reader = command.ExecuteReader() 'Executes the command when the query requires reading from the database.
            While Reader.Read() 'Reads the data from the database, loops until all values have been read from the database
                valueCount += 1 'Increments the size of the value array
                value = String.Format("{0}", Reader(0)) 'formats the value so that it is a readable and storable string
                valuesArray.Append(value) 'appends to the values array, which is later returned
            End While
        End If
        DataConnection.Close() 'Closes the connection, regardless of if the query was a read or a write one
        Return valuesArray
    End Function

    Function TextEntered()
        If SignUp_Username_txt.Text <> "" Then 'Three if statements, all check wether their respective field isn't empty
            If SignUp_InitialPassword_txt.Text <> "" Then
                If SignUp_ConfirmPassword_txt.Text <> "" Then
                    Return True 'True is returned when all fields aren't empty
                Else
                    Return False 'False is returned if any field is empty
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Function ValidData()
        number = False
        capital = False
        NotEmpty = TextEntered() 'Calls the TextEntered() function, stores the value in the NotEmpty variable
        If NotEmpty = True Then 'Checks that all fields have been filled
            Query = "SELECT UserName FROM dbo.NewUserInfo WHERE UserName = " + "'" + SignUp_Username_txt.Text + "'" 'Query used to select any usernames the same as the one entered by the user
            Users = applyQuery(Query, True) 'Calls the applyQuery() function, expecting to recieve a result, which will be stored under the Users variable, it's not an array because there will never be more than 1 user with the same username
            If valueCount = 0 Then 'valueCount is the global variable used to determine the size of the array returned from the query, it's intitially set to zero, so if it stays at zero it means the username is unique
                If SignUp_InitialPassword_txt.Text = SignUp_ConfirmPassword_txt.Text Then 'Makes sure the password has been confirmed properly by the user
                    passwordLength = SignUp_InitialPassword_txt.Text.Length() 'stores the length of the password in the variable passwordLength
                    If passwordLength >= 8 Then 'Makes sure the password is longer than eight characters
                        For n = 0 To passwordLength - 1
                            passwordArray.Append(SignUp_InitialPassword_txt.Text(n)) 'This for loop converts the characters within the password string into an array, for later access below
                        Next
                        For v = 0 To passwordLength - 1
                            If capital = True Then
                                capital = True 'Makes sure that once capital has been set to true, it will stay being recognised as true
                            Else
                                If 65 <= Convert.ToByte(passwordArray(v)) <= 90 Then 'Converts the character from the passwordArray into ASCII, then checks it's between 65 and 90, which means it's a capital letter
                                    capital = True
                                End If
                            End If
                            If number = True Then
                                number = True 'Makes sure that once number has been set to true, it will stay being recognised as true
                            Else
                                If 0 <= Convert.ToByte(passwordArray(v)) <= 9 Then ''Converts the character from the passwordArray into ASCII, then checks it's between 0 and 9, which means it's a number
                                    number = True
                                End If
                            End If
                        Next
                        If number = True And capital = True Then 'Checks number and capital are set to true, then returns true
                            Return True
                        Else
                            MessageBox.Show("Password must have at least one number and capital letter") 'Returns false and a message saying the user needs at least one number and capital letter
                            Return False
                        End If
                    Else
                        MessageBox.Show("Password must be at least 8 characters long") 'Returns false and sends a message that tells the user their password must be at least 8 characters long
                        Return False
                    End If
                Else
                    MessageBox.Show("Password and Confirm Password fields are not the same") 'Returns false and sends a message that tells the user their password and confirm password fields are not the same
                    Return False
                End If
            Else
                valueCount = 0 'Sets value count back to zero so the above functionality (relating to the if part of this if, else, endif statement) works properly.
                MessageBox.Show("This username is taken, please pick a different one") 'Returns false and sends a message that tells the user their username is taken and they must pick a different one.
                Return False
            End If
        Else
            MessageBox.Show("All fields must be filled") 'Returns false and sends a message to the user telling them that all fields must be filled
            Return False
        End If
    End Function

    Private Sub SignUp_SignUp_btn_Click(sender As Object, e As EventArgs) Handles SignUp_SignUp_btn.Click
        DataConnection = establishCon() 'establishes connection to the database, storing the connection in the variable DataConnection
        Correct = ValidData() 'calls the ValidData() function, stores the result (true if valid, false if not) in the variable Correct
        If Correct = True Then 'checks wether the username and password are valid, depending on the value of the Correct variable
            Query = "INSERT INTO dbo.NewUserInfo(UserName, Pass, UserLevel) VALUES(" + "'" + SignUp_Username_txt.Text + "', '" + SignUp_InitialPassword_txt.Text + "', '1')" 'Sets the next query to store the user's new username and password, and their level (always 1 for a new user) as a new record in the User_Table table
            applyQuery(Query, False) 'Calls the function applyQuery(), to apply the above query, no value is expected to be returned, hence not stored in a variable
            MessageBox.Show("Success") 'Tells the user they successfully logged in
            Me.Close() 'Closes the sign in form
            LogIn_Form.Show() 'Takes the user to the login form, where they can now log in
        End If
    End Sub

    Private Sub SignUp_Login_btn_Click(sender As Object, e As EventArgs) Handles SignUp_Login_btn.Click
        Me.Close() 'Closes the sign in form
        LogIn_Form.Show() 'Takes the user to the login form
    End Sub
End Class