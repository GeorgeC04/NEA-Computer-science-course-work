Imports System.Data.SqlClient
Public Class LogIn_Form

    'Dim SystemConnection As New SqlConnection() 'Creates a new connection to the database, will be used to call the establishCon function
    Dim Reader As SqlDataReader 'Creates an SQL reader, for queries that require data to be read from the database
    Dim Database As String = "" 'Stores the password for the user from the database
    Public Shared User As String = "" 'Stores the username for the user that they have inputted
    Dim UserPass As String = "" 'Stores the password for the user that they have inputted
    Dim filled As Boolean 'Stores the result returned from the function NotEmpty(), which determines wether all fields have been filled
    Dim UserQuery As String = "" 'Stores the query used to find the password for the username given by the user, from the database
    Dim AttemptCount = 1 'Stores the number of times the user has attempted to log in
    Dim same As Boolean 'Used to store the result of the function ComparePassword(), which will be true if the UserPass and Database are the same.
    Dim SystemConnection 'This variable stores the connection to the database


    Function establishCon() 'Takes in no parameters, used to establish a connection with the database
        Dim connectionData As String = "server=DESKTOP-H3A0KDT\SQLEXPRESS; Database=Logins;Integrated Security=true" 'All connection data needed to connect to the database
        Dim Connection As New SqlConnection(connectionData) 'creates a local variable that stores the connection to the database, this is then returned so that it can be passed to the global variable SystemConnection
        Return Connection
    End Function

    Function ApplyQuery(Query) 'Only needs Query, not return value, upon testing realised that login page never writes to the database
        'SizeCount isn't needed since will only ever get 1 return val at most
        Dim value As String = "" 'Creates a local variable to hold the values returned from the database
        Dim Command As New SqlCommand(Query, SystemConnection) 'Creates a new command, which holds the query and the connection to the databse, this is passed to the database.
        SystemConnection.Open() 'Opens connection to the database
        Reader = Command.ExecuteReader() 'Executes the reader, only reader needed since no data will ever be written to the database in the login form.
        While Reader.Read() 'Creates a while loop to read any data from the database, until there is no data left to be read.
            value = (String.Format("{0}", Reader(0))) 'Formats the value returned from the database into a string
        End While
        SystemConnection.Close() 'Closes the connection to the database
        Return value 'Returns the value from the database
    End Function

    Function NotEmpty()
        If Username_txt.Text <> "" Then
            If Password_txt.Text <> "" Then
                Return True 'Will return True when both Username and password slots are filled
            Else
                Return False
                MessageBox.Show("You need to fill in the form fully") 'Returns this message if either username or password or both slots are empty
            End If
        Else
            Return False
            MessageBox.Show("You need to fill in the form properly") 'Returns this message if either username or password or both slots are empty
        End If
    End Function

    Function ComparePassword()
        If Database = UserPass Then 'Checks that the database and username password are the same.
            Return True 'Returns true when databse and username password are the same
        Else
            Return False 'Returns false when username and database password are not the same
        End If
    End Function

    Private Sub LogIn_btn_Click(sender As Object, e As EventArgs) Handles LogIn_btn.Click
        User = Username_txt.Text 'Sets the variable User to hold the text entered in the username slot by the user
        MessageBox.Show(User)
        UserPass = Password_txt.Text 'Sets the variable UserPass to hold the text entered in the password slot by the user
        SystemConnection = establishCon() 'Runs the establishCon() function and the value returned is stored in the variable SystemConnection
        While AttemptCount <= 3 'Creates a while loop that allows the user three attempts to login
            filled = NotEmpty() 'Runs the function NotEmpty() and then stores the returned value (either positive or negative) in the variable filled
            If filled = True Then 'Checks to see wether both username and password has been filled, if they have been, then checks can be continued
                UserQuery = "SELECT Pass FROM dbo.NewUserInfo WHERE UserName = " + "'" + User + "'" 'Sets the next query to be sent to the databse, will select the stored password in the database from the username given
                Database = ApplyQuery(UserQuery) 'Applies the above query to the databse and stores any value returned under the varibale Database
                MessageBox.Show(Database)
                If Database <> "" Then 'Checks that Database isn't equal to nothing, if it is it means the user's username doesn't exist/
                    same = ComparePassword() 'Runs the ComparePassword() function and returns the result to the 'same' variable
                    If same = True Then 'Deteremines wether the databse and entered password are the same
                        AttemptCount = 4 'Increments the AttemptCount to be larger than the while loop's allowed value, hence stopping the while loop
                        MessageBox.Show("Success") 'Sends a message to the user to let them know they have successfully logged in
                        Me.Hide()
                        Dashboard_Form.Show()
                    Else
                        AttemptCount += 1 'Increments the AttemptCount, hence meaning the user has 1 less go at logging in correctly
                        MessageBox.Show("Your password is wrong") 'Sends a message to the user to let them know their password is wrong
                    End If
                Else
                    MessageBox.Show("Your username is wrong") 'Sends a message to the user to let them know their username is wrong
                End If
            Else
                MessageBox.Show("You need to fill in all positions") 'Sends a message to the user to let them know they must fill in all positions
            End If
        End While

    End Sub

    Private Sub SignIn_btn_Click(sender As Object, e As EventArgs) Handles SignIn_btn.Click
        Me.Hide() 'Hides the current form from view
        SignIn_Form.Show() 'Shows the sign in form
    End Sub


End Class