Imports System.Data
Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO
Imports System.Media
Imports System.Xml
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports DataTable = System.Data.DataTable
Public Class Common
    Sub CreateOutlookAppointment(ByVal subject As String, ByVal location As String)
        Try
            ' Create a new Outlook Application
            Dim outlookApp As New Outlook.Application()

            ' Create a new AppointmentItem
            Dim appointmentItem As Outlook.AppointmentItem = DirectCast(outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem), Outlook.AppointmentItem)

            ' Set the subject and location
            appointmentItem.Subject = subject
            appointmentItem.Location = location

            ' Display the appointment window
            appointmentItem.Display()

        Catch ex As System.Exception
            ' Handle any exceptions that may occur
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Public Function SendEmails(selMLSListingIDs As List(Of String), txtPriceList As String, emailTemplate As String, Optional fileOfferPDF As String = Nothing, Optional fromEmailAccountName As String = "", Optional strPurpose As String = "") As String
        Dim strSubject As String = String.Empty
        Dim strMessage As String = String.Empty
        Dim strAttachment As String = String.Empty
        Dim outlookApp As New Outlook.Application()

        Try
            Dim dtEmail As DataTable = SQLExecuteReader("select * from EmailTemplate")
            Dim rowsofEmail = From row In dtEmail.AsEnumerable()
                              Where row.Field(Of String)("DisplayName") = emailTemplate
                              Select row

            Dim idList As String = String.Join(",", selMLSListingIDs.Select(Function(id) $"'{id}'"))
            Dim dtData = SQLExecuteReader($"select * from Property where MLSListingID IN ({idList})")
            For Each rowData As DataRow In dtData.Rows
                Dim mailItem As Outlook.MailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem)
                For Each rowEmail As DataRow In rowsofEmail
                    strSubject = rowEmail("Subject").ToString()
                    strMessage = rowEmail("Message").ToString()
                    strAttachment = rowEmail("Attachment").ToString()
                Next

                ' Find the account that matches fromEmailAccountName
                Dim xmlDoc As New XmlDocument()
                Dim exeDirectory As String = AppContext.BaseDirectory
                Dim filePath As String = IO.Path.Combine(exeDirectory, "DB\login.xml")
                xmlDoc.Load(filePath)
                Dim DefaultEmailAccount As XmlNode = xmlDoc.SelectSingleNode("/Login/DefaultEmailAccount")
                If DefaultEmailAccount IsNot Nothing Then
                    fromEmailAccountName = DefaultEmailAccount.InnerText
                End If
                If fromEmailAccountName = "" Then
                Else
                    Dim fromAccount As Account = Nothing
                    For Each account As Account In outlookApp.Session.Accounts
                        If account.DisplayName.ToLower() = fromEmailAccountName.ToLower() OrElse account.SmtpAddress.ToLower() = fromEmailAccountName.ToLower() Then
                            fromAccount = account
                            Exit For
                        End If
                    Next
                    If fromAccount Is Nothing Then
                        Throw New System.Exception($"The account '{fromEmailAccountName}' was not found.")
                    End If
                    mailItem.SendUsingAccount = fromAccount
                End If

                Dim SellerEmail As String = rowData("SellerEmail").ToString()
                Dim SellerPhone As String = rowData("SellerPhone").ToString()
                Dim attachmentPath As String
                Dim Seller As String = "Seller"
                If (rowData("Seller").ToString().Length > 0) Then
                    Seller = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowData("Seller").ToString().Split(" ")(0).ToLower())
                End If


                ' Check if fileOfferPDF has a value and use it if available
                attachmentPath = ""
                If Not String.IsNullOrEmpty(fileOfferPDF) Then
                    mailItem.Attachments.Add(fileOfferPDF)
                End If
                If Not String.IsNullOrEmpty(strAttachment) Then
                    If InStr(strAttachment, ":") > 0 Then
                        attachmentPath = strAttachment
                    Else
                        attachmentPath = Path.Combine(AppContext.BaseDirectory, "DB\Attachment\" + strAttachment + "")
                    End If
                    mailItem.Attachments.Add(attachmentPath)
                End If

                'Phone number if available
                If SellerPhone Is String.Empty Then
                    SellerPhone = "000-000-0000"
                End If

                'To, Subject & Content
                mailItem.To = SellerEmail
                mailItem.Subject = strSubject.Replace("<<Address>>", rowData("SavedAddress").ToString()).Replace("<<Price>>", txtPriceList).Replace("<<First Name>>", Seller).Replace("<<Agentcontact>>", SellerPhone)
                mailItem.Body = strMessage.Replace("<<Address>>", rowData("SavedAddress").ToString()).Replace("<<Price>>", txtPriceList).Replace("<<First Name>>", Seller).Replace("<<Agentcontact>>", SellerPhone)
                mailItem.Send()
                ReleaseComObject(mailItem)

                'Update Propery DB
                Dim sqlString As String = ""
                If txtPriceList = "" Then
                    sqlString = "update property set Emailed =@Emailed where MLSListingID IN (@MLSListingID)"
                Else
                    sqlString = "update property set OfferPrice=@OfferPrice,Emailed =@Emailed where MLSListingID IN (@MLSListingID)"
                End If
                Dim SQLitecmd As New SQLiteCommand(sqlString)
                If Not (txtPriceList Is String.Empty) Then
                    SQLitecmd.Parameters.AddWithValue("@OfferPrice", txtPriceList)
                End If
                SQLitecmd.Parameters.AddWithValue("@Emailed", CInt(rowData("Emailed")) + 1)
                SQLitecmd.Parameters.AddWithValue("@MLSListingID", rowData("MLSListingID").ToString())
                SQLExecuteCommand(SQLitecmd)

                'Update Notes DB
                Dim SQLiteinsertcmd As New SQLiteCommand("insert into Notes values (@MLSListingID,@Date,@Notes)")
                SQLiteinsertcmd.Parameters.AddWithValue("@MLSListingID", rowData("MLSListingID").ToString())
                SQLiteinsertcmd.Parameters.AddWithValue("@Date", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                SQLiteinsertcmd.Parameters.AddWithValue("@Notes", strPurpose & "Email Sent: " & emailTemplate)
                SQLExecuteCommand(SQLiteinsertcmd)

            Next
            ReleaseComObject(outlookApp)
            Return "Emails sent successfully."

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show(ex.Message)
            Return ex.Message

        End Try

    End Function

    Public Sub InsertNote(MLSListingID As String, noteText As String, Optional noteDate As String = Nothing)
        If String.IsNullOrEmpty(noteDate) Then
            noteDate = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        End If
        Dim SQLiteinsertcmd As New SQLiteCommand("insert into Notes values (@MLSListingID,@Date,@Notes)")
        SQLiteinsertcmd.Parameters.AddWithValue("@MLSListingID", MLSListingID)
        SQLiteinsertcmd.Parameters.AddWithValue("@Date", noteDate)
        SQLiteinsertcmd.Parameters.AddWithValue("@Notes", noteText)
        SQLExecuteCommand(SQLiteinsertcmd)
    End Sub
    Public Function SendColabEmails(txtEmailId As String, txtEmailUser As String, emailTemplate As String, Optional fromEmailAccountName As String = "") As String
        Try
            Dim outlookApp As New Outlook.Application()
            Dim mailItem As Outlook.MailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem)

            Dim xmlDoc As New XmlDocument()
            Dim exeDirectory As String = AppContext.BaseDirectory
            Dim filePath As String = IO.Path.Combine(exeDirectory, "DB\login.xml")
            xmlDoc.Load(filePath)
            Dim DefaultEmailAccount As XmlNode = xmlDoc.SelectSingleNode("/Login/DefaultEmailAccount")
            If DefaultEmailAccount IsNot Nothing Then
                fromEmailAccountName = DefaultEmailAccount.InnerText
            End If

            ' Find the account that matches fromEmailAccountName
            If fromEmailAccountName = "" Then
            Else
                Dim fromAccount As Account = Nothing
                For Each account As Account In outlookApp.Session.Accounts
                    If account.DisplayName.ToLower() = fromEmailAccountName.ToLower() OrElse account.SmtpAddress.ToLower() = fromEmailAccountName.ToLower() Then
                        fromAccount = account
                        Exit For
                    End If
                Next

                If fromAccount Is Nothing Then
                    Throw New System.Exception($"The account '{fromEmailAccountName}' was not found.")
                End If

                ' Set the From account
                mailItem.SendUsingAccount = fromAccount
            End If

            ' Load email template data
            Dim dtEmail As DataTable = SQLExecuteReader("select * from EmailTemplate")
            Dim rowsofEmail = From row In dtEmail.AsEnumerable()
                              Where row.Field(Of String)("Name") = emailTemplate
                              Select row

            Dim strSubject As String = String.Empty
            Dim strMessage As String = String.Empty
            Dim strAttachment As String = String.Empty

            For Each rowEmail As DataRow In rowsofEmail
                strSubject = rowEmail("Subject").ToString()
                strMessage = rowEmail("Message").ToString()
                strAttachment = rowEmail("Attachment").ToString()
            Next

            ' Attach the file
            Dim attachmentPath As String
            If InStr(strAttachment, ":") > 0 Then
                attachmentPath = strAttachment
            Else
                attachmentPath = Path.Combine(AppContext.BaseDirectory, "DB\Attachment\" + strAttachment + "")
            End If
            If Not String.IsNullOrWhiteSpace(attachmentPath) Then
                mailItem.Attachments.Add(attachmentPath)
            End If

            ' Set mail properties
            mailItem.Subject = strSubject
            mailItem.Body = strMessage.Replace("<<First Name>>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(txtEmailUser))
            mailItem.To = txtEmailId

            ' Send the mail
            mailItem.Send()

            ' Clean up
            ReleaseComObject(mailItem)
            ReleaseComObject(outlookApp)

            Return "Emails sent successfully."
        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Error sending email: " & ex.Message)
            Return ex.Message
        End Try
    End Function

    Public Function GetOutlookEmailAccounts() As List(Of String)
        Dim emailAccounts As New List(Of String)
        Try
            Dim outlookApp As New Outlook.Application()
            Dim accounts As Accounts = outlookApp.Session.Accounts

            For Each account As Account In accounts
                'emailAccounts.Add(account.DisplayName & " (" & account.SmtpAddress & ")")
                emailAccounts.Add(account.SmtpAddress)
            Next

            'emailAccounts.Add("Kanagaraj.pkk@gmail.co)")
            'emailAccounts.Add("Kanagaraj@TEst.co)")
            'emailAccounts.Add("Kanagaraj@example.co)")

            ReleaseComObject(accounts)
            ReleaseComObject(outlookApp)

        Catch ex As System.Exception
            'MessageBox.Show("Error retrieving email accounts: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return emailAccounts
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As System.Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Function WriteDebug(sMsg As String) As String
        Dim logFilePath As String = Path.Combine(AppContext.BaseDirectory, "DB\log.txt")

        Using writer As System.IO.StreamWriter = System.IO.File.AppendText(logFilePath)
            writer.WriteLine($"Crash Time: {DateTime.Now}")
            writer.WriteLine($"Exception Type: {sMsg}")
        End Using
        Return ""
    End Function
    Public Function GenerateLog(exception As System.Exception) As String
        Dim logFilePath As String = Path.Combine(AppContext.BaseDirectory, "DB\log.txt")


        Using writer As System.IO.StreamWriter = System.IO.File.AppendText(logFilePath)
            writer.WriteLine($"Crash Time: {DateTime.Now}")
            writer.WriteLine($"Exception Type: {exception.GetType().FullName}")
            writer.WriteLine($"Exception Message: {exception.Message}")
            writer.WriteLine($"Stack Trace:{Environment.NewLine}{exception.StackTrace}")
            writer.WriteLine("----------")
        End Using
        Return ""
        ' Read and display the log content
        ' Dim logContent As String = File.ReadAllText(logFilePath)

    End Function

    Public Function SQLExecuteReader(sqlcmd As String) As DataTable
        Dim dataTable As New DataTable()
        ' Get the directory of the executable        
        Dim exeDirectory As String = AppContext.BaseDirectory
        Dim dbName As String = IO.Path.Combine(exeDirectory, "DB\crmls.db")
        If Not IO.File.Exists(dbName) Then
            MessageBox.Show("2a. Database file not found: " & dbName)
            Return dataTable
        End If
        'Dim connectionString As String = $"Data Source={dbName};MultipleActiveResultSets=true"
        Dim connectionString As String = $"Data Source={dbName};Version=3;"
        Dim SQLiteConn As New SQLiteConnection(connectionString)
        Dim SQLitecmd As New SQLiteCommand
        Dim SQLiteReader As SQLiteDataReader

        'SQLiteConn.ConnectionString = "Data Source=" + dbName + ";MultipleActiveResultSets=true;Integrated Security=true"        
        SQLiteConn.Open()
        SQLitecmd.Connection = SQLiteConn
        SQLitecmd.CommandText = sqlcmd
        SQLiteReader = SQLitecmd.ExecuteReader()
        dataTable.Load(SQLiteReader)

        SQLiteReader.Close()
        SQLiteConn.Close()
        Return dataTable
    End Function
    Public Function SQLExecuteQuery(sqlcmd As String) As String

        Dim Message As String
        Try

            Dim dbName As String = Path.Combine(AppContext.BaseDirectory, "DB\crmls.db")
            Dim SQLiteConn As New SQLiteConnection
            Dim SQLitecmd As New SQLiteCommand


            SQLiteConn.ConnectionString = "Data Source=" + dbName + ";MultipleActiveResultSets=true;Integrated Security=true"
            SQLiteConn.Open()

            SQLitecmd.Connection = SQLiteConn
            SQLitecmd.CommandText = sqlcmd
            Message = SQLitecmd.ExecuteNonQuery()

            SQLiteConn.Close()

        Catch ex As System.Exception
            Message = ex.Message
        End Try
        Return Message

    End Function

    Public Function SQLExecuteCommand(SQLitecmd As SQLiteCommand) As String
        Dim Message As String
        Try
            Dim dbName As String = Path.Combine(AppContext.BaseDirectory, "DB\crmls.db")
            Dim SQLiteConn As New SQLiteConnection

            SQLiteConn.ConnectionString = "Data Source=" + dbName + ";MultipleActiveResultSets=true;Integrated Security=true"
            SQLiteConn.Open()

            SQLitecmd.Connection = SQLiteConn
            Message = SQLitecmd.ExecuteNonQuery()
            SQLiteConn.Close()
            'Message = "1"
        Catch ex As System.Exception
            Message = ex.Message
        End Try
        Return Message

    End Function
End Class