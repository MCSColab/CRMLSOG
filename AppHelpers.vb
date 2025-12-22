Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO
Imports System.Media
Imports System.Text.RegularExpressions
Imports System.Windows.Controls.Primitives
Imports System.Xml
Imports HtmlAgilityPack
Imports Microsoft.Office.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf
Imports Microsoft.Win32

' =================================================================================================
' AppHelpers File
'
' Description:
'   This file contains all the helper classes that were extracted from MainWindow.vb.
'   Each class has a specific responsibility, following the Single Responsibility Principle.
'   This organization makes the code easier to read, maintain, and debug.
'
' Contents:
'   - PropertyData: A simple class to hold data scraped from web pages.
'   - Item: A simple class for ComboBox items.
'   - UIHelper: Manages common UI tasks like finding controls, filtering grids, and resizing.
'   - ConfigHelper: Handles reading and writing to the XML configuration file.
'   - WebViewHelper: Manages all interactions with WebView2 controls, including navigation,
'     JavaScript execution, and web scraping.
'   - DatabaseHelper: Encapsulates all database operations (SQLite).
'   - OutlookHelper: Manages integration with Microsoft Outlook.
'   - ExcelHelper: Handles exporting data to Excel files.
' =================================================================================================

#Region "Data Transfer Objects"

' DTO to hold property information scraped from the web or read from the DB.
Public Class PropertyData
    Public Property MLSListingID As String = ""
    Public Property ParcelNumber As String = ""
    Public Property SavedAddress As String = ""
    Public Property City As String = ""
    Public Property County As String = ""
    Public Property Seller As String = ""
    Public Property CoSeller As String = ""
    Public Property SellerEmail As String = ""
    Public Property CoSellerEmail As String = ""
    Public Property SellerPhone As String = ""
    Public Property CoSellerPhone As String = ""
    Public Property LADirect As String = ""
    Public Property LACell As String = ""
    Public Property Status As String = ""
    Public Property SellerOffice As String = ""
    Public Property CoSellerOffice As String = ""
    Public Property SellerDreLic As String = ""
    Public Property CoSellerDreLic As String = ""
    Public Property SellerOfficeDreLic As String = ""
    Public Property CoSellerOfficeDreLic As String = ""
    Public Property OfferPrice As String = ""
    Public Property ARVPrice As String = ""
    Public Property TabName As Integer = 0
    Public Property DateAdded As String = ""
End Class

' DTO for simple key-value pairs, used in ComboBoxes.
Public Class cboItem
    Public Property Id As Integer
    Public Property Name As String
End Class

' DTO for email template data
Public Class EmailTemplate
    Public Property DisplayName As String
    Public Property Subject As String
    Public Property Message As String
    Public Property Attachment As String
End Class


#End Region

#Region "UIHelper Class"

Public Class UIHelper
    Public Sub InitializeEmailTemplateComboBox(ByVal cmb As ComboBox)
        Dim data As New ObservableCollection(Of cboItem)()
        data.Add(New cboItem() With {.Id = 1, .Name = "Offer Email"})
        data.Add(New cboItem() With {.Id = 2, .Name = "Active Email"})
        data.Add(New cboItem() With {.Id = 3, .Name = "Pending Email"})
        data.Add(New cboItem() With {.Id = 4, .Name = "Closed Email"})
        cmb.ItemsSource = data
        cmb.DisplayMemberPath = "Name"
    End Sub

    Public Sub InitializeOutlookAccountsComboBox(ByVal cmb As ComboBox, ByVal defaultAccount As String)
        Dim outlookHelper As New OutlookHelper()
        cmb.ItemsSource = outlookHelper.GetOutlookEmailAccounts()
        cmb.SelectedIndex = 0
        If Not String.IsNullOrEmpty(defaultAccount) Then
            cmb.SelectedValue = defaultAccount
        End If
    End Sub

    Public Function GetSelectedGridItems(ByVal grid As DataGrid, ByVal columnName As String) As List(Of String)
        Dim selectedIds As New List(Of String)
        For Each item As Object In grid.Items
            Dim rowView = TryCast(item, DataRowView)
            If rowView IsNot Nothing Then
                Dim chkBox = FindVisualChild(Of CheckBox)(CType(grid.ItemContainerGenerator.ContainerFromItem(item), DataGridRow))
                If chkBox IsNot Nothing AndAlso chkBox.IsChecked = True Then
                    selectedIds.Add(rowView(columnName).ToString())
                End If
            End If
        Next
        Return selectedIds
    End Function

    Public Function GetAllGridItemIds(ByVal grid As DataGrid, ByVal columnName As String) As String
        Dim allIds As New List(Of String)
        If grid.Tag IsNot Nothing Then
            Dim dt = CType(grid.Tag, DataTable)
            For Each row As DataRow In dt.Rows
                allIds.Add(row(columnName).ToString())
            Next
        End If
        Return String.Join(",", allIds)
    End Function

    Public Sub FilterDataGrid(ByVal grid As DataGrid, ByVal searchText As String, ParamArray columnNames As String())
        Dim originalDataTable = TryCast(grid.Tag, DataTable)
        If originalDataTable Is Nothing Then Return

        Dim filteredView As DataView = originalDataTable.DefaultView
        If String.IsNullOrWhiteSpace(searchText) Then
            filteredView.RowFilter = String.Empty
        Else
            Dim filterParts = columnNames.Select(Function(col) $"{col} LIKE '%{searchText.Replace("'", "''")}%'")
            filteredView.RowFilter = String.Join(" OR ", filterParts)
        End If
        grid.ItemsSource = filteredView
    End Sub

    Public Sub FormatDataGridColumns(ByVal e As DataGridAutoGeneratingColumnEventArgs)
        If e.PropertyName = "DateAdded" Then
            Dim column = TryCast(e.Column, DataGridTextColumn)
            If column IsNot Nothing Then
                column.Binding.StringFormat = "dd-MMM-yy"
                e.Column.Width = New DataGridLength(1, DataGridLengthUnitType.Star)
            End If
        End If
    End Sub

    Public Sub OpenGoogleSearch(ByVal address As String)
        Dim searchUrl = "https://www.google.com/search?q=" & Uri.EscapeDataString(address)
        Dim psi As New ProcessStartInfo With {
            .FileName = searchUrl,
            .UseShellExecute = True
        }
        Process.Start(psi)
    End Sub

    Public Sub PopulateEmailFields(ByVal parentGrid As Grid, ByVal templates As List(Of EmailTemplate))
        Dim cnt As Integer = 1
        For Each template In templates
            FindTextBoxByName(parentGrid, "txtEmailSubject" & cnt).Text = template.Subject
            FindTextBoxByName(parentGrid, "txtEmailMessage" & cnt).Text = template.Message
            FindLabelByName(parentGrid, "lblEmailAttach" & cnt).Content = template.Attachment
            cnt += 1
        Next
    End Sub

    Public Sub PopulateSmsFields(ByVal txtBox As TextBox, ByVal message As String)
        txtBox.Text = message
    End Sub

    Public Function GetEmailTemplatesFromUI(ByVal parentGrid As Grid) As List(Of EmailTemplate)
        Dim templates As New List(Of EmailTemplate)
        For i = 1 To 3 ' Assuming 3 sets of email controls
            templates.Add(New EmailTemplate With {
                .DisplayName = FindLabelByName(parentGrid, "lblMessage" & i).Content.ToString(),
                .Subject = FindTextBoxByName(parentGrid, "txtEmailSubject" & i).Text,
                .Message = FindTextBoxByName(parentGrid, "txtEmailMessage" & i).Text,
                .Attachment = FindLabelByName(parentGrid, "lblEmailAttach" & i).Content.ToString()
            })
        Next
        Return templates
    End Function

    Public Function GetSmsTemplateFromUI(ByVal txtBox As TextBox) As String
        Return txtBox.Text
    End Function

    Public Sub HandleAttachmentBrowse(ByVal targetLabel As Label)
        Dim openFileDialog As New OpenFileDialog() With {
            .Filter = "All Files (*.*)|*.*"
        }
        If openFileDialog.ShowDialog() = True Then
            targetLabel.Content = openFileDialog.FileName
            Dim savePath = Path.Combine(AppContext.BaseDirectory, "DB\Attachment", openFileDialog.SafeFileName)
            If savePath <> openFileDialog.FileName Then
                File.Copy(openFileDialog.FileName, savePath, True)
            End If
        End If
    End Sub

    Public Function GetActiveWebView(ByVal tabControl As TabControl, ByVal mainView As WebView2, ByVal compView As WebView2) As WebView2
        If tabControl.SelectedIndex = 0 Then Return mainView
        If tabControl.SelectedIndex = 1 Then Return compView
        Return Nothing
    End Function

#Region "Visual Tree Helpers"
    Public Function FindVisualChild(Of T As DependencyObject)(parent As DependencyObject) As T
        If parent Is Nothing Then Return Nothing
        Dim count = VisualTreeHelper.GetChildrenCount(parent)
        For i = 0 To count - 1
            Dim child = VisualTreeHelper.GetChild(parent, i)
            If child IsNot Nothing AndAlso TypeOf child Is T Then
                Return CType(child, T)
            Else
                Dim childOfChild = FindVisualChild(Of T)(child)
                If childOfChild IsNot Nothing Then
                    Return childOfChild
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Function FindTextBoxByName(parent As Grid, name As String) As TextBox
        Return parent.Children.OfType(Of TextBox)().FirstOrDefault(Function(tb) tb.Name = name)
    End Function

    Private Function FindLabelByName(parent As Grid, name As String) As Label
        Return parent.Children.OfType(Of Label)().FirstOrDefault(Function(lbl) lbl.Name = name)
    End Function
#End Region

End Class

#End Region

#Region "ConfigHelper Class"

Public Class ConfigHelper
    Private ReadOnly _filePath As String

    Public Sub New()
        _filePath = Path.Combine(AppContext.BaseDirectory, "DB\login.xml")
    End Sub

    Private Function ReadXmlDoc() As XmlDocument
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(_filePath)
            Return xmlDoc
        Catch ex As Exception
            MessageBox.Show("Error loading configuration file: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function WriteXmlDoc(xmlDoc As XmlDocument) As Boolean
        Try
            xmlDoc.Save(_filePath)
            Return True
        Catch ex As Exception
            MessageBox.Show("Error saving configuration file: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function GetLoginCredentials(site As String) As (UserName As String, Password As String)
        Dim xmlDoc = ReadXmlDoc()
        If xmlDoc Is Nothing Then Return (String.Empty, String.Empty)

        Dim userNode = xmlDoc.SelectSingleNode($"/Login/{site.ToUpper()}/UserName")
        Dim passNode = xmlDoc.SelectSingleNode($"/Login/{site.ToUpper()}/Password")

        If userNode IsNot Nothing AndAlso passNode IsNot Nothing Then
            Return (userNode.InnerText, passNode.InnerText)
        Else
            MessageBox.Show($"Username or Password missing in login.xml for {site}")
            Return (String.Empty, String.Empty)
        End If
    End Function

    Public Function GetMlsStatusLists() As Dictionary(Of String, String)
        Dim statusLists As New Dictionary(Of String, String)
        Dim xmlDoc = ReadXmlDoc()
        If xmlDoc Is Nothing Then Return statusLists

        statusLists.Add("Active", xmlDoc.SelectSingleNode("/Login/MLSStatus/Active")?.InnerText)
        statusLists.Add("Pending", xmlDoc.SelectSingleNode("/Login/MLSStatus/Pending")?.InnerText)
        statusLists.Add("Sold", xmlDoc.SelectSingleNode("/Login/MLSStatus/Sold")?.InnerText)
        statusLists.Add("Hold", xmlDoc.SelectSingleNode("/Login/MLSStatus/Hold")?.InnerText)
        statusLists.Add("Other", xmlDoc.SelectSingleNode("/Login/MLSStatus/Other")?.InnerText)

        Return statusLists
    End Function

    Public Sub SaveRadius(radius As String)
        Dim xmlDoc = ReadXmlDoc()
        If xmlDoc Is Nothing Then Return
        Dim milesNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Miles")
        If milesNode IsNot Nothing Then
            milesNode.InnerText = radius.Trim()
            WriteXmlDoc(xmlDoc)
        Else
            MessageBox.Show("Miles tag not found in login.xml.")
        End If
    End Sub

    Public Function GetDefaultEmailAccount() As String
        Return ReadXmlDoc()?.SelectSingleNode("/Login/DefaultEmailAccount")?.InnerText
    End Function

    Public Sub SetDefaultEmailAccount(accountName As String)
        Dim xmlDoc = ReadXmlDoc()
        If xmlDoc Is Nothing Then Return
        Dim node = xmlDoc.SelectSingleNode("/Login/DefaultEmailAccount")
        If node IsNot Nothing Then
            node.InnerText = accountName
            WriteXmlDoc(xmlDoc)
        End If
    End Sub
End Class

#End Region

#Region "WebViewHelper Class"

Public Class WebViewHelper
    Public Property MLSListingMain As String = String.Empty
    Public Property MLSListingSub As String = String.Empty
    Private _bCodeProcessing As Boolean = False

    Public Function IsStillProcessing() As Boolean
        Return _bCodeProcessing
    End Function

    Public Sub SearchAddress(address As String, webView As WebView2)
        If String.IsNullOrWhiteSpace(address) Then
            MessageBox.Show("Enter Address to proceed")
            Return
        End If
        Dim url = $"https://matrix.crmls.org/matrix/search/residential/detail?address={Uri.EscapeDataString(address)}"
        webView.Source = New Uri(url)
    End Sub

    Public Async Function HandleMainDOMContentLoadedAsync(coreWv As CoreWebView2, searchAddress As String, codeProcessingFlag As Boolean) As Task(Of Boolean)
        Dim url = coreWv.Source.ToString().ToLower()
        _bCodeProcessing = codeProcessingFlag ' Sync state

        If _bCodeProcessing Then
            If url.Contains("signin.crmls.org") Then
                Await SiteLoginAsync(coreWv, "CRMLS")
            ElseIf url.Contains("matrix.crmls.org") AndAlso Not String.IsNullOrEmpty(MLSListingMain) Then
                Await coreWv.ExecuteScriptAsync($"document.getElementById('ctl02_m_ucSpeedBar_m_tbSpeedBar').value = '{MLSListingMain}';")
                Await coreWv.ExecuteScriptAsync("document.getElementById('ctl02_m_ucSpeedBar_m_lnkGo').click();")
                MLSListingMain = String.Empty
            End If
        ElseIf url.Contains("matrix.crmls.org/matrix/search/residential/detail") Then
            ' This is the complex automation sequence for comps
            Await coreWv.ExecuteScriptAsync($"document.getElementById('Fm2_Ctrl12_TB').value = '{searchAddress}'; MapSearchJs.geocode();")
            ' ... and so on for the rest of the automation logic (SetRadius, clickCheckbox, etc.)
        End If

        _bCodeProcessing = False ' Reset after processing
        Return _bCodeProcessing
    End Function

    Public Async Function HandleStatusDOMContentLoadedAsync(coreWv As CoreWebView2, mlsIdList As String) As Task(Of Boolean)
        Dim url = coreWv.Source.ToString().ToLower()
        Dim dbHelper As New DatabaseHelper()

        If url.Contains("signin.crmls.org") Then
            Await SiteLoginAsync(coreWv, "CRMLS")
        ElseIf url.Contains("matrix.crmls.org") AndAlso Not String.IsNullOrEmpty(mlsIdList) Then
            Await coreWv.ExecuteScriptAsync($"document.getElementById('ctl02_m_ucSpeedBar_m_tbSpeedBar').value = '{mlsIdList}';")
            Await coreWv.ExecuteScriptAsync("document.getElementById('ctl02_m_ucSpeedBar_m_lnkGo').click();")
        ElseIf url.Contains("results.aspx") AndAlso Not String.IsNullOrEmpty(mlsIdList) Then
            Dim html = Await GetHtmlAsync(coreWv)
            Dim dt = ParseStatusTable(html)
            dbHelper.UpdatePropertyStatuses(dt)
            MessageBox.Show("Statuses updated successfully.")
            Return True ' Signal that an update occurred
        ElseIf url.Contains("app.privy.pro/users/sign_in") Then
            Await SiteLoginAsync(coreWv, "PRIVY")
        End If
        Return False
    End Function

    Private Async Function SiteLoginAsync(coreWv As CoreWebView2, site As String) As Task
        Dim config = New ConfigHelper()
        Dim creds = config.GetLoginCredentials(site)
        If String.IsNullOrEmpty(creds.UserName) Then Return

        If site = "CRMLS" Then
            Await coreWv.ExecuteScriptAsync($"document.getElementById('Username').value = '{creds.UserName}';")
            Await coreWv.ExecuteScriptAsync($"document.getElementById('Password').value = '{creds.Password}';")
            Await coreWv.ExecuteScriptAsync("document.querySelector(""button[name='button']"").click();")
        ElseIf site = "PRIVY" Then
            Await coreWv.ExecuteScriptAsync($"document.getElementById('user_email').value = '{creds.UserName}';")
            Await coreWv.ExecuteScriptAsync($"document.getElementById('user_password').value = '{creds.Password}';")
            Await coreWv.ExecuteScriptAsync("document.getElementById('login_button').click();")
        End If
    End Function

    Public Async Function ExtractPropertyDataAsync(webView As WebView2) As Task(Of PropertyData)
        Dim html = Await GetHtmlAsync(webView.CoreWebView2)
        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)

        If doc.GetElementbyId("m_pnlDisplay") Is Nothing Then Return Nothing

        Dim prop As New PropertyData()
        prop.MLSListingID = GetInnerText(doc.DocumentNode.SelectSingleNode("//span[text()='LISTING ID: ']"), "nn")
        prop.SavedAddress = GetInnerText(doc.DocumentNode.SelectSingleNode("//input[@onclick='Dpy.onCheck(this,event)']"), "pn")?.Replace(" ", " ")
        prop.City = If(prop.SavedAddress?.Split(","c).Length > 1, prop.SavedAddress.Split(","c)(1).Trim(), "")
        prop.Seller = GetInnerText(doc.DocumentNode.SelectSingleNode("//span[contains(text(),'LA:')]"), "nnnn")
        prop.SellerEmail = GetInnerText(doc.DocumentNode.SelectSingleNode("//span[contains(text(),'Offers Email:')]"), "nnf")
        If String.IsNullOrEmpty(prop.SellerEmail) Then
            prop.SellerEmail = GetInnerText(doc.DocumentNode.SelectSingleNode("//span[contains(text(),'LA EMAIL:')]"), "nn")
        End If
        ' ... continue for all other properties ...
        Return prop
    End Function

    Public Async Function GetAddressFromActivePageAsync(webView As WebView2) As Task(Of String)
        Dim html = Await GetHtmlAsync(webView.CoreWebView2)
        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)
        Return GetInnerText(doc.DocumentNode.SelectSingleNode("//input[@onclick='Dpy.onCheck(this,event)']"), "pn")?.Replace(" ", " ")
    End Function

    Public Async Function GetMlsIdFromActivePageAsync(webView As WebView2) As Task(Of String)
        Dim html = Await GetHtmlAsync(webView.CoreWebView2)
        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)
        Return GetInnerText(doc.DocumentNode.SelectSingleNode("//span[text()='LISTING ID: ']"), "nn")
    End Function

    Private Async Function GetHtmlAsync(coreWv As CoreWebView2) As Task(Of String)
        Dim rawHtml = Await coreWv.ExecuteScriptAsync("document.documentElement.outerHTML")
        Return Regex.Unescape(rawHtml)
    End Function

    Private Function ParseStatusTable(html As String) As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("Listing Id")
        dt.Columns.Add("S") ' Status column
        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)
        Dim table = doc.DocumentNode.SelectNodes("//table[contains(@class, 'displayGrid')]")?.FirstOrDefault()
        If table IsNot Nothing Then
            For Each row As HtmlNode In table.SelectNodes(".//tbody/tr")
                Dim newRow = dt.NewRow()
                newRow("Listing Id") = row.SelectSingleNode("./td[contains(@class,'col_MLNumber')]")?.InnerText.Trim()
                newRow("S") = row.SelectSingleNode("./td[contains(@class,'col_Status')]")?.InnerText.Trim()
                dt.Rows.Add(newRow)
            Next
        End If
        Return dt
    End Function

    Private Function GetInnerText(hNode As HtmlNode, pattern As String) As String
        pattern = pattern.ToUpper()
        Try
            For Each tstep In pattern
                If hNode Is Nothing Then Return String.Empty
                Select Case tstep
                    Case "N"
                        hNode = hNode.NextSibling
                    Case "F"
                        hNode = hNode.FirstChild
                    Case "P"
                        hNode = hNode.ParentNode
                    Case Else
                        Return String.Empty
                End Select
            Next
            Return hNode?.InnerText.Trim()
        Catch
            Return String.Empty
        End Try
    End Function
End Class

#End Region

#Region "DatabaseHelper Class"

Public Class DatabaseHelper
    ' This class inherits or uses the existing fxCommon class for actual DB execution.
    ' For this example, I'm wrapping the logic.
    Private ReadOnly fxCommon As New Common()
    Private ReadOnly _configHelper As New ConfigHelper()

    Public Function LoadProperties(statusFilter As String, tabName As String) As DataTable
        Dim sFieldList = "MLSListingID,SavedAddress,City,Seller,SellerEmail,SellerPhone,CoSeller,LACell,LADirect,Emailed,OfferPrice,ARV,Status,DateAdded"
        Dim sCondition As String = ""

        If Not String.IsNullOrEmpty(statusFilter) Then
            Dim statusLists = _configHelper.GetMlsStatusLists()
            If statusLists.ContainsKey(statusFilter) AndAlso statusLists(statusFilter) IsNot Nothing Then
                Dim convertToQuotedString = Function(s As String) String.Join(",", s.Split(","c).Select(Function(val) $"'{val.Trim()}'"))
                Dim conditionList = convertToQuotedString(statusLists(statusFilter))
                sCondition = $" AND Status IN ({conditionList})"
            End If
        End If

        Dim sQuery = $"Select {sFieldList} from Property WHERE (TabName = '{tabName}' OR TabName = '' OR TabName is NULL) {sCondition} Order by Status"
        Return fxCommon.SQLExecuteReader(sQuery)
    End Function

    Public Function GetNotes(mlsId As String) As DataTable
        Return fxCommon.SQLExecuteReader($"select Date, Notes from Notes where MLSListingID='{mlsId}' ORDER BY rowid DESC")
    End Function

    Public Sub SaveNote(mlsId As String, noteText As String)
        Dim cmd As New SQLiteCommand("insert into Notes values (@MLSListingID, @Datetime, @Notes)")
        cmd.Parameters.AddWithValue("@MLSListingID", mlsId)
        cmd.Parameters.AddWithValue("@Datetime", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        cmd.Parameters.AddWithValue("@Notes", noteText)
        fxCommon.SQLExecuteCommand(cmd)
    End Sub

    Public Function SaveProperty(prop As PropertyData) As String
        Dim dt = fxCommon.SQLExecuteReader($"select * from Property where MLSListingID='{prop.MLSListingID}'")
        If dt.Rows.Count > 0 Then
            Return $"Property [{prop.SavedAddress}] - Already Exists."
        End If

        Dim cmd As New SQLiteCommand("INSERT INTO Property (MLSListingID, SavedAddress, City, Seller, SellerEmail, OfferPrice, Status, ARV, DateAdded) VALUES (@MLSListingID, @SavedAddress, @City, @Seller, @SellerEmail, @OfferPrice, @Status, @ARV, @DateAdded)")
        cmd.Parameters.AddWithValue("@MLSListingID", prop.MLSListingID)
        cmd.Parameters.AddWithValue("@SavedAddress", prop.SavedAddress)
        cmd.Parameters.AddWithValue("@City", prop.City)
        cmd.Parameters.AddWithValue("@Seller", prop.Seller)
        cmd.Parameters.AddWithValue("@SellerEmail", prop.SellerEmail)
        cmd.Parameters.AddWithValue("@OfferPrice", prop.OfferPrice)
        cmd.Parameters.AddWithValue("@Status", prop.Status)
        cmd.Parameters.AddWithValue("@ARV", prop.ARVPrice)
        cmd.Parameters.AddWithValue("@DateAdded", DateTime.Now.ToString("yyyy-MM-dd"))
        ' ... add all other parameters ...

        Dim result = fxCommon.SQLExecuteCommand(cmd)
        Return If(result = "1", $"Property [{prop.SavedAddress}] - Added successfully.", "Record not saved.")
    End Function

    Public Sub UpdatePropertyStatuses(dt As DataTable)
        For Each row As DataRow In dt.Rows
            Dim cmd As New SQLiteCommand("update property set Status=@Status where MLSListingID=@MLSListingID")
            cmd.Parameters.AddWithValue("@Status", row("S").ToString())
            cmd.Parameters.AddWithValue("@MLSListingID", row("Listing Id").ToString())
            fxCommon.SQLExecuteCommand(cmd)
        Next
    End Sub

    Public Sub MoveProperties(idList As List(Of String), targetTabName As String)
        If idList.Count = 0 Then Return
        Dim ids = String.Join(",", idList.Select(Function(id) $"'{id}'"))
        fxCommon.SQLExecuteQuery($"UPDATE Property SET TabName = '{targetTabName}' WHERE MLSListingID IN ({ids})")
    End Sub

    Public Sub DeleteProperties(idList As List(Of String))
        If idList.Count = 0 Then Return
        Dim ids = String.Join(",", idList.Select(Function(id) $"'{id}'"))
        fxCommon.SQLExecuteQuery($"DELETE FROM Property WHERE MLSListingID IN ({ids})")
        fxCommon.SQLExecuteQuery($"DELETE FROM Notes WHERE MLSListingID IN ({ids})")
    End Sub

    Public Function GetSmsText(address As String, seller As String) As String
        Dim dtSMS = fxCommon.SQLExecuteReader("Select * from EmailTemplate where Name like '%SMS%'")
        If dtSMS.Rows.Count > 0 Then
            Dim sellerFirstName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(seller.Split(" "c)(0).ToLower())
            Return dtSMS.Rows(0)("Message").ToString().Replace("<<Address>>", address).Replace("<<First Name>>", sellerFirstName)
        End If
        Return ""
    End Function

    Public Function GetEmailTemplates(emailType As String) As List(Of EmailTemplate)
        Dim templates As New List(Of EmailTemplate)
        Dim dt = fxCommon.SQLExecuteReader($"Select * from EmailTemplate where Name like '%{emailType}%'")
        For Each row As DataRow In dt.Rows
            templates.Add(New EmailTemplate With {
                .DisplayName = row("DisplayName").ToString(),
                .Subject = row("Subject").ToString(),
                .Message = row("Message").ToString(),
                .Attachment = row("Attachment").ToString()
            })
        Next
        Return templates
    End Function

    Public Function GetSmsTemplate() As String
        Dim dt = fxCommon.SQLExecuteReader("Select * from EmailTemplate where Name like '%SMS%'")
        Return If(dt.Rows.Count > 0, dt.Rows(0)("Message").ToString(), "")
    End Function

    Public Sub SaveEmailTemplates(templates As List(Of EmailTemplate))
        For Each t In templates
            Dim cmd As New SQLiteCommand("Update EmailTemplate Set Subject=@Subject, Message=@Message, Attachment=@Attachment where DisplayName = @DisplayName")
            cmd.Parameters.AddWithValue("@Subject", t.Subject)
            cmd.Parameters.AddWithValue("@Message", t.Message)
            cmd.Parameters.AddWithValue("@Attachment", t.Attachment)
            cmd.Parameters.AddWithValue("@DisplayName", t.DisplayName)
            fxCommon.SQLExecuteCommand(cmd)
        Next
    End Sub

    Public Sub SaveSmsTemplate(message As String)
        Dim cmd As New SQLiteCommand("Update EmailTemplate Set Message=@Message where Name = 'SMS'")
        cmd.Parameters.AddWithValue("@Message", message)
        fxCommon.SQLExecuteCommand(cmd)
    End Sub

    Public Sub GenerateLog(ex As Exception)
        ' Placeholder for actual logging logic, which might be in fxCommon
    End Sub

    Private Sub FindListPriceMLSID(ByVal targetMlsId As String)
        ' Call the function from your database helper to get the details.
        Dim propertyDetails As DataTable = GetMLSDetails(targetMlsId)

        ' Check if the function returned any data.
        If propertyDetails IsNot Nothing AndAlso propertyDetails.Rows.Count > 0 Then
            Dim propertyRow As DataRow = propertyDetails.Rows(0)
            Dim address As String = propertyRow("SavedAddress").ToString()
            Dim seller As String = propertyRow("Seller").ToString()
            Dim status As String = propertyRow("Status").ToString()

            ' Display or use the retrieved information.
            MessageBox.Show($"Found Property!{vbCrLf}" &
                        $"Address: {address}{vbCrLf}" &
                        $"Seller: {seller}{vbCrLf}" &
                        $"Status: {status}", "MLS ID Found")

        Else
            ' --- Failure: The MLS ID was not found in the table. ---
            MessageBox.Show($"The MLS ID '{targetMlsId}' was not found in the database.", "Not Found")

        End If
    End Sub
    ''' Retrieves all details for a single property from the database using its MLS ID.    
    ''' <param name="mlsId">The MLS Listing ID to search for.</param>
    ''' <returns>A DataTable containing the property details if found, otherwise an empty DataTable.</returns>
    Public Function GetMLSDetails(ByVal sMLSid As String) As System.Data.DataTable
        Dim fieldList As String = "*"
        Dim sQuery As String = $"SELECT {fieldList} FROM Property WHERE MLSListingID = {sMLSid}"
        'Execute the query and return the resulting DataTable.
        Dim dDataTable As System.Data.DataTable
        dDataTable = fxCommon.SQLExecuteReader(sQuery)
        If dDataTable.Rows.Count > 0 Then
            Return dDataTable
        Else
            Return New System.Data.DataTable()
        End If
    End Function
End Class

#End Region

#Region "OutlookHelper Class"

Public Class OutlookHelper
    Private ReadOnly fxCommon As New Common() ' Assuming fxCommon handles the interop

    Public Sub EnsureOutlookIsRunning()
        If Not IsOutlookRunning() Then
            StartOutlook()
        End If
    End Sub

    Private Function IsOutlookRunning() As Boolean
        Return Process.GetProcessesByName("OUTLOOK").Length > 0
    End Function

    Private Sub StartOutlook()
        Try
            Process.Start(Path.Combine(AppContext.BaseDirectory, "DB\outlook.bat"))
        Catch ex As Exception
            MessageBox.Show("Unable to start Outlook: " & ex.Message)
        End Try
    End Sub

    Public Function GetOutlookEmailAccounts() As List(Of String)
        ' This logic is assumed to be in fxCommon
        Return fxCommon.GetOutlookEmailAccounts()
    End Function

    Public Function SendPropertyEmails(idList As List(Of String), offerPrice As String, templateName As String) As String
        If idList.Count = 0 Then
            Return "No items selected to email."
        End If
        ' This logic is assumed to be in fxCommon
        Return fxCommon.SendEmails(idList, offerPrice, templateName)
    End Function

    Public Sub CreateAppointment(subject As String, location As String)
        ' This logic is assumed to be in fxCommon
        fxCommon.CreateOutlookAppointment(subject, location)
    End Sub
End Class

#End Region

#Region "ExcelHelper Class"

Public Class ExcelHelper
    Public Sub ExportToExcel(dataTable As DataTable)
        If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
            MessageBox.Show("No data to export.")
            Return
        End If

        Dim saveFileDialog As New SaveFileDialog() With {
            .Filter = "Excel Files|*.xlsx|All Files|*.*",
            .FileName = "Export.xlsx"
        }

        If saveFileDialog.ShowDialog() = True Then
            Try
                ' Using ClosedXML is generally safer and doesn't require Excel to be installed.
                ' The original code used Interop, which is shown here for compatibility.
                ExportWithInterop(dataTable, saveFileDialog.FileName)
                MessageBox.Show("Exported successfully!", "Success")
            Catch ex As Exception
                MessageBox.Show("Failed to export data: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub ExportWithInterop(dt As DataTable, filePath As String)
        Dim oExcel As Excel.Application = Nothing
        Dim oBook As Excel.Workbook = Nothing
        Try
            oExcel = New Excel.Application()
            oBook = oExcel.Workbooks.Add()
            Dim oSheet = CType(oBook.Worksheets(1), Excel.Worksheet)

            ' Headers
            For i = 0 To dt.Columns.Count - 1
                oSheet.Cells(1, i + 1) = dt.Columns(i).ColumnName
            Next

            ' Data
            For r = 0 To dt.Rows.Count - 1
                For c = 0 To dt.Columns.Count - 1
                    oSheet.Cells(r + 2, c + 1) = dt.Rows(r)(c).ToString()
                Next
            Next

            oSheet.Columns.AutoFit()
            oExcel.DisplayAlerts = False
            oBook.SaveAs(filePath)
            oBook.Close(False)
            oExcel.Quit()
        Finally
            ReleaseComObject(oBook)
            ReleaseComObject(oExcel)
        End Try
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class

#End Region
