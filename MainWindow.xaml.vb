Imports System.Collections.ObjectModel
Imports System.ComponentModel.DataAnnotations
Imports System.Data
Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO
Imports System.Media
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Xml
Imports ClosedXML.Excel
Imports HtmlAgilityPack
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Web.WebView2
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf
Imports Microsoft.Win32
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'Imports System.Media ' Assuming SystemSounds is still used for error messages
' Assuming CoreWebView2 is available in your context

Class MainWindow
    Private fxCommon As New Common()
    Private bCodeProcessing As Boolean = False
    Dim dataTable As System.Data.DataTable
    Dim dtEmail As System.Data.DataTable
    Dim message As String
    Dim MLSListingMain As String = String.Empty
    Dim MLSListingSub As String = String.Empty
    Dim MLSListingStatus As String = String.Empty
    Private tabChangeInitiated As Boolean = False
    Private tabB4Change As String = ""
    Private sActiveClassName As String = ""
    Public DragTabIndex As Integer
    Public DragButton As System.Windows.Controls.Button
    Private isAutofillEnabled As Boolean = True ' Session-only variable, defaults to ON.
    Private ReadOnly _settingsFilePath As String = Path.Combine(AppContext.BaseDirectory, "DB\login.xml")
    Private isSidebarHidden As Boolean = False
    Private origLeftSidebarWidth As Double
    Private origRightSidebarWidth As Double
    Private newWindowWidth As Double
    Private newWindowHeight As Double
    Public FlagAutoStatusChangeEmailsEnabled As Boolean = False
    Private LastOfferPrice As String = ""
    Private CurrentPropertyAddress As String = ""
    Private LastExtractedAddress As String = ""
    Private LastSentAddressToOfferGun As String = ""
    Private CurrentListPrice As String
    Private CurrentYearBuilt As String
    Private CurrentParcelNumber As String
    Private CurrentCounty As String

    Private CurrentLAName As String
    Private CurrentLACell As String
    Private CurrentLAEmail As String
    Private Currentbuildyr As String
    Private currentLO As String


    Private Sub CurrentDomain_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Dim exception As Exception = CType(e.ExceptionObject, Exception)
        fxCommon.GenerateLog(exception)

        ' Handle the exception here, 'exception' contains the exception object
        SystemSounds.Exclamation.Play()
        MessageBox.Show($"An unhandled exception occurred: {exception.Message}", "Unhandled Exception")
    End Sub
    Private Sub Window_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Escape Then
            Me.WindowState = WindowState.Minimized
            'e.Handled = True ' Mark the event as handled to prevent further processing
        End If

        'Check for Ctrl + B to toggle the sidebar
        If e.Key = Key.B AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
            ToggleSidebarVisibility()
            e.Handled = True ' Mark the event as handled to prevent further processing
        End If

    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        If Not IsOutlookRunning() Then
            StartOutlook()
        End If
        Me.WindowState = WindowState.Maximized
        InitializeComboBoxData()
        AddHandler WebViewMain.CoreWebView2InitializationCompleted, AddressOf WebView_CoreWebView2InitializationCompleted
        AddHandler WebViewComp.CoreWebView2InitializationCompleted, AddressOf WebView_CoreWebView2InitializationCompleted
        'AddHandler WebViewStatus.CoreWebView2InitializationCompleted, AddressOf WebViewStatus_CoreWebView2InitializationCompleted
        AddHandler WebViewPrivy.CoreWebView2InitializationCompleted, AddressOf WebViewStatus_CoreWebView2InitializationCompleted
        'AddHandler WebViewPipe.CoreWebView2InitializationCompleted, AddressOf WebViewStatus_CoreWebView2InitializationCompleted
        AddHandler WebViewog.CoreWebView2InitializationCompleted, AddressOf WebViewStatus_CoreWebView2InitializationCompleted
        AddHandler WebViewog2.CoreWebView2InitializationCompleted, AddressOf WebViewStatus_CoreWebView2InitializationCompleted
        AddHandler txtPrice.TextChanged, AddressOf TxtPrice_TextChanged

        bCodeProcessing = True
        chkAutoFill.IsChecked = isAutofillEnabled
        WebViewMain.Source = New Uri("https://matrix.crmls.org/Matrix/Default.aspx")
        WebViewPrivy.Source = New Uri("https://app.privy.pro/users/sign_in")
        'WebViewPipe.Source = New Uri("https://app.pipedrive.com/auth/login")
        WebViewog.Source = New Uri("https://www.offergun.com/generate")
        WebViewog2.Source = New Uri("https://www.offergun.com/offer-history")
        'TabControlMain.SelectedIndex = 5

        TabControlMain.SelectedIndex = 0
        cmbEmailAccount.ItemsSource = fxCommon.GetOutlookEmailAccounts()
        cmbEmailAccount.SelectedIndex = 0
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Dim DefaultEmailAccount As XmlNode = xmlDoc.SelectSingleNode("/Login/DefaultEmailAccount")
        If DefaultEmailAccount IsNot Nothing Then
            cmbEmailAccount.SelectedValue = DefaultEmailAccount.InnerText
        End If

        'LoadDataGrid()
        LoadSettings()
        'rdoMessage.IsChecked = True
        origLeftSidebarWidth = SidebarColumn.Width.Value
        origRightSidebarWidth = RightSidebarColumn.Width.Value
    End Sub

    Private Async Sub TxtPrice_TextChanged(
    sender As Object,
    e As TextChangedEventArgs
)
        Dim price = txtPrice.Text
        If price = "" Then Exit Sub

        LastOfferPrice = price

        ' Live push if already on OfferGun
        If WebViewog.CoreWebView2 IsNot Nothing Then
            Dim url = WebViewog.Source?.ToString().ToLower()
            If url IsNot Nothing AndAlso url.Contains("offergun.com/generate") Then
                Await ApplyOfferGunPrice()
            End If
        End If
    End Sub

    Private Async Function ApplyOfferGunPrice() As Task

        If LastOfferPrice = "" Then Exit Function

        ' Wait for React hydration
        Await Task.Delay(1000)

        Await WebViewog.CoreWebView2.ExecuteScriptAsync(
    "
    (function(){
        let i=document.querySelector('input[name=offerPrice]');
        if(!i) return;

        let s=Object.getOwnPropertyDescriptor(
            HTMLInputElement.prototype,'value'
        ).set;

        s.call(i,'" & LastOfferPrice & "');
        i.dispatchEvent(new Event('input',{bubbles:true}));
        i.dispatchEvent(new Event('change',{bubbles:true}));
    })();
    ")

    End Function

    Private Function IsOutlookRunning() As Boolean
        Dim processes() As Process = Process.GetProcessesByName("OUTLOOK")
        Return processes.Length > 0
    End Function
    Private Sub StartOutlook()
        Try
            Dim outlookProcess As New Process()
            outlookProcess.StartInfo.FileName = IO.Path.Combine(AppContext.BaseDirectory, "DB\outlook.bat")
            outlookProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            outlookProcess.Start()
        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Unable to start Outlook: " & ex.Message)
        End Try
    End Sub
    Private Sub RunJQuery(observerScript As String)
        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub

    Private Sub SaveRadiusToXml()
        Try
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(_settingsFilePath)
            Dim milesNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Miles")
            If milesNode IsNot Nothing Then
                milesNode.InnerText = txtRadius.Text.Trim()
                xmlDoc.Save(_settingsFilePath)
            Else
                MessageBox.Show("Miles tag not found in login.xml.")
            End If
        Catch ex As System.Exception
            MessageBox.Show("Error saving radius: " & ex.Message)
        End Try
    End Sub
    Private Sub SetRadius()
        Dim strRadius As String = txtRadius.Text
        Dim iRadius As String = "3"
        Dim radiusValue As Double
        Dim radiusMap As New Dictionary(Of Double, String) From {
        {0.25, "0"},
        {0.5, "1"},
        {1, "2"},
        {2, "3"},
        {3, "4"},
        {5, "5"},
        {10, "6"},
        {25, "7"}
    }

        If Double.TryParse(strRadius, radiusValue) Then
            Dim possibleMatches = radiusMap.Keys.Where(Function(k) k >= radiusValue).OrderBy(Function(k) k)
            Dim nextBest As Double

            If possibleMatches.Any() Then
                nextBest = possibleMatches.First()
            Else
                nextBest = radiusMap.Keys.Max()
            End If

            iRadius = radiusMap(nextBest)
        End If
        Dim observerScript As String = "
            var comboboxElement = document.getElementById('Fm2_Ctrl12_Radius');
            var optionElements = comboboxElement.querySelectorAll('option');
            for (var i = 0; i < optionElements.length; i++) {
                if (i === " & iRadius & ") {
                    optionElements[i].selected = true;break;
                }
            };
            MapSearchJs.updateRadius(true);"
        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub
    Private Sub clickCheckbox(strCheckBoxName As String)

        'If (checkboxElement.value === '6140' || checkboxElement.dataset.mtxTrack === 'Status - " & strCheckBoxName & "') {
        Dim observerScript As String = "
        var trElements = document.getElementsByTagName('tr');
        for (var i = 0; i < trElements.length; i++) {
            var trElement = trElements[i];
            var checkboxElement = trElement.querySelector('input[type=\" & Chr(34) & "checkbox\" & Chr(34) & "]');
            if (checkboxElement) {
                if (checkboxElement.dataset.mtxTrack === 'Status - " & strCheckBoxName & "') {
                    checkboxElement.click();break;
                }
            }
        }"

        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub

    Private Sub PropSubType(sSubType As String)
        ' 1. Split the input string into an array of individual types.
        Dim typesToSelect() As String = sSubType.Split(","c)
        Dim jsArrayString As String = $"[{String.Join(",", typesToSelect.Select(Function(t) $"'{t.Trim()}'"))}]"

        ' 2. Construct the JavaScript to loop through the array and select each item.
        Dim observerScript As String = $"
        var listboxElement = document.getElementById('Fm2_Ctrl4_LB');
        if (listboxElement) {{
            var typesToSelect = {jsArrayString};

            // First, clear all existing selections in the listbox
            for (var i = 0; i < listboxElement.options.length; i++) {{
                listboxElement.options[i].selected = false;
            }}

            // Next, loop through our array of types and select each one
            typesToSelect.forEach(function(type) {{
                var optionElement = listboxElement.querySelector('option[title=""' + type + '""]');
                if (optionElement) {{
                    optionElement.selected = true;
                }}
            }});
        }}
        "

        ' Execute the script on the WebView control
        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub
    Private Sub PropSubType1(sSubType As String)
        Dim observerScript As String = "
            var listboxElement = document.getElementById('Fm2_Ctrl4_LB');
            var optionElement = listboxElement.querySelector('option[title=""Single Family Residence""]');
            for (var i = 0; i < listboxElement.options.length; i++) {
                listboxElement.options[i].selected = false;
            }
            optionElement.selected = true;"
        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub


    Public Sub SelectCity(cityName As String)
        ' Escape the city name to prevent issues with special characters in the JavaScript string
        Dim escapedCityName = System.Text.Json.JsonSerializer.Serialize(cityName).Trim(""""c)

        Dim observerScript As String = $"
            var listboxElement = document.getElementById('Fm2_Ctrl10_LB');
            if (listboxElement) {{
                for (var i = 0; i < listboxElement.options.length; i++) {{
                    listboxElement.options[i].selected = false;
                }}
                var optionElement = listboxElement.querySelector('option[title=\""{escapedCityName}\""]');
                if (optionElement) {{
                    optionElement.selected = true;
                    listboxElement.dispatchEvent(new Event('change'));
                }}
            }}
        "
        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub

    Private Async Sub privylogin(CoreWV As CoreWebView2)

        Try
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(_settingsFilePath)

            Dim userNameNode = xmlDoc.SelectSingleNode("/Login/PRIVY/UserName")
            Dim passwordNode = xmlDoc.SelectSingleNode("/Login/PRIVY/Password")

            If userNameNode Is Nothing OrElse passwordNode Is Nothing Then
                SystemSounds.Exclamation.Play()
                MessageBox.Show("Username or Password is missing in Login.xml for PRIVY")
                Exit Sub
            End If

            Dim userName As String = JsEscape(userNameNode.InnerText)
            Dim password As String = JsEscape(passwordNode.InnerText)

            bCodeProcessing = False

            ' STEP 1: Enter email and click login
            Await CoreWV.ExecuteScriptAsync(
            $"document.getElementById('user_email').value = '{userName}';"
        )

            Await CoreWV.ExecuteScriptAsync(
            "document.getElementById('login_button').click();"
        )

            ' WAIT for password-only page to load
            Await Task.Delay(2000) ' adjust if needed

            ' STEP 2: Enter password and login
            Dim pwdScript As String =
$"
(function() {{
    const pwd = document.getElementById('user_password');
    const btn = document.getElementById('login_button');

    if (!pwd) {{
        console.log('Password field not found');
        return;
    }}

    pwd.focus();
    pwd.value = '{password}';
    pwd.dispatchEvent(new Event('input', {{ bubbles: true }}));
    pwd.dispatchEvent(new Event('change', {{ bubbles: true }}));

    if (btn) btn.click();
}})();
"

            Await CoreWV.ExecuteScriptAsync(pwdScript)

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Privy login failed." & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Function JsEscape(value As String) As String
        Return value.Replace("\", "\\").Replace("'", "\'")
    End Function

    Private Function offergunlogin(CoreWV As CoreWebView2)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Dim userNameNode As XmlNode = xmlDoc.SelectSingleNode("/Login/OFFERGUN/UserName")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("/Login/OFFERGUN/Password")
        'MessageBox.Show("Loading XML from: " & _settingsFilePath)
        If userNameNode IsNot Nothing AndAlso passwordNode IsNot Nothing Then
            Dim userName As String = userNameNode.InnerText
            Dim password As String = passwordNode.InnerText
            bCodeProcessing = False
            CoreWV.ExecuteScriptAsync("document.getElementById('email').value = '" & userName & "';")
            CoreWV.ExecuteScriptAsync("document.getElementById('password').value = '" & password & "';")
            CoreWV.ExecuteScriptAsync("document.querySelector(""button.py-4.px-16.m-2.text-xl.text-white.font-medium.transition-colors.duration-150.bg-yellow-400.rounded-lg.hover:bg-opacity-80"").click();")
            CoreWV.ExecuteScriptAsync("document.querySelector(""button[type='submit']"").click();")

        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Username or Password is missing in the Login.xml for CRMLS")
        End If
        Return String.Empty
    End Function
    Private Sub pipelogin1(CoreWV As CoreWebView2)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Dim userNameNode As XmlNode = xmlDoc.SelectSingleNode("/Login/PIPE/UserName")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("/Login/PIPE/Password")

        If userNameNode IsNot Nothing AndAlso passwordNode IsNot Nothing Then
            Dim userName As String = userNameNode.InnerText
            Dim password As String = passwordNode.InnerText
            bCodeProcessing = False

            ' --- Simulate interaction for the Email field ---
            CoreWV.ExecuteScriptAsync("var emailInput = document.getElementById('login');")
            CoreWV.ExecuteScriptAsync("emailInput.value = '" & userName.Replace("'", "\'") & "';")
            CoreWV.ExecuteScriptAsync("emailInput.dispatchEvent(new Event('focus', { bubbles: true }));")
            CoreWV.ExecuteScriptAsync("emailInput.dispatchEvent(new Event('input', { bubbles: true }));")
            CoreWV.ExecuteScriptAsync("emailInput.dispatchEvent(new Event('change', { bubbles: true }));")
            CoreWV.ExecuteScriptAsync("emailInput.dispatchEvent(new Event('blur', { bubbles: true }));")

            ' --- Simulate interaction for the Password field ---
            CoreWV.ExecuteScriptAsync("var passwordInput = document.getElementById('password');")
            CoreWV.ExecuteScriptAsync("passwordInput.value = '" & password.Replace("'", "\'") & "';")
            CoreWV.ExecuteScriptAsync("passwordInput.dispatchEvent(new Event('focus', { bubbles: true }));")
            CoreWV.ExecuteScriptAsync("passwordInput.dispatchEvent(new Event('input', { bubbles: true }));")
            CoreWV.ExecuteScriptAsync("passwordInput.dispatchEvent(new Event('change', { bubbles: true }));")
            CoreWV.ExecuteScriptAsync("passwordInput.dispatchEvent(new Event('blur', { bubbles: true }));")

            ' Click the login button
            CoreWV.ExecuteScriptAsync("document.querySelector('button.submit-button.bt').click();")
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Username or Password is missing in the Login.xml for CRMLS")
        End If
    End Sub

    Private Async Sub pipelogin(CoreWV As CoreWebView2)
        Dim xmlDoc As New XmlDocument()

        Try
            xmlDoc.Load(_settingsFilePath)
        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Error loading Login.xml: " & ex.Message)
            Exit Sub
        End Try

        Dim userNameNode As XmlNode = xmlDoc.SelectSingleNode("/Login/PIPE/UserName")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("/Login/PIPE/Password")

        If userNameNode IsNot Nothing AndAlso passwordNode IsNot Nothing Then
            Dim userName As String = userNameNode.InnerText
            Dim password As String = passwordNode.InnerText
            ' bCodeProcessing = False ' Unsure of context for this, keep if needed elsewhere

            ' Define a comprehensive JavaScript function to simulate typing with events and delays
            ' This function will be injected once and then called for each input field.
            Dim jsFunction As String = "
            async function simulateTyping(elementId, text) {
                const inputElement = document.getElementById(elementId);
                if (!inputElement) {
                    console.error('Element not found: ' + elementId);
                    return;
                }

                // 1. Focus the element
                inputElement.dispatchEvent(new Event('focus', { bubbles: true }));
                await new Promise(r => setTimeout(r, 100)); // Small delay after focus

                // 2. Clear existing value and trigger initial input event
                inputElement.value = '';
                inputElement.dispatchEvent(new Event('input', { bubbles: true }));
                await new Promise(r => setTimeout(r, 50)); // Delay after clearing

                // 3. Simulate character-by-character typing
                for (let i = 0; i < text.length; i++) {
                    const char = text[i];
                    const keyCode = char.charCodeAt(0); // ASCII value for key events

                    // Dispatch keydown event
                    inputElement.dispatchEvent(new KeyboardEvent('keydown', {
                        key: char,
                        code: 'Key' + char.toUpperCase(), // e.g., 'KeyA' for 'a'
                        keyCode: keyCode, // Deprecated but often still used
                        which: keyCode,   // Deprecated but often still used
                        bubbles: true,
                        cancelable: true // Allows event to be cancelled
                    }));
                    await new Promise(r => setTimeout(r, 30 + Math.random() * 20)); // Random delay for human-like timing

                    // Dispatch keypress event (only for printable characters)
                    // Keypress is deprecated but some older scripts might still rely on it.
                    if (char.length === 1 && char.match(/[\x20-\x7E]/)) { // Check for printable ASCII
                        inputElement.dispatchEvent(new KeyboardEvent('keypress', {
                            key: char,
                            code: 'Key' + char.toUpperCase(),
                            keyCode: keyCode,
                            which: keyCode,
                            bubbles: true,
                            cancelable: true
                        }));
                        await new Promise(r => setTimeout(r, 30 + Math.random() * 20)); // Random delay
                    }

                    // Update value and dispatch input event
                    inputElement.value += char;
                    inputElement.dispatchEvent(new Event('input', { bubbles: true }));
                    await new Promise(r => setTimeout(r, 50 + Math.random() * 50)); // Delay between characters

                    // Dispatch keyup event
                    inputElement.dispatchEvent(new KeyboardEvent('keyup', {
                        key: char,
                        code: 'Key' + char.toUpperCase(),
                        keyCode: keyCode,
                        which: keyCode,
                        bubbles: true,
                        cancelable: true
                    }));
                    await new Promise(r => setTimeout(r, 30 + Math.random() * 20)); // Random delay
                }

                // 4. Dispatch change event (value committed)
                inputElement.dispatchEvent(new Event('change', { bubbles: true }));
                await new Promise(r => setTimeout(r, 100)); // Delay after change

                // 5. Blur the element
                inputElement.dispatchEvent(new Event('blur', { bubbles: true }));
                await new Promise(r => setTimeout(r, 100)); // Delay after blur
            }
        "

            ' Inject the JavaScript function into the WebView2 page
            CoreWV.ExecuteScriptAsync(jsFunction)
            Await Task.Delay(200) ' Give a moment for the JavaScript function to be defined in the browser context

            ' Call the JavaScript function for the email field
            CoreWV.ExecuteScriptAsync("simulateTyping('login', '" & userName.Replace("'", "\'") & "');")
            ' Estimate the time needed for typing based on string length and per-character delays
            Await Task.Delay(userName.Length * 150 + 500) ' (approx. 150ms per char + 500ms buffer)

            ' Call the JavaScript function for the password field
            CoreWV.ExecuteScriptAsync("simulateTyping('password', '" & password.Replace("'", "\'") & "');")
            ' Estimate the time needed for typing based on string length and per-character delays
            Await Task.Delay(password.Length * 150 + 500) ' (approx. 150ms per char + 500ms buffer)

            ' Click the login button after all typing simulations are complete
            'CoreWV.ExecuteScriptAsync("document.querySelector('button.submit-button.bt').click();")

        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Username or Password is missing in the Login.xml for CRMLS")
        End If
    End Sub
    Private Function readsettingsxml() As XmlDocument
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Return xmlDoc
    End Function
    Private Function setxmltag(ByVal xmlDoc As XmlDocument, ByVal tagName As String, ByVal tagValue As String) As Boolean
        '<MLSSETTINGS>	
        '<Miles>0.25</Miles>
        '<Subcat>0.25</Subcat>
        '<Sqft>300</Sqft>
        '</MLSSETTINGS>

        Dim node As XmlNode = xmlDoc.SelectSingleNode("/Login/" & tagName)
        If node IsNot Nothing Then
            node.InnerText = tagValue
            Return writeSettingsXML(xmlDoc)
        Else
            MessageBox.Show("Tag not found: " & tagName)
            Return False
        End If
    End Function
    Private Function writeSettingsXML(ByVal xmlDoc As XmlDocument) As Boolean
        Try
            xmlDoc.Save(_settingsFilePath)
            Return True
        Catch ex As System.Exception
            MessageBox.Show("Error saving settings: " & ex.Message)
            Return False
        End Try
    End Function
    Private Function SiteLogin(CoreWV As CoreWebView2)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Dim userNameNode As XmlNode = xmlDoc.SelectSingleNode("/Login/CRMLS/UserName")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("/Login/CRMLS/Password")
        If userNameNode IsNot Nothing AndAlso passwordNode IsNot Nothing Then
            Dim userName As String = userNameNode.InnerText
            Dim password As String = passwordNode.InnerText
            bCodeProcessing = False
            CoreWV.ExecuteScriptAsync("document.getElementById('Username').value = '" & userName & "';")
            'CoreWV.ExecuteScriptAsync("document.getElementById('Password').value = '" & password & "';")
            CoreWV.ExecuteScriptAsync("document.getElementById('elPasswordText').value = '" & password & "';")
            CoreWV.ExecuteScriptAsync("document.querySelector(""button.py-4.px-16.m-2.text-xl.text-white.font-medium.transition-colors.duration-150.bg-yellow-400.rounded-lg.hover:bg-opacity-80"").click();")
            CoreWV.ExecuteScriptAsync("document.querySelector(""button[name='button']"").click();")
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Username or Password is missing in the Login.xml for CRMLS")
        End If
        Return String.Empty
    End Function

    Private Sub WebView_CoreWebView2InitializationCompleted(sender As Object, e As CoreWebView2InitializationCompletedEventArgs)
        Dim CurrentWV As Microsoft.Web.WebView2.Wpf.WebView2 = sender
        If e.IsSuccess Then
            AddHandler CurrentWV.CoreWebView2.DOMContentLoaded, AddressOf WebView_DOMContentLoaded
            'AddHandler CurrentWV.NavigationStarting, AddressOf WebView_NavigationStarting
            'AddHandler wvComp.NavigationCompleted, AddressOf WebView_NavigationCompleted        
            'AddHandler CurrentWV.CoreWebView2.DOMContentLoaded, AddressOf WebView_DOMContentLoaded
        End If
    End Sub

    Private Sub WebViewStatus_CoreWebView2InitializationCompleted(sender As Object, e As CoreWebView2InitializationCompletedEventArgs)
        Dim CurrentWV As Microsoft.Web.WebView2.Wpf.WebView2 = sender
        If e.IsSuccess Then
            AddHandler CurrentWV.CoreWebView2.DOMContentLoaded, AddressOf WebViewStatus_DOMContentLoaded
            'AddHandler CurrentWV.NavigationStarting, AddressOf WebView_NavigationStarting
            'AddHandler wvComp.NavigationCompleted, AddressOf WebView_NavigationCompleted        
            'AddHandler CurrentWV.CoreWebView2.DOMContentLoaded, AddressOf WebView_DOMContentLoaded
        End If
    End Sub
    Private Async Sub WebViewStatus_DOMContentLoaded(sender As Object, e As CoreWebView2DOMContentLoadedEventArgs)
        Dim CoreWV As Microsoft.Web.WebView2.Core.CoreWebView2 = sender
        Dim url As String = CoreWV.Source.ToString().ToLower()
        If url.Contains("https://signin.crmls.org/account/login") Then
            SiteLogin(CoreWV)
            bCodeProcessing = True
        ElseIf url.Contains("https://matrix.crmls.org/matrix/default.aspx") AndAlso MLSListingStatus IsNot String.Empty Then
            Await CoreWV.ExecuteScriptAsync("document.getElementById('ctl02_m_ucSpeedBar_m_tbSpeedBar').value = '" & MLSListingStatus & "';")
            Threading.Thread.Sleep(200)
            Await CoreWV.ExecuteScriptAsync("document.getElementById('ctl02_m_ucSpeedBar_m_lnkGo').click();")
            Threading.Thread.Sleep(200)
        ElseIf url.Contains("https://matrix.crmls.org/matrix/results.aspx") AndAlso MLSListingStatus IsNot String.Empty Then
            Await GetStatusData(CoreWV)
        ElseIf url.Contains("https://app.privy.pro/users/sign_in") Then
            privylogin(CoreWV)
            'ElseIf url.Contains("https://app.pipedrive.com/auth/login") Then
            'pipelogin(CoreWV)
        ElseIf url.Contains("https://www.offergun.com/login") Then
            offergunlogin(CoreWV)

        End If
        If url.Contains("https://www.offergun.com/generate") Then
            Await OfferGunAutoFill()
            Await UpdateOfferGunAddress()
        End If

    End Sub
    Private Async Function OfferGunAutoFill() As Task

        Dim price As String = txtPrice.Text
        ' Wait for React + Radix UI to be ready
        Await Task.Delay(1200)

        ' 1. Select Template
        Await WebViewog.CoreWebView2.ExecuteScriptAsync(
    "
    (async function(){
        let btn=[...document.querySelectorAll('button')]
            .find(b=>b.innerText.includes('Select a template'));
        if(!btn) return;
        btn.click();
        await new Promise(r=>setTimeout(r,600));
        let item=[...document.querySelectorAll('[role=menuitem],[role=option]')]
            .find(e=>e.innerText.trim()==='Cash Offer');
        if(item) item.click();
    })();
    ")

        Await Task.Delay(600)

        ' 2. Select Buyer
        Await WebViewog.CoreWebView2.ExecuteScriptAsync(
    "
    (async function(){
        let btn=[...document.querySelectorAll('button')]
            .find(b=>b.innerText.includes('Select a Buyer'));
        if(!btn) return;
        btn.click();
        await new Promise(r=>setTimeout(r,600));
        let item=[...document.querySelectorAll('[role=menuitem],[role=option]')]
            .find(e=>e.innerText.trim()==='John Smith');
        if(item) item.click();
    })();
    ")

        Await Task.Delay(600)

        ' 3. Fill Offer Price (React-safe)
        Await WebViewog.CoreWebView2.ExecuteScriptAsync(
    "
    (function(){
        let i=document.querySelector('input[name=offerPrice]');
        if(!i) return;
        let s=Object.getOwnPropertyDescriptor(
            HTMLInputElement.prototype,'value').set;
       s.call(i,'" & price & "');
        i.dispatchEvent(new Event('input',{bubbles:true}));
        i.dispatchEvent(new Event('change',{bubbles:true}));
    })();
    ")

    End Function
    Private Async Function UpdateOfferGunAddress() As Task

        Try
            ' ---- Safety checks ----
            If WebViewog Is Nothing OrElse WebViewog.CoreWebView2 Is Nothing Then Return

            Dim url As String = WebViewog.Source?.ToString()?.ToLower()
            If String.IsNullOrEmpty(url) OrElse Not url.Contains("offergun.com/generate") Then Return

            Dim addr As String = CurrentPropertyAddress
            If String.IsNullOrWhiteSpace(addr) Then Return

            ' Prevent duplicate injection
            'If addr = LastSentAddressToOfferGun Then Return

            ' Serialize address safely for JS
            Dim addressJson As String = System.Text.Json.JsonSerializer.Serialize(addr)

            Dim js As String =
$"(async function(){{

    function sleep(ms) {{ return new Promise(r => setTimeout(r, ms)); }}

    /* ---------- 1. Wait for address input ---------- */
    let input = null;
    for(let i=0;i<15;i++){{
        input = document.getElementById('search');
        if(input) break;
        await sleep(300);
    }}
    if(!input) return 'input-not-found';

    /* ---------- 2. Set address (React-safe) ---------- */
    let setter = Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,'value'
    ).set;

    setter.call(input, {addressJson});
    input.dispatchEvent(new Event('input', {{ bubbles:true }}));
    input.dispatchEvent(new Event('change', {{ bubbles:true }}));

    /* ---------- 3. Allow React to re-render ---------- */
    await sleep(800);

    /* ---------- 4. Find ENABLED Search button ---------- */
    let btn = null;
    for(let i=0;i<15;i++){{
        btn = [...document.querySelectorAll('button')]
            .find(b =>
                b.textContent &&
                b.textContent.trim().toLowerCase() === 'search' &&
                !b.disabled &&
                b.offsetParent !== null
            );
        if(btn) break;
        await sleep(300);
    }}

    if(!btn) return 'search-button-not-found';

    /* ---------- 5. Real user-like click ---------- */
    ['pointerdown','mousedown','mouseup','click'].forEach(type => {{
        btn.dispatchEvent(new MouseEvent(type, {{
            bubbles: true,
            cancelable: true,
            view: window
        }}));
    }});

    return 'ok';

}})();"

            Dim result As String = Await WebViewog.CoreWebView2.ExecuteScriptAsync(js)

            If Not String.IsNullOrWhiteSpace(result) AndAlso result.ToLower().Contains("ok") Then
                LastSentAddressToOfferGun = addr
                LogMessage("OfferGun search triggered for address: " & addr)
            Else
                LogMessage("UpdateOfferGunAddress failed: " & result)
            End If

        Catch ex As System.Exception
            fxCommon.GenerateLog(ex)
        End Try

    End Function

    Private Async Sub WebView_DOMContentLoaded(sender As Object, e As CoreWebView2DOMContentLoadedEventArgs)
        Dim CoreWV As Microsoft.Web.WebView2.Core.CoreWebView2 = sender
        Dim url As String = CoreWV.Source.ToString().ToLower()
        Dim strJScript As String

        If url.Contains("https://matrix.crmls.org/matrix/results") Then
            'Await ExtractMatrixAddress(CoreWV)
        End If
        ' LblURL.Text = url
        If bCodeProcessing = True Then

            If url.Contains("https://signin.crmls.org/account/login") Then
                SiteLogin(CoreWV)
            ElseIf url.Contains("https://matrix.crmls.org/matrix/results.aspx") AndAlso MLSListingSub IsNot String.Empty AndAlso MLSListingMain Is String.Empty Then
                strJScript = "var links = document.getElementsByTagName('a');var searchText = '" & MLSListingSub & "';for (var i = 0; i < links.length; i++) {if (links[i].innerText === searchText) {links[i].click();break;}}"
                WebViewMain.CoreWebView2.ExecuteScriptAsync(strJScript)
                Dim linkId As String = "m_DisplayCore" 'Specify the ID of the link element
                Dim jsCode As String = $"__doPostBack('{linkId}', 'Redisplay|581,,0');"
                CoreWV.ExecuteScriptAsync(jsCode)
                Threading.Thread.Sleep(1200)
                MLSListingSub = String.Empty

            ElseIf (url.Contains("https://matrix.crmls.org/matrix/mymatrix/home") Or url.Contains("https://matrix.crmls.org/matrix/results.aspx") Or url.Contains("https://matrix.crmls.org/matrix/default.aspx")) AndAlso MLSListingMain IsNot String.Empty Then
                strJScript = "var speedBarSearch = document.getElementById('ctl02_m_ucSpeedBar_m_tbSpeedBar');" &
             "if (speedBarSearch) {" &
             "    speedBarSearch.focus();" &
             "    speedBarSearch.value = '" & MLSListingMain & "';" &
             "    var event = new Event('keydown');" &
             "    speedBarSearch.dispatchEvent(event);" &
             "}" &
             "var speedBarSearchAlt = document.getElementById('ctl01_m_ucSpeedBar_m_tbSpeedBar');" &
             "if (speedBarSearchAlt) {" &
             "    speedBarSearchAlt.focus();" &
             "    speedBarSearchAlt.value = '" & MLSListingMain & "';" &
             "    var event = new Event('keydown');" &
             "    speedBarSearchAlt.dispatchEvent(event);" &
             "}"
                CoreWV.ExecuteScriptAsync(strJScript)
                MLSListingMain = String.Empty
                Threading.Thread.Sleep(1200)
                Dim jsCode As String = "SpeedBarJs.onKeyDown(new KeyboardEvent('keydown'));"
                CoreWV.ExecuteScriptAsync(jsCode)

                'return SpeedBarJs.checkGo();                
                CoreWV.ExecuteScriptAsync("document.getElementById('ctl01_m_ucSpeedBar_m_lnkGo').click();")
                CoreWV.ExecuteScriptAsync("document.getElementById('ctl02_m_ucSpeedBar_m_lnkGo').click();")
                CoreWV.ExecuteScriptAsync("SpeedBarJs.checkGo()")
                Threading.Thread.Sleep(200)
            End If


        ElseIf url.Contains("https://matrix.crmls.org/matrix/search/residential/detail") Then
            If isAutofillEnabled Then
                bCodeProcessing = False
                WebViewComp.CoreWebView2.ExecuteScriptAsync("document.getElementById('Fm2_Ctrl12_TB').focus();")
                WebViewComp.CoreWebView2.ExecuteScriptAsync("document.getElementById('Fm2_Ctrl12_TB').click();")
                Threading.Thread.Sleep(100)
                WebViewComp.CoreWebView2.ExecuteScriptAsync("document.getElementById('Fm2_Ctrl12_TB').value = '" & txtAddressSearch.Text & "';MapSearchJs.geocode();")
                Threading.Thread.Sleep(250)
                WaitForSuggestions()
                WebViewComp.CoreWebView2.ExecuteScriptAsync("document.querySelectorAll('.disambiguation li')[1].click()")

                'Now enable checkboxes
                Threading.Thread.Sleep(200)

                'set Radius
                SetRadius()
                Threading.Thread.Sleep(500)

                'Status
                clickCheckbox("Active")
                Threading.Thread.Sleep(100)
                clickCheckbox("Act Under Contract")
                Threading.Thread.Sleep(100)
                clickCheckbox("Pending")
                Threading.Thread.Sleep(100)
                clickCheckbox("Closed")
                Threading.Thread.Sleep(200)

                'PropertyAccessors Sub type
                Dim sSubType As String = ""
                If rdoFam.IsChecked Then
                    sSubType = "Single Family Residence"
                ElseIf rdoCondo.IsChecked Then
                    sSubType = "Condominium,Townhall"
                End If
                PropSubType(sSubType)
                Threading.Thread.Sleep(1000)

                'Choose City
                Dim fullAddress As String = txtAddressSearch.Text
                Dim addressParts() As String = fullAddress.Split(","c)
                Dim cityName As String = ""
                If addressParts.Length >= 2 Then
                    cityName = addressParts(addressParts.Length - 2).Trim()
                    If cityName <> "" Then
                        SelectCity(cityName)
                        Threading.Thread.Sleep(1000)
                    End If
                End If

                'Enter SqFt range
                If txtSqFtRange.Text.Trim() <> "" Then
                    Dim iSqftRange As Double = Val(txtSqFtRange.Text.Trim())
                    Dim iProSqFt As Double = 0
                    Dim liSqFT As New List(Of String)

                    liSqFT = GetGridSelectedItems("SqFt", "Checked")
                    If (liSqFT.Count > 0) Then
                        If (liSqFT.Count = 1) Then
                            iProSqFt = Val(liSqFT.Item(0).ToString())
                        Else
                            SystemSounds.Exclamation.Play()
                            MessageBox.Show("Please select one item to proceed.")
                        End If

                        'Create a sqft range with the calculation: SqFtRange = "(iProSqFt - iSqftRange) - (iProSqFt + iSqftRange)"
                        Dim sqftRange As String = (iProSqFt - iSqftRange).ToString() & "-" & (iProSqFt + iSqftRange).ToString()
                        'Enter this field in this field. 
                        '<input type="text" class="textbox" id="Fm2_Ctrl2012_TB" name="Fm2_Ctrl2012_TB" onkeypress="return SearchJs.onTextBoxKeyPress(event, this,SearchJs.isNumericKeyCode);" onpaste="return SearchJs.onTextBoxPaste(event, this,SearchJs.isNumericKeyCode);" onkeyup="SearchJs.setBackgroundColor(this, SearchJs.checkTextBox);" onchange="SearchJs.setBackgroundColor(this, SearchJs.checkTextBox);" data-mtx-track="Living Area" data-mtx-track-prop-type="NumericTextBox" data-mtx-track-prop-id="2012" maxlength="500" style="width:85px;background-color:#ffffff;" value="" title="The living area of the property, in square feet or square meters.  See the Living Area Units field to determine if this is Square Feet or Meters.  The default is Square Feet.">
                        lblSqFtRange.Content = "(" & sqftRange.ToString & ")"

                        WebViewComp.CoreWebView2.ExecuteScriptAsync("document.getElementById('Fm2_Ctrl2012_TB').focus();")
                        WebViewComp.CoreWebView2.ExecuteScriptAsync("document.getElementById('Fm2_Ctrl2012_TB').click();")
                        Threading.Thread.Sleep(100)
                        WebViewComp.CoreWebView2.ExecuteScriptAsync("document.getElementById('Fm2_Ctrl2012_TB').value = '" & sqftRange.ToString & "';")
                        Threading.Thread.Sleep(250)

                    End If
                End If

                'Click search
                RunJQuery("document.getElementById('m_ucSearchButtons_m_lbSearch').click();")
                Threading.Thread.Sleep(200)
            End If
        End If

    End Sub
    Private Sub WebView_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Dim CurrentWV As WebView2 = sender
        If e.TryGetWebMessageAsString() = "SuggestionsLoaded" Then
            CurrentWV.CoreWebView2.ExecuteScriptAsync("document.querySelectorAll('.disambiguation li')[1].click()")
        End If
    End Sub

    Private Sub WaitForSuggestions()
        Dim maxRetryCount As Integer = 3 ' Adjust as needed
        Dim retryInterval As Integer = 500 ' 1 second (adjust as needed)
        Dim retryCount As Integer = 0

        Dim observerScript As String = "
        var targetNode = document.body;
        var config = { childList: true, subtree: true };
        
        var callback = function(mutationsList, observer) {
            for(var mutation of mutationsList) {
                if (mutation.target.querySelector('.mapSearchDialog')) {
                    observer.disconnect();
                    window.external.notify('SuggestionsLoaded');
                }
            }
        };
        
        var observer = new MutationObserver(callback);
        observer.observe(targetNode, config);
    "

        ' Attach event handler for JavaScript notification
        AddHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived

        ' Inject MutationObserver script
        WebViewComp.CoreWebView2.ExecuteScriptAsync(observerScript)

        ' Poll for suggestions or wait for MutationObserver notification
        Do While retryCount < maxRetryCount
            Task.Delay(retryInterval).Wait()
            WebViewComp.CoreWebView2.ExecuteScriptAsync("if (document.querySelector('.mapSearchDialog')) window.external.notify('SuggestionsLoaded');")
            retryCount += 1
        Loop

        ' Cleanup
        RemoveHandler WebViewComp.CoreWebView2.WebMessageReceived, AddressOf WebView_WebMessageReceived
    End Sub
    'Private Async Sub SelectBuyer(wv As Microsoft.Web.WebView2.Wpf.WebView2, buyerName As String)

    '    Dim js As String =
    '"
    '(async function(){

    '    let btn = [...document.querySelectorAll('button')]
    '        .find(b => b.innerText.includes('Select a Buyer'));

    '    if (!btn) return 'buyer-button-not-found';

    '    btn.click();

    '    await new Promise(r => setTimeout(r, 800));

    '    let item = [...document.querySelectorAll('[role=menuitem],[role=option]')]
    '        .find(el => el.innerText.trim() === '" & buyerName & "');

    '    if (!item) return 'buyer-item-not-found';

    '    item.click();
    '    return 'ok';
    '})();
    '"

    '    Await wv.CoreWebView2.ExecuteScriptAsync(js)

    'End Sub

    'Private Async Sub FillOfferPrice(wv As Microsoft.Web.WebView2.Wpf.WebView2, price As String)

    '    Dim js As String =
    '"
    '(function(){
    '    let input = document.querySelector('input[name=offerPrice]');
    '    if (!input) return 'not-found';

    '    let setter = Object.getOwnPropertyDescriptor(
    '        HTMLInputElement.prototype, 'value'
    '    ).set;

    '    setter.call(input, '" & price & "');
    '    input.dispatchEvent(new Event('input', { bubbles: true }));
    '    input.dispatchEvent(new Event('change', { bubbles: true }));

    '    return 'ok';
    '})();
    '"

    '    Await wv.CoreWebView2.ExecuteScriptAsync(js)

    'End Sub

    'Private Async Sub SelectTemplate(wv As Microsoft.Web.WebView2.Wpf.WebView2, templateName As String)

    '    Dim js As String =
    '"
    '(async function(){

    '    // 1. Find Template button
    '    let btn = [...document.querySelectorAll('button')]
    '        .find(b => b.innerText.includes('Select a template'));

    '    if (!btn) return 'template-button-not-found';

    '    btn.click();

    '    // 2. Wait for dropdown
    '    await new Promise(r => setTimeout(r, 800));

    '    // 3. Click template option
    '    let item = [...document.querySelectorAll('[role=menuitem],[role=option]')]
    '        .find(el => el.innerText.trim() === '" & templateName & "');

    '    if (!item) return 'template-item-not-found';

    '    item.click();
    '    return 'ok';
    '})();
    '"

    '    Await wv.CoreWebView2.ExecuteScriptAsync(js)

    'End Sub


    Private Sub LoadDataGrid()
        Dim sFieldList As String
        Dim sCondition As String
        Dim sQuery As String
        Dim xmlDoc As New XmlDocument()
        Dim sConditionList As String

        'Get Settings
        Try
            xmlDoc.Load(_settingsFilePath)
        Catch ex As System.Exception
            MessageBox.Show("Error loading Settings XMl")
            Exit Sub
        End Try
        Dim sActiveList As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSStatus/Active")
        Dim sPendingList As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSStatus/Pending")
        Dim sSoldList As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSStatus/Sold")
        Dim sHoldList As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSStatus/Hold")
        Dim sOthList As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSStatus/Other")


        ' Lambda function to convert comma-separated values into a single-quoted, comma-separated string
        Dim convertToQuotedString As Func(Of String, String) = Function(s) String.Join(",", s.Split(New Char() {","c}).Select(Function(val) "'" & val.Trim() & "'"))

        'Set query        
        sFieldList = "MLSListingID,SavedAddress,City,Seller,SellerEmail,LACell,LADirect,SellerPhone,CoSeller,Emailed,OfferPrice,ARV,ListPrice,ClosePrice,Status,SqFt,DateAdded,pdDealId,pdStageName"
        sCondition = ""
        'If rdoActive.IsChecked = True Then
        '    sCondition = "Active"
        'ElseIf rdoPending.IsChecked = True Then
        '    sCondition = "Pending"
        'ElseIf rdoSold.IsChecked = True Then
        '    sCondition = "Sold"
        'ElseIf rdoHold.IsChecked = True Then
        '    sCondition = "Hold"
        'ElseIf rdoOth.IsChecked = True Then
        '    sCondition = "Other"
        'End If

        'Execute Query
        If sCondition IsNot "" Then
            Dim xNodeList As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSStatus/" & sCondition)
            sConditionList = convertToQuotedString(xNodeList.InnerText)
            sCondition = " AND Status in (" & sConditionList & ")"
        End If

        sQuery = "Select " & sFieldList & " from Property WHERE (TabName = 0 OR TabName = '' OR TabName is NULL ) " & sCondition & " Order by Status"
        dataTable = fxCommon.SQLExecuteReader(sQuery)
        Dim IsSelectedColumn As New DataColumn("IsSelected", GetType(Boolean))
        IsSelectedColumn.DefaultValue = False
        dataTable.Columns.Add(IsSelectedColumn)

        'dgDetails.ItemsSource = dataTable.DefaultView
        'dgDetails.Tag = dataTable

        'dgDetails.UpdateLayout()
        'dgDetails.ScrollIntoView(dgDetails.Items(0), dgDetails.Columns(dgDetails.Columns.Count - 1))
        'Dim selectedIds As New List(Of String)
        'selectedIds = GetGridSelectedItems("MLSListingID", "Checked")
        'lbldgStatus.Content = "Display/Selected: " & dgDetails.Items.Count & "/" & (selectedIds.Count + 1)
    End Sub

    'CheckMLSListingRow(dgDetails, "A12345")
    Public Sub CheckMLSListingRow(dgDetails As DataGrid, mlsListingID As String)
        For Each item In dgDetails.Items
            If item IsNot Nothing AndAlso item.GetType().GetProperty("MLSListingID") IsNot Nothing Then
                Dim idValue As String = item.GetType().GetProperty("MLSListingID").GetValue(item, Nothing)?.ToString()
                If item.GetType().GetProperty("IsChecked") IsNot Nothing Then
                    If idValue = mlsListingID Then
                        item.GetType().GetProperty("IsChecked").SetValue(item, True, Nothing)
                    Else
                        item.GetType().GetProperty("IsChecked").SetValue(item, False, Nothing)
                    End If
                End If
            End If
        Next
        dgDetails.Items.Refresh()
    End Sub

    Private Sub InitializeComboBoxData()
        Dim data As New ObservableCollection(Of Item)()
        data.Add(New Item() With {.Id = 1, .Name = "Offer Email"})
        data.Add(New Item() With {.Id = 1, .Name = "Active Email"})
        data.Add(New Item() With {.Id = 2, .Name = "Pending Email"})
        data.Add(New Item() With {.Id = 3, .Name = "Closed Email"})
        cmbEmailTemplate.ItemsSource = data
        cmbEmailTemplate.DisplayMemberPath = "Name"

    End Sub
    Private Function LoadEmailCombo() As String
        dataTable = fxCommon.SQLExecuteReader("select * from Property")
        dataTable.Columns.Add("Select", GetType(Boolean)).SetOrdinal(0)
        'dgDetails.ItemsSource = dataTable.DefaultView
        Return String.Empty
    End Function


    Private Sub TabControlMain_SelectionChanged(sender As Object, e As RoutedEventArgs)
        Dim tabName As String = TabControlMain.SelectedItem.Header.ToString()

        'If Not tabChangeInitiated Then Return
        If tabName <> tabB4Change Then
            If TabControlMain.SelectedItem.Header.ToString() = "Main" Then
                'More code???            
            ElseIf TabControlMain.SelectedItem.Header.ToString() = "Analyze" Then
                LoadSettingsFromXml()
            ElseIf TabControlMain.SelectedItem.Header.ToString() = "Opportunities" Then
                LoadDataGrid()
            ElseIf TabControlMain.SelectedItem.Header.ToString() = "Pipedrive" Then
                'LoadDataGridLeads()
            ElseIf TabControlMain.SelectedItem.Header.ToString() = "Email" Then
                'Do nothing
            ElseIf TabControlMain.SelectedItem.Header.ToString() = "Privy" Then
                'getPDFTableDetails()
            End If
        End If
        'Reset flag
        tabChangeInitiated = False
        tabB4Change = TabControlMain.SelectedItem.Header.ToString()
    End Sub
    Private Sub TabControlMain_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        tabChangeInitiated = True
    End Sub


    Private Sub dgDetails_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

        If (DirectCast(sender, System.Windows.FrameworkElement).Name = "TabControlMain") Then
            If TabControlMain.SelectedIndex = 0 Then
                ' Code for TabPage1
            ElseIf TabControlMain.SelectedIndex = 2 Then
                ' LoadDataGrid()
            ElseIf TabControlMain.SelectedIndex = 3 Then
                ' Do nothing
            End If

        ElseIf (DirectCast(sender, System.Windows.FrameworkElement).Name = "dgDetails") Then
            tabChangeInitiated = False
            'If dgDetails.SelectedItems.Count > 0 Then

            '    'LoadNotes
            '    Dim MLSID As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(0).ToString()
            '    Dim SavedAddress As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(1).ToString()
            '    Dim Seller As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(3).ToString()
            '    Dim dtNotes = fxCommon.SQLExecuteReader("select Date, Notes from Notes where MLSListingID='" & MLSID & "' ORDER BY rowid DESC")  'Order by strftime('%Y-%m-%d %H:%M:%S', date) ASC")' Order by strftime('%Y-%m-%d %H:%M:%S', date) ASC")
            '    dgNotes.ItemsSource = dtNotes.DefaultView

            '    'SMS
            '    Dim dtSMS = fxCommon.SQLExecuteReader("Select * from EmailTemplate where Name like '%SMS%'")
            '    Seller = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Seller.Split(" ")(0).ToLower())
            '    For Each rowData As DataRow In dtSMS.Rows
            '        txtSMS.Text = rowData("Message").ToString().Replace("<<Address>>", SavedAddress).Replace("<<First Name>>", Seller)
            '    Next

            'End If
        End If

        'Dim selectedIds As New List(Of String)
        'selectedIds = GetGridSelectedItems("MLSListingID", "Checked")
        'lbldgStatus.Content = "Display/Selected: " & dgDetails.Items.Count & "/" & (selectedIds.Count + 1)
    End Sub

    Private Sub ProcLoadNotesDG()
        'Try
        '    If dgDetails.SelectedItems.Count > 0 Then
        '        Dim MLSID As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(0).ToString()
        '        Dim dtNotes = fxCommon.SQLExecuteReader("select Date, Notes from Notes where MLSListingID='" & MLSID & "' ORDER BY rowid DESC")  'Order by strftime('%Y-%m-%d %H:%M:%S', date) ASC")
        '        dgNotes.ItemsSource = dtNotes.DefaultView
        '    End If
        'Catch ex As System.Exception
        '    MessageBox.Show("Error in loading Notes: " & ex.Message)
        'End Try
    End Sub

    Private Async Function MainSave_ClickAsync(sender As Object, e As RoutedEventArgs) As Task
        If TabControlMain.SelectedIndex = 0 Or TabControlMain.SelectedIndex = 1 Then
            Await ProcessRecord(False, String.Empty)
            SystemSounds.Exclamation.Play()
            MessageBox.Show(message)
            LoadDataGrid()
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Not Data exists to save")
        End If
        ' If LblURL.Text.Contains("https://matrix.crmls.org/matrix/results.aspx") Then
    End Function

    Private Async Sub MainSaveEmail_ClickAsync(sender As Object, e As RoutedEventArgs)
        Await sendEmailAndSave(True)
        ProcLoadNotesDG()
    End Sub

    Private Async Function sendEmailAndSave(email As Boolean) As Task
        Dim emailTemplate As String = ""
        Dim cbmVal As Integer = 0
        Try
            If TabControlMain.SelectedIndex = 0 Or TabControlMain.SelectedIndex = 1 Or TabControlMain.SelectedIndex = 2 Then
                If email = True Then
                    If cmbEmailTemplate.SelectedItem Is Nothing Then
                        SystemSounds.Exclamation.Play()
                        MessageBox.Show("Please select email template.")
                        Exit Function
                    End If
                End If
                Try
                    If cmbEmailTemplate.SelectedItem IsNot Nothing Then
                        cbmVal = cmbEmailTemplate.SelectedItem.id
                        emailTemplate = DirectCast(cmbEmailTemplate.SelectedItem, CRMLS.Item).Name
                    Else
                        cbmVal = 0
                    End If
                Catch ex As System.Exception
                    ' Log the error or handle it
                    cbmVal = 0
                End Try

                If cbmVal = 1 Then
                    If txtPrice.Text.Trim.Length = 0 Then
                        SystemSounds.Exclamation.Play()
                        MessageBox.Show("Please enter offer price.")
                    Else
                        Await ProcessRecord(email, emailTemplate)
                    End If
                Else
                    Await ProcessRecord(email, emailTemplate)
                End If

            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("No record selected for sending email.")
            End If

        Catch ex As System.Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Function
    Private Async Function sendEmailAutoStatusChange(email As Boolean) As Task
        Dim emailTemplate As String = ""

        Try
            emailTemplate = DirectCast(cmbEmailTemplate.SelectedItem, CRMLS.Item).Name
            Await ProcessRecord(email, emailTemplate)

        Catch ex As System.Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Function

    Private Sub SaveNotes_Click(sender As Object, e As RoutedEventArgs)

        'If (txtAddNotes.Text.Trim.Length > 0) Then
        '    If dgDetails.SelectedItems.Count > 0 Then
        '        Dim MLSID As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(0).ToString()
        '        Dim SQLiteinsertcmd As New SQLiteCommand("insert into Notes values (@MLSListingID,@Datetime,@Emailed)")

        '        SQLiteinsertcmd.Parameters.AddWithValue("@MLSListingID", MLSID)
        '        SQLiteinsertcmd.Parameters.AddWithValue("@Datetime", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        '        SQLiteinsertcmd.Parameters.AddWithValue("@Emailed", txtAddNotes.Text)
        '        fxCommon.SQLExecuteCommand(SQLiteinsertcmd)
        '        'fxCommon.SQLExecuteQuery("insert into Notes values ('" & MLSID & "','" & Date.Now.ToString("yyyy-MM-dd HH:mm:ss") & "','" & txtAddNotes.Text & "')")
        '        txtAddNotes.Text = String.Empty
        '        ' Dim dtNotes = fxCommon.SQLExecuteReader("select Date, Notes from Notes where MLSListingID='" & MLSID & "' Order by CONVERT(datetime, Date, 120) Desc")
        '        Dim dtNotes = fxCommon.SQLExecuteReader("select Date, Notes from Notes where MLSListingID='" & MLSID & "' ORDER BY rowid DESC") ' Order by strftime('%Y-%m-%d %H:%M:%S', date) DESC")
        '        dgNotes.ItemsSource = dtNotes.DefaultView
        '    End If

        'End If
    End Sub
    Public Sub sendemail(email As Boolean, emailTemplate As String, sTel As String)
        Dim MLSListingID As String = String.Empty
        Dim offerPrice As String
        Dim arvPrice As String
        offerPrice = txtPrice.Text
        'arvPrice = txtARVPrice.Text

        If (email) Then
            Dim selMLSListingID As New List(Of String)
            If TabControlMain.SelectedIndex = 2 Then
                selMLSListingID = GetGridSelectedItems("MLSListingID", "Checked")
            End If
            If (selMLSListingID.Count > 0) Then
                Dim Response As String = fxCommon.SendEmails(selMLSListingID, offerPrice, emailTemplate)
                txtPrice.Text = String.Empty
                'txtARVPrice.Text = String.Empty
                '''If TabControlMain.SelectedIndex < 2 Then
                '''LoadDataGrid()
                '''End If
                SystemSounds.Exclamation.Play()
                MessageBox.Show(Response)
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("MLSListingID is empty. Email(s) not sent")
            End If
        End If
    End Sub
    Public Async Function ProcessRecord(email As Boolean, emailTemplate As String) As Task
        Dim sHtml As String = String.Empty
        Dim MLSListingID As String = String.Empty
        Dim sTel As String '<<agentcontact>>
        Dim ParcelNumber As String
        Dim SavedAddress As String
        Dim City As String = String.Empty
        Dim County As String
        Dim Seller As String
        Dim SellerDreLic As String
        Dim CoSeller As String
        Dim CoSellerDreLic As String
        Dim CoSellerEmail As String
        Dim CoSellerPhone As String = String.Empty
        Dim SellerEmail As String
        Dim SellerPhone As String = String.Empty
        Dim LADirect As String
        Dim LACell As String = String.Empty
        Dim Status As String = String.Empty
        Dim SQFT As String = String.Empty
        Dim SellerOffice As String
        Dim CoSellerOffice As String
        Dim SellerOfficeDreLic As String
        Dim CoSellerOfficeDreLic As String
        Dim ListPrice As String = String.Empty
        Dim ClosePrice As String = String.Empty
        Dim t00 As Object
        Dim TabName As Integer

        'Init
        Dim message As String = ""
        Dim offerPrice As String
        Dim arvPrice As String
        offerPrice = txtPrice.Text
        'arvPrice = txtARVPrice.Text

        If TabControlMain.SelectedIndex = 0 Then
            sHtml = Await WebViewMain.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
        ElseIf TabControlMain.SelectedIndex = 1 Then
            sHtml = Await WebViewComp.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
        End If


        ' Dim allElements As String = Await webview.ExecuteScriptAsync("document.querySelectorAll('*');")
        Dim sHtmlDecoded As String = System.Text.RegularExpressions.Regex.Unescape(sHtml)
        Dim hNode As HtmlAgilityPack.HtmlNode
        Dim aHTML As New HtmlAgilityPack.HtmlDocument()
        aHTML.LoadHtml(sHtmlDecoded)
        If aHTML.GetElementbyId("m_pnlDisplay") IsNot Nothing Then
            If aHTML.GetElementbyId("m_pnlDisplay").SelectSingleNode($"//*[@class='hideDragHandle container-fluid']") IsNot Nothing Then

                'Extract
                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[text()='LISTING ID: ']")
                    'MLSListingID = hNode.NextSibling.NextSibling.InnerHtml
                    MLSListingID = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    MLSListingID = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[text()='PARCEL #:']")
                    ParcelNumber = GetInnerText(hNode, "nnf")
                    ' ParcelNumber = hNode.NextSibling.NextSibling.FirstChild.InnerHtml
                Catch ex As System.Exception
                    ParcelNumber = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//input[@type='checkbox' and @onclick='Dpy.onCheck(this,event)']")
                    SavedAddress = GetInnerText(hNode, "pn")
                    If SavedAddress <> String.Empty Then
                        'SavedAddress = hNode.ParentNode.NextSibling.InnerText
                        SavedAddress = SavedAddress.Replace("  ", " ")
                        City = If(SavedAddress.Split(",").Length > 1, SavedAddress.Split(",")(1).ToString(), String.Empty)
                    End If
                Catch ex As System.Exception
                    SavedAddress = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'COUNTY:')]")
                    County = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    County = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LA:')]")
                    Seller = GetInnerText(hNode, "nnnn")
                Catch ex As System.Exception
                    Seller = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CoLA:')]")
                    CoSeller = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    CoSeller = String.Empty
                End Try

                Try
                    hNode = Nothing
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CoLA CELL:')]")
                    CoSellerPhone = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    CoSellerPhone = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CoLA EMAIL:')]")
                    CoSellerEmail = GetInnerText(hNode, "nnf")
                Catch ex As System.Exception
                    CoSellerEmail = String.Empty
                End Try

                'LO
                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LO: ')]")
                    SellerOffice = GetInnerText(hNode, "nnnnf")
                Catch ex As System.Exception
                    SellerOffice = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CoLO:&nbsp;')]")
                    CoSellerOffice = GetInnerText(hNode, "n")
                Catch ex As System.Exception
                    CoSellerOffice = String.Empty
                End Try


                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LA&nbsp;State License: ')]")
                    SellerDreLic = GetInnerText(hNode, "nnf")
                Catch ex As System.Exception
                    SellerDreLic = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CoLA&nbsp;State License: ')]")
                    CoSellerDreLic = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    CoSellerDreLic = String.Empty
                End Try

                'LO
                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LO&nbsp;State License: ')]")
                    SellerOfficeDreLic = GetInnerText(hNode, "nnf")
                Catch ex As System.Exception
                    SellerOfficeDreLic = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CoLO&nbsp;State License: ')]")
                    CoSellerOfficeDreLic = GetInnerText(hNode, "nnf")
                Catch ex As System.Exception
                    CoSellerOfficeDreLic = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'Offers Email:')]")
                    SellerEmail = GetInnerText(hNode, "nnf")
                Catch ex As System.Exception
                    SellerEmail = String.Empty
                End Try


                '1.LA CELL:
                Try
                    t00 = Nothing
                    t00 = aHTML.DocumentNode.SelectNodes("//span[contains(text(),'LA CELL')]")
                    If t00 IsNot Nothing Then
                        hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LA CELL')]")
                        LACell = GetInnerText(hNode, "nn")
                        LACell = FormatPhoneNumber(LACell)
                    Else
                        LACell = String.Empty
                    End If
                Catch ex As System.Exception
                    LACell = String.Empty
                End Try

                '2.LA DIRECT:
                Try
                    t00 = Nothing
                    t00 = aHTML.DocumentNode.SelectNodes("//span[contains(text(),'LA DIRECT')]")
                    If t00 IsNot Nothing Then
                        hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LA DIRECT:')]")
                        LADirect = GetInnerText(hNode, "nn")
                        LADirect = FormatPhoneNumber(LADirect)
                    Else
                        LADirect = String.Empty
                    End If
                Catch ex As System.Exception
                    LADirect = String.Empty
                End Try


                Try
                    If SellerEmail = String.Empty Then
                        hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LA EMAIL:')]")
                        SellerEmail = GetInnerText(hNode, "nn")
                    End If
                Catch ex As System.Exception
                    SellerEmail = String.Empty
                End Try


                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LO&nbsp;PHONE:')]")
                    SellerPhone = GetInnerText(hNode, "nn")
                    If Not String.IsNullOrEmpty(SellerPhone) Then
                        Dim extIndex As Integer = SellerPhone.IndexOf("Ext:", StringComparison.OrdinalIgnoreCase)
                        If extIndex >= 0 Then
                            ' If "Ext:" is found, take the part of the string before it.
                            SellerPhone = SellerPhone.Substring(0, extIndex).Trim()
                        End If
                        ' Now, you can perform other formatting on the (potentially truncated) phone number.                        
                        SellerPhone = FormatPhoneNumber(SellerPhone)
                    End If
                Catch ex As System.Exception
                    SellerPhone = String.Empty
                End Try

                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'STATUS:')]")
                    Status = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    Status = String.Empty
                End Try
                If Status = "Active" Then
                    Status = "A"
                ElseIf Status = "Act Under Contract" Then
                    Status = "U"
                ElseIf Status = "Pending" Then
                    Status = "P"
                ElseIf Status = "Cancelled" Or Status = "Canceled" Then
                    Status = "K"
                ElseIf Status = "Closed" Or Status = "Sold" Then
                    Status = "S"
                ElseIf Status = "Withdrawn" Then
                    Status = "W"
                ElseIf Status = "Hold" Then
                    Status = "H"
                ElseIf Status = "Expired" Then
                    Status = "X"
                ElseIf Status = "Other" Then
                    Status = "O"
                End If


                'Price Details
                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'LIST PRICE:')]")
                    ListPrice = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    ListPrice = String.Empty
                End Try
                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'CLOSE PRICE:')]")
                    ClosePrice = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    ClosePrice = String.Empty
                End Try

                'SQFT
                Try
                    hNode = aHTML.DocumentNode.SelectSingleNode("//span[contains(text(),'SQFT(src):')]")
                    SQFT = GetInnerText(hNode, "nn")
                Catch ex As System.Exception
                    SQFT = String.Empty
                End Try

                If (SellerEmail = String.Empty) Then
                    message = "Seller Email is Empty.Record not saved."
                Else
                    'sFieldList = "MLSListingID,SavedAddress,City,Seller,SellerEmail,SellerPhone,CoSeller,LACell,LADirect,Emailed,OfferPrice,Status, "
                    'sCondition = ""
                    Dim dataTable As System.Data.DataTable = fxCommon.SQLExecuteReader("select * from Property where MLSListingID='" & MLSListingID & "'")

                    If dataTable.Rows.Count = 0 Then
                        Dim SQLitecmd As New SQLiteCommand("INSERT INTO Property (MLSListingID, ParcelNumber, SavedAddress, City, County, Seller, CoSeller, SellerEmail, SellerPhone, Emailed, OfferPrice, Status, ARV, LACell, LADirect, SellerOffice, CoSellerOffice, SellerDreLic, CoSellerDreLic, SellerOfficeDreLic, CoSellerOfficeDreLic, CoSellerPhone, CoSellerEmail, TabName, ListPrice, ClosePrice, SqFt, DateAdded) VALUES (@MLSListingID, @ParcelNumber, @SavedAddress, @City, @County, @Seller, @CoSeller, @SellerEmail, @SellerPhone, @count, @txtPrice, @Status, @txtARVPrice, @LACell, @LADirect, @SellerOffice, @CoSellerOffice, @SellerDreLic, @CoSellerDreLic, @SellerOfficeDreLic, @CoSellerOfficeDreLic, @CoSellerPhone, @CoSellerEmail, @TabName, @ListPrice, @ClosePrice, @SQFT, @DateAdded)")

                        SQLitecmd.Parameters.AddWithValue("@MLSListingID", MLSListingID)
                        SQLitecmd.Parameters.AddWithValue("@ParcelNumber", ParcelNumber)
                        SQLitecmd.Parameters.AddWithValue("@SavedAddress", SavedAddress)
                        SQLitecmd.Parameters.AddWithValue("@City", City)
                        SQLitecmd.Parameters.AddWithValue("@County", County)
                        SQLitecmd.Parameters.AddWithValue("@Seller", Seller)
                        SQLitecmd.Parameters.AddWithValue("@CoSeller", CoSeller)
                        SQLitecmd.Parameters.AddWithValue("@SellerEmail", SellerEmail)
                        SQLitecmd.Parameters.AddWithValue("@SellerPhone", SellerPhone)
                        SQLitecmd.Parameters.AddWithValue("@count", 0)
                        SQLitecmd.Parameters.AddWithValue("@txtPrice", offerPrice)
                        SQLitecmd.Parameters.AddWithValue("@Status", Status)
                        SQLitecmd.Parameters.AddWithValue("@txtARVPrice", arvPrice)
                        SQLitecmd.Parameters.AddWithValue("@LACell", LACell)
                        SQLitecmd.Parameters.AddWithValue("@LADirect", LADirect)
                        SQLitecmd.Parameters.AddWithValue("@SellerOffice", SellerOffice)
                        SQLitecmd.Parameters.AddWithValue("@CoSellerOffice", CoSellerOffice)
                        SQLitecmd.Parameters.AddWithValue("@SellerDreLic", SellerDreLic)
                        SQLitecmd.Parameters.AddWithValue("@CoSellerDreLic", CoSellerDreLic)
                        SQLitecmd.Parameters.AddWithValue("@SellerOfficeDreLic", SellerOfficeDreLic)
                        SQLitecmd.Parameters.AddWithValue("@CoSellerOfficeDreLic", CoSellerOfficeDreLic)
                        SQLitecmd.Parameters.AddWithValue("@CoSellerPhone", CoSellerPhone)
                        SQLitecmd.Parameters.AddWithValue("@CoSellerEmail", CoSellerEmail)
                        SQLitecmd.Parameters.AddWithValue("@TabName", TabName)
                        SQLitecmd.Parameters.AddWithValue("@ListPrice", ListPrice)
                        SQLitecmd.Parameters.AddWithValue("@ClosePrice", ClosePrice)
                        SQLitecmd.Parameters.AddWithValue("@SQFT", SQFT)
                        SQLitecmd.Parameters.AddWithValue("@DateAdded", DateTime.Now.ToString("yyyy-MM-dd"))

                        message = fxCommon.SQLExecuteCommand(SQLitecmd)
                        If (message = "1") Then
                            message = "Property [" & SavedAddress & "] - Added to Opportunities/PipeDrive"
                            txtPrice.Text = String.Empty
                            'txtARVPrice.Text = String.Empty
                            Await PipeDriveUpload(MLSListingID)
                        Else
                            message = "Record not saved."
                        End If
                    Else
                        message = "Property [" & SavedAddress & "] - Already Exists in Opportunities"
                        MessageBox.Show(message)
                    End If
                End If
            Else
                message = "Property is not listed in this page to save."
            End If
        Else
            message = "Property is not listed in this page to save."
        End If

        If (email) Then
            If LACell IsNot "" Then
                sTel = LACell
            End If
            If SellerPhone IsNot "" Then
                sTel = SellerPhone
            End If

            Dim selMLSListingID As New List(Of String)
            If TabControlMain.SelectedIndex = 2 Then
                selMLSListingID = GetGridSelectedItems("MLSListingID", "Checked")
            Else
                If MLSListingID IsNot String.Empty Then
                    selMLSListingID.Add(MLSListingID)
                End If
            End If
            If (selMLSListingID.Count > 0) Then
                Dim Response As String = fxCommon.SendEmails(selMLSListingID, offerPrice, emailTemplate)
                txtPrice.Text = String.Empty
                'txtARVPrice.Text = String.Empty
                'If TabControlMain.SelectedIndex < 2 Then
                'LoadDataGrid()
                'End If
                SystemSounds.Exclamation.Play()
                MessageBox.Show(Response)
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("MLSListingID is empty. Email(s) not sent")
            End If
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show(message)
        End If
        'Console.WriteLine("Generated SQL Query: " & sQuery)
    End Function

    Function GetInnerText(hNode As HtmlAgilityPack.HtmlNode, pattern As String) As String
        pattern = pattern.ToUpper()
        Try
            For Each tstep In pattern
                If hNode Is Nothing Then
                    Return String.Empty
                End If

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

            Return If(hNode IsNot Nothing, hNode.InnerText, String.Empty)
        Catch ex As System.Exception
            Return String.Empty
        End Try
    End Function


    Private Sub btnAddressSearch_Click(sender As Object, e As RoutedEventArgs)
        bCodeProcessing = True
        If Trim(txtAddressSearch.Text) = "" Then
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Enter Address to proceed")
        Else
            Dim addressToSearch As String = Uri.EscapeDataString(txtAddressSearch.Text)
            Dim strURL As String = $"https://matrix.crmls.org/matrix/search/residential/detail?address={addressToSearch}"
            WebViewComp.Source = New Uri(strURL)
        End If

        'Reset bCodeProcessing
        bCodeProcessing = False
    End Sub

    Private Sub CheckBox_Checked(sender As Object, e As RoutedEventArgs)
        ' When a checkbox is checked, update the count.
        UpdateSelectedCount()
    End Sub

    Private Sub CheckBox_Unchecked(sender As Object, e As RoutedEventArgs)
        ' When a checkbox is unchecked, update the count.
        UpdateSelectedCount()
    End Sub

    ' Helper function to update the displayed count
    Private Sub UpdateSelectedCount()
        Dim checkedCount As Integer = 0
        'Dim dataView As DataView = TryCast(dgDetails.ItemsSource, DataView)

        'If dataView IsNot Nothing Then
        '    checkedCount = dataView.Cast(Of DataRowView)().Count(Function(row)
        '                                                             Return Convert.ToBoolean(row("IsSelected"))
        '                                                         End Function)
        'End If

        'lbldgStatus.Content = $"Display/Selected: {dgDetails.Items.Count}/{checkedCount}"
    End Sub

    Private Function GetGridSelectedItems(ColumnName As String, Condition As String) As List(Of String)
        Dim selectedIds As New List(Of String)

        ' Safely get the data source from the DataGrid.
        'Dim dataView As DataView = TryCast(dgDetails.ItemsSource, DataView)

        'If dataView IsNot Nothing Then
        '    ' Cast the DataView to a generic collection of DataRowView.
        '    Dim allRows = dataView.Cast(Of DataRowView)()
        '    Dim filteredRows As IEnumerable(Of DataRowView)

        '    ' Apply the filter based on the Condition parameter.
        '    If Condition = "Checked" Then
        '        ' Get only the rows where the IsSelected column is True.
        '        filteredRows = allRows.Where(Function(row) Convert.ToBoolean(row("IsSelected")))
        '    ElseIf Condition = "All" Then
        '        filteredRows = allRows
        '    Else
        '        Return selectedIds
        '    End If

        '    ' Select the specified column's value from the filtered rows.
        '    selectedIds.AddRange(filteredRows.Select(Function(row) row(ColumnName).ToString()))
        'End If

        Return selectedIds
    End Function
    ' Your existing GetGridSelectedItems function, now potentially called from other places if needed.
    ' If you only need the count, UpdateSelectedCount is more direct.
    Private Function GetGridSelectedItemsOnlyChecked(ColumnName As String, Condition As String) As List(Of String)
        Dim selectedIds As New List(Of String)
        'Dim dataView As DataView = TryCast(dgDetails.ItemsSource, DataView)

        'If dataView IsNot Nothing Then
        '    Dim checkedRows = dataView.Cast(Of DataRowView)().Where(
        '        Function(row)
        '            Return Convert.ToBoolean(row("IsSelected"))
        '        End Function)

        '    selectedIds.AddRange(checkedRows.Select(Function(row) row(ColumnName).ToString()))
        'End If

        Return selectedIds
    End Function
    Private Function GetGridSelectedItemsOld(ColumnName As String, Condition As String)
        Dim rowIndex As Integer = 0
        Dim selectedIds As New List(Of String)
        Dim targetDataGrid As DataGrid
        'If TabControlMain.SelectedItem.Header.ToString() = "Opportunities" Then
        '    targetDataGrid = dgDetails
        'Else
        '    targetDataGrid = dgDetails
        'End If

        For Each item As Object In targetDataGrid.Items
            Dim Row As DataGridRow = DirectCast(targetDataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex), DataGridRow)
            If Row IsNot Nothing Then
                Dim presenter As DataGridCellsPresenter = FindVisualChild(Of DataGridCellsPresenter)(Row)
                If presenter IsNot Nothing Then
                    Dim cell As DataGridCell = DirectCast(presenter.ItemContainerGenerator.ContainerFromIndex(0), DataGridCell)
                    If cell IsNot Nothing Then
                        If Condition = "Checked" Then
                            Dim ChkBox As System.Windows.Controls.CheckBox = FindVisualChild(Of System.Windows.Controls.CheckBox)(cell)
                            If ChkBox.IsChecked Then
                                If TypeOf item Is System.Data.DataRowView Then
                                    selectedIds.Add(item(ColumnName))
                                End If
                            End If
                        End If
                        If Condition = "All" Then
                            If TypeOf item Is System.Data.DataRowView Then
                                selectedIds.Add(item(ColumnName))
                            End If
                        End If
                    End If
                End If
            End If
            rowIndex = rowIndex + 1
        Next

        Return selectedIds
    End Function


#Region "TabComp"
    Private Sub btnDgGridGoComp_Click(sender As Object, e As RoutedEventArgs)
        If TabControlMain.SelectedIndex = 2 Or TabControlMain.SelectedIndex = 3 Then
            Dim selectedIds As New List(Of String)
            selectedIds = GetGridSelectedItems("SavedAddress", "Checked")

            If (selectedIds.Count = 1) Then
                TabControlMain.SelectedIndex = 1
                txtAddressSearch.Text = selectedIds.Item(0).ToString()
                btnAddressSearch.RaiseEvent(New RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("Please select one item to proceed.")
            End If
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("No address selected for Comp.")
            TabControlMain.SelectedIndex = 2
        End If

    End Sub
    Private Async Sub btnCompGoListingfromPipeDrive_Click(sender As Object, e As RoutedEventArgs)
        Try
            'Dim sHtml As String = Await WebViewPipe.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
            'Dim sHtmlDecoded As String = System.Text.RegularExpressions.Regex.Unescape(sHtml)
            Dim aHTML As New HtmlAgilityPack.HtmlDocument()
            Dim searchAddress As String
            'aHTML.LoadHtml(sHtmlDecoded)

            ' Find all field rows
            Dim fieldRows As HtmlAgilityPack.HtmlNodeCollection = aHTML.DocumentNode.SelectNodes("//div[@data-testid='fields-list-row']")

            If fieldRows IsNot Nothing Then
                For Each rowNode As HtmlAgilityPack.HtmlNode In fieldRows
                    ' Find the div with the field name
                    Dim nameNode As HtmlAgilityPack.HtmlNode = rowNode.SelectSingleNode(".//div[@data-testid='field-name']")
                    If nameNode IsNot Nothing AndAlso nameNode.InnerText.Trim() = "Address" Then
                        ' Find the div that contains the value, which is a direct child of the
                        ' 'fields-list-row-field-components' div.
                        Dim addressValueNode As HtmlAgilityPack.HtmlNode = rowNode.SelectSingleNode(".//div[@data-testid='fields-list-row-field-components']/div/div")
                        If addressValueNode IsNot Nothing Then
                            searchAddress = addressValueNode.InnerText.Trim()
                            Exit For ' Found the address, so exit the loop
                        End If
                    End If
                Next
            End If

            If Not String.IsNullOrEmpty(searchAddress) Then
                Try
                    searchAddress = searchAddress.Replace(vbCr, "").Replace(vbLf, "").Trim()
                    searchAddress = searchAddress.Replace("  ", " ")
                    bCodeProcessing = True
                    TabControlMain.SelectedIndex = 1
                    txtAddressSearch.Text = searchAddress
                    btnAddressSearch.RaiseEvent(New RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
                Catch ex As System.Exception
                    searchAddress = String.Empty
                End Try
            End If

            If String.IsNullOrEmpty(searchAddress) Then
                SystemSounds.Exclamation.Play()
                MessageBox.Show("Exception. Property Address not found.")
            End If
        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception. Property Address not found.")
        End Try
    End Sub
    Private Async Sub btnCompGoListingfromPrivy_Click(sender As Object, e As RoutedEventArgs)
        Try
            ' Ensure CoreWebView2 is ready
            If WebViewPrivy.CoreWebView2 Is Nothing Then
                Await WebViewPrivy.EnsureCoreWebView2Async()
            End If

            ' Give Privy page time to render address
            Await Task.Delay(1500)

            Dim script As String =
"
(function(){
    let line1 = document.querySelector('h1.address-line1');
    let line2 = document.querySelector('div.address-line2');

    if(!line1) return null;

    let a1 = line1.innerText.trim();
    let a2 = line2 ? line2.innerText.trim() : '';

    let addr = (a1 + ' ' + a2).replace(/\s+/g,' ').trim();
    return addr === '' ? null : addr;
})();
"

            Dim result As String = Await WebViewPrivy.CoreWebView2.ExecuteScriptAsync(script)

            If String.IsNullOrWhiteSpace(result) OrElse result = "null" Then
                MessageBox.Show("Property Address not found.")
                Exit Sub
            End If

            ' Remove JS quotes safely
            Dim searchAddress As String =
            System.Text.Json.JsonSerializer.Deserialize(Of String)(result)

            bCodeProcessing = True
            TabControlMain.SelectedIndex = 2
            txtAddressSearch.Text = searchAddress
            btnAddressSearch.RaiseEvent(New RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception while extracting Privy address." & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Async Sub btnCompGoListingfromMain_Click(sender As Object, e As RoutedEventArgs)
        Try
            If TabControlMain.SelectedIndex = 1 Then
                Dim sHtml As String = Await WebViewMain.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
                Dim sHtmlDecoded As String = System.Text.RegularExpressions.Regex.Unescape(sHtml)
                Dim aHTML As New HtmlAgilityPack.HtmlDocument()
                Dim searchAddress As String
                aHTML.LoadHtml(sHtmlDecoded)
                If aHTML.GetElementbyId("m_pnlDisplay") IsNot Nothing Then
                    If aHTML.GetElementbyId("m_pnlDisplay").SelectSingleNode($"//*[@class='hideDragHandle container-fluid']") IsNot Nothing Then
                        Try
                            searchAddress = aHTML.DocumentNode.SelectSingleNode("//input[@type='checkbox' and @onclick='Dpy.onCheck(this,event)']").ParentNode.NextSibling.InnerText
                            searchAddress = searchAddress.Replace("  ", " ")
                        Catch ex As System.Exception
                            searchAddress = String.Empty
                        End Try
                        bCodeProcessing = True
                        TabControlMain.SelectedIndex = 2
                        txtAddressSearch.Text = searchAddress
                        btnAddressSearch.RaiseEvent(New RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
                    End If
                End If
            End If
        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception. Listing ID not exists.")
        End Try
    End Sub
    Private Async Sub btnOGmanualEntry_Click(sender As Object, e As RoutedEventArgs)

        Try
            If WebViewog.CoreWebView2 Is Nothing Then Exit Sub
            Dim price As String = CurrentListPrice
            If Not String.IsNullOrWhiteSpace(price) Then
                price = System.Text.RegularExpressions.Regex.Replace(price, "[^\d]", "")
            End If

            ' ================= 1. CLICK "Switch to Manual Entry" =================
            Await WebViewog.CoreWebView2.ExecuteScriptAsync("
            (function(){
                var btns = document.querySelectorAll('button');
                for (var i = 0; i < btns.length; i++) {
                    if (btns[i].innerText.trim() === 'Switch to Manual Entry') {
                        btns[i].click();
                        return 'CLICKED';
                    }
                }
                return 'NOT_FOUND';
            })();
        ")

            ' wait for manual form to load
            Await Task.Delay(1200)

            ' ================= 2. SPLIT ADDRESS =================
            Dim street As String = ""
            Dim unit As String = ""
            Dim city As String = ""
            Dim state As String = ""
            Dim zip As String = ""

            Dim addrPattern As String =
    "^(.*?)(?:\s+(?:Unit|Apt|#)\s*#?(\w+))?,\s*(.*?),\s*([A-Z]{2})\s*(\d{5})$"

            Dim m = System.Text.RegularExpressions.Regex.Match(
    CurrentPropertyAddress,
    addrPattern,
    System.Text.RegularExpressions.RegexOptions.IgnoreCase
)

            If m.Success Then
                street = m.Groups(1).Value.Trim()
                unit = m.Groups(2).Value.Trim()
                city = m.Groups(3).Value.Trim()
                state = m.Groups(4).Value.Trim()
                zip = m.Groups(5).Value.Trim()
            End If

            ' ================= 3. POPULATE BASIC FIELDS =================
            Await WebViewog.CoreWebView2.ExecuteScriptAsync($"
            (function(){{
                function setVal(id, val) {{
                    var el = document.getElementById(id);
                    if (!el) return;
                    
                    var setter = Object.getOwnPropertyDescriptor(
                        HTMLInputElement.prototype, 'value'
                    ).set;

                    setter.call(el, val);
                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}

                setVal('address', '{street}');
                setVal('unit', '{unit}');
                setVal('city', '{city}');
                setVal('state', '{state}');
                setVal('zip', '{zip}');
                setVal('county', '{CurrentCounty}');
                setVal('apn', '{CurrentParcelNumber}');
                setVal('yearBuilt', '{Currentbuildyr}');
                setVal('listPrice', '{price}');
               
            }})();
        ")

            Await Task.Delay(500)

            ' ================= 4. OPEN ADVANCED INFO =================
            Await WebViewog.CoreWebView2.ExecuteScriptAsync("
            (function(){
                var btns = document.querySelectorAll('button');
                for (var i = 0; i < btns.length; i++) {
                    if (btns[i].innerText.trim() === 'Advanced Info') {
                        btns[i].click();
                        return;
                    }
                }
            })();
        ")

            Await Task.Delay(600)

            ' ================= 5. POPULATE LISTING AGENT INFO =================
            Await WebViewog.CoreWebView2.ExecuteScriptAsync($"
            (function(){{
                function setVal(id, val) {{
                    var el = document.getElementById(id);
                    if (!el) return;
                    
                    var setter = Object.getOwnPropertyDescriptor(
                        HTMLInputElement.prototype, 'value'
                    ).set;

                    setter.call(el, val);
                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}

                setVal('listingAgentName', '{CurrentLAName}');
                setVal('listingAgentEmail', '{CurrentLAEmail}');
                setVal('listingAgentPhone', '{CurrentLACell}');
                setVal('listingBrokerage', '{currentLO}');
            }})();
        ")

        Catch ex As System.Exception
            MessageBox.Show("Manual entry failed" & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Async Sub btnOGanalyse_Click(sender As Object, e As RoutedEventArgs)

        Try
            ' Ensure Matrix WebView is ready
            If WebViewComp.CoreWebView2 Is Nothing Then Exit Sub

            ' Try extracting address from current Matrix page
            Await ExtractMatrix_All_DOM(WebViewComp.CoreWebView2)

            ' If still empty, extraction failed → stay on page
            If String.IsNullOrWhiteSpace(CurrentPropertyAddress) Then
                MessageBox.Show("Property address not found on this page.")
                Exit Sub
            End If

            ' Switch to OfferGun tab ONLY when address exists
            TabControlMain.SelectedIndex = 3   ' OfferGun tab index

            ' Small UI delay
            Await Task.Delay(200)

            ' SIMPLE REFRESH
            'If WebViewog.CoreWebView2 IsNot Nothing Then
            'WebViewog.CoreWebView2.Reload()
            'End If

            ' Small delay for page load
            Await Task.Delay(500)

            ' Populate OfferGun address input
            Await UpdateOfferGunAddress()

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Failed to copy property address.")
        End Try

    End Sub
    Private Async Sub btnOfferaddress_Click(sender As Object, e As RoutedEventArgs)

        Try
            If WebViewMain.CoreWebView2 Is Nothing Then Exit Sub

            If Not Await ExtractMatrix_All_DOM(WebViewMain.CoreWebView2) Then
                MessageBox.Show("Failed to extract Matrix details.")
                Exit Sub
            End If

            ' Move to OfferGun tab (NO REFRESH)
            TabControlMain.SelectedIndex = 3
            Await Task.Delay(300)

            Await UpdateOfferGunAddress()

        Catch ex As System.Exception
            fxCommon.GenerateLog(ex)
            MessageBox.Show("Extraction error.")
        End Try

    End Sub


    Private Async Function ExtractMatrix_All_DOM(Corewv As CoreWebView2) As Task(Of Boolean)

        Dim html = Await GetWebViewHtml(Corewv)
        If String.IsNullOrWhiteSpace(html) Then Return False

        Dim doc As New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)

        '========================
        ' ADDRESS
        '========================
        Dim addrNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'d-mega')]")
        If addrNode IsNot Nothing Then
            CurrentPropertyAddress = HtmlEntity.DeEntitize(addrNode.InnerText).Trim()
        End If

        '========================
        ' LIST PRICE
        '========================
        CurrentListPrice = GetValueByLabel(doc, "LIST PRICE:")

        '========================
        ' COUNTY
        '========================
        CurrentCounty = GetValueByLabel(doc, "COUNTY:")

        CurrentParcelNumber = ExtractParcel(doc)
        CurrentLAName = ExtractLAName(doc)
        Currentbuildyr = ExtractYearBuilt(doc)
        CurrentLACell = ExtractLACell(doc)
        currentLO = ExtractLOOffice(doc)
        CurrentLAEmail = ExtractLAEmail(doc)
        ' =========================================================
        ' REGEX Extraction from Text Content (Fallback/Override)
        ' =========================================================
        Dim sText As String = Await Corewv.ExecuteScriptAsync("document.body.innerText")
        If Not String.IsNullOrEmpty(sText) Then
            Dim sTextDecoded As String = System.Text.RegularExpressions.Regex.Unescape(sText)

            ' LA Name
            Dim sectionLA As String = System.Text.RegularExpressions.Regex.Match(sTextDecoded, "LA:\s*\(.*?\)\s*(.*)").Groups(1).Value.Trim()
            If Not String.IsNullOrEmpty(sectionLA) Then CurrentLAName = sectionLA

            ' LO Office
            Dim sectionLO As String = System.Text.RegularExpressions.Regex.Match(sTextDecoded, "LO:\s*\(.*?\)\s*(.*)").Groups(1).Value.Trim()
            If Not String.IsNullOrEmpty(sectionLO) Then currentLO = sectionLO

            ' LA Cell
            Dim sectionLACell As String = System.Text.RegularExpressions.Regex.Match(
        sTextDecoded,
        "LA\s*CELL\s*[:\-]?\s*([\(\)\d\.\-\s]{7,20})",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase
    ).Groups(1).Value.Trim()

            If String.IsNullOrEmpty(sectionLACell) Then
                sectionLACell =
        System.Text.RegularExpressions.Regex.Match(
            sTextDecoded,
            "LA\s*DIRECT\s*[:\-]?\s*([\(\)\d\.\-\s]{7,20})",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        ).Groups(1).Value.Trim()
            End If
            If Not String.IsNullOrEmpty(sectionLACell) Then CurrentLACell = sectionLACell



            ' LA Email
            Dim sectionLAEmail As String = System.Text.RegularExpressions.Regex.Match(sTextDecoded, "LA EMAIL:\s*(.*)").Groups(1).Value.Trim()
            If String.IsNullOrEmpty(sectionLAEmail) Then
                sectionLAEmail = System.Text.RegularExpressions.Regex.Match(sTextDecoded, "Offers Email:\s*(.*)").Groups(1).Value.Trim()
            End If
            If Not String.IsNullOrEmpty(sectionLAEmail) Then CurrentLAEmail = sectionLAEmail
        End If


        Return True
    End Function
    Private Function ExtractParcel(doc As HtmlDocument) As String

        Dim parcelNode = doc.DocumentNode.SelectSingleNode(
        "//span[contains(@class,'wrapped-field')]//a[contains(@href,'thirdpartyformpost.aspx')]"
    )

        If parcelNode IsNot Nothing Then
            Return HtmlEntity.DeEntitize(parcelNode.InnerText).Trim()
        End If

        Return Nothing
    End Function

    Private Function ExtractYearBuilt(doc As HtmlDocument) As String

        Dim td = doc.DocumentNode.SelectSingleNode(
        "//td[contains(.,'YEAR BUILT')]"
    )

        If td Is Nothing Then Return Nothing

        Dim spans = td.SelectNodes(".//span[contains(@class,'wrapped-field')]")
        If spans Is Nothing Then Return Nothing

        For Each span As HtmlAgilityPack.HtmlNode In spans
            Dim val = span.InnerText.Trim()
            If System.Text.RegularExpressions.Regex.IsMatch(val, "^\d{4}$") Then
                Return val
            End If
        Next

        Return Nothing
    End Function
    Private Function ExtractLAName(doc As HtmlDocument) As String

        Dim td = doc.DocumentNode.SelectSingleNode(
        "//span[contains(normalize-space(),'LA:')][not(contains(normalize-space(),'CoLA'))]/ancestor::td"
    )

        If td Is Nothing Then Return Nothing

        Dim links = td.SelectNodes(".//a")
        If links Is Nothing Then Return Nothing

        For Each a In links
            Dim txt = HtmlEntity.DeEntitize(a.InnerText).Trim()

            ' skip empty, agent-id, short codes
            If txt <> "" AndAlso
           Not txt.StartsWith("G") AndAlso
           txt.Any(AddressOf Char.IsLetter) AndAlso
           txt.Length > 5 Then

                Return txt
            End If
        Next

        Return Nothing
    End Function
    Private Function ExtractLACell(doc As HtmlDocument) As String

        Dim tds = doc.DocumentNode.SelectNodes("//td")
        If tds Is Nothing Then Return Nothing

        For Each td In tds

            Dim label = td.SelectSingleNode(
            ".//span[contains(translate(normalize-space(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'LA CELL')]"
        )

            If label Is Nothing Then Continue For

            Dim spans = td.SelectNodes(".//span")
            If spans Is Nothing Then Continue For

            For Each s In spans
                Dim txt = HtmlEntity.DeEntitize(s.InnerText).Trim()

                ' phone pattern
                If System.Text.RegularExpressions.Regex.IsMatch(txt, "\d{3}[-.\s]\d{3}[-.\s]\d{4}") Then
                    Return txt
                End If
            Next
        Next

        Return Nothing
    End Function
    Private Function ExtractLAEmail(doc As HtmlDocument) As String

        Dim tds = doc.DocumentNode.SelectNodes("//td")
        If tds Is Nothing Then Return Nothing

        For Each td In tds

            Dim label = td.SelectSingleNode(
            ".//span[contains(translate(normalize-space(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'EMAIL') or
                  contains(translate(normalize-space(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'OFFERS')]"
        )

            If label Is Nothing Then Continue For

            Dim mail = td.SelectSingleNode(".//a[starts-with(@href,'mailto:')]")
            If mail IsNot Nothing Then
                Return HtmlEntity.DeEntitize(mail.InnerText).Trim()
            End If
        Next

        ' FINAL fallback (very rare cases)
        Dim fallback = doc.DocumentNode.SelectSingleNode("//a[starts-with(@href,'mailto:')]")
        If fallback IsNot Nothing Then
            Return HtmlEntity.DeEntitize(fallback.InnerText).Trim()
        End If

        Return Nothing
    End Function
    Private Function ExtractLOOffice(doc As HtmlDocument) As String

        Dim td = doc.DocumentNode.SelectSingleNode(
        "//span[contains(normalize-space(),'LO:')][not(contains(normalize-space(),'CoLO'))]/ancestor::td"
    )

        If td Is Nothing Then Return Nothing

        Dim links = td.SelectNodes(".//a")
        If links Is Nothing Then Return Nothing

        Dim best As String = Nothing

        For Each a In links
            Dim txt = HtmlEntity.DeEntitize(a.InnerText).Trim()

            ' ignore office codes
            If txt.Length > 6 AndAlso txt.Any(AddressOf Char.IsLetter) Then
                If best Is Nothing OrElse txt.Length > best.Length Then
                    best = txt
                End If
            End If
        Next

        Return best
    End Function

    Private Function GetValueByLabel(doc As HtmlAgilityPack.HtmlDocument, labelText As String) As String
        Dim labelNode = doc.DocumentNode.SelectSingleNode(
        $"//span[contains(@class,'label') and normalize-space(text())='{labelText}']"
    )

        If labelNode Is Nothing Then Return Nothing

        Dim valueNode = labelNode.ParentNode.SelectSingleNode(
        ".//span[contains(@class,'field') and not(contains(@class,'label'))]"
    )

        If valueNode Is Nothing Then Return Nothing

        Return HtmlEntity.DeEntitize(valueNode.InnerText).Trim()
    End Function
    Private Async Function GetWebViewHtml(Corewv As CoreWebView2) As Task(Of String)

        If Corewv Is Nothing Then
            Throw New InvalidOperationException("CoreWebView2 is NOT initialized.")
        End If

        Dim html As String = Await Corewv.ExecuteScriptAsync(
        "document.documentElement.outerHTML"
    )

        If String.IsNullOrWhiteSpace(html) Then Return Nothing

        html = html.Trim(""""c)
        html = System.Text.RegularExpressions.Regex.Unescape(html)

        Return html
    End Function

    Private Function CleanText(node As HtmlNode) As String
        If node Is Nothing Then Return Nothing
        Return HtmlEntity.DeEntitize(node.InnerText).Replace(vbCr, "").Replace(vbLf, "").Trim()
    End Function






    ' Returns an empty string if the value is DBNull, otherwise returns the value as a string.
    Private Function SafeStr(value As Object) As String
        If IsDBNull(value) OrElse value Is Nothing Then
            Return String.Empty
        Else
            Return value.ToString()
        End If
    End Function
    Private Async Sub btnCompGoListing_Click(sender As Object, e As RoutedEventArgs)
        Await openMLSinMain()
    End Sub
    Private Async Sub btnUpdPipeDrive_Click(sender As Object, e As RoutedEventArgs)
        Await PipeDriveUpload()
    End Sub
    Private Async Function PipeDriveUpload(Optional sMLSListingID As String = Nothing) As Task(Of String)
        'Public Shared Async Function AddDealAsync(apiToken As String, companyDomain As String) As Task(Of String)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Dim apiToken As String = xmlDoc.SelectSingleNode("/Login/PIPE/Api").InnerText
        Dim companyDomain As String = xmlDoc.SelectSingleNode("/Login/PIPE/Password").InnerText

        Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v2/deals?api_token={apiToken}"

        'Get selected item from the listbox. 
        Dim selectedItem As String = String.Empty
        If Not String.IsNullOrEmpty(sMLSListingID) Then
            selectedItem = sMLSListingID
        Else
            If TabControlMain.SelectedIndex = 2 Or TabControlMain.SelectedIndex = 3 Then
                Dim selectedIds As New List(Of String)
                selectedIds = GetGridSelectedItems("MLSListingID", "Checked")
                If selectedIds.Count > 0 Then
                    selectedItem = selectedIds(0)
                Else
                    SystemSounds.Exclamation.Play()
                    MessageBox.Show("Please select a property to upload to PipeDrive.")
                End If
            End If
        End If

        If Not String.IsNullOrEmpty(selectedItem) Then
            ' Assuming selectedItem contains the MLSListingID, we can retrieve additional values from the database
            Dim dataTable As System.Data.DataTable = fxCommon.SQLExecuteReader("SELECT * FROM Property WHERE MLSListingID='" & selectedItem & "'")
            If dataTable.Rows.Count > 0 Then
                Dim propertyRow As DataRow = dataTable.Rows(0)

                ' Prepare the data to be sent in the request body                
                Dim sSQFT As String = SafeStr(propertyRow("SqFt"))
                Dim sStatus As String = SafeStr(propertyRow("Status"))
                Dim sARV As String = SafeStr(propertyRow("ARV"))
                Dim sLPrice As String = SafeStr(propertyRow("ListPrice"))
                Dim sClosePrice As String = SafeStr(propertyRow("ClosePrice"))
                Dim sOfferPrice As String = SafeStr(propertyRow("OfferPrice"))
                Dim sLACell As String = SafeStr(propertyRow("LACell"))
                Dim sLADirect As String = SafeStr(propertyRow("LADirect"))
                Dim sSellerPhone As String = SafeStr(propertyRow("SellerPhone"))
                Dim sSellerEmail As String = SafeStr(propertyRow("SellerEmail"))
                Dim sSellerName As String = SafeStr(propertyRow("Seller"))
                Dim sPropertyAddress As String = SafeStr(propertyRow("SavedAddress"))

                'if this field is empty propertyRow("SellerPhone") then assign '000-000-0000'
                If String.IsNullOrEmpty(sSQFT) Then sSQFT = "0"
                If String.IsNullOrEmpty(sARV) Then sARV = "0"
                If String.IsNullOrEmpty(sLPrice) Then sLPrice = "0"
                If String.IsNullOrEmpty(sOfferPrice) Then sOfferPrice = "0"
                If String.IsNullOrEmpty(sClosePrice) Then sClosePrice = "0"

                ' --- Setup ---
                Dim sLPriceFormatted As String = ""
                Dim sLOfferFormatted As String = ""
                Dim sARVFormatted As String = ""
                Dim priceIndicators As String = ""

                ' A list to hold the labels for prices that are not zero
                Dim indicators As New List(Of String)()
                Dim priceValue As Long

                ' --- Logic ---

                ' 1. Process List Price (LP) & Close Price
                ' The 'AndAlso priceValue > 0' ensures we only proceed if the parsed value is not zero.
                If Long.TryParse(sLPrice.Replace(",", "").Replace("$", "").Trim(), priceValue) AndAlso priceValue > 0 Then
                    sLPrice = priceValue
                    sLPriceFormatted = "LP: $" & priceValue.ToString("N0")
                    indicators.Add("LP")
                End If
                If Long.TryParse(sClosePrice.Replace(",", "").Replace("$", "").Trim(), priceValue) AndAlso priceValue > 0 Then
                    sClosePrice = priceValue
                End If


                ' 2. Process Offer Price (OP)
                If Long.TryParse(sOfferPrice.Replace(",", "").Replace("$", "").Trim(), priceValue) AndAlso priceValue > 0 Then
                    sLOfferFormatted = "OP: $" & priceValue.ToString("N0")
                    indicators.Add("OP")
                End If

                ' 3. Process ARV
                If Long.TryParse(sARV.Replace(",", "").Replace("$", "").Trim(), priceValue) AndAlso priceValue > 0 Then
                    sARVFormatted = "ARV: $" & priceValue.ToString("N0")
                    indicators.Add("ARV")
                End If

                ' 4. Create the final indicator string from the list
                ' This will only run if at least one price was greater than zero.
                If indicators.Any() Then
                    priceIndicators = $"[{String.Join(", ", indicators)}]"
                End If

                ' 5. Format the prices for the custom fields
                If String.IsNullOrEmpty(sSellerPhone) Then
                    sSellerPhone = "000-000-0000"
                End If
                If String.IsNullOrEmpty(sLACell) Then
                    sLACell = "000-000-0000"
                End If
                If String.IsNullOrEmpty(sLADirect) Then
                    sLADirect = "000-000-0000"
                End If


                Dim data As New Dictionary(Of String, Object) From {
                    {"title", sSellerName & "-" & sPropertyAddress},
                    {"value", Convert.ToDecimal(sLPrice)},
                    {"currency", "USD"},
                    {"custom_fields", New Dictionary(Of String, Object) From {
                        {"1452ff29b9a430f6ed4fa34d03bc3976e8c8310e", sStatus},
                        {"455b5e163c481fde5281cde7f39ebfd19c15abc0", sSellerName},
                        {"7924752563aef8a024ee2aa6d623ba6ed145c142", propertyRow("MLSListingID").ToString()},
                        {"0eddebbc40f548e8d347f4c51a1d6e1e0147a2c3", sSellerEmail},
                        {"d88f20459b047449f41455f5a2306138f729efce", sSellerPhone},
                        {"4dab1f2310bc4c53db7a82410560807b76b7bc44", sLACell},
                        {"941df5800d43c6bb7b598d5782784912fabc960c", sLADirect},
                        {"12466f54cf2399bab32a4a7d3ccf60939d27ee2d", sSQFT},
                        {"8033e63c4bfc1b97dd915f424da9c4401d65d34b", sPropertyAddress},
                        {"06c982ac5f06845323d1e234ed2a5e5791cea2b2", priceIndicators},
                        {"0654fc92dfa67f1c8801866d94c10c49c1835c3d", New Dictionary(Of String, Object) From
                            {{"value", Convert.ToDecimal(sLPrice)}, {"currency", "USD"}}},
                        {"782c0afe87597d97347ead892c1d8d71371d2515", New Dictionary(Of String, Object) From
                            {{"value", Convert.ToDecimal(sOfferPrice)}, {"currency", "USD"}}},
                        {"60489922d1c93011cf367b3ff2e24a0d9834122e", New Dictionary(Of String, Object) From
                            {{"value", Convert.ToDecimal(sClosePrice)}, {"currency", "USD"}}},
                        {"5ea31681a33100a1e935015e9482c65a168e5e5b", New Dictionary(Of String, Object) From
                            {{"value", Convert.ToDecimal(sARV)}, {"currency", "USD"}}}
                    }}
                }

                Dim jsonData As String = JsonConvert.SerializeObject(data)
                Using client As New HttpClient()
                    Dim content As New StringContent(jsonData, Encoding.UTF8, "application/json")
                    Dim response As HttpResponseMessage = Await client.PostAsync(url, content)
                    Dim result As String = Await response.Content.ReadAsStringAsync()
                    fxCommon.WriteDebug(result)

                    If response.IsSuccessStatusCode Then
                        Dim pdResponse As PipedriveResponse = JsonConvert.DeserializeObject(Of PipedriveResponse)(result)
                        Dim pdDealId As Integer = pdResponse.data.id
                        Dim pdStageId As Integer = pdResponse.data.stage_id
                        Dim sMsg As String = ""

                        'Add Agent details to Pipedrive Contact
                        Dim personId As Integer? = Await PipeDriveAddPersonContact(sSellerName, sSellerEmail, sSellerPhone, sLACell, sLADirect,)
                        If personId.HasValue Then
                            ' Update Property table with pdPersonID
                            Dim updatePersonCmd As New SQLiteCommand("UPDATE Property SET pdPersonID = @pdPersonID WHERE MLSListingID = @MLSListingID")
                            updatePersonCmd.Parameters.AddWithValue("@pdPersonID", personId.Value)
                            updatePersonCmd.Parameters.AddWithValue("@MLSListingID", selectedItem)
                            fxCommon.SQLExecuteCommand(updatePersonCmd)
                        End If

                        ' Now update the database with the new IDs
                        Dim SQLitecmd As New SQLiteCommand("UPDATE Property SET PipeDriveDate = @PipeDriveDate, pdDealId = @pdDealId, pdStageId = @pdStageId WHERE MLSListingID = @MLSListingID")
                        SQLitecmd.Parameters.AddWithValue("@PipeDriveDate", DateTime.Now.ToString("yyyy-MM-dd"))
                        SQLitecmd.Parameters.AddWithValue("@pdDealId", pdDealId)
                        SQLitecmd.Parameters.AddWithValue("@pdStageId", pdStageId)
                        SQLitecmd.Parameters.AddWithValue("@MLSListingID", selectedItem)
                        fxCommon.SQLExecuteCommand(SQLitecmd)

                        ' Add to Notes
                        fxCommon.InsertNote(selectedItem, "Uploaded to Pipedrive")
                        sMsg = "Deal created successfully!"

                        'Done
                        MessageBox.Show(sMsg)
                        Return "Deal created successfully! " & result
                    Else
                        MessageBox.Show($"Failed to create deal: {response.StatusCode} {result}")
                        Return $"Failed to create deal: {response.StatusCode} {result}"
                    End If
                End Using
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("No property details found for the selected item.")
                Return String.Empty
            End If
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("No property selected.")
            Return String.Empty
        End If
    End Function

    Private Async Function openMLSinMain() As Task
        Try
            If TabControlMain.SelectedIndex = 1 Then
                Dim sHtml As String = Await WebViewComp.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
                Dim sHtmlDecoded As String = System.Text.RegularExpressions.Regex.Unescape(sHtml)
                Dim aHTML As New HtmlAgilityPack.HtmlDocument()
                aHTML.LoadHtml(sHtmlDecoded)
                If aHTML.GetElementbyId("m_pnlDisplay") IsNot Nothing Then
                    If aHTML.GetElementbyId("m_pnlDisplay").SelectSingleNode($"//*[@class='hideDragHandle container-fluid']") IsNot Nothing Then
                        Try
                            MLSListingMain = aHTML.DocumentNode.SelectSingleNode("//span[text()='LISTING ID: ']").NextSibling.NextSibling.InnerHtml
                        Catch ex As System.Exception
                            MLSListingMain = String.Empty
                        End Try
                        bCodeProcessing = True
                        Dim strURL As String = "https://matrix.crmls.org/Matrix/Default.aspx"
                        WebViewMain.Source = New Uri(strURL)
                        WebViewMain.Reload()
                        Await WebViewMain.CoreWebView2.ExecuteScriptAsync("document.body.innerText")
                        TabControlMain.SelectedIndex = 0
                    End If
                Else
                    SystemSounds.Exclamation.Play()
                    MessageBox.Show("Property is not listed in this page to go for Listing.")
                End If
            ElseIf TabControlMain.SelectedIndex = 2 Or TabControlMain.SelectedIndex = 3 Then
                'New Proc
                TabControlMain.SelectedIndex = 0
                bCodeProcessing = True
                Await openMLSinMainSync()
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("Listing ID not exists. Go Listing works from Comp and List pages")
                TabControlMain.SelectedIndex = 2
            End If

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception. Listing ID not exists.")
        End Try
    End Function
    Private Async Sub btnCompGoGoogle_Click(sender As Object, e As RoutedEventArgs)
        Try

            Dim validAddress As Boolean = False
            Dim sHtml As String = String.Empty


            If TabControlMain.SelectedIndex = 0 Then
                sHtml = Await WebViewMain.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
                validAddress = True
            ElseIf TabControlMain.SelectedIndex = 1 Then
                sHtml = Await WebViewComp.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
                validAddress = True
            ElseIf TabControlMain.SelectedIndex = 2 Or TabControlMain.SelectedIndex = 3 Then
                validAddress = False
                Dim selectedIds As New List(Of String)
                'For Each item As Object In dgDetails.Items
                '    If TypeOf item Is System.Data.DataRowView AndAlso Not IsDBNull(item("Select")) Then
                '        If (item("Select")) Then
                '            selectedIds.Add(item("SavedAddress"))
                '        End If
                '    End If
                'Next
                For Each address As String In selectedIds
                    Dim searchUrl As String = "https://www.google.com/search?q=" & Uri.EscapeDataString(address)

                    Dim psi As New ProcessStartInfo
                    psi.UseShellExecute = True
                    psi.FileName = searchUrl
                    Process.Start(psi)
                Next
                If (selectedIds.Count = 0) Then
                    SystemSounds.Exclamation.Play()
                    MessageBox.Show("Property is not selected in this page to go for Google.")
                End If
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("No valid address found in the current page.")
                validAddress = False
            End If
            If validAddress Then
                Dim SavedAddress As String
                Dim sHtmlDecoded As String = System.Text.RegularExpressions.Regex.Unescape(sHtml)
                Dim aHTML As New HtmlAgilityPack.HtmlDocument()
                aHTML.LoadHtml(sHtmlDecoded)
                If aHTML.GetElementbyId("m_pnlDisplay") IsNot Nothing Then
                    If aHTML.GetElementbyId("m_pnlDisplay").SelectSingleNode($"//*[@class='hideDragHandle container-fluid']") IsNot Nothing Then
                        Try
                            SavedAddress = aHTML.DocumentNode.SelectSingleNode("//input[@type='checkbox' and @onclick='Dpy.onCheck(this,event)']").ParentNode.NextSibling.InnerText
                        Catch ex As System.Exception
                            SavedAddress = String.Empty
                        End Try

                        Dim searchUrl As String = "https://www.google.com/search?q=" & Uri.EscapeDataString(SavedAddress)

                        Dim psi As New ProcessStartInfo
                        psi.UseShellExecute = True
                        psi.FileName = searchUrl
                        Process.Start(psi)

                    End If
                Else
                    SystemSounds.Exclamation.Play()
                    MessageBox.Show("Property is not listed in this page to go for Google.")
                End If
            End If

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception. Property is not listed in this page to go for Google.")
        End Try



    End Sub
#End Region

    '#Region "TabEmail"

    Private Sub btnEmailBrowse1_Click(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "All Files (*.*)|*.*" ' Filter for specific file types
        Dim result As Boolean? = openFileDialog.ShowDialog() ' Show the dialog and get the result

        If result = True Then
            Select Case DirectCast(sender, System.Windows.FrameworkElement).Name
                'Case "btnEmailBrowse1" : lblEmailAttach1.Content = openFileDialog.FileName
                'Case "btnEmailBrowse2" : lblEmailAttach2.Content = openFileDialog.FileName
                'Case "btnEmailBrowse3" : lblEmailAttach3.Content = openFileDialog.FileName
            End Select

            Dim saveFilePath As String = IO.Path.Combine(AppContext.BaseDirectory, "DB\Attachment\" + openFileDialog.SafeFileName + "")
            If saveFilePath <> openFileDialog.FileName Then
                File.Copy(openFileDialog.FileName, saveFilePath, True)
            End If

            saveEmailDataToDB()
        End If

    End Sub
    Private Sub LoadEmail(EmailType As String)
        dtEmail = fxCommon.SQLExecuteReader($"Select * from EmailTemplate where Name like '%{EmailType}%'")
        Dim cnt As Int16 = 1
        For Each datarow As DataRow In dtEmail.Rows
            FindLabelByName("lblMessage" & cnt).Content = datarow("DisplayName").ToString()
            FindTextBoxByName("txtEmailSubject" & cnt).Text = datarow("Subject").ToString()
            FindTextBoxByName("txtEmailMessage" & cnt).Text = datarow("Message").ToString()
            FindLabelByName("lblEmailAttach" & cnt).Content = datarow("Attachment").ToString()
            cnt = cnt + 1
        Next

        Dim dtSMS = fxCommon.SQLExecuteReader("Select * from EmailTemplate where Name like '%SMS%'")
        'For Each datarow As DataRow In dtSMS.Rows
        '    txtSMSMessage1.Text = datarow("Message").ToString()
        'Next
    End Sub

    Private Function FindTextBoxByName(name As String) As System.Windows.Controls.TextBox
        Dim textBox As System.Windows.Controls.TextBox = Nothing
        Dim gridChildren As IEnumerable(Of UIElement) = LogicalTreeHelper.GetChildren(GridEmail).OfType(Of UIElement)()

        For Each child As UIElement In gridChildren
            If TypeOf child Is System.Windows.Controls.TextBox AndAlso CType(child, System.Windows.Controls.TextBox).Name = name Then
                textBox = CType(child, System.Windows.Controls.TextBox)
                Exit For
            End If
        Next

        Return textBox
    End Function
    Private Function FindLabelByName(name As String) As System.Windows.Controls.Label
        Dim textBox As System.Windows.Controls.Label = Nothing
        Dim gridChildren As IEnumerable(Of UIElement) = LogicalTreeHelper.GetChildren(GridEmail).OfType(Of UIElement)()

        For Each child As UIElement In gridChildren
            If TypeOf child Is System.Windows.Controls.Label AndAlso CType(child, System.Windows.Controls.Label).Name = name Then
                textBox = CType(child, System.Windows.Controls.Label)
                Exit For
            End If
        Next

        Return textBox
    End Function

    Private Sub btnEmailTempSave_Click(sender As Object, e As RoutedEventArgs)
        saveEmailDataToDB()
        saveSMSDataToDB()
        MessageBox.Show("Data saved successfully.")
    End Sub

    Private Sub btnSMSTempSave_Click(sender As Object, e As RoutedEventArgs)
        saveEmailDataToDB()
        saveSMSDataToDB()
        MessageBox.Show("Data saved successfully.")
    End Sub
    Sub saveEmailDataToDB()
        Dim cnt As Int16 = 1
        Dim EmailType As String
        Dim Subject As String
        Dim Message As String
        Dim Attachment As String
        For Each datarow As DataRow In dtEmail.Rows
            EmailType = FindLabelByName("lblMessage" & cnt).Content
            Subject = FindTextBoxByName("txtEmailSubject" & cnt).Text
            Message = FindTextBoxByName("txtEmailMessage" & cnt).Text
            Attachment = FindLabelByName("lblEmailAttach" & cnt).Content
            Dim SQLitecmd As New SQLiteCommand("Update EmailTemplate Set Subject=@Subject,Message=@Message,Attachment=@Attachment where DisplayName = @EmailType")
            SQLitecmd.Parameters.AddWithValue("@Subject", Subject)
            SQLitecmd.Parameters.AddWithValue("@Message", Message)
            SQLitecmd.Parameters.AddWithValue("@Attachment", Attachment)
            SQLitecmd.Parameters.AddWithValue("@EmailType", EmailType)
            fxCommon.SQLExecuteCommand(SQLitecmd)
            EmailType = String.Empty
            Subject = String.Empty
            Message = String.Empty
            Attachment = String.Empty
            cnt = cnt + 1
        Next
        SystemSounds.Exclamation.Play()
    End Sub
    Sub saveSMSDataToDB()
        Dim Message As String
        'Message = txtSMSMessage1.Text
        Dim SQLitecmd As New SQLiteCommand("Update EmailTemplate Set Message=@Message where Name = @EmailType")
        SQLitecmd.Parameters.AddWithValue("@Message", Message)
        SQLitecmd.Parameters.AddWithValue("@EmailType", "SMS")
        fxCommon.SQLExecuteCommand(SQLitecmd)
        SystemSounds.Exclamation.Play()
    End Sub
    Private Function FindVisualChild(Of T As DependencyObject)(parent As DependencyObject) As T
        Dim count As Integer = VisualTreeHelper.GetChildrenCount(parent)

        For i As Integer = 0 To count - 1
            Dim child As DependencyObject = VisualTreeHelper.GetChild(parent, i)

            If child IsNot Nothing AndAlso TypeOf child Is T Then
                Return DirectCast(child, T)
            Else
                Dim childOfChild As T = FindVisualChild(Of T)(child)
                If childOfChild IsNot Nothing Then
                    Return childOfChild
                End If
            End If
        Next

        Return Nothing
    End Function

    Private Sub txtPrice_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        e.Handled = Not IsTextAllowed(e.Text)
    End Sub

    Private Function IsTextAllowed(text As String)
        Dim regex = New Regex("[^0-9.]+")
        Return Not regex.IsMatch(text)
    End Function
    Private Sub txtPrice_LostFocus(sender As Object, e As RoutedEventArgs)
        'Dim txtbox As TextBox = DirectCast(sender, TextBox)
        'txtbox.Text = CDbl(Val(txtbox.Text.Replace(",", ""))).ToString("N0")
        Dim txtbox As System.Windows.Controls.TextBox = TryCast(sender, System.Windows.Controls.TextBox)
        If txtbox IsNot Nothing Then
            txtbox.Text = CDbl(Val(txtbox.Text.Replace(",", ""))).ToString("N0")
        End If
    End Sub

    Private Sub txtRadius_LostFocus(sender As Object, e As RoutedEventArgs)
        SaveRadiusToXml()
    End Sub

    Private Sub rdoEmailType_Checked(sender As Object, e As RoutedEventArgs)
        Dim radioButton As RadioButton = DirectCast(sender, RadioButton)
        LoadEmail(radioButton.Content)
    End Sub

    Private Sub Window_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        newWindowWidth = e.NewSize.Width
        newWindowHeight = e.NewSize.Height
        If isSidebarHidden Then
            adjustGridWidth(0)
        Else
            adjustGridWidth(125)
        End If
    End Sub

    Function FormatPhoneNumber(ByVal phoneNumber As String) As String
        ' Regular expression pattern to match digits
        Dim pattern As String = "\d"

        ' Extract all digit characters from the phone number string
        Dim digitMatches As MatchCollection = Regex.Matches(phoneNumber, pattern)
        Dim digitsOnly As String = String.Join("", digitMatches.Cast(Of Match)().Select(Function(m) m.Value))

        ' Handle empty or all-zero inputs
        If String.IsNullOrWhiteSpace(digitsOnly) OrElse digitsOnly = "0000000000" Then
            Return "0"
        End If

        ' Check if the phone number has 10 digits (assuming North American numbers)
        If digitsOnly.Length = 10 Then
            ' Format the phone number as nnn-nnn-nnnn
            Return String.Format("{0:###-###-####}", Long.Parse(digitsOnly))
        Else
            ' Handle numbers with a different number of digits
            Return digitsOnly
        End If
    End Function

    Private Sub btnAddAppointment_Click(sender As Object, e As RoutedEventArgs)
        If TabControlMain.SelectedIndex = 2 Or TabControlMain.SelectedIndex = 3 Then
            Dim selectedIds As New List(Of String)
            selectedIds = GetGridSelectedItems("SavedAddress", "Checked")

            If (selectedIds.Count = 1) Then
                fxCommon.CreateOutlookAppointment("Property Related Appointment", selectedIds.Item(0).ToString())
            Else
                SystemSounds.Exclamation.Play()
                MessageBox.Show("Please select one item to proceed.")
            End If
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("No address selected for Comp.")
            'TabControlMain.SelectedIndex = 2
        End If

    End Sub

    'Private Sub btnCreateOffer_Click(sender As Object, e As RoutedEventArgs)
    '    WebViewPDF.Source = New Uri("https://service.car.org/CAR_UMS/pages/accessComponents/zfpConnectPage_launch.jsp?status=FOUND&description=User+found+matching+only+authenticating+source+%2f+id&newproduct=1&mobile=")
    'End Sub

    Private Sub UpdateStatus_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedIds As New List(Of String)
        selectedIds = GetGridSelectedItems("MLSListingID", "All")
        MLSListingStatus = String.Join(",", selectedIds.Select(Function(id) $"{id}"))
        'WebViewStatus.Source = New Uri("https://matrix.crmls.org/Matrix/Default.aspx")
    End Sub

    Private Sub UpdateStatusMLSID(iMlsid As String)
        MLSListingStatus = iMlsid
        'WebViewStatus.Source = New Uri("https://matrix.crmls.org/Matrix/Default.aspx")
    End Sub

    Private Async Function GetStatusData(CoreWV As CoreWebView2) As Task
        Dim htmlCode As String
        Dim aHTML As New HtmlDocument()
        Dim rh As String = Await CoreWV.ExecuteScriptAsync("document.documentElement.outerHTML;")
        Dim dd As String = System.Text.RegularExpressions.Regex.Unescape(rh)
        dd = dd.Remove(0, 1)
        htmlCode = dd.Remove(dd.Length - 1, 1)
        aHTML.LoadHtml(htmlCode)
        Dim iCol As Integer

        'Pipedrive Deal - Stage ID and Status Name
        Dim stageDict As New Dictionary(Of String, String) From {
                        {"6", "Opportunity"},
                        {"7", "Contact Made"},
                        {"8", "Offer Made"},
                        {"9", "Front End Escrow"},
                        {"10", "Dispo"},
                        {"11", "Back End Escrow"},
                        {"12", "Closed"},
                        {"13", "Lost Pending"},
                        {"14", "Lost Close"}
                    }

        Try
            Dim table As HtmlNode = aHTML.DocumentNode.SelectNodes("//table[contains(@class, 'displayGrid')]")(0)
            Dim dt As New System.Data.DataTable()

            ' Extract headers from the table
            For Each headerCell As HtmlNode In table.SelectNodes(".//thead/tr/th")
                If headerCell.InnerText.Trim() <> "&nbsp;" Then
                    dt.Columns.Add(headerCell.InnerText.Trim())
                Else
                    dt.Columns.Add("Empty" & iCol)
                End If
                iCol += 1
            Next

            ' Extract rows from the table
            For Each row As HtmlNode In table.SelectNodes(".//tbody/tr")
                Dim newRow As DataRow = dt.NewRow()
                For i As Integer = 0 To row.SelectNodes("td").Count - 1
                    newRow(i) = row.SelectNodes("td")(i).InnerText.Trim()
                Next
                dt.Rows.Add(newRow)
            Next

            For Each row In dt.Rows
                Dim mlsId As String = row("Listing Id").ToString()
                Dim newStatus As String = row("S").ToString()

                ' **Step 1: Retrieve existing data from the database**
                ' This is the original, but vulnerable, approach.
                Dim existingDataQuery As String = "SELECT Status, pdStageID, pdStageName, pdDealId, pdStageHistory FROM Property WHERE MLSListingID='" & mlsId & "'"
                Dim existingDt As System.Data.DataTable = fxCommon.SQLExecuteReader(existingDataQuery)

                Dim prevStatus As String = ""
                Dim prevStageId As String = ""
                Dim prevStageName As String = ""
                Dim pdDealId As String = ""
                Dim pdStageHistory As String = ""
                Dim currStageId As String = ""

                If existingDt.Rows.Count > 0 Then
                    Dim existingRow As DataRow = existingDt.Rows(0)
                    prevStatus = GetSafeValue(existingRow, "Status")
                    prevStageId = GetSafeValue(existingRow, "pdStageID")
                    currStageId = prevStageId
                    prevStageName = GetSafeValue(existingRow, "pdStageName")
                    pdDealId = GetSafeValue(existingRow, "pdDealId")
                    pdStageHistory = GetSafeValue(existingRow, "pdStageHistory")
                End If

                ' **Step 2: Check if an update is needed (optional, but good practice)**
                LogMessage(message:="MLS ID: " & mlsId & " | Prev Status: " & prevStatus & " | New Status: " & newStatus & " | Prev Stage ID: " & prevStageId & " | PD Deal ID: " & pdDealId)
                If prevStatus <> newStatus Then

                    'Set Pipedrive Stage ID and Status
                    If prevStatus = "A" And newStatus = "S" Then 'Active to Sold
                        currStageId = "12"
                        If prevStageId = "13" Or prevStageId = "14" Then currStageId = prevStageId
                        Await PipedriveUpdateDealStatusAndStage(pdDealId, "lost", currStageId)
                        Await PipedriveAddNotesToDeal(pdDealId, "[CRMLS AutoUpdate] - Status change: Active to Sold")
                        pdStageHistory = prevStageId
                    ElseIf prevStatus = "P" And newStatus = "S" Then 'Pending to Sold
                        currStageId = "14"
                        Await PipedriveUpdateDealStatusAndStage(pdDealId, "lost", currStageId) 'Lost Closed
                        Await PipedriveAddNotesToDeal(pdDealId, "[CRMLS AutoUpdate] - Status change: Pending to Lost Sold")
                        pdStageHistory = prevStageId
                    ElseIf prevStatus = "A" And newStatus = "P" Then 'Active to Pending
                        currStageId = "13"
                        Await PipedriveUpdateDealStatusAndStage(pdDealId, "open", currStageId) 'Pending
                        Await PipedriveAddNotesToDeal(pdDealId, "[CRMLS AutoUpdate] - Status change: Active to Pending")
                        pdStageHistory = prevStageId
                    ElseIf prevStatus = "P" And newStatus = "A" Then 'Pending to Active
                        Await PipedriveUpdateDealStatusAndStage(pdDealId, "open", pdStageHistory) 'Previous Stage Made
                        Await PipedriveAddNotesToDeal(pdDealId, "[CRMLS AutoUpdate] - Status change: Pending to Active")
                        pdStageHistory = prevStageId
                    End If

                    'Update the database with the new status
                    Dim SQLitecmd As SQLiteCommand
                    If newStatus = "S" Then
                        SQLitecmd = New SQLiteCommand("UPDATE Property Set Status=@Status, ClosePrice=@ClosePrice, SqFt=@SQFT, pdStageHistory=@pdStageHistory, pdStageID=@pdStageID, pdStageName=@pdStageName WHERE MLSListingID=@MLSListingID")
                        SQLitecmd.Parameters.AddWithValue("@ClosePrice", row("L/C Price").ToString())
                    Else
                        SQLitecmd = New SQLiteCommand("UPDATE Property Set Status=@Status, ListPrice=@ListPrice, SqFt=@SQFT, pdStageHistory=@pdStageHistory, pdStageID=@pdStageID, pdStageName=@pdStageName WHERE MLSListingID=@MLSListingID")
                        SQLitecmd.Parameters.AddWithValue("@ListPrice", row("L/C Price").ToString())
                    End If
                    SQLitecmd.Parameters.AddWithValue("@Status", newStatus)
                    SQLitecmd.Parameters.AddWithValue("@SQFT", CInt(Math.Truncate(Val(row("Sqft").ToString()))))
                    SQLitecmd.Parameters.AddWithValue("@MLSListingID", mlsId)
                    SQLitecmd.Parameters.AddWithValue("@pdStageHistory", pdStageHistory)
                    SQLitecmd.Parameters.AddWithValue("@pdStageID", currStageId)
                    SQLitecmd.Parameters.AddWithValue("@pdStageName", stageDict(currStageId))
                    Dim rowsAffected As Integer = fxCommon.SQLExecuteCommand(SQLitecmd)

                End If
            Next

            ' Check for "Next" link and click if available
            Dim nextLinkNode = aHTML.DocumentNode.SelectSingleNode("//span[@id='m_upPaging']//span[@class='pagingLinks']/a[contains(text(),'Next') and not(@disabled)]")
            If nextLinkNode IsNot Nothing AndAlso nextLinkNode.Attributes("href") IsNot Nothing Then
                ' Click the Next link using JavaScript
                Dim jsClickNext As String = nextLinkNode.Attributes("href").Value.Replace("javascript:", "")
                Await CoreWV.ExecuteScriptAsync(jsClickNext)
                ' Wait for the next page to load
                Await Task.Delay(2000)
                ' Recursively call to process the next page
                Await GetStatusData(CoreWV)
                Return
            End If

            SystemSounds.Exclamation.Play()
            MessageBox.Show("Status updated successfully.")
            LoadDataGrid()

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception. Error is updating status.")
        End Try
    End Function
    ' Helper function to safely get a value from a DataRow
    Private Function GetSafeValue(ByVal row As DataRow, ByVal columnName As String) As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) Then
            Dim value = row(columnName)
            If value IsNot Nothing AndAlso Not Convert.IsDBNull(value) Then
                Return value.ToString()
            End If
        End If
        Return ""
    End Function
    Private Sub btnDgGridDelete_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedIds As New List(Of String)
        selectedIds = GetGridSelectedItems("MLSListingID", "Checked")
        Dim idList As String = String.Join(",", selectedIds.Select(Function(id) $"'{id}'"))
        Dim count = fxCommon.SQLExecuteQuery($"DELETE FROM Property WHERE MLSListingID IN ({idList})")
        fxCommon.SQLExecuteQuery($"DELETE FROM Notes WHERE MLSListingID IN ({idList})")
        If (count = "0") Then
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Record(s) not deleted.")
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Record(s) deleted successfully.")
        End If
        LoadDataGrid()
    End Sub
    Private Async Function openMLSinMainSync() As Task
        Try
            Dim selectedIds As New List(Of String)
            Dim strURL As String
            selectedIds = GetGridSelectedItems("MLSListingID", "Checked")
            MLSListingMain = String.Join(",", selectedIds.Select(Function(id) $"{id}"))
            MLSListingSub = MLSListingMain
            bCodeProcessing = True

            strURL = "https://matrix.crmls.org/matrix/mymatrix/home"
            WebViewMain.Source = New Uri(strURL)
            '  WebViewMain.Reload()
            'strWebContent = Await WebViewMain.CoreWebView2.ExecuteScriptAsync("document.body.innerText")
            Threading.Thread.Sleep(1500)
            Await Task.Delay(1000)
            'strURL = "https://matrix.crmls.org/Matrix/default.aspx"
            'WebViewMain.Source = New Uri(strURL)
            'WebViewMain.Reload()
            'strWebContent = Await WebViewMain.CoreWebView2.ExecuteScriptAsync("document.body.innerText")
            'Threading.Thread.Sleep(1500)

        Catch ex As System.Exception
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Exception. Listing ID not exists.")
        End Try
    End Function


    Private Sub btnSMS_Click(sender As Object, e As RoutedEventArgs)
        'If (txtSMS.Text.Trim.Length > 0) Then
        '    If dgDetails.SelectedItems.Count > 0 Then
        '        'For future reference
        '        '<a href="sms:1234567890?body=YourMessageHere">Send SMS</a>
        '        'Copy Text
        '        If txtSMS IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtSMS.Text) Then
        '            Clipboard.SetText(txtSMS.Text)
        '        End If

        '        'Save to Notes
        '        Dim MLSID As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(0).ToString()
        '        Dim sNotes As String = "SMS: " + txtSMS.Text.Trim()
        '        sNotes = System.Text.RegularExpressions.Regex.Replace(sNotes, "\s+", " ")
        '        ProcSaveNotes(MLSID, sNotes)
        '    End If
        'End If
    End Sub

    Private Sub btnCopyNumber_Click(sender As Object, e As RoutedEventArgs)
        'If dgDetails.SelectedItems.Count > 0 Then
        '    Dim sTel As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(5).ToString()
        '    sTel = sTel.Replace("-", "").Replace(".", "").Replace(" ", "")
        '    Clipboard.SetText(sTel)
        'End If
    End Sub
    Private Sub btnCall_Click(sender As Object, e As RoutedEventArgs)
        'If dgDetails.SelectedItems.Count > 0 Then
        '    Dim sTel As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(5).ToString()
        '    Dim startexternal As New Process()
        '    sTel = sTel.Replace("-", "").Replace(".", "").Replace(" ", "")
        '    Dim uri As New Uri("tel:" + sTel)
        '    Dim psi As New ProcessStartInfo()
        '    psi.UseShellExecute = True
        '    psi.FileName = uri.AbsoluteUri
        '    Process.Start(psi)

        '    'Save to notes
        '    Dim MLSID As String = DirectCast(dgDetails.SelectedItems(0), System.Data.DataRowView).Row.ItemArray(0).ToString()
        '    Dim sNotes As String = "Call: " + sTel
        '    ProcSaveNotes(MLSID, sNotes)
        'End If
    End Sub

    Sub ProcSaveNotes(MLSID As String, sNotes As String)
        Dim SQLiteinsertcmd As New SQLiteCommand("insert into Notes values (@MLSListingID,@Datetime,@Notes)")
        SQLiteinsertcmd.Parameters.AddWithValue("@MLSListingID", MLSID)
        SQLiteinsertcmd.Parameters.AddWithValue("@Datetime", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        SQLiteinsertcmd.Parameters.AddWithValue("@Notes", sNotes)
        fxCommon.SQLExecuteCommand(SQLiteinsertcmd)

        Dim dtNotes = fxCommon.SQLExecuteReader("select Date, Notes from Notes where MLSListingID='" & MLSID & "'  ORDER BY rowid DESC") ' Order by strftime('%Y-%m-%d %H:%M:%S', date) DESC")
        'dgNotes.ItemsSource = dtNotes.DefaultView
    End Sub

    Private Sub rdoAll_Checked(sender As Object, e As RoutedEventArgs)
        LoadDataGrid()
    End Sub

    Private Sub rdoLoadDatawithfilter()
        Dim tabName As String = TabControlMain.SelectedItem.Header.ToString()
        If tabName = "Opportunities" Then
            LoadDataGrid()
        End If
    End Sub
    Private Sub rdoActive_Checked(sender As Object, e As RoutedEventArgs)
        rdoLoadDatawithfilter()
    End Sub

    Private Sub rdoPending_Checked(sender As Object, e As RoutedEventArgs)
        rdoLoadDatawithfilter()
    End Sub

    Private Sub rdoSold_Checked(sender As Object, e As RoutedEventArgs)
        rdoLoadDatawithfilter()
    End Sub

    Private Sub rdoHold_Checked(sender As Object, e As RoutedEventArgs)
        rdoLoadDatawithfilter()
    End Sub

    Private Sub rdoOth_Checked(sender As Object, e As RoutedEventArgs)
        rdoLoadDatawithfilter()
    End Sub


    Private Sub btnMainBack_Click(sender As Object, e As RoutedEventArgs)
        If WebViewMain.CanGoBack Then
            WebViewMain.GoBack()
        End If
    End Sub
    Private Sub btnMainFwd_Click(sender As Object, e As RoutedEventArgs)
        If WebViewMain.CanGoForward Then
            WebViewMain.GoForward()
        End If
    End Sub

    'Privy
    Private Sub btnPrivyBack_Click(sender As Object, e As RoutedEventArgs)
        If WebViewPrivy.CanGoBack Then
            WebViewPrivy.GoBack()
        End If
    End Sub
    Private Sub btnPrivyFwd_Click(sender As Object, e As RoutedEventArgs)
        If WebViewPrivy.CanGoForward Then
            WebViewPrivy.GoForward()
        End If
    End Sub

    'Pipe
    Private Sub btnPipeBack_Click(sender As Object, e As RoutedEventArgs)
        'If WebViewPipe.CanGoBack Then
        '    WebViewPipe.GoBack()
        'End If
    End Sub
    Private Sub btnPipeFwd_Click(sender As Object, e As RoutedEventArgs)
        'If WebViewPipe.CanGoForward Then
        '    WebViewPipe.GoForward()
        'End If
    End Sub
    Private Sub btnCompBack_Click(sender As Object, e As RoutedEventArgs)
        If WebViewComp.CanGoBack Then
            WebViewComp.GoBack()
        End If
    End Sub

    Private Sub btnMoveToLeads_Click(sender As Object, e As RoutedEventArgs)
        ' Your code for the button click event

        ' Get the selected MLSListingIDs
        Dim selectedIds As New List(Of String)
        selectedIds = GetGridSelectedItems("MLSListingID", "Checked")
        Dim idList As String = String.Join(",", selectedIds.Select(Function(id) $"'{id}'"))

        ' Update the TabName column to '1' for the selected MLSListingIDs
        Dim updateQuery As String = $"UPDATE Property SET TabName = '1' WHERE MLSListingID IN ({idList})"
        Dim updateCount = fxCommon.SQLExecuteQuery(updateQuery)

        ' Check if any rows were updated
        If updateCount > 0 Then
            SystemSounds.Exclamation.Play()
            MessageBox.Show("Record Moved.")
        Else
            SystemSounds.Exclamation.Play()
            MessageBox.Show("No records updated.")
        End If

        ' Load the data grid
        '''LoadDataGrid()

    End Sub

    Private Sub btnCompFwd_Click(sender As Object, e As RoutedEventArgs)
        If WebViewComp.CanGoForward Then
            WebViewComp.GoForward()
        End If
    End Sub



    Private Sub btnExportLeads_Click(sender As Object, e As RoutedEventArgs)
        ExportToExcel()
    End Sub

    Sub ExportToExcel()
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Add(Type.Missing)
        oSheet = oBook.Worksheets(1)

        For Each dc In dataTable.Columns
            colIndex = colIndex + 1
            oSheet.Cells(1, colIndex) = dc.ColumnName
        Next

        'Export the rows to excel file
        For Each dr In dataTable.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In dataTable.Columns
                colIndex = colIndex + 1
                oSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next
        oSheet.Columns.AutoFit()
        'Save file in final path
        oExcel.DisplayAlerts = False

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*"
        If saveFileDialog.ShowDialog() = True Then
            oBook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Exported successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        End If

        oBook.Close(False, Type.Missing, Type.Missing)
        oExcel.Quit()
        ' Release resources
        ReleaseComObject(oBook)
        ReleaseComObject(oExcel)
    End Sub

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


    Private Sub ClosePopup_Click(sender As Object, e As RoutedEventArgs)
        'popup.IsOpen = False
    End Sub

    Private Async Sub MainAddToList_ClickAsync(sender As Object, e As RoutedEventArgs)
        Await sendEmailAndSave(False)
        ProcLoadNotesDG()
    End Sub


    Private Sub OnAutoGeneratingColumn(sender As Object, e As DataGridAutoGeneratingColumnEventArgs)
        ' Example: Modify the column header name
        e.Column.Header = e.PropertyName

        ' Example: To cancel the generation of a specific column
        'If e.PropertyName = "PropertyNameToExclude" Then
        ' e.Cancel = True
        'End If

        If e.PropertyName = "DateAdded" Then
            Dim column As DataGridTextColumn = DirectCast(e.Column, DataGridTextColumn)
            column.Binding.StringFormat = "dd-MMM-yy"
            e.Column.Width = New DataGridLength(1, DataGridLengthUnitType.Star)
        End If

        ' You can also modify the column further, for example, setting the width
        ' e.Column.Width = New DataGridLength(1, DataGridLengthUnitType.Star)
        ' e.Column.Width = New DataGridLength(1, DataGridLengthUnitType.Star)
    End Sub
    Private Sub btnExit_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub



    Private Sub cmbEmailAccount_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(_settingsFilePath)
        Dim DefaultEmailAccount As XmlNode = xmlDoc.SelectSingleNode("/Login/DefaultEmailAccount")
        If DefaultEmailAccount IsNot Nothing Then
            DefaultEmailAccount.InnerText = cmbEmailAccount.SelectedValue
            xmlDoc.Save(_settingsFilePath)
        End If
    End Sub

    Private Sub btnTabOpportunities_Click(sender As Object, e As RoutedEventArgs)
        'FloatTabItem(tabOpportunities, sender)
    End Sub

    Private Sub btnTabMain_Click(sender As Object, e As RoutedEventArgs)
        FloatTabItem(tabMain, sender)
    End Sub
    Private Sub btnTabAnalyze_Click(sender As Object, e As RoutedEventArgs)
        FloatTabItem(tabAnalyze, sender)
    End Sub
    Private Sub btnTabPrivy_Click(sender As Object, e As RoutedEventArgs)
        FloatTabItem(tabPrivy, sender)
    End Sub


    Private Sub FloatTabItem(tabItem As TabItem, button As System.Windows.Controls.Button)
        Dim floatingWindow As FloatingTabWindow = Application.Current.Windows.OfType(Of FloatingTabWindow)().FirstOrDefault()
        If floatingWindow Is Nothing Then
            floatingWindow = New FloatingTabWindow()
        End If

        For i As Integer = 0 To TabControlMain.Items.Count - 1
            If TabControlMain.Items(i).Equals(tabItem) Then

                If floatingWindow.Visibility = Visibility.Visible Then
                    floatingWindow.Close()
                    floatingWindow = New FloatingTabWindow()
                    If DragTabIndex <= i Then
                        i = i + 1
                    End If
                End If

                DragTabIndex = i
                DragButton = button
                DragButton.Visibility = Visibility.Hidden
                TabControlMain.SelectedIndex = i + 1
            End If
        Next

        TabControlMain.Items.Remove(tabItem)
        floatingWindow.FloatingTabControl.Items.Add(tabItem)
        floatingWindow.Show()
    End Sub

    Private Sub txtSearchOpportunities_TextChanged(sender As Object, e As TextChangedEventArgs)
        'Dim searchText As String = txtSearchOpportunities.Text.Trim().ToLower()

        ' Retrieve the original DataTable from the DataGrid's ItemsSource
        'Dim originalDataTable As System.Data.DataTable = TryCast(dgDetails.Tag, System.Data.DataTable)

        'If originalDataTable Is Nothing Then
        '    ' Save the original DataTable to the Tag property for restoring later
        '    originalDataTable = TryCast(dgDetails.ItemsSource, System.Data.DataView)?.Table
        '    If originalDataTable IsNot Nothing Then
        '        dgDetails.Tag = originalDataTable
        '    End If
        'End If

        'If originalDataTable IsNot Nothing Then
        '    Dim filteredView As System.Data.DataView = originalDataTable.DefaultView

        '    If String.IsNullOrWhiteSpace(searchText) Then
        '        ' Clear the RowFilter to show all rows when the search box is cleared
        '        filteredView.RowFilter = String.Empty
        '    Else
        '        ' Apply the filter to show rows matching the search text
        '        filteredView.RowFilter = $"SavedAddress LIKE '%{searchText}%' OR Seller LIKE '%{searchText}%'"
        '    End If

        '    dgDetails.ItemsSource = filteredView
        'End If
    End Sub



    ' Dummy handler for rdoFam Checked event
    Private Sub rdoFam_Checked(sender As Object, e As RoutedEventArgs)
        txtSettingsChanges()
    End Sub

    ' Dummy handler for rdoCondo Checked event
    Private Sub rdoCondo_Checked(sender As Object, e As RoutedEventArgs)
        txtSettingsChanges()
    End Sub

    ' Dummy handler for txtSqFtRange PreviewTextInput event
    Private Sub txtSqFtRange_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        ' TODO: Implement logic if needed        
    End Sub
    ' Dummy handler for txtSqFtRange LostFocus event
    Private Sub txtSqFtRange_LostFocus(sender As Object, e As RoutedEventArgs)
        txtSettingsChanges()
    End Sub
    Private Sub chkAutoFill_Checked(sender As Object, e As RoutedEventArgs)
        'Add code to set a global variable or setting to indicate that auto-fill is enabled
        isAutofillEnabled = True
    End Sub
    Private Sub chkAutoFill_Unchecked(sender As Object, e As RoutedEventArgs)
        ' Save the disabled state (False) to your application's settings.
        isAutofillEnabled = False
    End Sub

    Private Sub chkAutoFill0_Checked(sender As Object, e As RoutedEventArgs)
        chkAutoFill.IsChecked = True
    End Sub

    Private Sub chkAutoFill0_Unchecked(sender As Object, e As RoutedEventArgs)
        chkAutoFill.IsChecked = False
    End Sub

    Private Sub rdoFam0_Checked(sender As Object, e As RoutedEventArgs)
        rdoFam.IsChecked = True
        rdoCondo.IsChecked = False
    End Sub

    Private Sub rdoCondo0_Checked(sender As Object, e As RoutedEventArgs)
        rdoFam.IsChecked = False
        rdoCondo.IsChecked = True
    End Sub

    Private Sub txtRadius0_TextChanged(sender As Object, e As TextChangedEventArgs)
        If txtRadius IsNot Nothing Then
            txtRadius.Text = txtRadius0.Text
        End If
    End Sub

    Private Sub txtSqFtRange0_TextChanged(sender As Object, e As TextChangedEventArgs)
        If txtSqFtRange IsNot Nothing Then
            txtSqFtRange.Text = txtSqFtRange0.Text
        End If
    End Sub

    Private Sub txtSettingsChanges()
        Try
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(_settingsFilePath)

            ' Save Miles (Radius)
            Dim milesNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Miles")
            If milesNode IsNot Nothing Then
                milesNode.InnerText = txtRadius.Text.Trim()
            End If

            ' Save Sqft Range
            Dim sqftNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Sqft")
            If sqftNode IsNot Nothing Then
                sqftNode.InnerText = txtSqFtRange.Text.Trim()
            End If

            ' Save Subcategory
            Dim subcatNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Subcat")
            If subcatNode IsNot Nothing Then
                If rdoFam.IsChecked Then
                    subcatNode.InnerText = "Single Family"
                ElseIf rdoCondo.IsChecked Then
                    subcatNode.InnerText = "Condo/TownHouse"
                End If
            End If

            xmlDoc.Save(_settingsFilePath)
        Catch ex As System.Exception
            MessageBox.Show("Error saving settings: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadSettingsFromXml()
        Try
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(_settingsFilePath)

            ' Load Miles (Radius)
            Dim milesNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Miles")
            If milesNode IsNot Nothing Then
                txtRadius.Text = milesNode.InnerText
            End If

            ' Load Sqft Range
            Dim sqftNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Sqft")
            If sqftNode IsNot Nothing Then
                txtSqFtRange.Text = sqftNode.InnerText
            End If

            ' Load Subcategory
            Dim subcatNode As XmlNode = xmlDoc.SelectSingleNode("/Login/MLSSETTINGS/Subcat")
            If subcatNode IsNot Nothing Then
                If subcatNode.InnerText = "Single Family" Then
                    rdoFam.IsChecked = True
                ElseIf subcatNode.InnerText = "Condo/TownHouse" Then
                    rdoCondo.IsChecked = True
                End If
            End If
        Catch ex As System.Exception
            MessageBox.Show("Error loading settings: " & ex.Message)
        End Try
    End Sub

    Private Sub tabSettings_GotFocus(sender As Object, e As RoutedEventArgs)
        LoadSettings()
    End Sub

    Private Sub btnSaveSettings_Click(sender As Object, e As RoutedEventArgs)
        SaveSettings()
        MessageBox.Show("Settings saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub

    Private Sub LoadSettings()
        Try
            If Not File.Exists(_settingsFilePath) Then
                MessageBox.Show("Settings file not found: " & _settingsFilePath, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            Dim xdoc As XDocument = XDocument.Load(_settingsFilePath)
            Dim loginElement As XElement = xdoc.Element("Login")

            If loginElement Is Nothing Then Return

            ' Helper function to safely get a node's value.
            ' Using a function like this prevents errors if an XML tag is missing.
            Dim GetNodeValue = Function(parent As XElement, childName As String) As String
                                   If parent IsNot Nothing AndAlso parent.Element(childName) IsNot Nothing Then
                                       Return parent.Element(childName).Value
                                   End If
                                   Return String.Empty
                               End Function

            ' Load Credentials
            txtCrmlsUser.Text = GetNodeValue(loginElement.Element("CRMLS"), "UserName")
            txtCrmlsPassword.Text = GetNodeValue(loginElement.Element("CRMLS"), "Password")
            txtPrivyUser.Text = GetNodeValue(loginElement.Element("PRIVY"), "UserName")
            txtPrivyPassword.Text = GetNodeValue(loginElement.Element("PRIVY"), "Password")
            'txtPipeUser.Text = GetNodeValue(loginElement.Element("PIPE"), "UserName")
            'txtPipePassword.Text = GetNodeValue(loginElement.Element("PIPE"), "Password")
            'txtPipeApiKey.Text = GetNodeValue(loginElement.Element("PIPE"), "Api")
            'txtPipeCompanyName.Text = GetNodeValue(loginElement.Element("PIPE"), "Company")

            ' Load MLS Status
            Dim mlsStatusElement = loginElement.Element("MLSStatus")
            txtMlsActive.Text = GetNodeValue(mlsStatusElement, "Active")
            txtMlsSold.Text = GetNodeValue(mlsStatusElement, "Sold")
            txtMlsHold.Text = GetNodeValue(mlsStatusElement, "Hold")
            txtMlsPending.Text = GetNodeValue(mlsStatusElement, "Pending")
            txtMlsOther.Text = GetNodeValue(mlsStatusElement, "Other")

            ' Load MLS Settings
            Dim mlsSettingsElement = loginElement.Element("MLSSETTINGS")
            txtMlsMiles.Text = GetNodeValue(mlsSettingsElement, "Miles")
            txtMlsSubcat.Text = GetNodeValue(mlsSettingsElement, "Subcat")
            txtMlsSqft.Text = GetNodeValue(mlsSettingsElement, "Sqft")

            ' Load Default Email
            txtDefaultEmail.Text = GetNodeValue(loginElement, "DefaultEmailAccount")

            'Defaults
            Dim mlsDeafaultsSettingsElement = loginElement.Element("DefaultsSettings")
            Dim sTemp As String = GetNodeValue(mlsDeafaultsSettingsElement, "EmailStatusChangeUpdates")
            If sTemp = "True" Then
                FlagAutoStatusChangeEmailsEnabled = True
                chkAutoStatusChangeEmails.IsChecked = True
            Else
                FlagAutoStatusChangeEmailsEnabled = False
                chkAutoStatusChangeEmails.IsChecked = False
            End If
        Catch ex As System.Exception
            MessageBox.Show("Failed to load settings: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
    Private Sub tabSettings_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If tabSettings.IsSelected Then
            LoadSettings()
        End If
    End Sub


    Private Sub SaveSettings()
        Try
            If Not File.Exists(_settingsFilePath) Then
                MessageBox.Show("Cannot save. Settings file not found: " & _settingsFilePath,
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            Dim xdoc As XDocument = XDocument.Load(_settingsFilePath)
            Dim loginElement As XElement = xdoc.Element("Login")

            If loginElement Is Nothing Then
                MessageBox.Show("Invalid XML structure. <Login> not found.",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            ' ==============================
            ' HELPERS
            ' ==============================
            Dim GetOrCreate =
            Function(parent As XElement, name As String) As XElement
                Dim el = parent.Element(name)
                If el Is Nothing Then
                    el = New XElement(name)
                    parent.Add(el)
                End If
                Return el
            End Function

            Dim SetNodeValue =
            Sub(parent As XElement, childName As String, value As String)
                If parent Is Nothing Then Exit Sub
                Dim child = parent.Element(childName)
                If child Is Nothing Then
                    parent.Add(New XElement(childName, value))
                Else
                    child.Value = value
                End If
            End Sub

            ' ==============================
            ' CREDENTIALS
            ' ==============================
            Dim crmls = GetOrCreate(loginElement, "CRMLS")
            SetNodeValue(crmls, "UserName", txtCrmlsUser.Text.Trim())
            SetNodeValue(crmls, "Password", txtCrmlsPassword.Text.Trim())

            Dim privy = GetOrCreate(loginElement, "PRIVY")
            SetNodeValue(privy, "UserName", txtPrivyUser.Text.Trim())
            SetNodeValue(privy, "Password", txtPrivyPassword.Text.Trim())

            ' ==============================
            ' MLS STATUS
            ' ==============================
            Dim mlsStatus = GetOrCreate(loginElement, "MLSStatus")
            SetNodeValue(mlsStatus, "Active", txtMlsActive.Text.Trim())
            SetNodeValue(mlsStatus, "Sold", txtMlsSold.Text.Trim())
            SetNodeValue(mlsStatus, "Hold", txtMlsHold.Text.Trim())
            SetNodeValue(mlsStatus, "Pending", txtMlsPending.Text.Trim())
            SetNodeValue(mlsStatus, "Other", txtMlsOther.Text.Trim())

            ' ==============================
            ' MLS SETTINGS
            ' ==============================
            Dim mlsSettings = GetOrCreate(loginElement, "MLSSETTINGS")
            SetNodeValue(mlsSettings, "Miles", txtMlsMiles.Text.Trim())
            SetNodeValue(mlsSettings, "Subcat", txtMlsSubcat.Text.Trim())
            SetNodeValue(mlsSettings, "Sqft", txtMlsSqft.Text.Trim())

            ' ==============================
            ' DEFAULT SETTINGS
            ' ==============================
            Dim defaults = GetOrCreate(loginElement, "DefaultsSettings")
            Dim autoMail As String =
            If(chkAutoStatusChangeEmails.IsChecked = True, "True", "False")

            SetNodeValue(defaults, "EmailStatusChangeUpdates", autoMail)

            ' ==============================
            ' DEFAULT EMAIL
            ' ==============================
            SetNodeValue(loginElement, "DefaultEmailAccount", txtDefaultEmail.Text.Trim())

            ' ==============================
            ' SAVE FILE
            ' ==============================
            xdoc.Save(_settingsFilePath)

            'MessageBox.Show("Settings saved successfully!",
            '            "Success", MessageBoxButton.OK, MessageBoxImage.Information)

        Catch ex As system.Exception
            MessageBox.Show("Failed to save settings: " & ex.Message,
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub btnpipelogin_Click(sender As Object, e As RoutedEventArgs)
        'WebViewPipe.Source = New Uri("https://app.pipedrive.com/auth/login")
    End Sub

    Private Sub btnToggleSidebar_Click(sender As Object, e As RoutedEventArgs)
        ToggleSidebarVisibility()
    End Sub
    Private Sub ToggleSidebarVisibility()
        'LogLayoutInfo("ToggleSidebarB4")
        If isSidebarHidden Then
            SidebarColumn.Width = New GridLength(origLeftSidebarWidth)
            isSidebarHidden = False
        Else
            SidebarColumn.Width = New GridLength(0)
            isSidebarHidden = True
        End If
        adjustGridWidth(SidebarColumn.Width.Value)
        'LogLayoutInfo("ToggleSidebarAF")
    End Sub

    Private Sub adjustGridWidth(iadjust As Double)
        GridMain.ColumnDefinitions(1).Width = New GridLength(newWindowWidth - iadjust, GridUnitType.Pixel)
        'GridTab.ColumnDefinitions(0).Width = New GridLength(iadjust, GridUnitType.Pixel)
        TabControlMain.Width = newWindowWidth - iadjust
        WebViewMain.Width = newWindowWidth - iadjust - 10
        WebViewComp.Width = newWindowWidth - iadjust - 10
        'dgDetails.Width = newWindowWidth - iadjust - 50
    End Sub
    Private Sub LogLayoutInfo(context As String)
        Try
            Dim logPath As String = Path.Combine(AppContext.BaseDirectory, "layout_debug.log")
            Dim windowWidth As Double = Me.ActualWidth
            Dim windowHeight As Double = Me.ActualHeight
            Dim windowMaxWidth As Double = Me.MaxWidth
            Dim gridMainWidth As Double = If(GridMain IsNot Nothing, GridMain.ActualWidth, -1)
            Dim sidebarWidth As Double = If(SidebarColumn IsNot Nothing, SidebarColumn.ActualWidth, -1)
            Dim sidebarColumnWidthValue As Double = If(SidebarColumn IsNot Nothing, SidebarColumn.Width.Value, -1)
            Dim rightColumnWidth As Double = If(GridMain IsNot Nothing AndAlso GridMain.ColumnDefinitions.Count > 1, GridMain.ColumnDefinitions(1).ActualWidth, -1)
            Dim tabControlWidth As Double = If(TabControlMain IsNot Nothing, TabControlMain.ActualWidth, -1)
            'Dim detailsGridWidth As Double = If(dgDetails IsNot Nothing, dgDetails.ActualWidth, -1)

            'Dim logLine As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{context}] " &
            '    $"WindowWidth={windowWidth}, WindowHeight={windowHeight}, WindowMaxWidth={windowMaxWidth}, " &
            '    $"GridMainWidth={gridMainWidth}, SidebarColumn.ActualWidth={sidebarWidth}, SidebarColumn.Width.Value={sidebarColumnWidthValue}, " &
            '    $"RightColumn.ActualWidth={rightColumnWidth}, TabControlMain.ActualWidth={tabControlWidth}, dgDetails.ActualWidth={detailsGridWidth}"

            'File.AppendAllText(logPath, logLine & Environment.NewLine)
        Catch ex As System.Exception
            ' Ignore logging errors
        End Try
    End Sub

    Private Async Sub btnSyncPipeDrive_Click(sender As Object, e As RoutedEventArgs)
        Await PipedriveGetStageIDsandNames()
        Await PipedriveSyncDealsDataToPropertyTable()
        LoadDataGrid()
        'chkSelectAll.IsChecked = False
    End Sub

    Public Async Function PipedriveGetStageIDsandNames() As Task
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text

        ' Construct the API endpoint for stages
        'Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v1/stages?api_token={apiToken}"
        'Try
        '    Dim jsonResponse As String = Await PipedriveFetchApiJsonAsync(url)
        '    Dim stagesResponse As PipedriveStagesResponse = JsonConvert.DeserializeObject(Of PipedriveStagesResponse)(jsonResponse)

        '    If stagesResponse IsNot Nothing AndAlso stagesResponse.success AndAlso stagesResponse.data IsNot Nothing Then
        '        For Each stage As StageData In stagesResponse.data
        '            ' Prepare the SQLite UPSERT statement
        '            ' Note: The table 'PipeDriveStages' should have a PRIMARY KEY constraint on the 'id' column for INSERT OR REPLACE to work as an upsert.
        '            Dim sql As String = "INSERT OR REPLACE INTO PipeDriveStages (id, name, order_nr) VALUES (@id, @name, @order_nr)"
        '            Dim SQLitecmd As New SQLiteCommand(sql)
        '            SQLitecmd.Parameters.AddWithValue("@id", stage.id)
        '            SQLitecmd.Parameters.AddWithValue("@name", stage.name)
        '            SQLitecmd.Parameters.AddWithValue("@order_nr", stage.order_nr)
        '            fxCommon.SQLExecuteCommand(SQLitecmd)
        '        Next
        '        'MessageBox.Show("Pipedrive stages updated successfully!")
        '    Else
        '        MessageBox.Show("Failed to retrieve stages from Pipedrive API.")
        '    End If
        'Catch ex As Net.Http.HttpRequestException
        '    MessageBox.Show($"HTTP Request Error: {ex.Message}")
        'Catch ex As System.Exception
        '    MessageBox.Show($"An error occurred: {ex.Message}")
        'End Try
    End Function

    ' Fetches deals from Pipedrive and updates Property table with pdDealID, pdStageID, pdStageName using fxCommon.SQLExecuteReader/SQLExecuteCommand
    Public Async Function PipedriveSyncDealsDataToPropertyTable() As Task
        ' Load config.json for API key and domain
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text

        ' Fetch all stages from SQLite using fxCommon.SQLExecuteReader
        Dim stageDict As New Dictionary(Of Integer, String)
        Dim stageQuery As String = "SELECT id, name FROM PipeDriveStages"
        Dim dtStages As System.Data.DataTable = fxCommon.SQLExecuteReader(stageQuery)
        For Each row As DataRow In dtStages.Rows
            stageDict(Convert.ToInt32(row("id"))) = row("name").ToString()
        Next

        ' Fetch deals from Pipedrive
        'Dim dealsUrl = $"https://{companyDomain}.pipedrive.com/api/v2/deals?api_token={apiToken}&limit=500" '&status=open"
        'Dim dealsJson As String = Await PipedriveFetchApiJsonAsync(dealsUrl)
        'Dim dealsObj = Newtonsoft.Json.Linq.JObject.Parse(dealsJson)
        'If Not (dealsObj("success") IsNot Nothing AndAlso dealsObj("success").ToObject(Of Boolean)) Then
        '    MessageBox.Show("Failed to fetch deals from Pipedrive.")
        '    Return
        'End If

        'Dim deals = dealsObj("data")
        'If deals Is Nothing Then
        '    MessageBox.Show("No deals found.")
        '    Return
        'End If

        ' Update Property table using fxCommon.SQLExecuteCommand
        Dim iUpdateCount As Integer = 0
        'For Each deal In deals
        '    Dim mlsId = deal("custom_fields")?("7924752563aef8a024ee2aa6d623ba6ed145c142")?.ToString()
        '    Dim stageId = deal("stage_id")?.ToObject(Of Integer)()
        '    Dim pdId = deal("id")?.ToString()
        '    Dim pdStatus = deal("status")?.ToString()
        '    If String.IsNullOrEmpty(mlsId) OrElse stageId Is Nothing OrElse String.IsNullOrEmpty(pdId) Then
        '        Continue For
        '    End If
        '    Dim stageName As String = ""
        '    If stageDict.ContainsKey(stageId) Then
        '        stageName = stageDict(stageId)
        '    End If

        '    ' Update Property table (fxCommon.SQLExecuteCommand)
        '    Dim updateSql As String = "UPDATE Property SET pdDealID=@pdDealID, pdStageID=@pdStageID, pdStageName=@pdStageName WHERE MLSListingID=@MLSListingID"
        '    Dim SQLitecmd As New SQLite.SQLiteCommand(updateSql)
        '    SQLitecmd.Parameters.AddWithValue("@pdDealID", pdId)
        '    SQLitecmd.Parameters.AddWithValue("@pdStageID", stageId.ToString())
        '    SQLitecmd.Parameters.AddWithValue("@pdStageName", stageName)
        '    SQLitecmd.Parameters.AddWithValue("@MLSListingID", mlsId)
        '    Dim rowsAffected As Integer = fxCommon.SQLExecuteCommand(SQLitecmd)

        '    If rowsAffected = 0 Then
        '        ' Insert new record with as many details as possible from the deal JSON
        '        Dim insertSql As String = "INSERT INTO Property (MLSListingID, pdDealID, pdStageID, pdStageName, Seller, SellerEmail, SellerPhone, LACell, LADirect, SavedAddress, OfferPrice, ARV, ListPrice, ClosePrice, SqFt, Status) " &
        '                                     "VALUES (@MLSListingID, @pdDealID, @pdStageID, @pdStageName, @Seller, @SellerEmail, @SellerPhone, @LACell, @LADirect, @SavedAddress, @OfferPrice, @ARV, @ListPrice, @ClosePrice, @SqFt, @Status)"
        '        Dim insertCmd As New SQLite.SQLiteCommand(insertSql)
        '        insertCmd.Parameters.AddWithValue("@MLSListingID", mlsId)
        '        insertCmd.Parameters.AddWithValue("@pdDealID", pdId)
        '        insertCmd.Parameters.AddWithValue("@pdStageID", stageId.ToString())
        '        insertCmd.Parameters.AddWithValue("@pdStageName", stageName)

        '        insertCmd.Parameters.AddWithValue("@Seller", pdGetSafeStringValue(deal, "455b5e163c481fde5281cde7f39ebfd19c15abc0"))
        '        insertCmd.Parameters.AddWithValue("@SellerEmail", pdGetSafeStringValue(deal, "0eddebbc40f548e8d347f4c51a1d6e1e0147a2c3"))
        '        insertCmd.Parameters.AddWithValue("@SellerPhone", pdGetSafeStringValue(deal, "d88f20459b047449f41455f5a2306138f729efce"))
        '        insertCmd.Parameters.AddWithValue("@SavedAddress", pdGetSafeStringValue(deal, "8033e63c4bfc1b97dd915f424da9c4401d65d34b"))
        '        insertCmd.Parameters.AddWithValue("@OfferPrice", pdGetMoneyValue(deal, "782c0afe87597d97347ead892c1d8d71371d2515"))
        '        insertCmd.Parameters.AddWithValue("@ARV", pdGetMoneyValue(deal, "5ea31681a33100a1e935015e9482c65a168e5e5b"))
        '        insertCmd.Parameters.AddWithValue("@ListPrice", pdGetMoneyValue(deal, "0654fc92dfa67f1c8801866d94c10c49c1835c3d"))
        '        insertCmd.Parameters.AddWithValue("@ClosePrice", pdGetMoneyValue(deal, "60489922d1c93011cf367b3ff2e24a0d9834122e"))
        '        insertCmd.Parameters.AddWithValue("@SqFt", pdGetSafeStringValue(deal, "12466f54cf2399bab32a4a7d3ccf60939d27ee2d"))
        '        insertCmd.Parameters.AddWithValue("@Status", pdGetSafeStringValue(deal, "1452ff29b9a430f6ed4fa34d03bc3976e8c8310e"))

        '        insertCmd.Parameters.AddWithValue("@LACell", pdGetSafeStringValue(deal, "4dab1f2310bc4c53db7a82410560807b76b7bc44"))
        '        insertCmd.Parameters.AddWithValue("@LADirect", pdGetSafeStringValue(deal, "941df5800d43c6bb7b598d5782784912fabc960c"))
        '        '----
        '        fxCommon.SQLExecuteCommand(insertCmd)
        '    End If
        'Next

        MessageBox.Show("Pipedrive deals synced to Property table.")
    End Function

    ' Add a note to a Pipedrive deal
    Public Async Function PipedriveAddNotesToDeal(dealId As Integer, noteContent As String) As Task
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text
        'Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v1/notes?api_token={apiToken}"

        Dim data As New Dictionary(Of String, Object) From {
            {"deal_id", dealId},
            {"content", noteContent}
        }
        Using client As New Net.Http.HttpClient()
            Dim content As New Net.Http.StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(data), System.Text.Encoding.UTF8, "application/json")
            'Dim response = Await client.PostAsync(url, content)
            'If response.IsSuccessStatusCode Then
            '    'MessageBox.Show("Note added successfully.")
            'Else
            '    MessageBox.Show($"Failed to add note: {response.StatusCode}" & vbCrLf & Await response.Content.ReadAsStringAsync())
            'End If
        End Using
    End Function

    ' Update a Pipedrive deal's status and/or stage_id
    Public Async Function PipedriveUpdateDealStatusAndStage(dealId As Integer, Optional status As String = Nothing, Optional stageId As Integer = Nothing) As Task
        ' Load API key and domain from config.json
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text
        'Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v2/deals/{dealId}?api_token={apiToken}"
        Dim data As New Dictionary(Of String, Object)()

        If Not String.IsNullOrEmpty(status) Then data("status") = status
        If stageId <> 0 Then data("stage_id") = stageId
        If data.Count = 0 Then
            MessageBox.Show("Nothing to update.")
            Return
        End If
        Using client As New Net.Http.HttpClient()
            Dim content As New Net.Http.StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(data), System.Text.Encoding.UTF8, "application/json")
            'Dim response = Await client.PatchAsync(url, content)
            'If response.IsSuccessStatusCode Then
            '    'MessageBox.Show("Deal updated successfully.")
            'Else
            '    MessageBox.Show($"Failed to update deal: {response.StatusCode}" & vbCrLf & Await response.Content.ReadAsStringAsync())
            'End If
        End Using
    End Function

    ' Add or update a contact (person) in Pipedrive and return the person id. Optionally link to organization.
    Public Async Function PipeDriveAddPersonContact(name As String, email As String, Optional phone1 As String = Nothing, Optional phone2 As String = Nothing, Optional phone3 As String = Nothing, Optional organization As Integer? = Nothing) As Task(Of Integer?)
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text
        Dim personId As Integer? = Await PipeDriveSearchPersonByEmail(email)
        Dim phones As New List(Of Object)()
        If Not String.IsNullOrEmpty(phone1) Then phones.Add(New With {.value = phone1, .label = "work"})
        If Not String.IsNullOrEmpty(phone2) Then phones.Add(New With {.value = phone2, .label = "work"})
        If Not String.IsNullOrEmpty(phone3) Then phones.Add(New With {.value = phone3, .label = "work"})
        Dim data As New Dictionary(Of String, Object) From {
            {"name", name},
            {"email", email},
            {"phone", phones}
        }
        If organization.HasValue Then data("org_id") = organization.Value
        Using client As New Net.Http.HttpClient()
            If personId.HasValue Then
                ' Update existing
                'Dim updateUrl = $"https://{companyDomain}.pipedrive.com/api/v1/persons/{personId}?api_token={apiToken}"
                Dim content = New Net.Http.StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(data), System.Text.Encoding.UTF8, "application/json")
                'Dim response = Await client.PutAsync(updateUrl, content)
                'If response.IsSuccessStatusCode Then
                '    MessageBox.Show($"Person already exists. Updated. ID: {personId}")
                '    Return personId
                'Else
                '    MessageBox.Show($"Failed to update person: {response.StatusCode}" & vbCrLf & Await response.Content.ReadAsStringAsync())
                '    Return Nothing
                'End If
            Else
                ' Add new
                'Dim addUrl = $"https://{companyDomain}.pipedrive.com/api/v1/persons?api_token={apiToken}"
                Dim content = New Net.Http.StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(data), System.Text.Encoding.UTF8, "application/json")
                'Dim response = Await client.PostAsync(addUrl, content)
                'If response.IsSuccessStatusCode Then
                '    Dim respStr = Await response.Content.ReadAsStringAsync()
                '    Dim respObj = Newtonsoft.Json.Linq.JObject.Parse(respStr)
                '    Dim newId = respObj("data")?("id")
                '    MessageBox.Show($"Person added successfully. ID: {newId}")
                '    Return If(newId IsNot Nothing, newId.ToObject(Of Integer)(), Nothing)
                'Else
                '    MessageBox.Show($"Failed to add person: {response.StatusCode}" & vbCrLf & Await response.Content.ReadAsStringAsync())
                '    Return Nothing
                'End If
            End If
        End Using
    End Function

    ' Search for a person by email. Returns person id if found, else Nothing.
    Public Async Function PipeDriveSearchPersonByEmail(email As String) As Task(Of Integer?)
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text
        'Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v1/persons/search?api_token={apiToken}&term={email}&fields=email&exact_match=true"

        Using client As New Net.Http.HttpClient()
            'Dim response = Await client.GetAsync(url)
            'If response.IsSuccessStatusCode Then
            '    Dim respStr = Await response.Content.ReadAsStringAsync()
            '    Dim respObj = Newtonsoft.Json.Linq.JObject.Parse(respStr)
            '    Dim items = respObj("data")?("items")
            '    If items IsNot Nothing AndAlso items.HasValues AndAlso items.Count > 0 Then
            '        Return items(0)("item")("id").ToObject(Of Integer)()
            '    Else
            '        Return Nothing
            '    End If
            'Else
            '    MessageBox.Show($"Failed to search for person: {response.StatusCode}")
            '    Return Nothing
            'End If
        End Using
    End Function

    ' Get list of emails related to a person id from Pipedrive.
    Public Async Function PipeDriveGetPersonEmails(personId As Integer) As Task(Of List(Of JObject))
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text
        'Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v1/persons/{personId}/mailMessages?api_token={apiToken}"

        Using client As New Net.Http.HttpClient()
            'Dim response = Await client.GetAsync(url)
            'If response.IsSuccessStatusCode Then
            '    Dim respStr = Await response.Content.ReadAsStringAsync()
            '    Dim respObj = Newtonsoft.Json.Linq.JObject.Parse(respStr)
            '    Dim emails = respObj("data")
            '    If emails Is Nothing OrElse Not emails.HasValues Then
            '        MessageBox.Show($"No emails found for person {personId}.")
            '        Return New List(Of JObject)()
            '    End If
            '    Dim result As New List(Of JObject)()
            '    For Each email In emails
            '        result.Add(email)
            '    Next
            '    MessageBox.Show($"Found {result.Count} emails for person {personId}.")
            '    Return result
            'Else
            '    MessageBox.Show($"Failed to get emails: {response.StatusCode}" & vbCrLf & Await response.Content.ReadAsStringAsync())
            '    Return New List(Of JObject)()
            'End If
        End Using
    End Function

    ' Links email threads to deals if the email subject contains the property address
    Public Async Function PipeDriveMapEmailThreadsUsingPersonIds() As Task
        ' 1. Get all properties with pdPersonID and SavedAddress and pdDealID
        'Dim apiToken As String = txtPipeApiKey.Text
        'Dim companyDomain As String = txtPipeCompanyName.Text
        Dim dtProperties As System.Data.DataTable = fxCommon.SQLExecuteReader("SELECT MLSListingID, SavedAddress, pdPersonID, pdDealID FROM Property WHERE pdPersonID IS NOT NULL AND pdPersonID <> '' AND pdDealID IS NOT NULL AND pdDealID <> '' AND SavedAddress IS NOT NULL AND SavedAddress <> ''")
        If dtProperties.Rows.Count = 0 Then
            MessageBox.Show("No properties with pdPersonID and pdDealID found.")
            Return
        End If
        Using client As New Net.Http.HttpClient()
            For Each row As DataRow In dtProperties.Rows
                Dim personId As String = row("pdPersonID").ToString()
                Dim dealId As String = row("pdDealID").ToString()
                Dim address As String = row("SavedAddress").ToString()
                If String.IsNullOrEmpty(personId) OrElse String.IsNullOrEmpty(dealId) OrElse String.IsNullOrEmpty(address) Then Continue For

                ' 2. Get emails for this person
                'Dim url As String = $"https://{companyDomain}.pipedrive.com/api/v1/persons/{personId}/mailMessages?api_token={apiToken}"
                'Dim response = Await client.GetAsync(url)
                'If Not response.IsSuccessStatusCode Then Continue For
                'Dim respStr = Await response.Content.ReadAsStringAsync()
                'Dim respObj = Newtonsoft.Json.Linq.JObject.Parse(respStr)
                'Dim emails = respObj("data")
                '                If emails Is Nothing OrElse Not emails.HasValues Then Continue For
                '                For Each email In emails
                '                    Dim data = email("data")
                '                    If data Is Nothing Then Continue For
                '                    Dim subject As String = If(data("subject") IsNot Nothing, data("subject").ToString(), "")
                '                    Dim threadId As String = If(data("mail_thread_id") IsNot Nothing, data("mail_thread_id").ToString(), "")
                '                    If String.IsNullOrEmpty(subject) OrElse String.IsNullOrEmpty(threadId) Then Continue For

                '                    ' 3. If subject contains address, link thread to deal
                '                    If subject.ToLower().Contains(address.ToLower()) Then
                '                        ' Check if already linked in Property table
                '                        Dim checkSql = "SELECT pdEmailThreadID FROM Property WHERE pdDealID=@dealId"
                '                        Dim checkCmd As New SQLite.SQLiteCommand(checkSql)
                '                        checkCmd.Parameters.AddWithValue("@dealId", dealId)
                '                        Dim existingThreadId As Object = fxCommon.SQLExecuteCommand(checkCmd)
                '                        If existingThreadId IsNot Nothing AndAlso Not IsDBNull(existingThreadId) AndAlso existingThreadId.ToString() = threadId Then
                '                            ' Already linked, skip
                '                            Continue For
                '                        End If

                '                        ' Link thread to deal using endpoint: PUT /api/v1/mailbox/mailThreads/{threadId}/deal/{dealId}?api_token=xxx
                '                        ' 1. Form the correct URL, which only contains the threadId in the path.
                '                        '    The api_token is a query parameter, which is correct.
                '                        'Dim linkUrl = $"https://{companyDomain}.pipedrive.com/api/v1/mailbox/mailThreads/{threadId}?api_token={apiToken}"

                '                        ' 2. Create a JSON payload for the request body.
                '                        '    The API expects the deal_id to be in the request body, not the URL.
                '                        Dim payload As Object = New With {
                '    .deal_id = dealId
                '}

                '                        ' 3. Serialize the payload to a StringContent object.
                '                        Dim jsonPayload As String = JsonConvert.SerializeObject(payload)
                '                        Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

                '                        ' 4. Make the PUT request with the correct URL and content.
                '                        Dim linkResp = Await client.PutAsync(linkUrl, content)
                '                        If linkResp.IsSuccessStatusCode Then
                '                            MessageBox.Show($"Linked thread {threadId} to deal {dealId} for address '{address}'")
                '                            ' Optionally update Property table with thread id
                '                            Dim updateSql = "UPDATE Property SET pdEmailThreadID=@threadId WHERE pdDealID=@dealId"
                '                            Dim SQLitecmd As New SQLite.SQLiteCommand(updateSql)
                '                            SQLitecmd.Parameters.AddWithValue("@threadId", threadId)
                '                            SQLitecmd.Parameters.AddWithValue("@dealId", dealId)
                '                            fxCommon.SQLExecuteCommand(SQLitecmd)
                '                        Else
                '                            MessageBox.Show($"Failed to link thread {threadId} to deal {dealId}: {linkResp.StatusCode}")
                '                        End If
                '                    End If
                '                Next
            Next
        End Using
        MessageBox.Show("Email thread mapping completed.")
    End Function
    Private Function pdGetSafeStringValue(ByVal deal As JObject, ByVal fieldKey As String) As String
        ' Safely get the custom_fields JObject
        Dim customFields As JObject
        If Not deal.TryGetValue("custom_fields", customFields) OrElse customFields Is Nothing Then
            Return ""
        End If

        ' Safely get the value for the specific field key
        Dim fieldValue As JToken
        If customFields.TryGetValue(fieldKey, fieldValue) AndAlso fieldValue IsNot Nothing Then
            Return fieldValue.ToString()
        End If

        Return ""
    End Function
    Private Function pdGetMoneyValue(ByVal deal As JObject, ByVal fieldKey As String) As String
        ' Safely get the custom_fields JObject
        Dim customFields As JObject
        If Not deal.TryGetValue("custom_fields", customFields) OrElse customFields Is Nothing Then
            ' If the field is not found, return a formatted string for zero.
            Return 0.ToString("C0")
        End If

        ' Safely get the token for the specified field key
        Dim moneyFieldToken As JToken
        If customFields.TryGetValue(fieldKey, moneyFieldToken) AndAlso moneyFieldToken IsNot Nothing Then
            ' Check if the token is an object (for money fields)
            If moneyFieldToken.Type = JTokenType.Object Then
                Dim moneyField As JObject = moneyFieldToken.ToObject(Of JObject)()

                ' Safely get the 'value' token from the money field JObject
                Dim valueToken As JToken
                If moneyField.TryGetValue("value", valueToken) AndAlso valueToken IsNot Nothing Then
                    ' Convert the token to a Decimal
                    Dim value As Decimal = valueToken.ToObject(Of Decimal)()

                    ' Format the decimal with currency symbol and thousands separator
                    ' "C0" format specifier means Currency with 0 decimal places
                    Return value.ToString("C0")
                End If
            End If
        End If

        ' If any part of the value extraction fails, return a formatted string for zero
        Return 0.ToString("C0")
    End Function
    Private Function pdGetMoneyValueOld(ByVal deal As JObject, ByVal fieldKey As String) As Decimal
        ' Safely get the custom_fields JObject
        Dim customFields As JObject
        If Not deal.TryGetValue("custom_fields", customFields) OrElse customFields Is Nothing Then
            Return 0
        End If

        ' Safely get the token for the specified field key
        Dim moneyFieldToken As JToken
        If customFields.TryGetValue(fieldKey, moneyFieldToken) AndAlso moneyFieldToken IsNot Nothing Then
            ' Check if the token is an object (for money fields)
            If moneyFieldToken.Type = JTokenType.Object Then
                Dim moneyField As JObject = moneyFieldToken.ToObject(Of JObject)()

                ' Safely get the 'value' token from the money field JObject
                Dim valueToken As JToken
                If moneyField.TryGetValue("value", valueToken) AndAlso valueToken IsNot Nothing Then
                    ' Convert the token to a Decimal
                    Return valueToken.ToObject(Of Decimal)()
                End If
            End If
        End If

        Return 0
    End Function

    ' Helper function to fetch JSON from a Pipedrive API endpoint
    Private Async Function PipedriveFetchApiJsonAsync(url As String) As Task(Of String)
        Using client As New Net.Http.HttpClient()
            Return Await client.GetStringAsync(url)
        End Using
    End Function

    Private Sub chkAutoStatusChangeEmails_Checked(sender As Object, e As RoutedEventArgs)
        'Save this status change to the login.xml settings
        SaveSettings()
    End Sub

    Private Async Sub btnLinkEmails_Click(sender As Object, e As RoutedEventArgs)
        Await PipeDriveMapEmailThreadsUsingPersonIds()
    End Sub

    Private Sub chkSelectAll_Checked(sender As Object, e As RoutedEventArgs)
        CheckAllCheckboxes(True)
    End Sub

    Private Sub chkSelectAll_UnChecked(sender As Object, e As RoutedEventArgs)
        CheckAllCheckboxes(False)
    End Sub
    Public Sub CheckAllCheckboxes(bCheck As Boolean)
        ' Ensure the DataGrid has an ItemsSource to work with.
        'If Me.dgDetails.ItemsSource IsNot Nothing Then
        '    ' Loop through each item in the DataGrid's collection.
        '    For Each item As Object In Me.dgDetails.ItemsSource
        '        Try
        '            ' This line attempts to set the value using the column indexer
        '            ' of a DataRowView object. This is the correct way for DataTables.
        '            item.Item("IsSelected") = bCheck
        '        Catch ex As System.Exception
        '            ' If an exception occurs, it means the item is not a DataRowView or
        '            ' the 'IsSelected' column does not exist. You can add more specific
        '            ' handling here if needed.
        '            System.Diagnostics.Debug.WriteLine($"Failed to set 'IsSelected' for an item: {ex.Message}")
        '        End Try
        '    Next
        'End If
    End Sub


    ''' <summary>
    ''' Appends a new message to the top of the log TextBox with a date and time stamp.
    ''' This function is thread-safe and can be called from any background thread.
    ''' </summary>
    ''' <param name="message">The text message to be logged.</param>
    Public Sub LogMessage(ByVal message As String)
        Dim timestamp As String = DateTime.Now.ToString("[dd-MMM-yy HH:mm:ss] ")
        Dim newLogEntry As String = timestamp & message
        ' Use the Dispatcher to ensure the UI update is done on the main thread.
        ' This prevents "The calling thread cannot access this object" errors.
        'If Me.txtLog.Dispatcher.CheckAccess() Then
        '    ' If we are already on the UI thread, update the TextBox directly.
        '    ' vbCrLf ensures a new line is added between log entries.
        '    Me.txtLog.Text = newLogEntry & vbCrLf & Me.txtLog.Text
        'Else
        '    ' If we are on a different thread, use Invoke to update the UI safely.
        '    Me.txtLog.Dispatcher.Invoke(Sub()
        '                                    Me.txtLog.Text = newLogEntry & vbCrLf & Me.txtLog.Text
        '                                End Sub)
        'End If
    End Sub
End Class

Public Class Item
    Public Property Id As Integer
    Public Property Name As String
End Class

Public Class PipedriveDeal
    Public Property id As Integer
    Public Property stage_id As Integer
    ' You can add other properties from the JSON if needed
End Class

' Class to represent a single stage object
Public Class StageData
    Public Property id As Integer
    Public Property name As String
    Public Property order_nr As Integer
End Class

' Class to represent the overall JSON response structure
Public Class PipedriveStagesResponse
    Public Property success As Boolean
    Public Property data As List(Of StageData)
End Class
Public Class PipedriveResponse
    Public Property success As Boolean
    Public Property data As PipedriveDeal
End Class