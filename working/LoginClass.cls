VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private Main As Object
Private Csv As Object
'
Private Const FormTypeLogin = "aoLoginForm"
Private Const cr = vbCrLf & vbTab
Private Const cr2 = vbCrLf & vbTab & vbTab
Private Const cr3 = vbCrLf & vbTab & vbTab & vbTab
Private Const cr4 = vbCrLf & vbTab & vbTab & vbTab & vbTab
Private Const cr5 = vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
Private Const ButtonLogin = " Login "
Private Const maxVisitLoginAttempts = 10
'
'========================================================================
'   v3.3 Add-on Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
End Function
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(MainObject As Object)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("LoginClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Link As String
    Dim VisibleName As String
    Dim CS As Long
    '
    If Main.IsAuthenticated() Then
        VisibleName = Trim(Main.MemberName)
        If VisibleName = "" Then
            CS = Main.OpenCSContentRecord("people", Main.MemberID)
            If Main.IsCSOK(CS) Then
                VisibleName = Trim(Main.GetCSText(CS, "FirstName") & " " & Main.GetCSText(CS, "LastName"))
                If VisibleName = "" Then
                    VisibleName = "Unnamed user"
                End If
            End If
            Call Main.CloseCS(CS)
        End If
        Link = ModifyQueryString(Main.refreshQueryString, "method", "logout")
        Link = Main.ServerPage & "?" & Link
        GetContent = GetContent & "<P>You are currently logged in as " & VisibleName & ". If this is not you, please <a href=""" & Link & """>click here</a>.</P>"
    ElseIf (Main.GetSiteProperty("BUILDVERSION", "") <= "4.1.562") Then
        '
        ' old main.getLoginForm just includes the form
        '
        GetContent = GetContent & vbCrLf & vbTab & "<!-- legacy loginform mode - main returns just the un/pw forms -->"
        GetContent = GetContent & "<P>To login, enter your username and password.</P>"
        GetContent = GetContent & Main.GetLoginForm()
        '
        If kmaEncodeBoolean(Main.GetSiteProperty("AllowPasswordEmail", False)) Then
            GetContent = GetContent & "<P>Forget your password? If you are a member of this site and can not remember your password, enter your email address below and we will email your matching username and password.</P>"
            GetContent = GetContent & Main.GetSendPasswordForm()
        End If
    Else
        '
        ' new login main.getLoginForm returns the systems default form
        '
        GetContent = GetContent & GetLoginForm_Default()
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("LoginClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

'
'========================================================================
'   default login form
'========================================================================
'
Friend Function GetLoginForm_Default() As String
    On Error GoTo ErrorTrap
    '
    Dim returnHtml As String
    Dim Panel As String
    Dim loginAddonID As Long
    Dim usernameMsg As String
    Dim QueryString As String
    Dim loginForm As String
    Dim Caption As String
    Dim formSn As String
    Dim formType As String
    Dim needLoginForm As Boolean
    
    '
    ' ----- process the previous form, if login OK, return blank (signal for page refresh)
    '
    needLoginForm = True
    formSn = Main.GetStreamText("ccFormSN")
    If formSn <> "" Then
        formType = Main.GetStreamText(kmaEncodeText(formSn) & "type")
        If formType = FormTypeLogin Then
            If ProcessLoginFormDefault(kmaEncodeText(formSn)) Then
                returnHtml = ""
                needLoginForm = False
                '
                ' temp workaround - eventually contensive needs to recognize the blank return
                '
                If Main.renderedpageId <> 0 Then
                    Call Main.redirect("?" & Main.refreshQueryString)
                End If
            End If
        End If
    End If
    If needLoginForm Then
        '
        ' ----- When page loads, set focus on login username
        '
        Call Main.AddRefreshQueryString("method", "")
        loginForm = ""
        formSn = GetFormSN
        'Call AddPageTitle2("Login", "login")
        Call Main.AddOnLoadJavascript2("document.getElementById('aoLoginFormUsernameInput').focus()", "login")
        '
        ' ----- Error Messages
        '
        If kmaEncodeBoolean(Main.GetSiteProperty("allowEmailLogin", False)) Then
            usernameMsg = "<b>To login, enter your username or email address with your password.</b></p>"
        Else
            usernameMsg = "<b>To login, enter your username and password.</b></p>"
        End If
        '
        QueryString = Main.ServerQueryString
        'QueryString = ModifyQueryString(QueryString, "method", "", False)
        QueryString = ModifyQueryString(QueryString, "requestbinary", "", False)
        '
        ' ----- Username
        '
        If kmaEncodeBoolean(Main.GetSiteProperty("allowEmailLogin", False)) Then
            Caption = "Username&nbsp;or&nbsp;Email"
        Else
            Caption = "Username"
        End If
        '
        loginForm = loginForm _
            & cr & "<tr>" _
            & cr2 & "<td style=""text-align:right;vertical-align:middle;width:30%;padding:4px"" align=""right"" width=""30%"">" & Caption & "&nbsp;</td>" _
            & cr2 & "<td style=""text-align:left;vertical-align:middle;width:70%;padding:4px"" align=""left""  width=""70%""><input ID=""aoLoginFormUsernameInput"" NAME=""" & "aoLoginFormUsername"" VALUE=""" & kmaEncodeHTML(Main.MemberLoginUsername) & """ SIZE=""20"" MAXLENGTH=""50"" ></td>" _
            & cr & "</tr>"
        '
        ' ----- Password
        '
        If kmaEncodeBoolean(Main.GetSiteProperty("allowNoPasswordLogin", False)) Then
            Caption = "Password&nbsp;(optional)"
        Else
            Caption = "Password"
        End If
        loginForm = loginForm _
            & cr & "<tr>" _
            & cr2 & "<td style=""text-align:right;vertical-align:middle;width:30%;padding:4px"" align=""right"">" & Caption & "&nbsp;</td>" _
            & cr2 & "<td style=""text-align:left;vertical-align:middle;width:70%;padding:4px"" align=""left"" ><input NAME=""" & "aoLoginFormPassword"" VALUE="""" SIZE=""20"" MAXLENGTH=""50"" type=""password""></td>" _
            & cr & "</tr>" _
            & ""
        '
        ' ----- autologin support
        '
        If kmaEncodeBoolean(Main.GetSiteProperty2("AllowAutoLogin", False)) Then
            loginForm = loginForm _
                & cr & "<tr>" _
                & cr2 & "<td align=""right"">&nbsp;</td>" _
                & cr2 & "<td align=""left"" >" _
                & cr3 & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">" _
                & cr4 & "<tr>" _
                & cr5 & "<td valign=""top"" width=""20""><input type=""checkbox"" name=""" & formSn & "autologin"" value=""ON"" checked></td>" _
                & cr5 & "<td valign=""top"" width=""100%"">Login automatically from this computer</td>" _
                & cr4 & "</tr>" _
                & cr3 & "</table>" _
                & cr2 & "</td>" _
                & cr & "</tr>"
        End If
        loginForm = loginForm _
            & cr & "<tr>" _
            & cr2 & "<td colspan=""2"">&nbsp;</td>" _
            & cr & "</tr>" _
            & ""
        loginForm = "" _
            & cr & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">" _
            & kmaIndent(loginForm) _
            & cr & "</table>" _
            & ""
        loginForm = loginForm _
            & Main.GetFormInputHidden("ccFormSN", formSn) _
            & Main.GetFormInputHidden(formSn & "Type", FormTypeLogin) _
            & Main.GetFormInputHidden("email", Main.MemberLoginEmail) _
            & Main.GetPanelButtons(ButtonLogin, "Button") _
            & ""
        loginForm = "" _
            & Main.GetFormStart(QueryString) _
            & kmaIndent(loginForm) _
            & cr & "</form>" _
            & ""
    
    '-------
        
        Panel = "" _
            & Main.GetUserError _
            & cr & "<p class=""ccAdminNormal"">" & usernameMsg _
            & loginForm _
            & ""
'        Panel = "" _
'            & GetUserError _
'            & cr & "<p class=""ccAdminNormal"">You are attempting to enter an access controlled area. Continue only if you have authority to enter this area. Information about your visit will be recorded for security purposes.</p>" _
'            & cr & "<p class=""ccAdminNormal"">" & usernameMsg _
'            & loginForm _
'            & ""
        '
        ' ----- Password Form
        '
        If kmaEncodeBoolean(Main.GetSiteProperty2("AllowPasswordEmail", True)) Then
            Panel = "" _
                & Panel _
                & cr & "<p class=""ccAdminNormal""><b>Forget your password?</b></p>" _
                & cr & "<p class=""ccAdminNormal"">If you are a member of the system and can not remember your password, enter your email address below and we will email your matching username and password.</p>" _
                & Main.GetSendPasswordForm() _
                & ""
        End If
        '
        returnHtml = "" _
            & cr & "<div class=""ccLoginFormCon"">" _
            & kmaIndent(Panel) _
            & cr & "</div>" _
            & ""
'        '
'        ' --- Print credits
'        '
'        returnHtml = "" _
'            & GetPanel(returnHtml, "ccPanel", "ccPanelHilite", "ccPanelShadow", "400", 15) _
'            & cr & "<p>&nbsp;</p>" _
'            & cr & "<p>&nbsp;</p>" _
'            & cr & "<p style=""text-align:center""><a href=""http://www.Contensive.com"" target=""_blank""><img src=""/ccLib/images/ccLibLogin.GIF"" width=""80"" height=""33"" border=""0"" alt=""Contensive Content Control"" " & TagClose & "</A></p>" _
'            & cr & "<p style=""text-align:center"" class=""ccAdminSmall"">The content on this web site is managed and delivered by the Contensive Site Management Server. If you do not have member access, please use your back button to return to the public area.</p>" _
'            & ""
'        '
'        ' --- create an outer table to hold the form
'        '
'        returnHtml = "" _
'            & cr & "<div class=""ccCon"" style=""width:400px;margin:auto auto auto auto;"">" _
'            & kmaIndent(GetPanelHeader("Login")) _
'            & kmaIndent(returnHtml) _
'            & "</div>"
    End If
    '
    GetLoginForm_Default = returnHtml
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    'Call HandleClassErrorAndBubble_TrapPatch1("getLoginForm_default")
End Function

'
'========================================================================
' ----- Process the login form
'========================================================================
'
Private Function ProcessLoginFormDefault(formSn As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim MethodName As String
    Dim Button As String
    Dim CS As Long
    Dim LoginErrorMessage As String
    Dim LocalMemberID As Long
    Dim MemberLoginUsername  As String
    Dim MemberLoginPassword As String
    Dim MemberLoginAutoLogin As String
    '
    MethodName = "ProcessLoginFormDefault()"
    ProcessLoginFormDefault = False
    '
    If True Then
        '
        ' Processing can happen
        '   1) early in init() -- legacy
        '   2) as well as at the front of getLoginForm - to support addon Login forms
        ' This flag prevents the default form from processing twice
        '
        'loginFormDefaultProcessed = True
        MemberLoginUsername = Main.GetStreamText("aoLoginFormUsername")
        MemberLoginPassword = Main.GetStreamText("aoLoginFormPassword")
        MemberLoginAutoLogin = Main.GetStreamBoolean2(formSn & "autologin")
        '
        If (Main.VisitCookieSupport) Then
        'If (Main.VisitLoginAttempts < maxVisitLoginAttempts) And (Main.VisitCookieSupport) Then
            LocalMemberID = Main.GetLoginMemberID(MemberLoginUsername, MemberLoginPassword)
            If LocalMemberID = 0 Then
                Main.VisitLoginAttempts = Main.VisitLoginAttempts + 1
                Call Main.SaveVisit
            Else
                ProcessLoginFormDefault = Main.LoginMemberByID(LocalMemberID, MemberLoginAutoLogin)
                If ProcessLoginFormDefault Then
                    Call Main.LogActivity2("successful username/password login", Main.MemberID, Main.MemberOrganizationID)
                Else
                    Call Main.LogActivity2("bad username/password login", Main.MemberID, Main.MemberOrganizationID)
                End If
            End If
        End If
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    'Call HandleClassErrorAndResume_TrapPatch1(MethodName)
End Function
'
'
'
Private Function GetFormSN() As String
    GetFormSN = CStr(GetRandomInteger)
End Function
