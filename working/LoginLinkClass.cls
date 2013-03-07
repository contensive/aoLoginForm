VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginLinkClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private Main As Object
Private Csv As Object
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
    Dim QS As String
    '
    QS = Main.RefreshQueryString
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
        QS = ModifyQueryString(QS, "method", "logout")
        Link = Main.ServerPage & "?" & QS
        GetContent = GetContent & "<div class=""LoginContainer"">You are currently logged in as " & VisibleName & ". If this is not you, please <a href=""" & Link & """>click here</a>.</div>"
    Else
        '
        ' ----- Login form
        '
        QS = ModifyQueryString(QS, "method", "login")
        Link = Main.ServerPage & "?" & QS
        GetContent = GetContent & "<div class=""LoginContainer""><a href=""" & Link & """>Login</a></div>"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("LoginClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function






