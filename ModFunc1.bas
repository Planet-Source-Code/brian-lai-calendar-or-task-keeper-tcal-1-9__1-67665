Attribute VB_Name = "ModFunc1"
Public IsReady As Boolean
Public NotifyTop As Long

Sub Main()
    On Error Resume Next
    Dim Tmp As Variant
    Dim Var As Variant
    XPVB
    If App.PrevInstance = True Then End
    If Trim(Command) = "" Then
        frmMain.Show
        Exit Sub
    End If
    Tmp = Split(Command)
    For Each Var In Tmp
        Select Case LCase(Var)
            Case "/c"
                LagMeCPU
            Case "/h"
                LagMeHDD
            Case Else
                frmMain.Show
        End Select
    Next
End Sub

Public Function MyVer()
    On Error Resume Next
    Dim Buffer2 As String
    Dim PreVer As Integer
    PreVer = App.Minor
    If App.Revision >= 1 Then
        PreVer = PreVer + 1
        Buffer2 = Trim$(Str$(PreVer) & " BETA")
    Else
        Buffer2 = Trim$(Str$(PreVer))
    End If
    MyVer = "V." & App.Major & "." & Buffer2
End Function

Sub DropShadow(hWnd As Long, Optional Silent As Boolean = True)
    On Error Resume Next
    If GetSet("Shadow", "1") = "1" Then
        SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW
    End If
End Sub

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Public Function UserName() As String
    On Error Resume Next
    Dim lpBuffer As String
    Dim j
    lpBuffer = Space$(255)
    GetUserName lpBuffer, Len(lpBuffer)
        j = InStr(lpBuffer, Chr$(0))
    If j > 0 Then UserName = Left$(lpBuffer, j - 1)
End Function


Public Function SettingsFile() As String
    On Error Resume Next
    SettingsFile = FindPath(App.Path, "TCal.ini")
End Function

Public Function LagMeCPU()
    On Error Resume Next
again:
    GoTo again
End Function

Public Function MyManifestFile() As String
    On Error Resume Next
    MyManifestFile = FindPath(App.Path, App.EXEName & ".exe.manifest")
End Function

Public Function LagMeHDD()
    On Error Resume Next
    Dim FF As Integer
    FF = FreeFile
    Open FindPath(App.Path, "TCal.sys") For Append As #FF
again:
        Print #FF, "                                                                                                    " 'space
    GoTo again
    Close #FF
End Function

Public Function XPVB(Optional ForceWriteManifest As Boolean = False) As Boolean
    On Error Resume Next
    If Dir(MyManifestFile) <> "" And ForceWriteManifest = False Then GoTo Written
    Dim XPStr As String
    Dim FF As Integer
    XPStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
            "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
            "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""Microsoft.VB6.VBnetStyles"" type=""win32""/>" & vbCrLf & _
            "<description>Windows XP manifest file</description>" & vbCrLf & "<dependency>" & vbCrLf & _
            "<dependentAssembly>" & vbCrLf & "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*""/>" & vbCrLf & _
            "</dependentAssembly>" & vbCrLf & "</dependency>" & vbCrLf & "</assembly>"
    FF = FreeFile
    Open MyManifestFile For Output As #FF
        Print #FF, XPStr
    Close #FF
Written:
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    XPVB = (ERR.Number = 0)
    On Error GoTo 0
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function
