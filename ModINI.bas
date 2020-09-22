Attribute VB_Name = "ModINI"
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    On Error Resume Next
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, getprivateprofilestring(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    On Error Resume Next
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function

Function GetSet(Key As String, Optional Default As String) As String
    On Error Resume Next
    Dim buffer As String
    buffer = ReadINI(UserName, Key, SettingsFile)
    If Len(buffer) = 0 Then
        If Len(Default) > 0 Then
            WriteINI UserName, Key, Default, SettingsFile
        End If
    End If
    GetSet = ReadINI(UserName, Key, SettingsFile)
End Function

Function SaveSet(Key As String, Value As String) As String
    On Error Resume Next
        WriteINI UserName, Key, Value, SettingsFile
    SaveSet = Key
End Function

Function GetItemData(ItemNumber As Integer, WhichItem As String, Optional Default As String) As String
    On Error Resume Next
    Dim buffer As String
    buffer = ReadINI(UserItem(ItemNumber), WhichItem, SettingsFile)
    'Debug.Print buffer
    If Len(buffer) = 0 Then
        If Len(Default) > 0 Then
            WriteINI UserItem(ItemNumber), WhichItem, Default, SettingsFile
        End If
    End If
    GetItemData = ReadINI(UserItem(ItemNumber), WhichItem, SettingsFile)
End Function

Function SetItemData(ItemNumber As Integer, WhichItem As String, Value As String) As String
    On Error Resume Next
    Dim buffer As String
    'Debug.Print "write of entry " & ItemNumber & "/" & WhichItem & " as " & Value
    buffer = WriteINI(UserItem(ItemNumber), WhichItem, Value, SettingsFile)
    SetItemData = buffer
End Function

Public Function UserItem(Index As Integer) As String
    On Error Resume Next
    UserItem = UserName & IIf(Index < 10, "0", "") & Index
End Function
