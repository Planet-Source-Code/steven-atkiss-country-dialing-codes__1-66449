Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


'----[ Constants ]----'
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY



'----[ Enums ]----'
Public Enum rcMainKey       'root keys constants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Function StripNonChar(Text As String, Optional cReplace As String = "", Optional sTrim As Boolean = False, Optional RemoveCr As Boolean = False) As String

    Dim LP As Single
    Dim TempStr As String
    Dim sSplit() As String
    Dim sStart As Single, sEnd As Single
        
    If sTrim = False Then
        For LP = 1 To Len(Text)
            If Asc(Mid$(Text, LP, 1)) > 31 And Asc(Mid$(Text, LP, 1)) < 127 Then
                TempStr = TempStr & Mid$(Text, LP, 1)
            Else
                If cReplace <> "" Then
                    TempStr = TempStr & cReplace
                End If
            End If
        Next LP
    Else
        For LP = 1 To Len(Text)
            If Asc(Mid$(Text, LP, 1)) > 31 And Asc(Mid$(Text, LP, 1)) < 127 Then
                sStart = LP
                Exit For
            End If
        Next LP
        For LP = Len(Text) To 1 Step -1
            
            If Asc(Mid$(Text, LP, 1)) < 31 Or Asc(Mid$(Text, LP, 1)) > 127 Then
                sEnd = (LP - sStart) '+ 1
            End If
        Next LP
        If sStart = 0 And sEnd = 0 Then Exit Function
        TempStr = Mid$(Text, sStart, sEnd)
        
    End If
        
    If RemoveCr = True Then
        Text = TempStr
        If InStr(Text, vbCrLf) Then
            TempStr = ""
        
            sSplit = Split(Text, vbCrLf)
            
            For LP = 0 To UBound(sSplit)
                If Trim(sSplit(LP)) <> "" Then
                    TempStr = TempStr & sSplit(LP) & vbCrLf
                End If
            Next LP
           StripNonChar TempStr, "", True
        End If
    End If
        
    StripNonChar = Trim(TempStr)

End Function

Public Function GetREGSZVal(Key As String, PropertyName As String) As String

    Dim Hkey As Long
    Dim C As Long
    Dim r As Long
    Dim S As String
    Dim T As Long
    
    r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key, 0, KEY_READ, Hkey)
    C = 255
    S = String(C, Chr(0))
    r = RegQueryValueEx(Hkey, PropertyName, 0, T, S, C)
   
    GetREGSZVal = Trim(Left(S, C - 1))
    
    RegCloseKey Hkey
    
End Function

