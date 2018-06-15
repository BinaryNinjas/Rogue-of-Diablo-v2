Attribute VB_Name = "modFunc"
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_FRAMECHANGED = &H20
Public Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function SetTimer Lib "User32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "User32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public TestedCharCount As Long
Public sckX As Integer
      Public TimerID As Long

'''''''''''
Public Type BNCS_Info
D2DVKey As String
D2XPKey As String
BnetAccount As String
Server As String
Verbyte As String
End Type

Public BNCS_Stuff As BNCS_Info
Public InProcessOfConnecting As Boolean
Public flags As Integer
Public ClassFlag As Integer

Public ListCnt As Integer
Public Keycnt As Integer
Public EXPKeycnt As Integer
Public Uper(500) As Integer
Public BNCSCnt As Integer
Public CharNamesTested As Long
Public AllBotsConnected As Boolean
Public CharFullRC As Boolean
Public AutoSwapFull As Boolean
Public Type WhosLive
User As String ' BNCS username whose account hash is in realm
Live As Boolean ' Whether that hash is or has been connected
End Type

Public CharDuring0x02CreatedCounter(500) As Integer

Public Spawn As String
Public CharList(500) As New clsHashTable
Public D2DVList As New clsHashTable
Public D2XPList As New clsHashTable
Public BNCSList As New clsHashTable

'''''''''''


'Following two functions are used to move the windows around by clicking directly on
'the window in abscence of the standard window title bars
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Function CheckFlags(flags As Integer)
 Select Case flags

Case "0"
frmRoD.lblType.Caption = "Type: Classic"
IsExpansion = False

Case "40"
frmRoD.lblType.Caption = "Type: Classic Ladder"
IsExpansion = False

Case "4"
frmRoD.lblType.Caption = "Type: Classic Hardcore"
IsExpansion = False

Case "44"
frmRoD.lblType.Caption = "Type: Classic Hardcore Ladder"
IsExpansion = False

Case "20"
frmRoD.lblType.Caption = "Type: Expansion"
IsExpansion = True
Case "60"
frmRoD.lblType.Caption = "Type: Expansion Ladder"
IsExpansion = True

Case "24"
frmRoD.lblType.Caption = "Type: Expansion Hardcore"
IsExpansion = True

Case "64"
frmRoD.lblType.Caption = "Type: Expansion Hardcore Ladder"
IsExpansion = True

End Select

End Function

Public Function WriteStuff(appname As String, Key As String, sString As String, Optional strIni As String) As Boolean
Dim sFile As String
Dim L As Long
WriteStuff = False
On Error GoTo WriteStuff_Error
If strIni = vbNullString Then
    sFile = App.Path & "\Config.ini"
Else
    sFile = App.Path & "\" & strIni
End If
L = WritePrivateProfileStringA(appname, Key, sString, sFile)
WriteStuff = True

WriteStuff_Error:
If Err.Number <> 0 Then
MsgBox Err.Description
End If
End Function
Public Function GetStuff(appname As String, Key As String, Optional strIni As String) As String
Dim sFile As String
Dim sDefault As String
Dim lSize As Integer
Dim L As Long
Dim sUser As String
sUser = Space$(128)
lSize = Len(sUser)
If strIni = vbNullString Then
    sFile = App.Path & "\Config.ini"
Else
    sFile = strIni
End If
sDefault = vbNullString
L = GetPrivateProfileStringA(appname, Key, sDefault, sUser, lSize, sFile)
sUser = Mid(sUser, 1, InStr(sUser, Chr(0)) - 1)
GetStuff = sUser
End Function
Public Function LoadData(FiileName As String)
Dim whole_file As String
Dim animals() As String

Dim i As Integer
 

    ' Get the whole file.
    whole_file = GrabFile(App.Path & "\" & FiileName)
    ' Break the file into lines.
    animals = Split(whole_file, vbCrLf)
     
    Do Until i = UBound(animals) + 1
    
'skip blank lines

''
    If Mid$(Trim(animals(i)), 1) <> "" Then
If CharList(ListCnt).Exists(Str(CharList(ListCnt).HashCode("Char" & i))) = True Then
'Chat.AddChat vbWhite, "Person already in UserDB"
 
Else
 CharList(ListCnt).Add Str(CharList(ListCnt).HashCode("Char" & i)), Mid$(Trim(animals(i)), 1)
  'AddChat "Added " & Mid$(Trim(animals(i)), 1) & " to Char List.", vbRed
 
 End If
    End If
   i = i + 1
   DoEvents
   Loop
   AddChat "Loaded " & CharList(ListCnt).Count & " Character names.", vbRed
ListCnt = ListCnt + 1

End Function
Public Function WriteInvalid(NameChar As String)
      Dim sFileText As String
  
      Dim iFileNo As Integer
   
        iFileNo = FreeFile
   
            'open the file for writing
   
        Open App.Path & "\Invalid.txt" For Append As #iFileNo
   
      'please note, if this file already exists it will be overwritten!
  
            'write some example text to the file

        Print #iFileNo, NameChar
 
       

            'close the file (if you dont do this, you wont be able to open it again!)
     Close #iFileNo
End Function
Public Function WriteCreated(NameChar As String)
      Dim sFileText As String
  
      Dim iFileNo As Integer
   
        iFileNo = FreeFile
   
            'open the file for writing
   
        Open App.Path & "\Created.txt" For Append As #iFileNo
   
      'please note, if this file already exists it will be overwritten!
  
            'write some example text to the file

        Print #iFileNo, NameChar
 
       

            'close the file (if you dont do this, you wont be able to open it again!)
     Close #iFileNo
End Function
Public Function LoadBNCS(FiileName As String)
Dim whole_file As String
Dim animals() As String

Dim i As Integer
 

    ' Get the whole file.
    whole_file = GrabFile(App.Path & "\" & FiileName)
    ' Break the file into lines.
    animals = Split(whole_file, vbCrLf)
     
    Do Until i = UBound(animals) + 1
    
    'skip blank lines
'MsgBox animals(i)
''


    
    If Mid$(Trim(animals(i)), 1) <> "" Then
    
    If InStr(1, Mid$(Trim(animals(i)), 1), "//") = False Then
MsgBox animals(i) & " does not have // seperating User//Pass. Close and bot and fix please."
Exit Function
End If

If BNCSList.Exists(Str(BNCSList.HashCode("BNCS" & i))) = True Then
'Chat.AddChat vbWhite, "Person already in UserDB"
 
Else
 BNCSList.Add Str(BNCSList.HashCode("BNCS" & i)), LCase(Mid$(Trim(animals(i)), 1))
  'AddChat "Added " & Mid$(Trim(animals(i)), 1) & " to Char List.", vbRed
 
 End If
    End If
   i = i + 1
   DoEvents
   Loop
   AddChat "Loaded " & BNCSList.Count & " Battle.net accounts.", vbRed
End Function
Public Function LoadD2Keys(FiileName As String)
Dim whole_file As String
Dim animals() As String

Dim i As Integer
 

    ' Get the whole file.
    whole_file = GrabFile(App.Path & "\" & FiileName)
    ' Break the file into lines.
    animals = Split(whole_file, vbCrLf)
     
    Do Until i = UBound(animals) + 1
    
    'skip blank lines

''
    If Mid$(Trim(animals(i)), 1) <> "" Then
If D2DVList.Exists(Str(D2DVList.HashCode("D2DV" & i))) = True Then
'Chat.AddChat vbWhite, "Person already in UserDB"
 
Else
 D2DVList.Add Str(D2DVList.HashCode("D2DV" & i)), LCase(Mid$(Trim(animals(i)), 1))
  'AddChat "Added " & Mid$(Trim(animals(i)), 1) & " to Char List.", vbRed
 
 End If
 End If
   i = i + 1
   DoEvents
   Loop
   AddChat "Loaded " & D2DVList.Count & " D2DV CdKeys.", vbRed
End Function
Public Function LoadD2XPKeys(FiileName As String)
Dim whole_file As String
Dim animals() As String

Dim i As Integer
 

    ' Get the whole file.
    whole_file = GrabFile(App.Path & "\" & FiileName)
    ' Break the file into lines.
    animals = Split(whole_file, vbCrLf)
     
    Do Until i = UBound(animals) + 1
    
    'skip blank lines

''
    If Mid$(Trim(animals(i)), 1) <> "" Then
If D2XPList.Exists(Str(D2XPList.HashCode("D2XP" & i))) = True Then
'Chat.AddChat vbWhite, "Person already in UserDB"
 
Else
 D2XPList.Add Str(D2XPList.HashCode("D2XP" & i)), LCase(Mid$(Trim(animals(i)), 1))
  'AddChat "Added " & Mid$(Trim(animals(i)), 1) & " to Char List.", vbRed
 
 End If
    End If
   i = i + 1
   DoEvents
   Loop
   AddChat "Loaded " & D2XPList.Count & " D2XP CdKeys.", vbRed
End Function

   Private Function GrabFile(ByVal file_name As String) As _
    String
Dim fnum As Integer

    On Error GoTo NoFile
    fnum = FreeFile
    Open file_name For Input As fnum
    GrabFile = Input$(LOF(fnum), fnum)
    Close fnum
    Exit Function

NoFile:
    GrabFile = ""
    MsgBox "Couldnt Grab File :( something's wrong!! (" & file_name & ")"
End Function

Public Function AddCharAccs(Uxer As String)
Uxer = Trim(Uxer)
If CharList(ListCnt).Exists(Str(CharList(ListCnt).HashCode(LCase(Uxer)))) = True Then
'Chat.AddChat vbWhite, "Person already in UserDB"
 
Else
 CharList(ListCnt).Add Str(CharList(ListCnt).HashCode(LCase(Uxer))), LCase(Uxer)
 modBnetPKTS.Send0x0E Uxer & " has been added to the database."
 modBnetPKTS.Send0x0E "[Database=" & CharList(ListCnt).Count & " users]"
 
 End If
End Function


Public Sub ConnectBNCS(watsock As Integer)
'sckClosed = 0
'sckClosing = 8
'sckConnected = 7
'sckConnecting = 6
'sckConnectionPending = 3
'sckError = 9
'sckHostResolved = 5
'sckListening = 2
'sckOpen = 1
'sckResolvingHost = 4
Select Case frmRoD.sckBnet(watsock).State


    Case "0"
    AddChat "BNCS(" & watsock & ") Connecting to " & BNCS_Stuff.Server, vbRed
    frmRoD.sckBnet(watsock).Close
        frmRoD.sckBnet(watsock).Connect BNCS_Stuff.Server, 6112
    Case "7"
AddChat "BNCS(" & watsock & ") Connecting to " & BNCS_Stuff.Server, vbRed
    frmRoD.sckBnet(watsock).Close
        frmRoD.sckBnet(watsock).Connect BNCS_Stuff.Server, 6112
        
        Case "3"
        AddChat "BNCS(" & watsock & ") is still trying to connect..", vbRed

Case "9"
        AddChat "BNCS(" & watsock & ") Socket Error. Make sure you're connecting to the correct server.", vbRed
Case Else
AddChat "Winsock Error: Socket State: " & frmRoD.sckBnet(watsock).State, vbRed

        
        
End Select

End Sub
Sub pause(interval As String)

    Dim X
    X = Timer


    Do While Timer - X < Val(interval)


        DoEvents
        Loop

    End Sub
Public Sub ConnectMCP(watsock As Integer)
'frmRoD.sckMCP(watsock).Close
frmRoD.sckMCP(watsock).RemoteHost = IP
frmRoD.sckMCP(watsock).RemotePort = Port
frmRoD.sckMCP(watsock).Connect IP, Port
Chat.AddChat "Realm IP: " & IP & ":" & Port, vbRed
End Sub


Public Function SanityCheck()

'Check too see if everything is ok
Dim ErrorList(1) As String
Dim AnyError As Boolean
Dim ServerErr As Boolean
Dim KeyErr As Boolean
Dim SockErr As Boolean
Dim VerErr As Boolean
Dim BNCSActErr As Boolean
Dim ErrX As Integer
Dim i As Integer

If BNCS_Stuff.Server = "" Then 'blank server

ErrorList(ErrX) = "[You have not filled out which server to connect to.]"
SockErr = True
ErrX = ErrX + 1
AnyError = True
Else
'ErrorList(ErrX) = "[Server Good]"

End If

If BNCS_Stuff.Verbyte = "" Then
ErrorList(ErrX) = "[You have not inputted a Version Byte]"
VerErr = True
ErrX = ErrX + 1
AnyError = True
Else
'ErrorList(ErrX) = "[Verbyte Good]"

End If



For i = 0 To UBound(ErrorList)

AddChat ErrorList(i), vbBlue


Next i



If AnyError = True Then
frmSettings.Show


End If

End Function

Function TimerProc(hwnd As Long, uMsg As Long, EventID As Long, dwTime As Long) As Long
  Dim MSG As VbMsgBoxResult
    'This code will display every designated interval
    
    AddChat "5 secs have elasped - Still Waiting...", vbBlue
    
    
    If AllBotsConnected = True Then
        Call KillTimer(frmRoD.hwnd, 1000)
    Exit Function
    End If
    
End Function

