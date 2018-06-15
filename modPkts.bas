Attribute VB_Name = "modPkts"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub calcHashBuf Lib "BNCSutil.dll" (ByVal Data As String, ByVal Length As Long, ByVal Hash As String)
Public Declare Function checkRevisionFlat Lib "BNCSutil.dll" (ByVal ValueString As String, ByVal File1 As String, ByVal File2 As String, ByVal File3 As String, ByVal mpqNumber As Long, ByRef Checksum As Long) As Long

Public Declare Function kd_quick Lib "BNCSutil.dll" _
    (ByVal CDKey As String, ByVal ClientToken As Long, _
    ByVal ServerToken As Long, PublicValue As Long, _
    Product As Long, ByVal HashBuffer As String, _
    ByVal BufferLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)

Public Declare Function kd_product Lib "BNCSutil.dll" (ByVal decoder As Long) As Long
Public Declare Function CheckRevisionEx Lib "CheckRevision.dll" (ByVal GameFile1 As String, ByVal GameFile2 As String, ByVal GameFile3 As String, ByVal ValueString As String, ByRef version As Long, ByRef Checksum As Long, ByVal exeinfo As String, ByVal PathToDLL As String, ByVal sUnused As String, ByVal PathToVideoBin As String) As Long

Private Const KeyCodes As String = "246789BCDEFGHJKMNPRTVWXZ"
'this is how we are going to use the packet buffer
Public buf2 As New clsBuffer
    Public Port As String
    Public IP As String
Public tmpFileName As String
Public tmpFormula As String
Public IsExpansion As Boolean
Public UDPToken As Long
Public ServerToken As Long
Public Exp As Boolean
Public RealmTitle As String
Public MCPcookie As Long
Public MCPStatus As Long
Public MCPChunk1 As String
Public MCPChunk2 As String


Public Sub BNLS_0x1A(ByVal FileTime As String)
Chat.AddChat "Sending BNLS_0x1A", &HC0C0C0
'(DWORD) Product ID.*
'(DWORD) Flags.**
'(DWORD) Cookie.
'(ULONGLONG) Timestamp for version check archive.
'(String) Version check archive filename.
'(String) Checksum formula.

'* Valid product IDs are:

'#define PRODUCT_STARCRAFT             (0x01)
'#define PRODUCT_BROODWAR              (0x02)
'#define PRODUCT_WAR2BNE               (0x03)
'#define PRODUCT_DIABLO2               (0x04)
'#define PRODUCT_LORDOFDESTRUCTION     (0x05)
'#define PRODUCT_JAPANSTARCRAFT        (0x06)
'#define PRODUCT_WARCRAFT3             (0x07)
'#define PRODUCT_THEFROZENTHRONE       (0x08)

'** The flags field is currently reserved and must be set to zero or you will be disconnected.

'MsgBox "Filetime = " & frmmain.StrToHex(FileTime)
With buf2


If Exp = False Then

'Chat.AddChat Index, "BNLS for D2", vbYellow
.InsertDWORD &H4  'BW product ID
Else
'Chat.AddChat Index, "BNLS for D2XP", vbYellow
.InsertDWORD &H5  'BW product ID

End If




.InsertDWORD &H0 'Flags
.InsertDWORD &H1 ' Cookie
'.InsertDWORD ServerToken 'Timestamp for version check archive
.InsertSTRING FileTime
.InsertNTString frmMain.tmpFileName 'filename
.InsertNTString frmMain.tmpFormula 'formula
.InsertbnlsHEADER &H1A
.SendPacket frmMain.sckBNLS
End With

End Sub


 
 Public Function DecodeD2(ByVal CDKey As String) As String
    Dim tmpByte As Byte, i%, A&, B&, R&, Key$(15)
    For i = 1 To 16 'Fill array
        Key(i - 1) = UCase(Mid$(CDKey, i, 1))
    Next i
    Dim IntStr%(1), i2%
    R = 1 'base flag
    For i = 0 To 14 Step 2
        For i2 = 0 To 1
            IntStr(i2) = InStr(1, KeyCodes, Key(i + i2)) - 1
            If IntStr(i2) = -1 Then IntStr(i2) = &HFF
            If i2 = 0 Then A = IntStr(i2) * 3 Else A = IntStr(i2) + A * 8
        Next i2
        If A >= &H100 Then
            A = A - &H100
            tmpByte = tmpByte Or R 'set flag
        End If
        B = ((RShift(A, 4) And &HF) + &H30)
        A = ((A And &HF) + &H30)
        If B > &H39 Then B = B + &H7
        If A > &H39 Then A = A + &H7
        Key(i) = Chr$(B)
        Key(i + 1) = Chr$(A)
        R = R * 2 'upgrade flag
    Next i
    Erase IntStr()
    '//Valid Check
    R = 3
    For i = 0 To 15
        R = R + (GetNumValue(Key(i)) Xor (R * 2))
    Next i
    R = R And &HFF
    If Not R = tmpByte Then
        'Cdkey is shit
    End If
    '//Shuffling
    Dim tmpD As String * 1
    For i = 15 To 0 Step -1
        If i > 8 Then tmpByte = ((i - 9) And &HF) Else tmpByte = ((i + 7) And &HF)
        tmpD = Key(i)
        Key(i) = Key(tmpByte)
        Key(tmpByte) = tmpD
    Next i
    '//hash Values
    Dim HashKey&
    HashKey = &H13AC9741
    For i = 15 To 0 Step -1
        tmpByte = Asc(Key(i))
        If tmpByte <= &H37 Then
            Key(i) = Chr$(((HashKey And &HFF) And 7) Xor tmpByte)
            HashKey = RShift(HashKey, 3)
        ElseIf tmpByte < &H41 Then
            Key(i) = Chr$((i And 1) Xor tmpByte)
        Else
            Key(i) = Chr$(tmpByte)
        End If
    Next i
    '//return key
    DecodeD2 = Join(Key, vbNullString)
    Erase Key()
End Function


Public Function EncodeD2(ByVal CDKey As String) As String
    Dim tmpByte As Byte, i%, A&, B&, R&, Key$(15)
    For i = 1 To 16 'Fill array
        Key(i - 1) = UCase(Mid$(CDKey, i, 1))
    Next i
    '//unhashsing
    Dim HashKey&
    HashKey = &H13AC9741
    For i = 15 To 0 Step -1
        tmpByte = Asc(Key(i))
        If tmpByte <= &H37 Then
            Key(i) = Chr$(((HashKey And &HFF) And 7) Xor tmpByte)
            HashKey = RShift(HashKey, 3)
        ElseIf Val(tmpByte) < &H41 Then
            Key(i) = Chr$((i And 1) Xor tmpByte)
        Else
            Key(i) = Chr$(tmpByte)
        End If
    Next i
    '//unshuffling
    Dim tmpD As String * 1
    For i = 0 To 15
        If i > 8 Then tmpByte = ((i - 9) And &HF) Else tmpByte = ((i + 7) And &HF)
        tmpD = Key(i)
        Key(i) = Key(tmpByte)
        Key(tmpByte) = tmpD
    Next i
    '//flag extract
    R = 3
    For i = 0 To 15
        R = R + (GetNumValue(Key(i)) Xor (R * 2))
    Next i
    R = R And &HFF
    tmpByte = &H80 'seed the flag
    '//convert hex to KeyCodes
    For i = 14 To 0 Step -2
        A = GetNumValue(Key(i))
        B = GetNumValue(Key(i + 1))
        A = CLng("&H" & Hex(A) & Hex(B))
        If R And tmpByte Then A = A + &H100
        Call KeyCodeOffSets(A, B)
        Key(i) = Mid(KeyCodes, B + 1, 1)
        Key(i + 1) = Mid(KeyCodes, A + 1, 1)
        tmpByte = tmpByte / 2 'downgrade flag
    Next i
    '//return encoded key
    EncodeD2 = Join(Key, vbNullString)
    Erase Key()
End Function

Private Sub KeyCodeOffSets(Bit1&, Bit2&)
    Bit2 = 0
    While Bit1 >= &H18
        Bit2 = Bit2 + 1
        Bit1 = Bit1 - &H18
    Wend
End Sub

Public Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
    On Error Resume Next
    RShift = CDbl(pnValue \ (2 ^ pnShift))
End Function

Public Function GetNumValue(ByVal C As String) As Long
    On Error Resume Next
    C = UCase(C)
    If IsNumeric(C) Then
        GetNumValue = Asc(C) - &H30
    Else
        GetNumValue = Asc(C) - &H37
    End If
End Function

Public Sub Send0x40(Index As Integer)
''0x3A, 0x40, 0x3E, 0x01, 0x19, 0x07, abc
'[2:15:29 PM] - : Pro_Tech : -  once u send 0x19 and get the character list, thats when u would  create/delete charcter, ifu dont wanna logon one just yet
Chat.AddChat "Sending 0x40 Querying Realm", vbRed
With buf2
.InsertHEADER &H40
.SendPacket frmRoD.sckBnet(Index)
End With


End Sub

Public Sub Send0x3E(Index As Integer)
Chat.AddChat "BNCS(" & Index & ") Sending 0x3E Logging onto Realm", vbRed
Dim ClientToken As Long
 ClientToken = GetTickCount 'any old number (cant be 0)
    Dim pHash As String * 20
    Dim InHash As String
    Dim outhash As String
    
outhash = String(20, 0)
InHash = "password"

Call calcHashBuf(InHash, Len(InHash), outhash)

InHash = buf2.MakeDWORD(ClientToken) & buf2.MakeDWORD(ServerToken) & outhash

Call calcHashBuf(InHash, Len(InHash), outhash)


With buf2

.InsertDWORD ClientToken
.InsertSTRING outhash
.InsertNTString RealmTitle
.InsertHEADER &H3E
.SendPacket frmRoD.sckBnet(Index)

End With


End Sub

Public Sub Send0x01MCP(Index As Integer)
Chat.AddChat "MCP(" & Index & ") Sending 0x01 MCP", vbRed
Dim User As String
 
User = Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1)

With buf2

.InsertDWORD MCPcookie
.InsertDWORD MCPStatus
.InsertNonNTString MCPChunk1
.InsertNonNTString MCPChunk2
.InsertNTString User
.InsertMCPHEADER &H1
.SendPacket frmRoD.sckMCP(Index)

End With

End Sub

Public Sub Send0x02MCP(Index As Integer)
' Chat.AddChat "MCP(" & Index & ") Sending 0x02 MCP", vbRed
'AddChat "MCP(" & Index & ") Creating: " & CharList(ListCnt).Item(Str(CharList(ListCnt).HashCode("Char" & Uper(ListCnt)))), vbBlack
 
If CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index)))) = "" Then
Uper(Index) = "0"
'AddChat "count reset", vbRed
End If

With buf2

.InsertDWORD "&H" & ClassFlag 'Class
.InsertWORD "&H" & flags 'Flags
.InsertNTString CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index))))     'char name
.InsertMCPHEADER &H2
.SendPacket frmRoD.sckMCP(Index)

End With
 Uper(Index) = Uper(Index) + 1

End Sub
Public Sub Send0x07MCP()
Chat.AddChat "Sending 0x07 MCP", vbWhite

With buf2


.InsertNTString User
.InsertMCPHEADER &H7
.SendPacket frmMain.sckMCP

End With

End Sub

Public Sub Send0x19MCP(Index As Integer)
Chat.AddChat "MCP(" & Index & ") Sending 0x19 MCP", vbWhite

With buf2

.InsertDWORD &H8
.InsertMCPHEADER &H19
.SendPacket frmRoD.sckMCP(Index)

End With

End Sub

Public Sub Send0x12MCP()
Chat.AddChat "Sending 0x12 MCP", vbWhite

With buf2

.InsertMCPHEADER &H12
.SendPacket frmMain.sckMCP

End With

End Sub
Public Sub Send0x50(Index As Long)
Chat.AddChat "BNCS(" & Index & ") Sending 0x50", &HC0C0C0

    With buf2
        .InsertDWORD &H0
        .InsertDWORD &H49583836 '68XI
        
        If frmRoD.chkEx = vbChecked Then
         .InsertDWORD .GetDWORD("PX2D")  'GetDWORD("NB2W") 'ClientID)  'turn string into number into packer buffer
        Else
         .InsertDWORD .GetDWORD("VD2D")  'GetDWORD("NB2W") 'ClientID)  'turn string into number into packer buffer
        End If
        
        .InsertDWORD "&H" & BNCS_Stuff.Verbyte   'version byte for sc/bw
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertNTString "GBR" 'null byte terminated string
        .InsertNTString "United Kingdom"
        .InsertHEADER &H50 'add the header FF 50 and 2 byte lengh of packet buffer
        .SendPacket frmRoD.sckBnet(Index)   'hand the winsock to the sub so it can use it to send the data down it

    End With
    
End Sub
Public Sub Send0x51(Index As Integer, tmpFormula As String, tmpFileName As String)
AddChat "BNCS(" & Index & ") Sending 0x51", vbRed
 Dim Product As Long
Dim PublicValue As String
Dim pubVal As String
Dim PrivateValue As String
Dim CDKey As String, Decode As String

'Exp Keys
Dim Decode2 As String
Dim Cdkey2 As String
Dim Product2 As Long
Dim PublicValue2 As String
Dim pubVal2 As String
Dim PrivateValue2 As String

CDKey = D2DVList.Item(Str(D2DVList.HashCode("D2DV" & Keycnt)))
Cdkey2 = D2XPList.Item(Str(D2XPList.HashCode("D2XP" & EXPKeycnt)))

'''''''''''''''MAKE HASH CHECK FUNCTION
If Dir$(App.Path & "\D2DV\Game.exe") = "" Then
    MsgBox "Hashes Not Found, Please put Hash Files in /D2DV"
    Exit Sub
Else
'MsgBox "Hashes Found"
End If
If Dir$(App.Path & "\D2XP\Game.exe") = "" Then
    MsgBox "Hashes Not Found, Please put Hash Files in /D2XP"
    Exit Sub
Else
'MsgBox "Hashes Found"
End If

Dim Checksum As Long
Dim exeversion As Long
Dim exeinfo As String
'tmpFileName = Replace(tmpFileName, "ver-", "")
 'MsgBox tmpFileName
 
 Select Case IsExpansion
 Case True
If checkrevision(App.Path & "\D2XP\Game.exe", App.Path & "\D2XP\BNClient.dll", App.Path & "\D2XP\D2Client.dll", tmpFormula, exeversion, Checksum, exeinfo, tmpFileName) = 0 Then
    MsgBox "CheckRevision Failed!"
    Exit Sub
Else
    'MsgBox "CheckRevision PASSED!!"
End If
Case Else
If checkrevision(App.Path & "\D2DV\Game.exe", App.Path & "\D2DV\BNClient.dll", App.Path & "\D2DV\D2Client.dll", tmpFormula, exeversion, Checksum, exeinfo, tmpFileName) = 0 Then
    MsgBox "CheckRevision Failed!"
    Exit Sub
Else
    'MsgBox "CheckRevision PASSED!!"
End If

End Select
'MsgBox exeinfo
 
 
 
 
Select Case IsExpansion
 
 


Case True

Decode = DecodeD2(CDKey)
Product = Mid$(Decode, 1, 2)
 pubVal = Mid$(Decode, 3, 6)
PrivateValue = Mid$(Decode, 9, 8)
''
Decode2 = DecodeD2(Cdkey2)
Product2 = CLng("&H" & Mid$(Decode2, 1, 2))
 pubVal2 = Mid$(Decode2, 3, 6)
PrivateValue2 = Mid$(Decode2, 9, 8)
Debug.Print "Product = " & Product2
Debug.Print "Public = " & pubVal2
Debug.Print "Private = " & PrivateValue2

Case False
 
'If Len(CDKey) = "26" Then
'Chat.AddChat "Using NEW D2Keys", vbWhite
'Call Send0x51WAR3(tmpFormula, tmpFileName, Checksum, exeversion, exeinfo)
'Exit Sub
'End If

Decode = DecodeD2(CDKey)
Product = Mid$(Decode, 1, 2)
 pubVal = Mid$(Decode, 3, 6)
PrivateValue = Mid$(Decode, 9, 8)


End Select

 Dim ClientToken As Long
 ClientToken = GetTickCount 'any old number (cant be 0)

Dim outhash As String * 20
Dim outhash2 As String * 20
If kd_quick(CDKey, ClientToken, ServerToken, buf2.GetDWORD(pubVal), Product, outhash, 20) = 0 Then
AddChat "CDKey is Invalid, switching Keys", &HBDB76B
Keycnt = Keycnt + 1
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Exit Sub
End If

If IsExpansion = True Then
If kd_quick(Cdkey2, ClientToken, ServerToken, buf2.GetDWORD(pubVal2), Product2, outhash2, 20) = 0 Then
AddChat "Exp CDKey is Invalid, switching Keys", &HBDB76B
EXPKeycnt = EXPKeycnt + 1
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Exit Sub
End If
    End If
    
    
DoEvents

Dim KeyHash As String
Dim Keyhash2 As String

Select Case IsExpansion

 
Case True
Keyhash2 = buf2.MakeDWORD(ClientToken) & buf2.MakeDWORD(ServerToken) & buf2.MakeDWORD(Product2) & buf2.MakeDWORD(ConvertHexToLong(pubVal2)) & buf2.MakeDWORD(&H0) & buf2.MakeDWORD(ConvertHexToLong(PrivateValue2))

 
Case Else


KeyHash = buf2.MakeDWORD(ClientToken) & buf2.MakeDWORD(ServerToken) & buf2.MakeDWORD(Product) & buf2.MakeDWORD(ConvertHexToLong(pubVal)) & buf2.MakeDWORD(&H0) & buf2.MakeDWORD(ConvertHexToLong(PrivateValue))

End Select



 
 
    


With buf2
.InsertDWORD ClientToken 'GetTickCount()

 .InsertDWORD exeversion
 .InsertDWORD Checksum    'EXE hash


If IsExpansion = True Then
.InsertDWORD &H2 ' 2 keys
Else
.InsertDWORD &H1 '1 cdkey
End If


 .InsertDWORD &H0 'no spawning
 
 
 .InsertDWORD 16 'd2, w2 key len
 
 
 
.InsertDWORD Product  'Product Value of Key
'Product = 6
'[8:51:25 PM] Public Val = D3E323
'[8:51:25 PM] Private = F6F45582

 
  .InsertDWORD ConvertHexToLong(pubVal)   'buf2.GetDWORD(pubVal)

 



 .InsertDWORD &H0 'Null
.InsertSTRING outhash 'Hashed Key Data

If IsExpansion = True Then

.InsertDWORD 16
.InsertDWORD Product2  'Product Value of Key
.InsertDWORD ConvertHexToLong(pubVal2)   'buf2.GetDWORD(pubVal)
.InsertDWORD &H0
.InsertSTRING outhash2
End If
.InsertNTString exeinfo '"game.exe 03/04/07 18:24:51 57344" 'exeinfo  'Exe Info
 .InsertNTString "RoDv2"  'Owner Name
.InsertHEADER &H51
 .SendPacket frmRoD.sckBnet(Index)
 
 End With
 
 
    


End Sub
Public Sub Send0x51WAR3(tmpFormula As String, tmpFileName As String, Checksum As Long, exeversion As Long, exeinfo As String)
Call Chat.AddChat("Sending 0x51", vbWhite)
Dim Product As Long
Dim PublicValue As Long
Dim PrivateValue As Long
Dim CDKey As String

'CDKey = modFunc.D2DVKey
Dim ClientToken As Long
 ClientToken = GetTickCount 'any old number (cant be 0)
 

 Dim outhash As String * 20
'outhash = String(20, 0)
 'Call calcHashBuf(KeyHash, Len(KeyHash), outhash)

If kd_quick(CDKey, ClientToken, ServerToken, PublicValue, Product, outhash, 20) = 0 Then
MsgBox "Failed decode"
Exit Sub
End If
DoEvents
 
 Dim Check As Long
 
  
 
With buf2
.InsertDWORD ClientToken 'GetTickCount()
 .InsertDWORD exeversion  'version of the exe file
 .InsertDWORD Checksum  'EXE hash
.InsertDWORD &H1 '1 cdkey
 .InsertDWORD &H0 'no spawning
 .InsertDWORD 26 'length of key
.InsertDWORD Product  'Product Value of Key i.e 01 or 02
.InsertDWORD PublicValue  'Public Value of Cd Key 7digit number
 .InsertDWORD &H0 'Null
.InsertSTRING outhash 'Hashed Key Data
.InsertNTString exeinfo 'exeinfo 'Exe Info
 .InsertNTString "Rogue"  'Owner Name
.InsertHEADER &H51
 .SendPacket frmMain.sckBnet
 End With
 
End Sub

Public Sub Send0x3A(Index As Integer)   'Logon
Call AddChat("BNCS(" & Index & ") Sending 0x3A", vbRed)
Dim ClientToken As Long
 ClientToken = GetTickCount 'any old number (cant be 0)
    Dim pHash As String * 20
    Dim InHash As String
    Dim outhash As String
    
outhash = String(20, 0)
Dim User As String
Dim Pass As String

User = Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1)
Pass = Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), Len(User) + 3)
 
InHash = LCase(Pass)
Call calcHashBuf(InHash, Len(InHash), outhash)

InHash = buf2.MakeDWORD(ClientToken) & buf2.MakeDWORD(ServerToken) & outhash

Call calcHashBuf(InHash, Len(InHash), outhash)


'(DWORD) Client Token
'(DWORD) Server Token
'(DWORD[5]) Password Hash
'(STRING)  Username
With buf2
.InsertDWORD ClientToken
.InsertDWORD ServerToken
.InsertSTRING outhash
.InsertNTString User
.InsertHEADER &H3A

  .SendPacket frmRoD.sckBnet(Index)
  

 End With

End Sub
Public Sub Send0x14(Index As Integer, udpstamp As Long)
'Call Chat.ShowChat(i, vbGreen, "Sending 0x14")
Chat.AddChat "BNCS(" & Index & ") Sending 0x14", vbRed
With buf2
.InsertDWORD udpstamp
.InsertHEADER &H14
 .SendPacket frmRoD.sckBnet(Index)
 End With
End Sub



Public Function ConvertHexToLong(sHex As String) As Long

On Error GoTo errHandler:
    Dim n As Integer
    Dim sTemp As String * 1
    Dim nTemp As Integer
    Dim nFinal() As Integer
    Dim bNegative As Boolean
    ReDim nFinal(0)
    If LenB(sHex) = 0 Then
        ConvertHexToLong = 0
        Exit Function
    End If
    bNegative = False
    For n = Len(sHex) To 1 Step -1
        sTemp = Mid$(sHex, n, 1)
        nTemp = Val(sTemp)
        If nTemp = 0 Then
            Select Case UCase(sTemp)
                Case "A"
                    nTemp = 10
                Case "B"
                    nTemp = 11
                Case "C"
                    nTemp = 12
                Case "D"
                    nTemp = 13
                Case "E"
                    nTemp = 14
                Case "F"
                    nTemp = 15
                Case "-"
                    bNegative = True
                Case Else
                    nTemp = 0
            End Select
        End If
        ReDim Preserve nFinal(UBound(nFinal) + 1)
        nFinal(UBound(nFinal)) = nTemp
    Next
    ConvertHexToLong = nFinal(1)
    For n = 2 To UBound(nFinal)
        ConvertHexToLong = ConvertHexToLong + (nFinal(n) * (4 ^ (n * 2 - 2)))
    Next
    If bNegative Then ConvertHexToLong = ConvertHexToLong - (ConvertHexToLong * 2)
Exit Function
errHandler:

ConvertHexToLong = 0

End Function
Public Sub Send0x3D(Index As Integer)
Dim User As String
Dim Pass As String
User = Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1)
Pass = Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), Len(User) + 3)
 

Chat.AddChat "BNCS(" & Index & ") Attempting to make account: " & User, vbRed
    'Dim pHash As String * 20
    Dim InHash As String
    Dim outhash As String * 20
    'Call Chat.ShowChat(i, vbGreen, "Sending 0x3D")
outhash = String(20, 0)


InHash = Pass

Call calcHashBuf(InHash, Len(InHash), outhash)


With buf2
.InsertSTRING outhash
.InsertNTString User
.InsertHEADER &H3D
.SendPacket frmRoD.sckBnet(Index)
End With



End Sub
