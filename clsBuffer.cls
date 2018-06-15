VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function send Lib "ws2_32.dll" (ByVal sckHandle As Long, ByRef InBuf As Any, ByVal BufLen As Long, ByVal flags As Long) As Long

Private Buffer As String

Public Sub InsertWORD(ByVal Num2Byte As Integer)
    'turns a integer number into a 2 byte string and adds it to the buffer
    Buffer = Buffer & MakeWORD(Num2Byte)
End Sub
Public Sub InsertDWORD(ByVal Num4Byte As Long)
    'turns a long number into a 4 byte string and adds it to the buffer
    Buffer = Buffer & MakeDWORD(Num4Byte)
End Sub
Public Sub InsertDWORD64bit(ByVal Num8Byte As Long)
    'turns a long number into a 8 byte string and adds it to the buffer
    Buffer = Buffer & MakeDWORD(Num8Byte)
End Sub
Public Function InsertBYTE(Data As Integer)
Buffer = Buffer & Chr(Data)
End Function
Public Sub InsertSTRING(ByVal NewData As String, Optional Terminator As String = "")
    'adds a string to the buffer with a terminator (optional)
    Buffer = Buffer & NewData & Terminator
End Sub
Public Sub InsertNTString(ByVal NewData As String)
    'adds a string to the buffer with a null byte at the end
    Buffer = Buffer & NewData & Chr(&H0)
End Sub
Public Sub InsertNT_0DString(ByVal NewData As String)
    'adds a string to the buffer with a null byte at the end
    Buffer = Buffer & NewData & Chr(&HD)
End Sub
Public Sub InsertHEADER(ByVal PacketID As Byte)
    'adds a packet header to start of the buffer string
    Buffer = Chr(&HFF) & Chr(PacketID) & MakeWORD(Len(Buffer) + 4) & Buffer
End Sub
Public Sub InsertMCPHEADER(ByVal PacketID As Byte)
    'adds a packet header to start of the buffer string
    Buffer = MakeWORD(Len(Buffer) + 3) & Chr(PacketID) & Buffer
End Sub
Public Sub InsertbnlsHEADER(ByVal PacketID As Byte)
    'adds a packet header to start of the buffer string
    Buffer = MakeWORD(Len(Buffer) + 3) & Chr(PacketID) & Buffer
End Sub
Public Sub Clear()
    'clears are packet buffer to nothing so we can make another one
    '(called when a packet is sent, see SendPacket function at bottom)
    On Error Resume Next
    Buffer = ""
End Sub


Public Function InsertNonNTString(Data As String)
    Buffer = Buffer & Data
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'these are the fain functions* that return converted variables
'so the insert* subs can add them to the buffer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function MakeWORD(ByVal Num As Integer) As String
    'takes a integer number and returns a 2 byte string
    MakeWORD = String(2, &H0)
    Call CopyMemory(ByVal MakeWORD, Num, 2)
End Function
Public Function MakeDWORD(ByVal Num As Long) As String
    'takes a long number and returns a 4 byte string
    MakeDWORD = String(4, &H0)
    Call CopyMemory(ByVal MakeDWORD, Num, 4)
End Function

Public Function StrToHex(ByVal Data As String, Optional Splitter As String = " ") As String
    ';this function turns data (string) into a hex byte string
    Dim aLen As Long: aLen = Len(Data)
    Dim eLen As Long: eLen = 2 + Len(Splitter)
    Dim iPos As Long
    StrToHex = String((aLen * eLen), "~")
    For i = 1 To (aLen * eLen) Step eLen
        iPos = iPos + 1
        Mid(StrToHex, i, eLen) = Right("00" & Hex(Asc(Mid(Data, iPos, 1))), 2) & Splitter
    Next i
    StrToHex = Trim(StrToHex)
End Function
Public Function HexToStr(ByVal Data As String) As String
    'this one takes hex byte string and turns it back to data
    Dim tmpData As String: tmpData = Replace(Data, " ", "") 'remove any spaces in the bytes
    HexToStr = String(Len(tmpData) / 2, 0) 'create smaler space to fit the converted characters in
    Dim iPos As Long
    For Z = 1 To Len(tmpData) Step 2
        iPos = iPos + 1
        Mid(HexToStr, iPos, 1) = Chr(Val("&H" & Mid(tmpData, Z, 2)))
    Next Z
End Function

Public Function HexToLong(ByVal Data As String) As String
    
        Dim tmpData As String: tmpData = Replace(Data, " ", "")  'remove any spaces in the bytes
 

      HexToLong = CLng("&H" & Mid(tmpData, 1, Len(tmpData)))

End Function

Public Function GetUserIp(HexIp As String)
    Dim CombinedParts As String
    Dim Part1, Part2, Part3, Part4 As String
   
    CombinedParts = Replace(HexIp, " ", "")
    
    Part1 = Mid$(CombinedParts, 1, 2)
    
    Part2 = Mid$(CombinedParts, 3, 2)
    
    Part3 = Mid$(CombinedParts, 5, 2)
    
    Part4 = Mid$(CombinedParts, 7, 2)
    
    
    

    GetUserIp = Trim$(CLng("&H" & Part1) & "." & CLng("&H" & Part2) & "." & CLng("&H" & Part3) & "." & CLng("&H" & Part4))
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'these 3 functions are used to get from strings (and not add to are buffer string)
'so we use these mostly for getting values from battle.net packets etc
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function GetSTRING(ByVal Data As String, Optional Terminator As String = vbNullChar) As String
    'gets a string from with in a string
    'by defalt, the string it will get will be as long as the 1st null byte it comes to.
    Dim tmpInt As Integer
    tmpInt = InStr(1, Data, Terminator)
    If tmpInt = 0 Then
        GetSTRING = Data
    Else
        GetSTRING = Mid(Data, 1, tmpInt - 1)
    End If
End Function
Public Function GetSTRING20(ByVal Data As String, Optional Terminator As String = vbNullChar) As String
    'gets a string from with in a string
    'by defalt, the string it will get will be as long as the 1st null byte it comes to.
    Dim tmpInt As Integer
    tmpInt = InStr(1, Data, Chr$(&H20))
    If tmpInt = 0 Then
        GetSTRING20 = Data
    Else
        GetSTRING20 = Mid(Data, 1, tmpInt - 1)
    End If
End Function
Public Function GetSTRING0D(ByVal Data As String, Optional Terminator As String = vbNullChar) As String
    'gets a string from with in a string
    'by defalt, the string it will get will be as long as the 1st null byte it comes to.
    Dim tmpInt As Integer
    tmpInt = InStr(1, Data, Chr$(&HD))
    If tmpInt = 0 Then
        GetSTRING0D = Data
    Else
        GetSTRING0D = Mid(Data, 1, tmpInt - 1)
    End If
End Function
Public Function GetSTRING2C(ByVal Data As String, Optional Terminator As String = vbNullChar) As String
    'gets a string from with in a string
    'by defalt, the string it will get will be as long as the 1st null byte it comes to.
    Dim tmpInt As Integer
    tmpInt = InStr(1, Data, Chr$(&H2C))
    If tmpInt = 0 Then
        GetSTRING2C = Data
    Else
        GetSTRING2C = Mid(Data, 1, tmpInt - 1)
    End If
End Function
Public Function GetWORD(ByVal Data As String) As Integer
    'takes 2 bytes of data (string) and converts it into its integer value
    If Data = "" Then Exit Function
    Call CopyMemory(GetWORD, ByVal Data, 2)
End Function
Public Function GetDWORD(ByVal Data As String) As Long
    'takes 4 bytes of data (string) and converts it into its long value
    If Data = "" Then Exit Function
    Call CopyMemory(GetDWORD, ByVal Data, 4)
End Function

Public Function GetWHOLEDWORD(ByVal Data As String, LenofGRAB As Long) As Long
    'takes 4 bytes of data (string) and converts it into its long value
    If Data = "" Then Exit Function
    Call CopyMemory(GetWHOLEDWORD, ByVal Data, LenofGRAB)
End Function


Public Function MakeLong(X As String) As Long
    If Len(X) < 4 Then
        Exit Function
    End If
    CopyMemory MakeLong, ByVal X, 4
End Function




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'and we need a sub to send the data down a winsock
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub SendPacket(SOCK As Winsock)

    'put are packet buffer into a tmp buffer
    Dim tmpBuffer As String
    tmpBuffer = Buffer
    'delete the buffer
    Call Clear
    If tmpBuffer = "" Then Exit Sub   'avoid sending nothing ""
    If Not SOCK.State = sckConnected Then Exit Sub  'only send data if socket is in connected state
    'Call SOCK.SendData(tmpBuffer)
       Call send(SOCK.SocketHandle, ByVal tmpBuffer, Len(tmpBuffer), 0&)
     DoEvents 'this should stop any winsock controll bugs :/
End Sub

'we can call this to check are packet out when testing
 

 
 Private Function gBF(xChar As Long) As String

 
  Select Case xChar
    Case &H20 To &H7E: gBF = Chr(xChar)
  Case Else: gBF = "."
  End Select
  End Function
Public Sub SendPacketUDP(SOCK As Winsock)
    'put are packet buffer into a tmp buffer
    Dim tmpBuffer As String
    tmpBuffer = Buffer
    'delete the buffer
    Call Clear
    If tmpBuffer = "" Then
    Call Chat.AddChat("Buffer is Empty!..Not sending last UDP Packet", vbRed)
    Exit Sub   'avoid sending nothing ""
    End If
    
    Call SOCK.SendData(tmpBuffer)
    'Chat.AddChat vbBlack, buf2.HexToStr(tmpBuffer)
    DoEvents 'this should stop any winsock controll bugs :/
End Sub
Public Sub CalcBUFFERlength(SOCK As Winsock)
    'put are packet buffer into a tmp buffer
    Dim tmpBuffer As String
    tmpBuffer = Buffer
    'delete the buffer
    Call Clear
    If tmpBuffer = "" Then
   ' Call Chat.AddChat(frmMain.TabStrip1.SelectedItem.Index, "Buffer is Empty!", vbRed)
    Exit Sub      'avoid sending nothing ""
    End If
    
    
modStuff2.LenPKT = Len(tmpBuffer)

    DoEvents 'this should stop any winsock controll bugs :/
End Sub
Public Sub GETUDPCheckSum(SOCK As Winsock)
    'put are packet buffer into a tmp buffer
    Dim tmpBuffer As String
    tmpBuffer = Buffer
    'delete the buffer
    Call Clear
    If tmpBuffer = "" Then
  '  Call Chat.AddChat(frmMain.TabStrip1.SelectedItem.Index, "Buffer is Empty!", vbRed)
    Exit Sub      'avoid sending nothing ""
    End If
    
'Mid(Buffer, 5, 2) = MakeWORD(UDPCheckSum(Mid(Buffer, 5)))

    
modStuff2.UDPCheckSummm = UDPCheckSum(Mid(tmpBuffer, 5))
 'MsgBox GameStuff.cheksum
 
    DoEvents 'this should stop any winsock controll bugs :/
End Sub
Private Function SubCheckSum(ByVal Buf As String, ByVal Length As Integer) As Long
    Dim sum1, sum2
    Dim i As Integer, iY As Integer
        For iY = 0 To Length - 1
            i = Length - iY
            sum2 = sum2 + Asc(Mid(Buf, i, 1))
            If sum2 > &HFF Then
                sum2 = sum2 - &HFF
            End If
            sum1 = sum1 + sum2
        Next iY
        SubCheckSum = (LShift((sum2 And &HFF), 8)) Or ((sum1 Mod &HFF) And &HFF)
End Function

Private Function UDPCheckSum(Buf As String) As Integer
    Dim subsum As Long, Length As Integer
    Dim A As Long, B As Long, Ret As Integer
        CopyMemory Length, ByVal Mid$(Buf, 3, 2), 2
        Length = Length - 2
        subsum = SubCheckSum(Mid$(Buf, 3), Length)
        A = &HFF - ((subsum And &HFF) + (RShift(subsum, 8))) Mod &HFF
        B = CLng((((&HFF - (A + RShift(subsum, 8)) Mod &HFF) And &HFF) Or LShift(A, 8)))
        CopyMemory Ret, B, 2
        UDPCheckSum = Ret
        
End Function

Private Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
On Error Resume Next
    RShift = CDbl(pnValue \ (2 ^ pnShift))
End Function

Private Function LShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
On Error Resume Next
    LShift = CDbl(pnValue * (2 ^ pnShift))
End Function
