VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRoD 
   BorderStyle     =   0  'None
   Caption         =   "Rogue of Diablo v2"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRoD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   4440
   End
   Begin VB.PictureBox TrayIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1440
      Picture         =   "frmRoD.frx":F172
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4440
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Left            =   3960
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckMCP 
      Index           =   0
      Left            =   3480
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckBnet 
      Index           =   0
      Left            =   3000
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkClassic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Classic"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox chkEx 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Expansion"
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox chkHard 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hardcore"
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox chkLad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ladder"
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton optAma 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amazon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optSorc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sorceress"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton optAss 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Assassin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.OptionButton optDru 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Druid"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton optBarb 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Barbarian"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.OptionButton optPala 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Paladin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton optNecro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Necromancer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2175
      Left            =   660
      TabIndex        =   12
      Top             =   1050
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmRoD.frx":1E2E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   10440
      Picture         =   "frmRoD.frx":1E360
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   10080
      Picture         =   "frmRoD.frx":1F63A
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   10440
      Picture         =   "frmRoD.frx":20914
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   10080
      Picture         =   "frmRoD.frx":21BEE
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   9840
      TabIndex        =   27
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   7560
      TabIndex        =   26
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   4600
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rogue of Diablo v2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   165
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   480
      TabIndex        =   24
      Top             =   495
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      Height          =   255
      Left            =   9720
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9600
      TabIndex        =   22
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      Height          =   255
      Left            =   8760
      TabIndex        =   21
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   8640
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   4200
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1680
      X2              =   6720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   600
      X2              =   4560
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   600
      X2              =   1320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line7 
      X1              =   600
      X2              =   600
      Y1              =   960
      Y2              =   3360
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speed: 0 p/s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Line Line6 
      X1              =   5760
      X2              =   6720
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line5 
      X1              =   6720
      X2              =   6720
      Y1              =   960
      Y2              =   3360
   End
   Begin VB.Label lblClass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Class: Amazon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type: Classic"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label lblFlags 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   0
      Top             =   4560
      Width           =   615
   End
End
Attribute VB_Name = "frmRoD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Dim u As Integer

 Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


      'types
      Private Const flgClassic = &H0
      Private Const flgHardcore = &H4
      Private Const flgExpansion = &H20
      Private Const flgLadder = &H40
      'classes
      Private Const flgAmazon = &H0
      Private Const flgSorc = &H1
      Private Const flgNecro = &H2
      Private Const flgPala = &H3
      Private Const flgBarb = &H4
      Private Const flgDruid = &H5
      Private Const flgAss = &H6
      
 
      
      Private Const GWL_STYLE = (-16)
      Private Const GWL_EXSTYLE = (-20)
      Private Const WS_EX_LAYERED = &H80000
      Private Const LWA_COLORKEY = &H1
        Const WS_EX_TRANSPARENT = &H20&
        
 '       [01:04] <@lord2800> classic, hardcore, ladder, expansion, expansion hardcore, expansion ladder, classic hardcore ladder, expansion hardcore ladder





Private Sub chkClassic_Click()
If chkClassic = 1 Then
chkEx = 0
flags = flags + Hex(flgClassic)
'lblFlags.Caption = Flags
End If
If chkClassic = 0 Then
flags = flags - Hex(flgClassic)
'lblFlags.Caption = Flags
End If
CheckFlags (flags)

End Sub

Private Sub chkEx_Click()
If chkEx = 1 Then
chkClassic = 0
flags = flags + Hex(flgExpansion)
'lblFlags.Caption = Flags
End If
If chkEx = 0 Then
flags = flags - Hex(flgExpansion)
'lblFlags.Caption = Flags
End If
CheckFlags (flags)
End Sub

Private Sub chkHard_Click()
If chkHard = 1 Then
flags = flags + Hex(flgHardcore)
'lblFlags.Caption = Flags
End If
If chkHard = 0 Then
flags = flags - Hex(flgHardcore)
'lblFlags.Caption = Flags
End If
CheckFlags (flags)
End Sub

Private Sub chkLad_Click()
If chkLad = 1 Then
flags = flags + Hex(flgLadder)
'lblFlags.Caption = Flags
End If
If chkLad = 0 Then
flags = flags - Hex(flgLadder)
'lblFlags.Caption = Flags
End If
CheckFlags (flags)
End Sub







Private Sub Command2_Click()
MsgBox CharList(0).Item(Str(CharList(0).HashCode("Char0")))
End Sub



Private Sub Command1_Click()
'  MsgBox CharList(5).Item(Str(CharList(5).HashCode("Char0")))
'MsgBox ClassFlag
'MsgBox flags
 'MsgBox Uper(0)
'LoadData ("Accounts.txt")
'MsgBox Replace(BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), "//") - 1), Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), "//") - 1) & "RoD")
'MsgBox Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS0"))), "//") - 1) & "RoD"
'CharDuring0x02CreatedCounter(0) = 5
'CharDuring0x02CreatedCounter(1) = 2
'MsgBox CharDuring0x02CreatedCounter(0)
 MsgBox D2DVList.Count
 
         
'Call SetTimer(Me.hwnd, 1000, 5000, AddressOf TimerProc)
'MsgBox buf2.HexToStr("31 31")
        End Sub

Private Sub Command3_Click()
LoadData ("Accounts2.txt")

End Sub

Private Sub Command4_Click()
LoadData ("Realm_Accounts.txt")

End Sub

'[01:04] <@lord2800> basically there's two basic types: classic or expansion, then there's two basic modes: hardcore and softcore, then there's two basic realms: nonladder and ladder
Private Sub Form_Load()
'SetWindowLong rtb.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
Me.BackColor = vbCyan 'Set the backcolor of the tobe transparent form to w/e color

          SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

'now everything is transparent so you're gonna have to use this to only make what is a certain color transparent

          SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY


AddChat "Rogue of Diablo Version 2 [Build 1]", &H408080
AddChat "Created by Myst", &H408080
AddChat "www.DarkBlizz.org", &H408080
AddChat "Op W@R / Op FallenArms [USEast]", &HCD69DD
AddChat "BETA EDITION - Still Under Development!!", &HFF4EE5
BNCS_Stuff.Server = GetSetting(App.ProductName, "Config", "Server")
BNCS_Stuff.Verbyte = GetSetting(App.ProductName, "Config", "Verbyte")
Spawn = GetSetting(App.ProductName, "Config", "Spawn")


'LoadData ("Realm_Accounts.txt")
LoadBNCS ("BNCS_Accounts.txt")
LoadD2Keys ("D2DV_Keys.txt")
LoadD2XPKeys ("D2XP_Keys.txt")
If Spawn = "" Then
Spawn = BNCSList.Count ''Spawn amount of bots per bncs name
End If


AddChat "Spawning: " & Spawn, &H123321

  
'AddChat "[Server: " & BNCS_Stuff.Server & "]", &H548956
 
 Call SanityCheck
 
 

 End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub



Private Sub Image5_Click()
NoSysIcon INTRAY
End Sub

Private Sub Image6_Click()
''''''''''''''
''''''''''''''
''''''''''''''
''''''''''''''
 If MsgBox("Are you sure you want to quit?", vbOKCancel) = vbOK Then
MsgBox "Thanks for being part of the Rogue of Diablo v2 Beta" & vbCrLf & "Visit www.DarkBlizz.org"
Unload Me
KillProcess.KillProcess ("RogueOfDiablo.exe")

Else
'nothing
End If


End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image5.Visible = True
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = False
Image6.Visible = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label7.ForeColor = vbBlack
 
End Sub

Private Sub Label10_Click()
frmSettings.Show
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbBlue

End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlack
Label7.ForeColor = vbBlack
Label10.ForeColor = vbBlack

End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image5.Visible = False

Image8.Visible = True
Image6.Visible = False
End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbBlack

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlack
End Sub

Private Sub Label6_Click()
'start
 Dim Norealm As Boolean

Dim i As Integer

For i = 0 To Spawn - 1

If CharList(i).Count = 0 Then
AddChat Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & i))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & i))), "//") - 1) & " doesn't have any characters to make.", vbRed
Norealm = True
End If
 
 Next i

If Norealm = False Then
ConnectBNCS sckX
Else
AddChat "Load your character lists through the settings and then click Apply.", vbBlue
End If

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
On Error Resume Next
AddChat "Disconnected All Connections", vbRed
Dim X As Integer
Do Until X = Spawn
sckBnet(X).Close
sckMCP(X).Close
X = X + 1
DoEvents
Loop

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlue
End Sub

Private Sub lblFlags_Change()
'MsgBox lblFlags.Caption
Select Case lblFlags.Caption

Case "0"
lblType.Caption = "Type: Classic"

Case "40"
lblType.Caption = "Type: Classic Ladder"

Case "4"
lblType.Caption = "Type: Classic Hardcore"

Case "44"
lblType.Caption = "Type: Classic Hardcore Ladder"

Case "20"
lblType.Caption = "Type: Expansion"

Case "60"
lblType.Caption = "Type: Expansion Ladder"

Case "24"
lblType.Caption = "Type: Expansion Hardcore"

Case "64"
lblType.Caption = "Type: Expansion Hardcore Ladder"

End Select


End Sub

Private Sub optAma_Click()
lblClass.Caption = "Class: Amazon"
ClassFlag = Hex(flgAmazon)

End Sub

Private Sub optAss_Click()
lblClass.Caption = "Class: Assassin"
ClassFlag = Hex(flgAss)

End Sub

Private Sub optBarb_Click()
lblClass.Caption = "Class: Barbarian"
ClassFlag = Hex(flgBarb)

End Sub

Private Sub optDru_Click()
lblClass.Caption = "Class: Druid"
ClassFlag = Hex(flgDruid)
 
End Sub

Private Sub optNecro_Click()
lblClass.Caption = "Class: Necromancer"
ClassFlag = Hex(flgNecro)

End Sub

Private Sub optPala_Click()
lblClass.Caption = "Class: Paladin"
ClassFlag = Hex(flgPala)

End Sub

Private Sub optSorc_Click()
lblClass.Caption = "Class: Sorceress"
ClassFlag = Hex(flgSorc)

End Sub

Private Sub sckBnet_Close(Index As Integer)
AddChat "BNCS(" & Index & ") sckbnet Closed", vbRed
End Sub

Private Sub sckBnet_Connect(Index As Integer)
AddChat "BNCS(" & Index & ") Connected", vbRed
buf2.InsertBYTE &H1
buf2.SendPacket sckBnet(Index)

modPkts.Send0x50 (Index)
InProcessOfConnecting = True
End Sub

Private Sub sckBnet_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Static RecvBuffer As String 'we hold the data sent to us from bnet in this string
    Dim PacketLengh As Integer
    Dim Data As String
    Call sckBnet(Index).GetData(Data, vbString)
    RecvBuffer = RecvBuffer & Data 'add the recv data to are buffer string from bnet
    'now lets split each packet up and parse it baced on the 2 bytes lengh in the header
    
    
    While Len(RecvBuffer) >= 4
        PacketLengh = buf2.GetWORD(Mid(RecvBuffer, 3, 2))
        If Len(RecvBuffer) < PacketLengh Then Exit Sub 'we have to wait for the missing data to come before we handle it (packet may have got split up abit)
        'cut out the packet from the buffer and put it in are tmp data string
        Data = Left(RecvBuffer, PacketLengh)
        RecvBuffer = Mid(RecvBuffer, 1 + PacketLengh) 'skip the packet we just cut out
        Call ParseBnetData(Index, Data)   'see a few subs below where we hand the packet to the parse sub
    Wend
    
    
End Sub
Private Sub ParseBnetData(Index As Integer, ByVal Data As String)
    Dim PacketID As Byte
     'Dim tmpfilename As String
      '          Dim tmpformula As String
    PacketID = Asc(Mid(Data, 2, 1))
  
    
    Select Case PacketID
        Case &H25 'bnet sent us a ping packet!
            'FF 25 08 00 D7 EB DC 03
            'add the only data in the ping packet into are own so we can ecco it back
          
            buf2.InsertDWORD buf2.GetDWORD(Mid(Data, 5, 4))
            buf2.InsertHEADER &H25
            buf2.SendPacket sckBnet(Index)
 



        Case &H50 'Bnet sent us a responce from are &H50 packet
        
                      Call AddChat("BNCS(" & Index & ") Recieved 0x50", vbRed)

             If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then 'passed challenge
                ' MsgBox "Begining 0x50 Parse"
                'now to parse out the important infomation from the packet
                Dim tmpFileTime As String * 8

                'this is used later on, so we must remember it, thats why the "ServerToken" String is declared publicly in the modStuff
                
                modPkts.ServerToken = buf2.GetDWORD(Mid(Data, 9, 4)) ' 09 9F 06 73
                
                tmpFileName = buf2.GetSTRING(Mid(Data, 25))  'will look for a null byte as from 25th byte on and cut out the string when it finds the null byte
                
                tmpFileTime = Mid$(Data, 17, 8)

                tmpFormula = buf2.GetSTRING(Mid(Data, 26 + Len(tmpFileName))) 'get the string behind the filename string (useing the lengh of the filename)
         
         'Call modPkts.BNLS_0x1A(tmpFileTime)
      Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
            Else 'failed challenge (unless this is w3..)
          MsgBox "Failed 0x50 Parse"
                sckBnet(Index).Close
            
            End If
        
        
        
        
        
        
        Case &H51  'CdKey
Chat.AddChat "BNCS(" & Index & ") Recieved 0x51 (" & Hex(buf2.GetDWORD(Mid$(Data, 5, 4))) & ")", vbRed


        If buf2.GetDWORD(Mid(Data, 5, 2)) = &H0 Then    'Working
        

        Call Chat.AddChat("BNCS(" & Index & ") Cd-Key is Good", vbRed)

        Call modPkts.Send0x3A(Index)
        End If
        
        If buf2.GetDWORD(Mid(Data, 5, 2)) = &H200 Then 'Invalid
        Call Chat.AddChat("BNCS(" & Index & ") Cd-Key is Invalid, you are now Ip-Banned from this server.", vbRed)

                Exit Sub

        End If
 
        If buf2.GetDWORD(Mid(Data, 5, 2)) = &H202 Then 'Banned
        Call Chat.AddChat("BNCS(" & Index & ") Cd-Key is Banned!!", vbRed)
Keycnt = Keycnt + 1
If Keycnt < D2DVList.Count Then
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Else
AddChat "You don't have anymore CDKeys", vbRed
End If

                 Exit Sub

         End If
   If buf2.GetDWORD(Mid(Data, 5, 2)) = &H203 Then 'Wrong Product
   Call Chat.AddChat("BNCS(" & Index & ") Wrong Products NumbNuts", vbRed)
Keycnt = Keycnt + 1
If Keycnt < D2DVList.Count Then
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Else
AddChat "You don't have anymore CDKeys", vbRed
End If

                 Exit Sub
End If
    If buf2.GetDWORD(Mid(Data, 5, 2)) = &H201 Then  'CdKey InUse
     
       Dim KeyinUsePerson As String
   KeyinUsePerson = buf2.GetSTRING(Mid$(Data, 9))
      Chat.AddChat "BNCS(" & Index & ") CdKey is In-Use by " & KeyinUsePerson, vbRed
      
      Keycnt = Keycnt + 1
If Keycnt < D2DVList.Count Then
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Else
AddChat "You don't have anymore CDKeys", vbRed
End If

      
                 End If
          If buf2.GetDWORD(Mid(Data, 5, 2)) = &H101 Then 'Invalid Version
          
        Chat.AddChat "BNCS(" & Index & ") Invalid Version!!", vbRed
                 End If
             If buf2.GetDWORD(Mid(Data, 5, 2)) = &H100 Then  'Old Game Version
     Chat.AddChat "BNCS(" & Index & ") Old Game Version", vbRed
                 End If

If buf2.GetDWORD(Mid(Data, 5, 2)) = &H102 Then 'Game Version must be downgraded
      Chat.AddChat "BNCS(" & Index & ") Game Version must be downgraded", vbRed
                 End If

                 
              If buf2.GetDWORD(Mid(Data, 5, 2)) = &H210 Then 'EXp key invalid
      Chat.AddChat "BNCS(" & Index & ") Expansion key is invalid", vbRed
                 End If
                 If buf2.GetDWORD(Mid(Data, 5, 2)) = &H211 Then 'exp key inuse
        Chat.AddChat "BNCS(" & Index & ") Expansion key in use", vbRed
        EXPKeycnt = EXPKeycnt + 1
If Keycnt < D2XPList.Count Then
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Else
AddChat "You don't have anymore Expansion CDKeys", vbRed
End If

                 End If
                 
                          If buf2.GetDWORD(Mid(Data, 5, 2)) = &H212 Then 'exp key banned
     Chat.AddChat "BNCS(" & Index & ") Expansion key banned", vbRed
             EXPKeycnt = EXPKeycnt + 1
If Keycnt < D2XPList.Count Then
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Else
AddChat "You don't have anymore Expansion CDKeys", vbRed
End If
                 End If
                 
                          If buf2.GetDWORD(Mid(Data, 5, 2)) = &H213 Then 'exp key another game
     Chat.AddChat "BNCS(" & Index & ") Expansion key for another game", vbRed
             EXPKeycnt = EXPKeycnt + 1
If Keycnt < D2XPList.Count Then
Call modPkts.Send0x51(Index, tmpFormula, tmpFileName)
Else
AddChat "You don't have anymore Expansion CDKeys", vbRed
End If
                 End If
                 

         
    
Case &H3D 'Creating Account
'Call Chat.ShowChat(i, vbGreen, "Recieved 0x3D")
If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then 'Account Created
Chat.AddChat "BNCS(" & Index & ") Account was Created.", vbRed
Call modPkts.Send0x3A(Index)
Else

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H2 Then 'invalid
Chat.AddChat "BNCS(" & Index & ") Account has invalid characters.", vbRed
Exit Sub
Else


If buf2.GetDWORD(Mid(Data, 5, 4)) = &H3 Then 'Name contained bad stuff
Chat.AddChat "BNCS(" & Index & ") Account contains bad words.", vbRed
Exit Sub
Else

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H4 Then 'Account Exists
Chat.AddChat "BNCS(" & Index & ") Account already exists.", vbRed
Exit Sub

Else


If buf2.GetDWORD(Mid(Data, 5, 4)) = &H6 Then 'Account doesnt have enough alphanumeric
Chat.AddChat "BNCS(" & Index & ") Account doesn't have enough charecters.", vbRed
Exit Sub

End If
End If
End If
End If
End If
    


Case &H3A ' account login
'Chat.AddChat "Received 0x3a", vbGreen

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then 'Success
Call Chat.AddChat("BNCS(" & Index & ") Password is Good", vbRed)

Call modPkts.Send0x40(Index)

End If


If buf2.GetDWORD(Mid(Data, 5, 4)) = &H1 Then 'does not exist
Call Chat.AddChat("BNCS(" & Index & ") Account does not Exist", vbRed)
Call modPkts.Send0x3D(Index)
End If

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H2 Then 'PW Bad

Call Chat.AddChat("BNCS(" & Index & ") Password is Wrong!", vbRed)

   

End If

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H6 Then 'account closure
 Dim aclose As String
   aclose = buf2.GetSTRING(Mid$(Data, 9))
 Chat.AddChat "BNCS(" & Index & ") " & aclose, vbRed
Exit Sub
End If


Case &H40 'Realm Querying
Chat.AddChat "BNCS(" & Index & ") Recieved 0x40", vbRed
Chat.AddChat buf2.GetSTRING(Mid$(Data, 17)), vbRed
RealmTitle = buf2.GetSTRING(Mid$(Data, 17))

Call modPkts.Send0x3E(Index)

Case &H3E 'Logon Realm
Chat.AddChat "BNCS(" & Index & ") Recieved 0x3E", vbRed

MCPcookie = buf2.GetDWORD(Mid$(Data, 5, 4))
MCPStatus = buf2.GetDWORD(Mid$(Data, 9, 4))
MCPChunk1 = Mid$(Data, 13, 8)
MCPChunk2 = Mid$(Data, 29, 48)

  
    
IP = buf2.GetUserIp(buf2.StrToHex(Mid$(Data, 21, 4)))
Port = buf2.HexToLong(buf2.StrToHex(buf2.GetSTRING(Mid$(Data, 25, 4))))
    
   
Timer1.Enabled = True
ConnectMCP (Index)
'sckBnet(Index).Close
InProcessOfConnecting = False
End Select


End Sub

Private Sub sckMCP_Close(Index As Integer)
Chat.AddChat "MCP(" & Index & ") Disconnected!", vbRed

End Sub

Private Sub Timer1_Timer()

Label8.Caption = "Tested: " & CharNamesTested


End Sub
Private Sub sckMCP_Connect(Index As Integer)
Chat.AddChat "MCP(" & Index & ") Connected to Realm", vbRed

 buf2.InsertBYTE &H1
 buf2.SendPacket frmRoD.sckMCP(Index)
 
Call modPkts.Send0x01MCP(Index)


End Sub

Private Sub sckMCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  Dim Data As String
   Call sckMCP(Index).GetData(Data, vbString)

    Dim PacketID As Byte
     'Dim tmpfilename As String
      '          Dim tmpformula As String
    PacketID = Asc(Mid(Data, 3, 1))
    Select Case PacketID
    ''0x3A, 0x40, 0x3E, 0x01, 0x19, 0x07, abc
'[2:15:29 PM] - : Pro_Tech : -  once u send 0x19 and get the character list, thats when u would  create/delete charcter, ifu dont wanna logon one just yet

    Case &H1
    
    Chat.AddChat "MCP(" & Index & ") Recieved 0x01", vbRed
        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &H0 Then 'success
        Chat.AddChat "MCP(" & Index & ") Successfully Logged onto Realm", vbRed
        
        Select Case CharFullRC
        
        Case False
        If Index <> Spawn - 1 Then
        Load sckMCP(Index + 1)
        Load sckBnet(Index + 1)
        pause (5)
        Call modFunc.ConnectBNCS(Index + 1)
        'Call modFunc.ConnectMCP(Index + 1)
        Else
        AllBotsConnected = True
        Chat.AddChat "All bots spawned.", vbBlue
        End If
        
        Case True
        AddChat "[Pausing 5 seconds.]", vbGreen
        pause (5)
                If Index = Spawn - 1 Then
AllBotsConnected = True
End If
       ' modFunc.ConnectBNCS Index
        CharFullRC = False
        End Select
        
        'Call modStuff2.Send0x19MCP(Index)
           Call modPkts.Send0x19MCP(Index)
         
         
        End If
        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &HC Then  'No bnet connect
        Chat.AddChat "MCP(" & Index & ") No Battle.net Connection Detected.", vbRed
        End If
        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &H7E Then 'keybanned
        Chat.AddChat "MCP(" & Index & ") CdKey is banned from Realm", vbGreen
        End If
        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &H7F Then 'ipban
        Chat.AddChat "MCP(" & Index & ") You are Ip-Banned from the server. Try again later.", vbRed
        End If
        
        
          Case &H2
      'Chat.AddChat "Recieved 0x02", vbGreen
CharNamesTested = CharNamesTested + 1

        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &H0 Then 'success
        Chat.AddChat "MCP(" & Index & ") Successfully Created Character! - " & CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index) - 1))), vbBlack
                WriteCreated Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1) & " => " & (CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index) - 1))))
CharDuring0x02CreatedCounter(Index) = CharDuring0x02CreatedCounter(Index) + 1
If CharDuring0x02CreatedCounter(Index) >= "8" Then  'Char full
modPkts.Send0x19MCP Index
Else
        modPkts.Send0x02MCP (Index)
        End If
     'Call modPkts.Send0x07MCP
        End If
        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &H14 Then 'char exists
         Debug.Print "MCP(" & Index & ") Character already exists, or you reached max char limit of 8."
         modPkts.Send0x02MCP (Index)
         
        End If
        If buf2.GetDWORD(Mid$(Data, 4, 4)) = &H15 Then 'invalid
        Debug.Print "MCP(" & Index & ") Invalid Char name: " & CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index) - 1)))
        Chat.AddChat "MCP(" & Index & ") Invalid Char name: " & CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index) - 1))), vbRed
        WriteInvalid (CharList(Index).Item(Str(CharList(Index).HashCode("Char" & Uper(Index) - 1))))
        
        CharList(Index).Remove Str(CharList(Index).HashCode("Char" & Uper(Index) - 1))
        modPkts.Send0x02MCP (Index)
        End If
       
       
 
  Case &H19
  
            Chat.AddChat "MCP(" & Index & ") Recieved 0x19", vbGreen


Dim AmtChar As Integer

AmtChar = buf2.GetWORD(Mid$(Data, 6))
CharDuring0x02CreatedCounter(Index) = AmtChar
If AmtChar = "8" Then
AddChat "MCP(" & Index & ") Character is Full, Switching Account.", vbBlue
 Dim switchedAct As String
 switchedAct = Replace(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1), Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1) & "RoD")
 BNCSList.Remove Str(BNCSList.HashCode("BNCS" & Index))
 BNCSList.Add Str(BNCSList.HashCode("BNCS" & Index)), switchedAct
 AddChat "Switched to: " & Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1), vbBlue
If AutoSwapFull = False Then
AddChat "AutoSwap on Account Full isn't enabled.", vbRed

End If

Debug.Print "CHARACTER FULL! " & BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & Index))), "//") - 1
CharFullRC = True
sckBnet(Index).Close
sckMCP(Index).Close
If AllBotsConnected <> True Then
     AddChat "Waiting 5 secs, not all bots have connected.", vbRed
    Call SetTimer(Me.hwnd, 1000, 5000, AddressOf TimerProc)
'Timer2.Enabled = True
    Else
'Timer2.Enabled = False
If AutoSwapFull = True Then
    ConnectBNCS Index
End If
    End If
Else
         Call modPkts.Send0x02MCP(Index)

End If

'Dim AmtChar As Integer

'AmtChar = buf2.GetWORD(Mid$(Data, 6))
'frmD2.Show
'Call frmD2.D2CharScreen(AmtChar)
'frmD2.d2RTB.SelColor = vbRed
'frmD2.d2RTB.SelText = Mid$(Data, 16)

 



'0030  ff aa 46 9c 00 00 67 00  19 08 00 02 00 00 00 02   ..F...g. ........
'0040  00 c5 0a 56 4b 54 68 69  65 46 70 72 6f 00 84 80   ...VKThi eFpro...
'0050  ff ff ff ff ff ff ff ff  ff ff ff 01 ff ff ff ff   ........ ........
'0060  ff ff ff ff ff ff ff 01  81 80 80 80 ff ff ff 00   ........ ........
'0070  e0 0a 56 4b 44 69 6d 65  4f 70 00 84 80 ff ff ff   ..VKDime Op......
'0080  ff ff ff ff ff ff ff ff  01 ff ff ff ff ff ff ff   ........ ........
'0090  ff ff ff ff 01 81 80 80  80 ff ff ff 00            ........ .....

        'Call modStuff2.Send0x07MCP(Index)
        
    End Select

End Sub

Private Sub Timer2_Timer()
If InProcessOfConnecting = True Then
'wait
Else
ConnectBNCS 0
InProcessOfConnecting = False
End If

End Sub

Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'###########################################################
' This Function Tells what Event happened to icon in
' system tray
'###########################################################

    Dim MSG As Long
    MSG = (X And &HFF) * &H100

    Select Case MSG
        Case 0 'mouse moves
        
        Case &HF00  'left mouse button down
        
        Case &H1E00 'left mouse button up
        
        Case &H3C00  'right mouse button down
        'PopupMenu mnuOP, 2, , , mnuCTRL 'show the popoup menu
        
        Case &H2D00 'left mouse button double click
        NoSysIcon True    'Show App on double clicking Mouse's Left Button
        
        Case &H4B00 'right mouse button up
        
        Case &H5A00 'right mouse button double click
        
    End Select
   
End Sub
 
'###########################################################
' This Function Show TrayIcon.Picture in System Tray
' Don't Bother what it says
' To change TrayIcon's ToolTip goto 2nd Last Line
'###########################################################
Public Function ShowProgramInTray()
INTRAY = True   'Means App is now in Tray
    
    NI.cbSize = Len(NI) 'set the length of this structure
    NI.hwnd = TrayIcon.hwnd 'control to receive messages from
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'operation flags
    NI.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
    NI.hIcon = TrayIcon.Picture  'the location of the icon to display
  
' Change System Tray Icon's Tool Tip Here bt don't delete chr$(0) [its line carriage here]
    
    NI.szTip = "Rogue of Diablo v2 - Created by Myst" + Chr$(0) 'LoadResString(Language) + Chr$(0)  'the tool tip to display"
    result = Shell_NotifyIconA(NIM_ADD, NI)    'add the icon to the system tray
End Function


'###########################################################
' This Function Delete TrayIcon.Picture from System Tray
' Don't Bother what it says
'###########################################################
Private Sub DeleteIcon(pic As Control)
INTRAY = False  'Means app is unloaded or Max mode
    
    ' On remove, we only have to give enough information for Windows
    ' to locate the icon, then tell the system to delete it.
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.cbSize = Len(NI)
    NI.hwnd = pic.hwnd
    NI.uCallbackMessage = WM_MOUSEMOVE
    result = Shell_NotifyIconA(NIM_DELETE, NI)
End Sub

'###########################################################
' This Function controls 3 funtions
' 1] Visibility of App Form
' 2] Visibility of Tray Icon
' 3] Menu Caption Control
'###########################################################

Public Function NoSysIcon(maxIcon As Boolean)
    Select Case maxIcon
    Case False   'Case App in Min Mode
        Me.Visible = False
        ShowProgramInTray               'Now show TrayIcon PictureBox's Picture in SysTray as icon
        'mnuCTRL.Caption = "E&xpand Application"
    
    Case Else   'Case App in Max Mode
        Me.Visible = True
        DeleteIcon TrayIcon
        'mnuCTRL.Caption = "Minimize App to System Tray"
    
    End Select

End Function

