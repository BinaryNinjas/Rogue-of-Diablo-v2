VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton cmdSet 
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Set"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   128
      FCOLO           =   0
      EMBOSSM         =   16744576
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmSettings.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.CheckBox chkSwapBNCS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Swap BNCS Login Names"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CheckBox chkAltKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alternate CdKeys"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2055
      Left            =   2400
      TabIndex        =   12
      Top             =   4560
      Width           =   3975
      Begin VB.CheckBox chkActFull 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Swap on Account Full"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2400
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstActList 
      Height          =   2895
      Left            =   6600
      TabIndex        =   9
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BNCS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account List"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.TextBox txtSpawn 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Text            =   "5"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtVerbyte 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Text            =   "0D"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtserver 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Left            =   3360
      TabIndex        =   2
      Text            =   "europe.battle.net"
      Top             =   2400
      Width           =   2415
   End
   Begin LVbuttons.LaVolpeButton cmdBrowse 
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Browse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   128
      FCOLO           =   0
      EMBOSSM         =   16744576
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmSettings.frx":001C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Command1 
      Height          =   495
      Left            =   7560
      TabIndex        =   18
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Apply"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   128
      FCOLO           =   0
      EMBOSSM         =   16744576
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmSettings.frx":0038
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Command2 
      Height          =   495
      Left            =   9000
      TabIndex        =   19
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   128
      FCOLO           =   0
      EMBOSSM         =   16744576
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmSettings.frx":0054
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Image Image4 
      Height          =   1605
      Left            =   2400
      Picture         =   "frmSettings.frx":0070
      Top             =   6585
      Width           =   2145
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   7680
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   7680
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label lblBrow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Realm_Accounts.txt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label lblSet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BNCSselcted"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   5390
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Account Lists"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   435
      Left            =   5880
      Picture         =   "frmSettings.frx":0FFD
      Top             =   3795
      Width           =   450
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   5880
      Picture         =   "frmSettings.frx":155F
      Top             =   3090
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   5880
      Picture         =   "frmSettings.frx":1AC1
      ToolTipText     =   "Input the Battle.net server you want to connect to."
      Top             =   2340
      Width           =   450
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Verbyte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3165
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   1400
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2390
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Index           =   0
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   8895
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
 
 Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
 Dim ActListCnt As Integer
      
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

Private Sub cmdBrowse_Click()
Dim WhatFile As String
cd.ShowOpen
WhatFile = cd.FileTitle

lblBrow.Caption = WhatFile
End Sub

Private Sub cmdSet_Click()
lstActList.SelectedItem.SubItems(1) = lblBrow.Caption
 
End Sub
Private Function ValidationCheck() As Boolean



End Function
Private Sub Command1_Click()

'////Validation Check
If ValidationCheck = False Then
'Exit Sub
Else
End If


'////
SaveSetting App.ProductName, "Config", "Server", txtserver.Text
SaveSetting App.ProductName, "Config", "Verbyte", txtVerbyte.Text
SaveSetting App.ProductName, "Config", "Spawn", txtSpawn.Text
BNCS_Stuff.Server = txtserver.Text
BNCS_Stuff.Verbyte = txtVerbyte.Text
Spawn = txtSpawn.Text

For ActListCnt = 1 To lstActList.ListItems.Count
SaveSetting App.ProductName, "Config", "ActList" & ActListCnt, lstActList.ListItems.Item(ActListCnt).SubItems(1)
Next ActListCnt


'////Load AccountLists for each account
Dim i As Integer
Dim O As Integer
ListCnt = 0

For i = 1 To lstActList.ListItems.Count
    If CharList(O).Count >= 1 Then
    CharList(O).RemoveAll
    End If
    
LoadData lstActList.ListItems.Item(i).SubItems(1)
O = O + 1
DoEvents
Next i

If chkActFull = 1 Then
AutoSwapFull = True
End If

If Spawn > BNCSList.Count Then
MsgBox "You can't spawn more bots than you have loaded. Lower the amount of spawns or close bot and add more BNCS names and then restart."
Else
Unload Me
End If
'////



End Sub

Private Sub Command2_Click()
Unload Me

End Sub

 
 

 

Private Sub Command3_Click()
MsgBox CharList(6).Item(Str(CharList(6).HashCode("Char0")))

End Sub

Private Sub Form_Load()
'SetWindowLong rtb.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
frmSettings.Show
Me.BackColor = vbCyan 'Set the backcolor of the tobe transparent form to w/e color

          SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

'now everything is transparent so you're gonna have to use this to only make what is a certain color transparent

          SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
          
          
          txtserver.Text = GetSetting(App.ProductName, "Config", "Server")
          txtVerbyte.Text = GetSetting(App.ProductName, "Config", "Verbyte")
          txtSpawn.Text = GetSetting(App.ProductName, "Config", "Spawn")
          'If txtserver.Text = vbNull Then
          
        'If txtserver.Text = "" Then
        'txtserver.BackColor = &HFF00FF
        'End If
        'If txtVerbyte.Text = "" Then
        'txtVerbyte.BackColor = &HFF00FF
        'End If
        
        'If txtSpawn.Text = "" Then
        'txtSpawn.BackColor = &HFF00FF
        'End If
        
        'lstActList.Width = lstActList.ColumnHeaders.Item(2).Left
        
        Dim u As Integer
        Do Until u = BNCSList.Count
        lstActList.ListItems.Add , , Mid$(BNCSList.Item(Str(BNCSList.HashCode("BNCS" & u))), 1, InStr(1, BNCSList.Item(Str(BNCSList.HashCode("BNCS" & u))), "//") - 1)
        lstActList.ListItems.Item(u + 1).SubItems(1) = GetSetting(App.ProductName, "Config", "ActList" & u + 1)
        
      u = u + 1
      DoEvents
      Loop
      
       
        
        
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub

Private Sub Image1_Click()
MsgBox "Input the battle.net server you wish to connect too."
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = False
Image6.Visible = True
End Sub
Private Sub Image6_Click()
''''''''''''''
''''''''''''''
''''''''''''''
''''''''''''''

Unload Me

End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image8.Visible = True
Image6.Visible = False
End Sub

Private Sub Label8_Click()

End Sub

Private Sub Image2_Click()
MsgBox "Input the current Version Byte of D2." & vbCrLf & "As of June 2010, it is 0D"

End Sub

Private Sub Image3_Click()
MsgBox "Input how many bots you want to use, the amount of bots can not exceed the amount of BNCS names you have loaded." & vbCrLf & "You have " & BNCSList.Count & " BNCS names loaded."

End Sub

Private Sub lstActList_ItemClick(ByVal Item As MSComctlLib.ListItem)
  lblSet.Caption = lstActList.SelectedItem
End Sub

Private Sub txtserver_Change()
If txtserver.Text = Null Then
txtserver.BackColor = vbRed
End If
End Sub

Private Sub txtSpawn_Change()
txtSpawn.BackColor = vbWhite
End Sub
