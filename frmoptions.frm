VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmoptions 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Options"
   ClientHeight    =   3480
   ClientLeft      =   4920
   ClientTop       =   4515
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1965
      Left            =   4440
      Picture         =   "frmoptions.frx":0000
      ScaleHeight     =   1905
      ScaleWidth      =   3750
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Help"
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Index           =   2
      Left            =   2240
      TabIndex        =   27
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Traffic"
      Height          =   375
      Index           =   1
      Left            =   1120
      TabIndex        =   26
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Settings"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   6780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O&k"
      Height          =   330
      Left            =   3240
      TabIndex        =   5
      Top             =   3000
      Width           =   1035
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Frame Frame4 
         Caption         =   "Uploads total:"
         Height          =   1695
         Left            =   2275
         TabIndex        =   17
         Top             =   360
         Width           =   2175
         Begin VB.Label Label6 
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Downloads total:"
         Height          =   1695
         Left            =   60
         TabIndex        =   16
         Top             =   360
         Width           =   2115
         Begin VB.Label Label5 
            Height          =   735
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   4455
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1965
         Left            =   300
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   20
         Top             =   0
         Width           =   3810
      End
      Begin VB.Label Label7 
         Caption         =   "© Daniel Räinä 2002"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1980
         Width           =   2715
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      Begin VB.CheckBox Check3 
         Caption         =   "Graphical display only"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmoptions.frx":17552
         Left            =   2520
         List            =   "frmoptions.frx":17574
         TabIndex        =   28
         Text            =   "Transparency"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000A&
         Caption         =   "Start when windows starts"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000A&
         Caption         =   "Always on top"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   1695
         Left            =   2520
         TabIndex        =   9
         Top             =   60
         Width           =   1875
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00BEAA49&
            Height          =   315
            Left            =   180
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   14
            Top             =   960
            Width           =   315
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H0000FF00&
            Height          =   315
            Left            =   180
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   12
            Top             =   600
            Width           =   315
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H000000FF&
            Height          =   315
            Left            =   180
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   10
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Background"
            Height          =   195
            Left            =   660
            TabIndex        =   15
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Upload"
            Height          =   195
            Left            =   660
            TabIndex        =   13
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Download"
            Height          =   195
            Left            =   660
            TabIndex        =   11
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Connection speed"
         Height          =   1695
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   2355
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter your connection speed. For example, if you have a 56k modem, you enter 56000."
            Height          =   795
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Transparency"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   4455
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmoptions.frx":175B5
         Top             =   0
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
If Check2.Value = 1 Then
Call SetStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NetMon", App.Path + "\" + App.EXEName + ".exe")
End If
If Check2.Value = 0 Then
Call DelStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NetMon") ', App.Path + "\" + App.EXEName + ".exe")
End If
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Select Case Combo1.ListIndex
    Case 0
    Call MakeTransparent(frmmain.hWnd, 10 * 255 / 100)
    Case 1
    Call MakeTransparent(frmmain.hWnd, 20 * 255 / 100)
    Case 2
    Call MakeTransparent(frmmain.hWnd, 30 * 255 / 100)
    Case 3
    Call MakeTransparent(frmmain.hWnd, 40 * 255 / 100)
    Case 4
    Call MakeTransparent(frmmain.hWnd, 50 * 255 / 100)
    Case 5
    Call MakeTransparent(frmmain.hWnd, 60 * 255 / 100)
    Case 6
    Call MakeTransparent(frmmain.hWnd, 70 * 255 / 100)
    Case 7
    Call MakeTransparent(frmmain.hWnd, 80 * 255 / 100)
    Case 8
    Call MakeTransparent(frmmain.hWnd, 90 * 255 / 100)
    Case 9
    Call MakeTransparent(frmmain.hWnd, 100 * 255 / 100)
    
End Select
Call SaveSetting("netmon", "options", "visibility", Str(Combo1.ListIndex))
End Sub

Private Sub Command1_Click()
conspeed = Text1.Text
If Check2.Value = 0 Then SaveSetting "netmon", "setting", "autostart", "0"
If Check2.Value = 1 Then SaveSetting "netmon", "setting", "autostart", "1"
SaveSetting "netmon", "options", "graphics", Str(Check3.Value)
If Check3.Value = 1 Then grphx = 265
If Check3.Value = 0 Then grphx = 0
frmmain.height = frmmain.height + 15
frmmain.width = frmmain.width + 15
SaveSetting "netmon", "setting", "connectionspeed", Text1.Text
If Check1.Value = 1 Then
If ontop = True Then GoTo xit
ontop = True
StayOnTop frmmain, frmmain.Left, frmmain.Top, frmmain.width, frmmain.height
SaveSetting "netmon", "setting", "ontop", "true"
End If
If Check1.Value = 0 Then
removefromtop frmmain, frmmain.Left, frmmain.Top, frmmain.width, frmmain.height
If ontop = False Then GoTo xit
ontop = False
SaveSetting "netmon", "setting", "ontop", "false"
End If
xit:
Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
    Picture1.Visible = True
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
Case 1
    Picture1.Visible = False
    Picture2.Visible = True
    Picture3.Visible = False
    Picture4.Visible = False
Case 2
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = True
    Picture4.Visible = False
Case 3
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = True
End Select
DoEvents
Call drawsplash
End Sub

Private Sub Form_Load()
On Error Resume Next
Check3.Value = GetSetting("netmon", "options", "graphics", "0")
Label9.Caption = "Total downloads since " & installdate
CommonDialog1.CancelError = True
Combo1.ListIndex = GetSetting("netmon", "options", "visibility", "9")
Dim dltemp2 As Double, ultemp2 As Double
If GetSetting("netmon", "setting", "autostart", "0") = "1" Then Check2.Value = 1
StayOnTop Me, (Screen.width / 2 - Me.width / 2) / 15, (Screen.height / 2 - Me.height / 2) / 15, Me.width, Me.height
Text1.Text = conspeed
dltemp2 = Round(DLtot, 0)
dltemp = GiveByteValues(dltemp2)
Label5.Caption = Round(dltemp, 2) & " " & what
ultemp2 = Round(ULtot, 0)
ultemp = GiveByteValues(ultemp2)
Label6.Caption = Round(ultemp, 2) & " " & what
Picture5.BackColor = dlcolor
Picture6.BackColor = ulcolor
Picture7.BackColor = bgcolor
If ontop = True Then Check1.Value = 1
SetNumber Text1, True
End Sub

Private Sub Picture5_Click()
On Error GoTo error
CommonDialog1.ShowColor
Picture5.BackColor = CommonDialog1.Color
dlcolor = Picture5.BackColor
Call SaveSetting("netmon", "setting", "dlcolor", dlcolor)
error:
End Sub

Private Sub Picture6_Click()
On Error GoTo error
CommonDialog1.ShowColor
Picture6.BackColor = CommonDialog1.Color
ulcolor = Picture6.BackColor
Call SaveSetting("netmon", "setting", "ulcolor", ulcolor)
error:
End Sub

Private Sub Picture7_Click()
On Error GoTo error
CommonDialog1.ShowColor
Picture7.BackColor = CommonDialog1.Color
bgcolor = Picture7.BackColor
Call SaveSetting("netmon", "setting", "bgcolor", bgcolor)
frmmain.picbuffer.BackColor = bgcolor
error:
End Sub


Private Sub drawsplash()
Dim crcolor As Long

For x = 1 To Picture8.ScaleWidth
    For y = 1 To Picture8.ScaleHeight
    DoEvents
    crcolor = GetPixel(Picture9.hdc, x, y)
    Picture8.Line (x, y)-(Picture8.ScaleWidth, y), crcolor
    Next y
Next x
Picture8.Picture = Picture9.Picture
End Sub
