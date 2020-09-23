VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      Height          =   2040
      Left            =   0
      ScaleHeight     =   1980
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox picdisplay 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00BEAA49&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H000000FF&
         Height          =   1980
         Left            =   0
         ScaleHeight     =   132
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   3
         Top             =   -15
         Width           =   4500
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3510
      Top             =   2745
   End
   Begin VB.PictureBox picbuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00BEAA49&
      ForeColor       =   &H000000FF&
      Height          =   2040
      Left            =   600
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   4560
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2835
      Top             =   2595
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":20D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2100
      Width           =   5610
   End
   Begin VB.Menu menu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu options 
         Caption         =   "Inställningar"
      End
      Begin VB.Menu quit 
         Caption         =   "Avsluta"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declarations for trayicon
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private nid As NOTIFYICONDATA

Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private m_objIpHelper As CIpHelper

Dim l_sent(300) As Double, l_rec(300) As Double
Dim oOld As Long
Dim oNew As Long
Dim sOld As Long
Dim sNew As Long
Dim trayed As Boolean
Dim svalue As Long, tvalue As Long
Dim xpos As Long, ypos As Long, windowswidth As Long, windowsheight As Long

Private Sub Form_Load()
'Initialize settings
On Error Resume Next
If GetSetting("netmon", "options", "graphics", "0") = 1 Then grphx = 265
Dim vis As Integer
installdate = GetSetting("netmon", "options", "installdate", "null")
If installdate = "null" Then
    installdate = FormatDateTime(Now, vbLongDate)
    SaveSetting "netmon", "options", "installdate", installdate
End If
vis = GetSetting("netmon", "options", "visibility", "9")
Select Case vis
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


dlcolor = GetSetting("netmon", "setting", "dlcolor", "&H0000FF")
ulcolor = GetSetting("netmon", "setting", "ulcolor", "&H00FF00")
bgcolor = GetSetting("netmon", "setting", "bgcolor", "&HBEAA49")
conspeed = GetSetting("netmon", "setting", "connectionspeed", "1024000")
xpos = GetSetting("netmon", "setting", "windowsposx", "0")
ypos = GetSetting("netmon", "setting", "windowsposy", "0")
windowswidth = GetSetting("netmon", "setting", "windowswidth", "4650")
windowsheight = GetSetting("netmon", "setting", "windowsheight", "2460")
DLtot = GetSetting("netmon", "setting", "DLtot", "1")
ULtot = GetSetting("netmon", "setting", "ULtot", "1")
ontop = GetSetting("netmon", "setting", "ontop", "true")
picbuffer.BackColor = bgcolor
OldWindowProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WindowProc)
If ontop = True Then
StayOnTop Me, xpos, ypos, windowswidth, windowsheight
Else
Me.Move xpos / 15, ypos / 15, windowswidth / 15, windowsheight / 15
End If
Form_Resize
Set m_objIpHelper = New CIpHelper
With nid
.cbSize = Len(nid)
.hWnd = Me.hWnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = ImageList1.ListImages(4).Picture
End With
nid.szTip = "NetMon" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Resize()
On Error Resume Next
picdisplay.Left = Me.width - (picdisplay.width) - 160
picdisplay.height = Me.height - 360 + grphx
picbuffer.height = Me.height - 360 + grphx
If grphx = 265 Then
Picture1.height = Me.height - 390 + grphx + 15
Else
Picture1.height = Me.height - 390
End If
Label8.Top = Me.height - 370 + grphx
Label8.width = Me.width - 135
If Me.width > 2220 Then
Label8.FontSize = 10
End If
If Me.width < 2220 And Me.width > 2000 Then
Label8.FontSize = 8
End If
If Me.width < 2000 Then
Label8.FontSize = 6
End If

draw (True)
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hWnd, &H112, 61458, 0
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hWnd, &H112, 61458, 0
End Sub

Private Sub options_Click()
frmoptions.Show 1, frmmain
End Sub

Private Sub picdisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hWnd, &H112, 61458, 0
End If
If Button = 2 Then
Me.PopupMenu menu, options
End If
End Sub

Private Sub quit_Click()
SaveSetting "netmon", "setting", "DLtot", Round(DLtot, 0)
SaveSetting "netmon", "setting", "ULtot", Round(ULtot, 0)
SaveSetting "netmon", "setting", "windowsposx", frmmain.Left
SaveSetting "netmon", "setting", "windowsposy", frmmain.Top
SaveSetting "netmon", "setting", "windowswidth", frmmain.width
SaveSetting "netmon", "setting", "windowsheight", frmmain.height
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub



Public Sub draw(refresh As Boolean)
Dim temp As Double
If refresh = False Then
For sort = 298 To 1 Step -1
    l_sent(sort + 1) = l_sent(sort)
    l_rec(sort + 1) = l_rec(sort)
Next sort

If tvalue >= 0 Then
temp = (tvalue / 1024)
l_rec(1) = temp  ' * picdisplay.ScaleHeight
End If

If tvalue >= 0 Then
temp = (svalue / 1024)
l_sent(1) = temp ' * picdisplay.ScaleHeight
End If

End If
If trayed = False Then
For showval = 0 To 300
    If l_rec(showval) > l_sent(showval) Then
    picbuffer.ForeColor = dlcolor
    picbuffer.Line _
    (picbuffer.ScaleWidth - (showval), picdisplay.ScaleHeight - (picdisplay.ScaleHeight * l_rec(showval) / Round((conspeed * 1.3) / 10000, 0))) _
    -(picbuffer.ScaleWidth - (showval), picbuffer.ScaleHeight)
    picbuffer.ForeColor = ulcolor
    picbuffer.Line _
    (picbuffer.ScaleWidth - (showval), picdisplay.ScaleHeight - (picdisplay.ScaleHeight * l_sent(showval) / Round((conspeed * 1.3) / 10000, 0))) _
    -(picbuffer.ScaleWidth - (showval), picbuffer.ScaleHeight)
    Else
    picbuffer.ForeColor = ulcolor
    picbuffer.Line _
    (picbuffer.ScaleWidth - (showval), picdisplay.ScaleHeight - (picdisplay.ScaleHeight * l_sent(showval) / Round((conspeed * 1.3) / 10000, 0))) _
    -(picbuffer.ScaleWidth - (showval), picbuffer.ScaleHeight)
    picbuffer.ForeColor = dlcolor
    picbuffer.Line _
    (picbuffer.ScaleWidth - (showval), picdisplay.ScaleHeight - (picdisplay.ScaleHeight * l_rec(showval) / Round((conspeed * 1.3) / 10000, 0))) _
    -(picbuffer.ScaleWidth - (showval), picbuffer.ScaleHeight)
    'Debug.Print showval, l_sent(showval) / Round(conspeed / 10000, 0)
    End If
Next showval
'Debug.Print l_rec(1), temp
Call BitBlt(picdisplay.hdc, 0, 0, picdisplay.ScaleWidth, picdisplay.ScaleHeight, picbuffer.hdc, 0, 0, vbSrcCopy)
picdisplay.refresh
picbuffer.Cls
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As Long

Msg = x / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDBLCLK
    If Me.Visible = True Then
        trayed = True
        Me.Visible = False
    Else
        Me.Visible = True
        trayed = False
        Me.SetFocus
    End If
End Select
End Sub
Private Function GetTraffic()
Dim objInterface2 As CInterface
Dim obJHelper As CInterface
Set objInterface2 = New CInterface
Set obJHelper = m_objIpHelper.Interfaces(1)
oNew = m_objIpHelper.BytesReceived
sNew = m_objIpHelper.BytesSent

svalue = sNew - sOld
tvalue = oNew - oOld

If tvalue > 200000 Or tvalue < 0 Then
tvalue = 0
svalue = 0
GoTo around
End If
If svalue > 200000 Or svalue < 0 Then
GoTo around
tvalue = 0
svalue = 0
End If
ULtot = ULtot + svalue '/ 1024
DLtot = DLtot + tvalue '/ 1024
around:
Label8.Caption = "DL: " & Round((tvalue / 1024), 1) & " kb/s - UL: " & Round((svalue / 1024), 1) & " kb/s"
If Not nid.hIcon = ImageList1.ListImages(4).Picture Then
If Round((tvalue / 1024), 1) > 0 And Round((svalue / 1024), 1) > 0 Then nid.hIcon = ImageList1.ListImages(4).Picture
End If
If Not nid.hIcon = ImageList1.ListImages(2).Picture Then
If Round((tvalue / 1024), 1) > 0 And Round((svalue / 1024), 1) = 0 Then nid.hIcon = ImageList1.ListImages(2).Picture
End If
If Not nid.hIcon = ImageList1.ListImages(3).Picture Then
If Round((tvalue / 1024), 1) = 0 And Round((svalue / 1024), 1) > 0 Then nid.hIcon = ImageList1.ListImages(3).Picture
End If
If Not nid.hIcon = ImageList1.ListImages(1).Picture Then
If Round((tvalue / 1024), 1) = 0 And Round((svalue / 1024), 1) = 0 Then nid.hIcon = ImageList1.ListImages(1).Picture
End If
nid.szTip = "NetMon - DL: " & Round((tvalue / 1024), 1) & " kb/s - UL: " & Round((svalue / 1024), 1) & " kb/s" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
oOld = oNew
sOld = sNew
draw (False)
End Function

Private Sub Timer2_Timer()
GetTraffic
End Sub
