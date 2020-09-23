VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   7920
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
      Begin VB.VScrollBar scrllYScan 
         Height          =   855
         LargeChange     =   5
         Left            =   1200
         Max             =   20
         Min             =   1
         TabIndex        =   22
         Top             =   240
         Value           =   2
         Width           =   255
      End
      Begin VB.HScrollBar scrllXScan 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   21
         Top             =   840
         Value           =   5
         Width           =   975
      End
      Begin VB.HScrollBar scrllSleep 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   1000
         TabIndex        =   19
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   1335
         Begin VB.OptionButton Option1 
            Caption         =   "Solid"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Transparent"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdFadeIn 
         Caption         =   "Fade In"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdFadeOut 
         Caption         =   "Fade InOut"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblYScan 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblXScan 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "           Mask          XScan        YScan"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Sleep"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   495
      End
   End
   Begin VB.PictureBox picSpriteWhtBkg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   4200
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox picSpriteBlkBkg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   240
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox picToFadeIn 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   4080
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   480
      Width           =   2925
   End
   Begin VB.PictureBox picHidden1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   10080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMask 
      Height          =   360
      Left            =   7920
      Picture         =   "frmMain.frx":55A1
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   840
      Width           =   360
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   10080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   120
      Picture         =   "frmMain.frx":563B
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   480
      Width           =   3750
   End
   Begin VB.Label Label1 
      Caption         =   "(click for new mask)"
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   18
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Mask Black/White"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Sprite ""And"""
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sprite ""OR"""
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sprite (click to load new sprite)"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Main Screen (click to load new picture)"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
rtnInitialize
End Sub

Private Sub rtnInitialize()
intMaxPicSize = 250 'Height and Width
intPlaceX = 0: intPlaceY = 0 'Sprite placement
'Mask scan Scrollbar Values
lblXScan = scrllXScan.Value
lblYScan = scrllYScan.Value
intScanX = scrllXScan.Value
intScanY = scrllYScan.Value
Option1(0).Value = True 'Select transparency
booReDraw = True 'Whether or not to redraw sprites (avoid redundancy)
frmMain.Show
frmMain.Refresh
frmMain.ScaleMode = 3
picMain.ScaleMode = 3
picMain.AutoRedraw = True
picMain.AutoSize = True
picMain.BackColor = vbWhite
If picMain.Width > 300 Then picMain.Width = 300
If picMain.Height > 300 Then picMain.Height = 300
picMain.Picture = picMain.Image
Set pixMainPicture = picMain.Picture
Set pixMainPictureWithSprite = picMain.Picture 'In case of FadeInOut click
picEdit.ScaleMode = 3
picEdit.AutoRedraw = True
picEdit.AutoSize = True
picEdit.BackColor = vbWhite
picEdit.Picture = picEdit.Image
picMask.Picture = picMask.Image
End Sub

Private Sub cmdFadeIn_Click()
If Option1(0).Value = True And booReDraw = True Then rtnShowPictureDraw
rtnFadeIn
booReDraw = False
End Sub

Private Sub cmdFadeOut_Click()
If Option1(0).Value = True And booReDraw = True Then rtnShowPictureDraw
rtnFadeIn
rtnFadeOut
booReDraw = False
End Sub

Private Sub picMain_Click()
'Get New Main Screen Picture
rtnCommonDialog
If stFilename = "" Then Exit Sub
booReDraw = True 'Redraw sprites
picMain.Picture = LoadPicture(stFilename)
If picMain.Width > intMaxPicSize Then picMain.Width = intMaxPicSize
If picMain.Height > intMaxPicSize Then picMain.Height = intMaxPicSize
If picToFadeIn.Width > picMain.Width Then picToFadeIn.Width = picMain.Width
If picToFadeIn.Height > picMain.Height Then picToFadeIn.Height = picMain.Height
picMain.Picture = picMain.Image
Set pixMainPicture = picMain.Picture
picToFadeIn.Picture = picToFadeIn.Image
End Sub
Private Sub picToFadeIn_Click()
'Get New Sprite
rtnCommonDialog
If stFilename = "" Then Exit Sub
booReDraw = True 'Redraw sprites
picToFadeIn.Picture = LoadPicture(stFilename)
If picToFadeIn.Width > picMain.Width Then picToFadeIn.Width = picMain.Width
If picToFadeIn.Height > picMain.Height Then picToFadeIn.Height = picMain.Height
picToFadeIn.Picture = picToFadeIn.Image
End Sub
Private Sub picMask_Click()
'Get New Mask
Dim stTemp As String
rtnCommonDialog
If stFilename = "" Then Exit Sub
picMask.Picture = LoadPicture(stFilename)
picMask.Picture = picMask.Image
stTemp = Right(stFilename, 10) 'mask**.***
stTemp = Mid(stTemp, 5, 2)
'Select X,Y default scan values per each mask to avoid user confusion
Select Case stTemp
    Case "01"
    scrllXScan.Value = 5
    scrllYScan.Value = 2
    Case "02"
    scrllXScan.Value = 8
    scrllYScan.Value = 8
    Case "03"
    scrllXScan.Value = 6
    scrllYScan.Value = 7
    Case "04"
    scrllXScan.Value = 14
    scrllYScan.Value = 1
    Case "05"
    scrllXScan.Value = 18
    scrllYScan.Value = 1
    Case "06"
    scrllXScan.Value = 10
    scrllYScan.Value = 1
    Case "07"
    scrllXScan.Value = 10
    scrllYScan.Value = 1
    Case "08"
    scrllXScan.Value = 11
    scrllYScan.Value = 1
    Case "10"
    scrllXScan.Value = 6
    scrllYScan.Value = 7
    Case "11"
    scrllXScan.Value = 12
    scrllYScan.Value = 1
    Case "12"
    scrllXScan.Value = 20
    scrllYScan.Value = 1
    Case "13"
    scrllXScan.Value = 10
    scrllYScan.Value = 1
    Case "14"
    scrllXScan.Value = 6
    scrllYScan.Value = 1
End Select
End Sub

Private Sub rtnCommonDialog()
'Common Dialog Window
CommonDialog1.CancelError = True  'Enable on error or cancel GoTo
On Error GoTo cancelPressed
CommonDialog1.Flags = cdlOFNHideReadOnly  'disable read only chk box
CommonDialog1.DialogTitle = "Open Main Picture"        'Title displayed
CommonDialog1.InitDir = ""                'Start Directory
       
'Format     object.Filter [= description1 |filter1 |description2 |filter2...]
'Example    Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico
CommonDialog1.Filter = "Pictures (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|(All Files)|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
stFilename = CommonDialog1.FileName
Exit Sub
cancelPressed:
stFilename = CommonDialog1.FileName
End Sub

Private Sub scrllXScan_Change()
lblXScan = scrllXScan.Value
intScanX = scrllXScan.Value
End Sub

Private Sub scrllYScan_Change()
lblYScan = scrllYScan.Value
intScanY = scrllYScan.Value
End Sub
