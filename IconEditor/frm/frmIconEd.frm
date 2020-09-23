VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmIconEd 
   BorderStyle     =   0  'None
   Caption         =   "DreamVB's IconED"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6660
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   46
      Top             =   105
      Width           =   480
   End
   Begin VB.PictureBox SrcBrush2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6660
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   45
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox SrcBrush1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6690
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   44
      Top             =   1140
      Visible         =   0   'False
      Width           =   480
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   9
      Left            =   6105
      TabIndex        =   43
      ToolTipText     =   "Stamp"
      Top             =   1320
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":0000
   End
   Begin VB.PictureBox pSrc2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6690
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      Top             =   645
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   975
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraProp 
      Height          =   2415
      Left            =   5640
      TabIndex        =   33
      Top             =   2565
      Width           =   1545
      Begin VB.ComboBox cboBrush 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1995
         Width           =   1320
      End
      Begin VB.ComboBox cboFlip 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1455
         Width           =   1320
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   930
         Width           =   1320
      End
      Begin VB.ComboBox cboWidth 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   390
         Width           =   1320
      End
      Begin VB.Label lblBrush 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brush:"
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lblFlip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flip Image:"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DrawStyle:"
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   735
         Width           =   765
      End
      Begin VB.Label lblDraw 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DrawWidth:"
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   165
         Width           =   840
      End
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   8
      Left            =   6120
      TabIndex        =   32
      ToolTipText     =   "Color Picker"
      Top             =   900
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":0112
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   31
      Top             =   5085
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
      EndProperty
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   5
      Left            =   5670
      TabIndex        =   28
      ToolTipText     =   "Circel Filled"
      Top             =   2160
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":0464
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   1
      Left            =   5670
      TabIndex        =   24
      ToolTipText     =   "Line"
      Top             =   480
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":07B6
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   0
      Left            =   5670
      TabIndex        =   23
      ToolTipText     =   "Pencil"
      Top             =   75
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":0B08
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5DCD7&
      Height          =   240
      Index           =   32
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   22
      Top             =   4710
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   240
      Index           =   31
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   21
      Top             =   4470
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   240
      Index           =   14
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   20
      Top             =   4200
      Width           =   600
   End
   Begin VB.PictureBox pBackC 
      BackColor       =   &H00F5DCD7&
      Height          =   705
      Left            =   375
      ScaleHeight     =   645
      ScaleWidth      =   195
      TabIndex        =   19
      ToolTipText     =   "Backcolor"
      Top             =   90
      Width           =   255
   End
   Begin VB.PictureBox pForeC 
      BackColor       =   &H00000000&
      Height          =   705
      Left            =   75
      ScaleHeight     =   645
      ScaleWidth      =   195
      TabIndex        =   18
      ToolTipText     =   "Forecolor"
      Top             =   90
      Width           =   255
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000080&
      Height          =   240
      Index           =   13
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   17
      Top             =   3960
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   12
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   16
      Top             =   3720
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   240
      Index           =   11
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   15
      Top             =   3495
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   240
      Index           =   10
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   14
      Top             =   3270
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008080&
      Height          =   240
      Index           =   9
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   13
      Top             =   3045
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      Height          =   240
      Index           =   8
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   12
      Top             =   2805
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   240
      Index           =   7
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   11
      Top             =   2580
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   240
      Index           =   6
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   10
      Top             =   2340
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800080&
      Height          =   240
      Index           =   5
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   9
      Top             =   2115
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   240
      Index           =   4
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   8
      Top             =   1890
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      Height          =   240
      Index           =   3
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   7
      Top             =   1665
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      Height          =   240
      Index           =   2
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   1395
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   1
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   1140
      Width           =   600
   End
   Begin VB.PictureBox pColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   4
      Top             =   870
      Width           =   600
   End
   Begin VB.PictureBox pHolder 
      Height          =   4905
      Left            =   690
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   323
      TabIndex        =   2
      Top             =   75
      Width           =   4905
      Begin VB.PictureBox pDraw 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   15
         ScaleHeight     =   321
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   321
         TabIndex        =   3
         Top             =   15
         Width           =   4815
      End
   End
   Begin MSComctlLib.ImageList ImgSaveIco 
      Left            =   7305
      Top             =   8340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16112855
      _Version        =   393216
   End
   Begin VB.PictureBox p2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   4635
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   1
      Top             =   7380
      Width           =   15
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   180
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   7770
      Width           =   4815
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   2
      Left            =   5670
      TabIndex        =   25
      ToolTipText     =   "Box"
      Top             =   900
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":0E5A
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   3
      Left            =   5670
      TabIndex        =   26
      ToolTipText     =   "Box Filled"
      Top             =   1320
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":11AC
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   4
      Left            =   5670
      TabIndex        =   27
      ToolTipText     =   "Circle"
      Top             =   1740
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":14FE
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   6
      Left            =   6120
      TabIndex        =   29
      ToolTipText     =   "Fillcan"
      Top             =   75
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":1850
   End
   Begin Project1.dFlatButton cmdTools 
      Height          =   390
      Index           =   7
      Left            =   6120
      TabIndex        =   30
      ToolTipText     =   "Erase"
      Top             =   495
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmIconEd.frx":1BA2
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   24
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   24
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuGrid 
         Caption         =   "Grid"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mnuRotate 
         Caption         =   "Rotate"
         Begin VB.Menu mnuRight 
            Caption         =   "&Right"
         End
         Begin VB.Menu mnuLeft 
            Caption         =   "&Left"
         End
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "&Flip"
         Begin VB.Menu mnuHoz 
            Caption         =   "&Horizontal"
         End
         Begin VB.Menu mnuVer 
            Caption         =   "&Vertical"
         End
      End
      Begin VB.Menu mnuMirror 
         Caption         =   "&Mirror"
         Begin VB.Menu mnuLeft1 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuRight1 
            Caption         =   "&Right"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmIconEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum DrawTools
    tPen = 0
    tLine = 1
    tBox1 = 2
    tBox2 = 3
    tEllipse1 = 4
    tEllipse2 = 5
    tFillcan = 6
    tErase = 7
    tColorPicker = 8
    tStamp = 9
End Enum

Private Tool As DrawTools
Private Draw As Boolean
Private oColor1 As OLE_COLOR
Private sX As Single
Private sY As Single
Private oX As Single
Private oY As Single
Private dInitDir As String

Private Sub SaveIcon(SaveFileName As String)
Dim iPic As IPictureDisp
    'Save the icon
    Set iPic = New StdPicture
    'Store image
    ImgSaveIco.ListImages.Add , , pIcon.Image
    'Convert to icon
    Set iPic = ImgSaveIco.ListImages(1).ExtractIcon
    'Save the icon
    Call SavePicture(iPic, SaveFileName)
    ImgSaveIco.ListImages.Clear
    Set iPic = Nothing
End Sub

Private Sub DrawBrush()
Dim x As Long
Dim y As Long
Dim col As Long
Dim iRet As Long

    'Set the brush to be copyied
    SrcBrush2.Picture = SrcBrush1.Picture
    
    For x = 0 To SrcBrush2.ScaleWidth
        For y = 0 To SrcBrush2.ScaleHeight
            col = GetPixel(SrcBrush2.hdc, x, y)
            If (col = vbBlack) Then
                iRet = SetPixelV(SrcBrush2.hdc, x, y, pIcon.ForeColor)
            End If
        Next y
    Next x
    
    SrcBrush2.Refresh
    
End Sub

Private Sub LoadBrushs()
Dim xFile As String
    'This sub Loads in the brushs
    xFile = Dir(AppData & "brushs\*.bmp")

    Do Until (xFile = "")
        cboBrush.AddItem xFile
        'Get next filename
        xFile = Dir()
        DoEvents
    Loop
    
    If (cboBrush.ListCount) Then
        cboBrush.ListIndex = 0
    End If
    
End Sub


Private Sub RotateIcon(RRight As Boolean)
Dim x As Integer
Dim y As Integer
Dim col As Long
Dim iRet As Long
    'Used to rotate the icon.
    Set pSrc2.Picture = pIcon.Image
    
    For x = 0 To (32 - 1)
        For y = 0 To (32 - 1)
            col = GetPixel(pIcon.hdc, x, y)
            If (RRight) Then
                'Right
                iRet = SetPixelV(pSrc2.hdc, (32 - 1 - y), x, col)
            Else
                'Left
                iRet = SetPixelV(pSrc2.hdc, y, (32 - 1 - x), col)
            End If
        Next y
    Next x
    
    Set pIcon.Picture = pSrc2.Image
    'Update the icon
    Call UpdateIcon
End Sub

Private Sub FlipIcon(FlipOp As Integer)

    If (FlipOp = 0) Then
        StretchBlt pIcon.hdc, (DrawArea - 1), 0, -DrawArea, DrawArea, _
        pIcon.hdc, 0, 0, DrawArea, DrawArea, vbSrcCopy
    End If
    
    If (FlipOp = 1) Then
        StretchBlt pIcon.hdc, 0, (DrawArea - 1), DrawArea, -DrawArea, _
        pIcon.hdc, 0, 0, DrawArea, DrawArea, vbSrcCopy
    End If
    
    pIcon.Refresh
    Call UpdateIcon
End Sub

Private Function GetDLGName(Optional dShowOpen As Boolean = True, _
Optional dTitle As String = "Open") As String
On Error GoTo OpenErr:

    With CD1
        .CancelError = True
        .DialogTitle = dTitle
        .Filter = "Icon Files(*.ico)|*.ico|"
        .InitDir = dInitDir
        .FileName = ""
        
        If (dShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        
        'Set InitDir
        dInitDir = GetFilePath(.FileName)
        'Return filename
        GetDLGName = .FileName
    End With
    
    Exit Function
OpenErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub NewIcon()
    'Erase any picture
    Set pIcon.Picture = Nothing
    'Set Icon mask color
    pIcon.BackColor = ImgSaveIco.MaskColor
    'Update icon
    Call UpdateIcon
End Sub

Private Sub SetupGrid()
Dim Count As Long

    'This just draws the grid style lines
    For Count = 0 To DrawArea Step ColSize
        Call DottedLine(1, Count, p1)
        Call DottedLine(2, Count, p2)
    Next Count
    
    Call p1.Refresh
    Call p2.Refresh
    
    'Update the drawing area.
    Call UpdateIcon
End Sub

Private Sub UpdateIcon()
Dim iRet As Long
Dim Count As Long

    iRet = StretchBlt(pDraw.hdc, 0, 0, DrawArea, DrawArea, pIcon.hdc, 0, 0, 32, 32, vbSrcCopy)
    
    'Check if showing grid
    If (mnuGrid.Checked) Then
        'Draw the grid
        For Count = 0 To pDraw.ScaleWidth Step ColSize
            BitBlt pDraw.hdc, 0, Count, (pDraw.ScaleWidth - 1), 1, p1.hdc, 0, 0, vbSrcCopy
            BitBlt pDraw.hdc, Count, 0, 1, (pDraw.ScaleHeight - 1), p2.hdc, 0, 0, vbSrcCopy
        Next Count
    End If
    
    Call pDraw.Refresh
End Sub

Private Sub cmdTool_Click(Index As Integer)
    Tool = Index
End Sub

Private Sub cboBrush_Click()
    'Load custom brush
    SrcBrush1.Picture = LoadPicture(AppData & "Brushs\" & cboBrush.Text)
End Sub

Private Sub cboFlip_Click()
    'Flip Icon
    Call FlipIcon(cboFlip.ListIndex)
End Sub

Private Sub cboStyle_Click()
    'Set Drawwidth
    pIcon.DrawStyle = cboStyle.ListIndex
End Sub

Private Sub cboWidth_Click()
    pIcon.DrawWidth = Val(cboWidth.Text)
End Sub

Private Sub cmdTools_Click(Index As Integer)
    lblBrush.Visible = (Index = tStamp)
    cboBrush.Visible = (Index = tStamp)
    Tool = Index
End Sub

Private Sub Form_Load()
    AppData = FixPath(App.Path)
    Call NewIcon
    Call SetupGrid
    'Add draw widths
    cboWidth.AddItem "1"
    cboWidth.AddItem "2"
    cboWidth.AddItem "3"
    cboWidth.AddItem "4"
    cboWidth.AddItem "5"
    cboWidth.AddItem "6"
    cboWidth.AddItem "7"
    cboWidth.AddItem "8"
    cboFlip.AddItem "Horizontal"
    cboFlip.AddItem "Vertical"
    'Add some draw styles
    cboStyle.AddItem "Soild"
    cboStyle.AddItem "Dash"
    cboStyle.AddItem "Dot"
    cboStyle.AddItem "Dash-Dot"
    cboWidth.ListIndex = 0
    cboStyle.ListIndex = 0
    cboFlip.ListIndex = 0
    'Load brushs
    Call LoadBrushs
    'Select first tool
    Call cmdTools_Click(0)
End Sub

Private Sub Form_Resize()
    lnTop(0).X2 = frmIconEd.ScaleWidth
    lnTop(1).X2 = frmIconEd.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIconEd = Nothing
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmIconEd.Caption & " Ver 1.0", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
    Unload frmIconEd
End Sub

Private Sub mnuGrid_Click()
    mnuGrid.Checked = (Not mnuGrid.Checked)
    Call UpdateIcon
End Sub

Private Sub mnuHoz_Click()
    Call FlipIcon(0)
End Sub

Private Sub mnuLeft_Click()
    'Rotate Left
    Call RotateIcon(False)
End Sub

Private Sub mnuLeft1_Click()
    'Mirror left
    Call Mirror(pIcon, 0)
    'Update icon
    Call UpdateIcon
End Sub

Private Sub mnuNew_Click()
    If MsgBox("Are you sure you want to start a new icon?", vbYesNo Or vbExclamation, frmIconEd.Caption) = vbYes Then
        Call NewIcon
    End If
End Sub

Private Sub mnuOpen_Click()
Dim lFile As String
    'Open Icon Filename
    lFile = GetDLGName()
    'Load the icon on to pIcon
    pIcon.Picture = LoadPicture(lFile)
    'Update the icon
    Call UpdateIcon
End Sub

Private Sub mnuRight_Click()
    'Rotate Right
    Call RotateIcon(True)
End Sub

Private Sub mnuRight1_Click()
    'Mirror Right
    Call Mirror(pIcon, 0)
    'Update icon
    Call UpdateIcon
End Sub

Private Sub mnuSave_Click()
Dim lFile As String
    'Get save filename
    lFile = GetDLGName(False, "Save")
    
    If Len(lFile) Then
        'Save the Icon
        Call SaveIcon(lFile)
    End If
    
End Sub

Private Sub mnuVer_Click()
    Call FlipIcon(1)
End Sub

Private Sub pColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbLeftButton) Then pForeC.BackColor = pColor(Index).BackColor
    If (Button = vbRightButton) Then pBackC.BackColor = pColor(Index).BackColor
End Sub

Private Sub pDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim iRet As Long
    
    'Store mouse positions
    sX = (x \ ColSize)
    sY = (y \ ColSize)
        
    oX = sX
    oY = sY

    If (Button = vbLeftButton) Then pIcon.ForeColor = pForeC.BackColor
    If (Button = vbRightButton) Then pIcon.ForeColor = pBackC.BackColor
        
    Select Case Tool
        Case tPen
            'Pencil Tool
            pIcon.PSet (sX, sY)
            Call UpdateIcon
            sX = oX
            sY = oY
        Case tFillcan
            'Fillcan Tool
            pIcon.FillColor = pIcon.ForeColor
            pIcon.FillStyle = vbFSSolid
            iRet = ExtFloodFill(pIcon.hdc, sX, sY, pIcon.Point(sX, sY), &H1)
            'Update Icon
            Call UpdateIcon
        Case tColorPicker
            'Color Picker Tool
            If (Button = vbLeftButton) Then pForeC.BackColor = pIcon.Point(sX, sY)
            If (Button = vbRightButton) Then pBackC.BackColor = pIcon.Point(sX, sY)
        Case tStamp
            'Stamp Tool
            Call DrawBrush
            iRet = TransparentBlt(pIcon.hdc, (sX - SrcBrush2.Width \ 2), (sY - SrcBrush2.Height \ 2), SrcBrush2.Width, SrcBrush2.Height, _
            SrcBrush2.hdc, 0, 0, SrcBrush2.Width, SrcBrush2.Height, RGB(255, 0, 255))
            'Update Icon
            Call UpdateIcon
    End Select
        
    'Allow drawing
    Draw = True
 
End Sub

Private Sub pDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim iRet As Long
Dim mRgb As RGBQUAD

    'Store mouse positions
    oX = (x \ ColSize)
    oY = (y \ ColSize)
        
    'Get RGB Color
    Call LongToRGB(GetPixel(pIcon.hdc, oX, oY), mRgb)
    'Update statusbar text
    sBar1.Panels(1).Text = "X: " & oX & ", Y: " & oY
    sBar1.Panels(2).Text = "RGB: [" & mRgb.Red & ", " & mRgb.Green & ", " & mRgb.Blue & "]"
    
    'Check drawing tool in use.
    If (Draw = True) Then
        Select Case Tool
            Case tPen
                'Pencil Tool
                pIcon.Line (sX, sY)-(oX, oY)
                Call UpdateIcon
                sX = oX
                sY = oY
            Case tLine
                'Line Tool
                pIcon.AutoRedraw = False
                pIcon.Refresh
                pIcon.Line (sX, sY)-(oX, oY)
                Call UpdateIcon
            Case tBox1
                'Box Tool
                pIcon.AutoRedraw = False
                pIcon.Refresh
                pIcon.Line (sX, sY)-(oX, oY), , B
                Call UpdateIcon
            Case tBox2
                'Box Tool
                pIcon.AutoRedraw = False
                pIcon.Refresh
                pIcon.Line (sX, sY)-(oX, oY), , BF
                Call UpdateIcon
                'Ellipse tool
            Case tEllipse1
                pIcon.AutoRedraw = False
                pIcon.Refresh
                iRet = Ellipse(pIcon.hdc, sX, sY, oX, oY)
                Call UpdateIcon
                'Filled Ellipse Tool
            Case tEllipse2
                pIcon.AutoRedraw = False
                pIcon.FillColor = pIcon.ForeColor
                pIcon.FillStyle = 0
                pIcon.Refresh
                iRet = Ellipse(pIcon.hdc, sX, sY, oX, oY)
                Call UpdateIcon
            Case tErase
                'Erase Tool
                iRet = SetPixelV(pIcon.hdc, sX, sY, ImgSaveIco.MaskColor)
                Call UpdateIcon
                sX = oX
                sY = oY
            Case tStamp
                'Stamp Tool
                Call DrawBrush
                iRet = TransparentBlt(pIcon.hdc, (sX - SrcBrush2.Width \ 2), (sY - SrcBrush2.Height \ 2), SrcBrush2.Width, SrcBrush2.Height, _
                SrcBrush2.hdc, 0, 0, SrcBrush2.Width, SrcBrush2.Height, RGB(255, 0, 255))
        End Select
    End If
    
End Sub

Private Sub pDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim iRet As Long
    pIcon.AutoRedraw = True
    
    'Line Tool
    Select Case Tool
        Case tLine
            'Line Tool
            pIcon.Line (sX, sY)-(oX, oY)
        Case tBox1
            'Box Tool
            pIcon.Line (sX, sY)-(oX, oY), , B
        Case tBox2
            'Filled Box Tool
            pIcon.Line (sX, sY)-(oX, oY), , BF
        Case tEllipse1
            'Ellipse Tool
            iRet = Ellipse(pIcon.hdc, sX, sY, oX, oY)
        Case tEllipse2
            'Filled Ellipse Tool
            pIcon.FillStyle = 0
            iRet = Ellipse(pIcon.hdc, sX, sY, oX, oY)
        End Select
        
        pIcon.FillStyle = vbFSTransparent
        pIcon.Refresh
        Draw = False
        Call UpdateIcon
End Sub
