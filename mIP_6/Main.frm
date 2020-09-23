VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8910
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ME"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   11
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3540
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EE"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   10
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3540
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DF"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   9
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3540
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SS"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3540
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OU"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   7
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3540
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CO"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   6
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3540
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DI"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   1650
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BW"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1650
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BR"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1650
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RE"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   1650
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GR"
      ForeColor       =   &H00008000&
      Height          =   225
      Index           =   1
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1650
      Width           =   375
   End
   Begin VB.OptionButton optEff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BL"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1650
      Width           =   375
   End
   Begin VB.CommandButton cmdSepia 
      Caption         =   "Sepia"
      Height          =   285
      Left            =   1485
      TabIndex        =   52
      Top             =   960
      Width           =   645
   End
   Begin VB.CommandButton cmdGrey 
      Caption         =   "Grey"
      Height          =   285
      Left            =   795
      TabIndex        =   51
      Top             =   960
      Width           =   660
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x8"
      Height          =   270
      Index           =   7
      Left            =   8325
      TabIndex        =   50
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x7"
      Height          =   270
      Index           =   6
      Left            =   8010
      TabIndex        =   49
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x6"
      Height          =   270
      Index           =   5
      Left            =   7695
      TabIndex        =   48
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x5"
      Height          =   270
      Index           =   4
      Left            =   7380
      TabIndex        =   47
      Top             =   45
      Width           =   300
   End
   Begin VB.PictureBox picSelect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   9120
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   46
      Top             =   45
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select OFF"
      Height          =   270
      Index           =   1
      Left            =   3375
      TabIndex        =   44
      Top             =   420
      Width           =   960
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select ON"
      Height          =   270
      Index           =   0
      Left            =   2355
      TabIndex        =   43
      Top             =   420
      Width           =   960
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x4"
      Height          =   270
      Index           =   3
      Left            =   7065
      TabIndex        =   42
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x3"
      Height          =   270
      Index           =   2
      Left            =   6750
      TabIndex        =   41
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x2"
      Height          =   270
      Index           =   1
      Left            =   6435
      TabIndex        =   40
      Top             =   45
      Width           =   300
   End
   Begin VB.PictureBox picPB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   6945
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   39
      Top             =   525
      Width           =   1575
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "x1"
      Height          =   270
      Index           =   0
      Left            =   6120
      TabIndex        =   38
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdFlute 
      Caption         =   "Both"
      Height          =   240
      Index           =   2
      Left            =   1350
      TabIndex        =   37
      Top             =   6930
      Width           =   540
   End
   Begin VB.CommandButton cmdFlute 
      Caption         =   "Vert"
      Height          =   240
      Index           =   1
      Left            =   1650
      TabIndex        =   36
      Top             =   6645
      Width           =   510
   End
   Begin VB.CommandButton cmdFlute 
      Caption         =   "Horz"
      Height          =   240
      Index           =   0
      Left            =   1110
      TabIndex        =   35
      Top             =   6645
      Width           =   510
   End
   Begin VB.CommandButton cmdSwaps 
      Caption         =   "GR"
      Height          =   225
      Index           =   2
      Left            =   405
      TabIndex        =   33
      Top             =   6945
      Width           =   360
   End
   Begin VB.CommandButton cmdSwaps 
      Caption         =   "BR"
      Height          =   225
      Index           =   1
      Left            =   615
      TabIndex        =   32
      Top             =   6660
      Width           =   345
   End
   Begin VB.CommandButton cmdSwaps 
      Caption         =   "BG"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   31
      Top             =   6660
      Width           =   360
   End
   Begin VB.CommandButton cmdMirrors 
      Caption         =   "B"
      Height          =   270
      Index           =   3
      Left            =   450
      TabIndex        =   26
      Top             =   6030
      Width           =   285
   End
   Begin VB.CommandButton cmdMirrors 
      Caption         =   "R"
      Height          =   270
      Index           =   2
      Left            =   765
      TabIndex        =   25
      Top             =   5865
      Width           =   285
   End
   Begin VB.CommandButton cmdMirrors 
      Caption         =   "L"
      Height          =   270
      Index           =   1
      Left            =   135
      TabIndex        =   24
      Top             =   5865
      Width           =   285
   End
   Begin VB.CommandButton cmdMirrors 
      Caption         =   "T"
      Height          =   270
      Index           =   0
      Left            =   450
      TabIndex        =   23
      Top             =   5715
      Width           =   285
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      Height          =   345
      Left            =   60
      TabIndex        =   22
      ToolTipText     =   " Reload last opened file "
      Top             =   225
      Width           =   660
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix"
      Height          =   345
      Left            =   1470
      TabIndex        =   21
      ToolTipText     =   " Fix Effects as New Image "
      Top             =   225
      Width           =   660
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   11
      LargeChange     =   4
      Left            =   2010
      Max             =   6
      Min             =   1
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3795
      Value           =   3
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   10
      LargeChange     =   4
      Left            =   1605
      Max             =   -3
      Min             =   3
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3795
      Value           =   1
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   9
      LargeChange     =   4
      Left            =   1245
      Max             =   1
      Min             =   16
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3795
      Value           =   1
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   8
      LargeChange     =   4
      Left            =   885
      Max             =   7
      Min             =   1
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3795
      Value           =   7
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   4
      LargeChange     =   4
      Left            =   1635
      Max             =   255
      Min             =   1
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1905
      Value           =   100
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   5
      LargeChange     =   4
      Left            =   1965
      Max             =   16
      Min             =   48
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1905
      Value           =   16
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   6
      LargeChange     =   4
      Left            =   165
      Max             =   -66
      Min             =   66
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3795
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   7
      LargeChange     =   4
      Left            =   525
      Max             =   0
      Min             =   100
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3795
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   3
      LargeChange     =   4
      Left            =   1260
      Max             =   0
      Min             =   100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1905
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   2
      LargeChange     =   4
      Left            =   885
      Max             =   0
      Min             =   100
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1905
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   1
      LargeChange     =   4
      Left            =   510
      Max             =   0
      Min             =   100
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1905
      Width           =   240
   End
   Begin VB.VScrollBar VSRGB 
      Height          =   1500
      Index           =   0
      LargeChange     =   4
      Left            =   165
      Max             =   0
      Min             =   100
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1905
      Width           =   240
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert"
      Height          =   285
      Left            =   135
      TabIndex        =   8
      ToolTipText     =   " Invert displayed image "
      Top             =   960
      Width           =   630
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   345
      Left            =   765
      TabIndex        =   7
      ToolTipText     =   " Reset current starting image "
      Top             =   225
      Width           =   660
   End
   Begin VB.CommandButton cmdFlips 
      Caption         =   "Horz"
      Height          =   255
      Index           =   0
      Left            =   1350
      TabIndex        =   6
      Top             =   5745
      Width           =   600
   End
   Begin VB.CommandButton cmdFlips 
      Caption         =   "Vert"
      Height          =   255
      Index           =   1
      Left            =   1350
      TabIndex        =   5
      Top             =   6045
      Width           =   600
   End
   Begin VB.HScrollBar scrZoom 
      Height          =   225
      Left            =   2310
      Max             =   32
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Value           =   1
      Width           =   1770
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1530
      LargeChange     =   10
      Left            =   7965
      Max             =   0
      Min             =   156
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Value           =   1
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   10
      Left            =   2640
      Max             =   156
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5895
      Value           =   1
      Width           =   1530
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   2430
      ScaleHeight     =   332
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   840
      Width           =   5175
      Begin VB.Shape shpSelect 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         DrawMode        =   7  'Invert
         Height          =   405
         Left            =   735
         Top             =   840
         Width           =   450
      End
      Begin VB.Line LineX 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         DrawMode        =   7  'Invert
         X1              =   148
         X2              =   148
         Y1              =   28
         Y2              =   71
      End
      Begin VB.Line LineY 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         DrawMode        =   7  'Invert
         X1              =   76
         X2              =   130
         Y1              =   23
         Y2              =   23
      End
   End
   Begin VB.Label LabContStep 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Continuous"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1185
      TabIndex        =   66
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label LabEffVals 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BL  1"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   53
      Top             =   1305
      Width           =   795
   End
   Begin VB.Shape shpBorder 
      Height          =   135
      Left            =   4890
      Top             =   945
      Width           =   120
   End
   Begin VB.Label LabDims 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4440
      TabIndex        =   45
      Top             =   435
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   480
      Picture         =   "Main.frx":0442
      Stretch         =   -1  'True
      Top             =   8010
      Width           =   1140
   End
   Begin VB.Label LabFlute 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Flutes"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1245
      TabIndex        =   34
      Top             =   6345
      Width           =   780
   End
   Begin VB.Label LabSwap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Swaps"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   195
      TabIndex        =   30
      Top             =   6345
      Width           =   780
   End
   Begin VB.Label LabMirrors 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mirrors"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   195
      TabIndex        =   29
      Top             =   5415
      Width           =   780
   End
   Begin VB.Label LabFilter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filters"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   28
      Top             =   645
      Width           =   1995
   End
   Begin VB.Label LabFlips 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Flips"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   27
      Top             =   5415
      Width           =   780
   End
   Begin VB.Label LabZoom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zoom"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4065
      TabIndex        =   3
      Top             =   45
      Width           =   1920
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu FileOps 
         Caption         =   "&Open picture file"
         Index           =   0
      End
      Begin VB.Menu FileOps 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu FileOps 
         Caption         =   "&Save displayed viewport as BMP or JPEG"
         Index           =   2
      End
      Begin VB.Menu FileOps 
         Caption         =   "Save &whole image as BMP or JPEG"
         Index           =   3
      End
      Begin VB.Menu FileOps 
         Caption         =   "Save s&election as BMP or JPEG"
         Index           =   4
      End
      Begin VB.Menu FileOps 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu FileOps 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Filter &options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Continuous scroll bar actions."
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Stepped action (Press Filter buttons to apply)"
         Index           =   1
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "File Info"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mini-Image Processor  by Robert Rayment  August 2006

' Updates
' 3 Aug: Pictures with width or height <=15 were disallowed
'        now changed to <=2 pixels disallowed.

' 4 Aug: Added saving to JPEG using the Ron van Tilburg (John Korejwa)
'        class at CodeId=50351.

' 5 Aug: Minor correction to vertical scrollbar position when zooming in
'        with the horizontal cross-wire at the bottom of the picture.
'
'        Turn Select OFF when a scrollbar turns off.
'
'11 Aug  Moved vertical scroll bar to normal position on the right of
'        the viewport.  Also New pictures now open at top-left instead of
'        at bottom-left.  Correction to edging for SmoothSharp effect.
'        Option to change filter effects continuously or stepped.
'
'12 Aug  Buttons above the filter scrollbars will now act on the
'        scrollbar value without having to press the scrollbars,
'        in continuous or stepped mode.
'
'13 Aug  Added routine to cJpeg.cls to save whole image from byte array,
'        with RvT's help.

' Note
' 1.  The standard VB scrollbars can behave a bit oddly at times
'     (particularly with a manifest file) - seems to sort itself
'     out eventually.
'     For very large pictures (say > 1500x1500) alter Filter options
'     (menu item) to stepped.

Option Explicit

' For XP manifest
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long
Private m_hMod As Long
''

Private FWidORG As Long          ' Form starting & minimum W & H
Private FHitORG As Long
Private RightMargin As Long, BottomMargin As Long  ' Hard twip values

Private PathSpec$
Private CurrentPath$, FileSpec$
Private SavePath$, SaveSpec$

Private aScroll As Boolean       ' Booleans to restrict scroll bar actions
Private aZoom As Boolean
Private aRGB As Boolean
Private aMouseDown As Boolean

Private aEffAction As Boolean    ' T Continuous, F Stepped

Private aSelect As Boolean

Dim CommonDialog1 As OSDialog

'#### EFFECTS ####

Private Sub cmdSepia_Click()
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Sepia
   DISPLAY
End Sub
Private Sub cmdSepia_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabFilter = "Sepia"
End Sub

Private Sub cmdGrey_Click()
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Grey
   DISPLAY
End Sub
Private Sub cmdGrey_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabFilter = "Grey"
End Sub

Private Sub cmdInvert_Click()
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Invert
   DISPLAY
End Sub
Private Sub cmdInvert_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabFilter = "Invert displayed image"
End Sub

Private Sub cmdMirrors_Click(Index As Integer)
' 0,1,2,3  - > T,L,R,B
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Mirrors Index
   DISPLAY
End Sub
Private Sub cmdMirrors_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
   Case 0: LabFilter = "Mirror Top"
   Case 1: LabFilter = "Mirror Left"
   Case 2: LabFilter = "Mirror Right"
   Case 3: LabFilter = "Mirror Bottom"
   End Select
End Sub

Private Sub cmdFlips_Click(Index As Integer)
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Flips Index
   DISPLAY
End Sub
Private Sub cmdFlips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
   Case 0: LabFilter = "Flip Horz"
   Case 1: LabFilter = "Flip Vert"
   End Select
End Sub


Private Sub cmdSwaps_Click(Index As Integer)
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Swaps Index
   DISPLAY
End Sub
Private Sub cmdSwaps_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
   Case 0: LabFilter = "Swap Blue Green"
   Case 1: LabFilter = "Swap Blue Red"
   Case 2: LabFilter = "Swap Green Red"
   End Select
End Sub

Private Sub cmdFlute_Click(Index As Integer)
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Flute Index
   DISPLAY
End Sub
Private Sub cmdFlute_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
   Case 0: LabFilter = "Horizontal Fluted Window"
   Case 1: LabFilter = "Vertical Fluted Window"
   Case 2: LabFilter = "H && V Fluted Window"
   End Select
End Sub

Private Sub optEff_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim svaEffAction As Boolean
   svaEffAction = aEffAction
   optEff(Index).Value = False
   aEffAction = True
   VSRGB_Change Index
   aEffAction = svaEffAction
   Picture1.SetFocus
End Sub

Private Sub optEff_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
   Case 0: LabFilter = "Vary Blue"
   Case 1: LabFilter = "Vary Green"
   Case 2: LabFilter = "Vary Red"
   Case 3: LabFilter = "Vary Brightness"
   Case 4: LabFilter = "Black && White"
   Case 5: LabFilter = "Black && White Dither"
   Case 6: LabFilter = "Contrast"
   Case 7: LabFilter = "Outline"
   Case 8: LabFilter = "Sharp - Smooth"
   Case 9: LabFilter = "Diffuse"
   Case 10: LabFilter = "Emboss - Engrave"
   Case 11: LabFilter = "Melt"
   End Select
   LabFilter.Refresh
End Sub

Private Sub SetFilterLimits()
   VSRGB(0).Max = 0: VSRGB(0).Min = 20    ' B
   VSRGB(0).SmallChange = 1: VSRGB(0).LargeChange = 2
   
   VSRGB(1).Max = 0: VSRGB(1).Min = 20    ' G
   VSRGB(1).SmallChange = 1: VSRGB(1).LargeChange = 2
   
   VSRGB(2).Max = 0: VSRGB(2).Min = 20    ' R
   VSRGB(2).SmallChange = 1: VSRGB(2).LargeChange = 2
   
   VSRGB(3).Max = 0: VSRGB(3).Min = 20    ' Brightness
   VSRGB(3).SmallChange = 1: VSRGB(3).LargeChange = 2
   
   VSRGB(4).Max = 255: VSRGB(4).Min = 1   ' Black & White
   VSRGB(4).SmallChange = 1: VSRGB(4).LargeChange = 4
   
   VSRGB(5).Max = 48: VSRGB(5).Min = 16   ' Black & White Dither
   VSRGB(5).SmallChange = 1: VSRGB(5).LargeChange = 4
   
   VSRGB(6).Max = -66: VSRGB(6).Min = 66  ' Contrast
   VSRGB(6).SmallChange = 1: VSRGB(6).LargeChange = 4
   
   VSRGB(7).Max = 0: VSRGB(7).Min = 100   ' Outline
   VSRGB(7).SmallChange = 1: VSRGB(7).LargeChange = 4
   
   VSRGB(8).Max = 7: VSRGB(8).Min = 1     ' Sharp Smooth
   VSRGB(8).SmallChange = 1: VSRGB(8).LargeChange = 1
   
   VSRGB(9).Max = 1: VSRGB(9).Min = 16    ' Diffuse
   VSRGB(9).SmallChange = 1: VSRGB(9).LargeChange = 2
   
   VSRGB(10).Max = -3: VSRGB(10).Min = 3  ' Emboss Engrave
   VSRGB(10).SmallChange = 1: VSRGB(10).LargeChange = 1
   
   VSRGB(11).Max = 8: VSRGB(11).Min = -8   ' Melt
   VSRGB(11).SmallChange = 1: VSRGB(11).LargeChange = 2
End Sub

Private Sub SetVSRGBValues()
Dim k As Long
   aRGB = False
   For k = 0 To 3
      VSRGB(k).Value = 4
   Next k
   VSRGB(4).Value = 127
   VSRGB(5).Value = 32
   VSRGB(6).Value = 0
   VSRGB(7).Value = 50
   VSRGB(8).Value = 5
   VSRGB(9).Value = 2
   VSRGB(10).Value = 0
   VSRGB(11).Value = 2
   aRGB = True
End Sub


Private Sub mnuOptions_Click(Index As Integer)
   Select Case Index
   Case 0   ' Continuous Scrollbar Effects
      aEffAction = True
      mnuOptions(0).Checked = True
      mnuOptions(1).Checked = False
      LabContStep = "Continuous"
   Case 1   ' Stepped      "         "
      aEffAction = False
      mnuOptions(0).Checked = False
      mnuOptions(1).Checked = True
      LabContStep = "Stepped"
   End Select
End Sub


Private Sub VSRGB_Scroll(Index As Integer)
' Thumb changes value on mouse_up
   Call VSRGB_Change(Index)
End Sub

Private Sub VSRGB_Change(Index As Integer)
Dim Par As Long
Dim zPar As Single

   If Not aRGB Then
      VSRGB(Index).Refresh
      Picture1.SetFocus
      Exit Sub
   End If
   
   If Not aPicLoaded Then
      Picture1.SetFocus
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      VSRGB(Index).Refresh
      Picture1.SetFocus
      Exit Sub
   End If
   
   optEff_MouseMove Index, 0, 0, 0, 0
   
   If aEffAction Then
      Screen.MousePointer = vbHourglass
   End If
   
   Select Case Index
   Case 0 ' BL             0 -> 20
      ' Index O-B  1-G  2-R
      zPar = CSng(VSRGB(Index).Value) / 4  '0.0 -> 5.0
      LabEffVals = "BL " & Str$(zPar * 20)
      If aEffAction Then VaryRGB Index, zPar
   Case 1 ' GR             0 -> 20
      ' Index O-B  1-G  2-R
      zPar = CSng(VSRGB(Index).Value) / 4  '0.0 -> 5.0
      LabEffVals = "GR " & Str$(zPar * 20)
      If aEffAction Then VaryRGB Index, zPar
   Case 2 ' RE             0 -> 20
      ' Index O-B  1-G  2-R
      zPar = CSng(VSRGB(Index).Value) / 4  '0.0 -> 5.0
      LabEffVals = "RE " & Str$(zPar * 20)
      If aEffAction Then VaryRGB Index, zPar
   Case 3  ' BR Brightness           0 -> 20
      zPar = CSng(VSRGB(3).Value) / 4      '0.0 -> 5.0
      LabEffVals = "BR " & Str$(zPar * 20)
      If aEffAction Then Brightness zPar
   
   Case 4   ' BW Black & White       1 -> 255
      Par = CLng(VSRGB(4).Value)
      LabEffVals = "BW " & Str$(Par)
      If aEffAction Then BlackWhite Par
   
   Case 5   ' DI Black & White Dither  16 -> 48
      zPar = CSng(VSRGB(5).Value)
      LabEffVals = "DI " & Str$(zPar)
      If aEffAction Then BlackWhiteDither zPar
   
   Case 6   ' CO Contrast          -66 -> +66
      Par = CLng(VSRGB(6).Value)
      LabEffVals = "CO " & Str$(Par)
      If aEffAction Then Contrast Par
   Case 7   ' OU Outline             0 -> 100
      Par = CLng(VSRGB(7).Value)
      LabEffVals = "OU " & Str$(Par)
      If aEffAction Then OutLine Par
   Case 8   ' SS Sharp-Smooth  765-4321
            '                  123-4321
      Par = CLng(VSRGB(8).Value)
      LabEffVals = "SS " & Str$(Par - 4)
      If aEffAction Then SharpSmooth Par
   Case 9   ' DF Diffuse            1  ->  16
      Par = CLng(VSRGB(9).Value)
      LabEffVals = "DF " & Str$(Par)
      If aEffAction Then Diffuse Par
   Case 10  ' EE Emboss Engrave     -3 -> +3
      Par = CLng(VSRGB(10).Value)
      LabEffVals = "EE " & Str$(Par)
      If aEffAction Then EmbossEngrave Par
   Case 11  ' ME Melt                -8 -> +8
      Par = CLng(VSRGB(11).Value)
      LabEffVals = "ME " & Str$(Par)
      If aEffAction Then Melt Par
   End Select
   DISPLAY
   Screen.MousePointer = vbDefault
End Sub
'#### END EFFECTS ####

Private Sub cmdReload_Click()
' Reload last opened file
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      LabFilter = "Filters"
      Exit Sub
   End If
   LoadThePicture FileSpec$, PicDataORG()
   ReDim PicData(0 To 3, 0 To W - 1, 0 To H - 1)
   PicData() = PicDataORG()
   SetVSRGBValues
   DISPLAY
End Sub
Private Sub cmdReload_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabFilter = ""
End Sub

Private Sub cmdReset_Click()
' Reset current starting image
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      LabFilter = "Filters"
      Exit Sub
   End If
   PicData() = PicDataORG()
   SetVSRGBValues
   DISPLAY
End Sub
Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabFilter = ""
End Sub

Private Sub cmdFix_Click()
' Fix as New Image
   If Not aPicLoaded Then
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      LabFilter = "Filters"
      Exit Sub
   End If
   PicDataORG() = PicData()
   SetVSRGBValues
End Sub
Private Sub cmdFix_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   LabFilter = ""
End Sub

Private Sub mnuInfo_Click()
   MsgBox "Shows file information", vbInformation, "m-IP by Robert Rayment"
End Sub

'#### FILE STUFF ####
Private Sub FileOps_Click(Index As Integer)
Dim Title$, Filt$, Indir$
Dim FIndex As Long
Dim A$
Dim k As Long
   
   Screen.MousePointer = vbHourglass
   
   Select Case Index
   Case 0   ' Open
      Title$ = "Open a picture file"
      Filt$ = "Pics bmp,jpg,gif|*.bmp;*.jpg;*.gif"
      FileSpec$ = ""
      Indir$ = CurrentPath$ 'Pathspec$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, Indir$, "", Me.hwnd, FIndex
   '   FIndex = 1 bmp
   '   FIndex = 2 jpg
   '   etc
      Set CommonDialog1 = Nothing
      ' Avoid click through
      SetCursorPos Form1.Left \ STX + 130, Form1.Top \ STY + 150
      
      If Len(FileSpec$) = 0 Then
         ' Cancel or Error
         Screen.MousePointer = vbDefault
         LabFilter = "Filters"
         Exit Sub
      End If
      CurrentPath$ = GetPath(FileSpec$)
      Picture1.Picture = LoadPicture
      Picture1.Cls
      Picture1.Refresh
      If Not LoadThePicture(FileSpec$, PicDataORG()) Then Exit Sub
            
      If W <= 2 Or H <= 2 Then
         MsgBox "Picture size too small ie <= 2x2 ", vbCritical, "m-IP"
         W = 20
         H = 20
         ReDim PicDataORG(1, W - 1, H - 1)
         ReDim PicData(1, W - 1, H - 1)
         mnuInfo.Caption = ""
         aPicLoaded = True    ' To avoid extra message box
         cmdSelect_Click (1)  ' Selection off
         SetPicBox
         aScroll = True
         aZoom = True
         scrZoom.Value = 4  ' 4-3=1
         scrZoom_Change     ' To remove any scrollbars
         Picture1.BackColor = RGB(224, 224, 224)
         FileSpec$ = ""
         aPicLoaded = False
         Exit Sub
      End If
      
      aPicLoaded = True
      ReDim PicData(0 To 3, 0 To W - 1, 0 To H - 1)
      PicData() = PicDataORG()
      
      SetPicBox
      aScroll = True
      aZoom = True
      xlonew = 0
      ylonew = 0
      xlo = 0
      ylo = 0

      scrZoom.Value = 4  ' 4-3=1
      scrZoom_Change
      
      A$ = " " & GetFileName(FileSpec$) & " "
      A$ = A$ & "WxH =" & Str$(W) & " x" & Str$(H) & " "
      A$ = A$ & Str$(FileLen(FileSpec$)) & "  B"
      mnuInfo.Caption = A$
      cmdSelect_Click (1)  ' Selection off
      SetVSRGBValues

   Case 1   ' Break
   Case 2, 3, 4  ' Save viewport, Save whole image, Save selection
      If Not aPicLoaded Then
         Screen.MousePointer = vbDefault
         LabFilter = "Filters"
         MsgBox "No picture loaded yet", vbInformation, "m-IP"
         Exit Sub
      End If
      Select Case Index
      Case 2
         Title$ = "Save displayed viewport (BMP or JPEG)"
         Filt$ = "Pics bmp|*.bmp|Pics jpeg|*.jpg"
      Case 3
         Title$ = "Save whole image (BMP or JPEG)"
         Filt$ = "Pics bmp|*.bmp|Pics jpeg|*.jpg"
      Case 4
         Title$ = "Save selection (BMP or JPEG)"
         If shpSelect.Width < 4 Or shpSelect.Height < 4 Then
            MsgBox "Selection too small (ie < 4)", vbCritical, "m-IP"
            Screen.MousePointer = vbDefault
            Picture1.MousePointer = vbCrosshair
            Exit Sub
         End If
         Filt$ = "Pics bmp|*.bmp|Pics jpeg|*.jpg"
      End Select
      
      SaveSpec$ = ""
      Indir$ = SavePath$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, Indir$, "", Me.hwnd, FIndex
      Set CommonDialog1 = Nothing
      ' Avoid click through
      SetCursorPos Form1.Left \ STX + 130, Form1.Top \ STY + 150
      
      If Len(SaveSpec$) = 0 Then
         Screen.MousePointer = vbDefault
         LabFilter = "Filters"
         Exit Sub
      End If
      '   FIndex = 1 bmp
      '   FIndex = 2 jpg
      
      Select Case Index
      Case 2, 3, 4
         If FIndex = 1 Then
            FixExtension SaveSpec$, ".bmp"
         Else
            FixExtension SaveSpec$, ".jpg"
         End If
      End Select
      
      SavePath$ = GetPath(SaveSpec$)
      
      Select Case Index
      Case 2
         If FIndex = 1 Then         ' Save viewport, bmp
            Picture1.Picture = Picture1.Image
            SavePicture Picture1.Image, SaveSpec$
         Else                       ' Save viewport, jpg
            Picture1.Picture = Picture1.Image
            SaveHDCJPEG SaveSpec$, Picture1.HDC, wp, hp
         End If
      Case 3
         If FIndex = 1 Then         ' Save whole image, bmp
            If Not SaveBMP24(SaveSpec$, PicData(), W, H) Then
               MsgBox "Saving whole image - failed", vbCritical, "m-IP"
            End If
         Else                       ' Save whole image, jpg
            SaveDATAJPEG SaveSpec$, PicData(), W, H, 0, 0
         End If
      Case 4
         picSelect.Width = shpSelect.Width
         picSelect.Height = shpSelect.Height
         Call BitBlt(picSelect.HDC, 0, 0, shpSelect.Width, shpSelect.Height, _
              Picture1.HDC, xp1, yp1, vbSrcCopy)
         picSelect.Picture = picSelect.Image
         If FIndex = 1 Then         ' Save selection, bmp
            SavePicture picSelect.Image, SaveSpec$
         Else                       ' Save selection, jpg
            SaveHDCJPEG SaveSpec$, picSelect.HDC, shpSelect.Width, shpSelect.Height
         End If
      End Select
      picSelect.Picture = LoadPicture
      picSelect.Width = 4
      picSelect.Height = 4
   Case 5   ' Break
   Case 6   ' Exit
      Unload Me     ' >> Form_Unload
   End Select
   LabFilter = "Filters"
   Screen.MousePointer = vbDefault
End Sub

Private Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   p = InStr(1, FSpec$, ".")
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub

Private Function GetFileName(FSpec$) As String
' VB5 also
Dim L As Long
Dim k As Long
   GetFileName = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k = 0 Then
      GetFileName = FSpec$
   Else
      GetFileName = Right$(FSpec$, L - k)
   End If
End Function

Private Function FileExists(FSpec$) As Boolean
   If Dir(FSpec$) <> "" Then FileExists = True
End Function

Private Function GetPath(FSpec$) As String
' VB5 also
Dim L As Long
Dim k As Long
   GetPath = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k <> 0 Then
      GetPath = Left$(FSpec$, k)  ' NB includes last \
   End If
End Function
'#### END FILE STUFF ####

'#### FORM STUFF ####
Private Sub Form_Initialize()
   m_hMod = LoadLibrary("shell32.dll")
   InitCommonControls
End Sub

Private Sub Form_Load()
Dim k As Long
    If App.LogMode <> 1 Then
        MsgBox "Much faster when compiled", vbExclamation, "m-IP"
    End If

   On Error Resume Next
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   ' Hard code some initial sizes
   Me.Height = 8145
   Me.Width = 8985
   FWidORG = Me.Width
   FHitORG = Me.Height
   RightMargin = 255 + 240
   BottomMargin = 1200 + 30
   Me.BackColor = &H808080
   LabContStep.BackColor = RGB(100, 130, 220)
   LabContStep = "Continuous"
   Zoom = 1
   Show     ' >> Form_Resize
   
   picPB.DrawWidth = 3
   LineY.Visible = False
   LineX.Visible = False
   Zoom = 4  ' actual zoom = 4-3 = 1
   aZoom = False
   HScroll1.Min = 0
   HScroll1.Visible = False
   VScroll1.Min = 0
   VScroll1.Visible = False
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   SavePath$ = PathSpec$
   
   aScroll = False
   scrZoom.Left = Picture1.Left - 5
   LabZoom.Left = scrZoom.Left + scrZoom.Width + 1
   scrZoom.Min = 1   '  1/4 = 0.25 min
   scrZoom.Max = 23  ' 23-3 = 20   max
   scrZoom.Value = 4 ' actual zoom = 4-3
   aScroll = True
   aRGB = False
   ' Position effects scrollbars
   For k = 0 To 11
      VSRGB(k).Left = optEff(k).Left + 4
   Next k
   SetFilterLimits
   SetVSRGBValues
   aRGB = True
   aPicLoaded = False
   mnuOptions(0).Checked = True
   mnuOptions(1).Checked = False
   aEffAction = True
   
   aSelect = False
   LabDims = " Select OFF "
   shpSelect.Visible = False
   FileOps(4).Enabled = False
   
   LabZoom = "<- ZOOM ->"
   Caption = "mini-Image Processor"
   On Error GoTo 0
End Sub

Private Sub Form_Resize()
   If WindowState <> vbMinimized Then
      
      If Me.Width < FWidORG Then Me.Width = FWidORG
      If Me.Height < FHitORG Then Me.Height = FHitORG
      
      SetPicBox
      Surround
      If aPicLoaded Then
         Zoomer
      End If
      LabZoom = "Z=" & Str$(Zoom) & "  wxh =" & Str$(Int(Picture1.Width)) & "x" & LTrim$(Str$(Int(Picture1.Height)))

   End If
End Sub

Private Sub Surround()
Dim y As Long
   y = cmdSelect(0).Top + cmdSelect(0).Height + 3
   Me.Line (0, 0)-(Picture1.Left - 6, Me.Height), RGB(100, 130, 220), BF
   Me.Line (Picture1.Left - 5, y + 2)-(Picture1.Left - 5, Me.Height / STY)
   Me.Line (Picture1.Left - 6, y + 2)-(Picture1.Left - 6, Me.Height / STY), vbWhite
   
   Me.Line (Picture1.Left - 5, 0)-(Me.Width / STX, y + 1), RGB(100, 130, 220), BF
   Me.Line (Picture1.Left - 5, y + 2)-(Me.Width / STX, y + 2), vbWhite
   Me.Line (Picture1.Left - 5, y + 3)-(Me.Width / STX, y + 3)
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = vbDefault
   Erase PicDataORG()
   Erase PicData()
   FreeLibrary m_hMod
   Set Form1 = Nothing
   End
End Sub
'#### END FORM STUFF ####

'#### ZOOMING ####
Private Sub cmdZoom_Click(Index As Integer)
' Set zoom = 1
   If Not aPicLoaded Then
      Picture1.SetFocus
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Exit Sub
   End If
   Zoom = (Index + 1) * 4 - 3
   scrZoom.Value = (Index + 1) + 3 'Zoom ' actual zoom = 4-3 =1
   Picture1.SetFocus
End Sub

Private Sub scrZoom_Scroll()
' Thumb
   Call scrZoom_Change
End Sub

Private Sub scrZoom_Change()
' End buttons, body & Value
Dim TZoom As Single
   If Not aPicLoaded Then
      scrZoom.Value = 4
      Exit Sub
   End If
   If scrZoom.Value < 4 Then
      TZoom = scrZoom.Value / 4
   Else
      TZoom = scrZoom.Value - 3
   End If
   
   If TZoom * W <= 2 Or TZoom * H <= 2 Then
      Exit Sub
   End If
   
   Zoom = TZoom
   If Not aZoom Then Exit Sub
   If aPicLoaded Then
      Zoomer
   End If
   If scrZoom.Value < 4 Then
      LabZoom = "Z=" & Format(Zoom, "Fixed") & "  wxh =" & Str$(Int(Picture1.Width)) & "x" & LTrim$(Str$(Int(Picture1.Height)))
   Else
      LabZoom = "Z=" & Str$(Zoom) & "  wxh =" & Str$(Int(Picture1.Width)) & "x" & LTrim$(Str$(Int(Picture1.Height)))
   End If
End Sub

Private Sub Zoomer() ' {PictureBox, HScroll, VScroll
   aScroll = False
   Call PZoomer(Picture1, HScroll1, VScroll1)
   aScroll = True

   DISPLAY
   ' Picture1 border
   With shpBorder
      .Left = Picture1.Left - 1
      .Top = Picture1.Top - 1
      .Width = Picture1.Width + 2
      .Height = Picture1.Height + 2
   End With
   
   If HScroll1.Visible = False Or VScroll1.Visible = False Then
      cmdSelect_Click (1)  ' Turn Select off
   End If
End Sub
'#### END ZOOMING ####


Private Sub SetPicBox()
   Picture1.Width = (Me.Width - RightMargin) \ STX - Picture1.Left
   Picture1.Height = (Me.Height - BottomMargin) \ STY - Picture1.Top
   CurrPicWID = Picture1.Width
   CurrPicHIT = Picture1.Height
   wp = Picture1.Width
   hp = Picture1.Height
End Sub


'#### MOUSE ON PICBOX & SELECT ####

Private Sub cmdSelect_Click(Index As Integer)
   If Not aPicLoaded Then
      Screen.MousePointer = vbDefault
      MsgBox "No picture loaded yet", vbInformation, "m-IP"
      Picture1.SetFocus
      Exit Sub
   End If

   If Index = 0 Then
      aSelect = True
      LabDims = " Select ON "
      FileOps(4).Enabled = True
      Picture1.MousePointer = vbCrosshair
   Else
      aSelect = False
      LabDims = " Select OFF "
      shpSelect.Visible = False
      FileOps(4).Enabled = False
   End If
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Public XTL As Long, YTL As Long ' Top left coords in PicData()
' Public xp1 As Long, yp1 As Long ' MouseDown x,y
   If Not aPicLoaded Then Exit Sub
   If aSelect Then
      Picture1.MousePointer = vbCrosshair
      LabDims = " X,Y " & Str$(x) & "," & Str$(y) & " "
      If Button = vbLeftButton Then
         aMouseDown = True
         With shpSelect
            .Visible = True
            .Left = x
            .Top = y
            .Width = 1
            .Height = 1
         End With
         XTL = xlo
         YTL = ylo + hp / Zoom
         xp1 = x
         yp1 = y
      End If
   Else
      Screen.MousePointer = vbDefault
      SetCursor LoadCursor(0, 32649&)     ' Hand
      If HScroll1.Visible Or VScroll1.Visible Then
         If Button = vbLeftButton Then
            aMouseDown = True
            XTL = xlo
            YTL = ylo + hp / Zoom 'yhi
            xp1 = x
            yp1 = y
         Else
            aMouseDown = False
         End If
      Else
         aMouseDown = False
      End If
   End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not aPicLoaded Then Exit Sub
   If aSelect Then
      Picture1.MousePointer = vbCrosshair
      LabDims = " X,Y " & Str$(x) & "," & Str$(y) & " "
      If Button = vbLeftButton And aMouseDown Then
         If x > xp1 And y > yp1 Then   ' No negatives allowed
         LabDims = " WxH " & Str$(shpSelect.Width) & " x" & Str$(shpSelect.Height) & " "
            If x < wp And y < hp Then  ' Viewport limits
               shpSelect.Width = (x - xp1)
               shpSelect.Height = (y - yp1)
               LabDims = " WxH " & Str$(shpSelect.Width) & " x" & Str$(shpSelect.Height) & " "
            End If
         End If
      End If
   Else
      SetCursor LoadCursor(0, 32649&)     ' Hand
      aScroll = False
      If Button = vbLeftButton And aMouseDown Then
      
         Call MouseMoveCalcs(x, y, HScroll1, VScroll1)
      
         DISPLAY
      End If
      aScroll = True
   End If
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not aPicLoaded Then Exit Sub
   aMouseDown = False
End Sub
'#### END MOUSE ON PICBOX & SELECT ####


'#### PICTURE SCROLLBARS ####
Private Sub HScroll1_Scroll()
' Thumb
   Call HScroll1_Change
End Sub

Private Sub HScroll1_Change()
' End buttons, body & Value
   xlo = HScroll1.Value
   If Not aScroll Then Exit Sub
   DISPLAY
End Sub

Private Sub VScroll1_Scroll()
' Thumb
   Call VScroll1_Change
End Sub

Private Sub VScroll1_Change()
' End buttons, body & Value
   ylo = VScroll1.Value
   yhi = ylo + hp / Zoom
   If Not aScroll Then Exit Sub
   DISPLAY
End Sub
'#### END PICTURE SCROLLBARS ####

Private Sub DISPLAY()    '<<<<<<<<<<
' Public BHI As BITMAPINFOHEADER
   
   Call CrossHairs(Picture1, HScroll1, VScroll1, LineX, LineY)
   
   SetStretchBltMode Picture1.HDC, HALFTONE
   ' Would need COLORONCOLOR for setting pixels
   ' if used in a paint program.
   
   Call StretchDIBits(Picture1.HDC, _
   0, 0, _
   wp, hp, _
   xlo, ylo, _
   wp / Zoom, hp / Zoom, _
   PicData(0, 0, 0), _
   BHI, 0, vbSrcCopy)
   
   Picture1.Refresh
End Sub

Private Sub mnuHelp_Click()
Dim A$, C$
   C$ = vbCrLf
   A$ = ""
   A$ = A$ & "m-IP by Robert Rayment" & C$ & C$
   A$ = A$ & "Filters Buttons:-" & C$
   A$ = A$ & " BL  Vary blue             GR  Vary green           RE  Vary red" & C$
   A$ = A$ & " BR  Vary brightness    BW  Black & White      DI  Black & White Dither  " & C$
   A$ = A$ & " CO  Contrast              OU  Outline                SS  Sharp-Smooth" & C$
   A$ = A$ & " DF  Diffuse                 EE  Emboss-Engrave   ME  Melt" & C$
   A$ = A$ & "These buttons run default and also stepped operations." & C$ & C$
   A$ = A$ & "Saved images are 24bpp BMP or JPEG for the viewport, whole" & C$
   A$ = A$ & "image and a selection.  A selection cannot be larger than the" & C$
   A$ = A$ & "viewport.  For example with a 1024x768 screen the maximum" & C$
   A$ = A$ & "size of the saved viewport will be about 840x630." & C$
   A$ = A$ & "The viewport can be larger or smaller than the whole image." & C$
   A$ = A$ & C$
   A$ = A$ & "When both cross-wires show, zooming will be roughly" & C$
   A$ = A$ & "at the cross-over point.  So, to zoom in on a point on" & C$
   A$ = A$ & "the picture, move it to the cross-over." & C$
   A$ = A$ & "Zoom sizes are in 1/4, 1/2, 3/4, 1, 2,,,20." & C$
   A$ = A$ & "Z shows the zoom multiplier & wxh is the viewport size." & C$
   A$ = A$ & C$
   A$ = A$ & "Multiple effects can be done by pressing the Fix button" & C$
   A$ = A$ & "between using the effects." & C$
   MsgBox A$, vbInformation, "m-IP Help"
End Sub


