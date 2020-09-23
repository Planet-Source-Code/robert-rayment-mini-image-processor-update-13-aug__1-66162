Attribute VB_Name = "ModZoomScroll"
' ModZoomScroll.bas

Option Explicit

Public XTL As Long, YTL As Long ' Top left coords in PicData()
Public xp1 As Long, yp1 As Long ' MouseDown x,y
Public wp As Long, hp As Long   ' Picture1 width & height
Public xlo As Long, xhi As Long ' H Scroll bar values
Public ylo As Long, yhi As Long ' V Scroll bar values
Public Zoom As Single           ' Zoom value 1 - 23 = .25,.5,.75,1,2,3,,20

Public PicDataORG() As Byte     ' 32bpp image data
Public PicData() As Byte        ' 32bpp image data

Public aPicLoaded As Boolean

' Relative position of scroll bar thumbs
Public xlonew As Single
Public ylonew As Single

Public CurrPicWID As Long       ' Current PIC W & H
Public CurrPicHIT As Long
Public STX As Long, STY As Long ' Screen.TwipsPerPixelX/Y


Public Sub PZoomer(PIC As PictureBox, HS As HScrollBar, VS As VScrollBar)
' Form1 refs
' PIC    Picture1
' HS,VS  HScroll1, VScroll1
   
   If wp > CurrPicWID Or W * Zoom > CurrPicWID Then
      PIC.Width = CurrPicWID
      wp = PIC.Width
      xlonew = 0
      If HS.Max <> 0 Then
         xlonew = xlo / HS.Max '<<<<<<<<<<
      End If
      
      With HS
         .Visible = True
         .Left = PIC.Left
         .Width = PIC.Width
         '.Top = PIC.Top + PIC.Height + 2
         .Max = W - wp / Zoom  'maxiHorz
         .Min = 0
         .Value = 0   ' -> HS_Change

         If xlonew <> 0 Then
            .Value = xlonew * (HS.Max) ' <<<<<<<<<<
         End If
         
      End With
   Else
      HS.Visible = False
      If aPicLoaded Then
         PIC.Width = W * Zoom
         wp = PIC.Width
         HS.Value = 0
      End If
   End If
   
   If hp > CurrPicHIT Or H * Zoom > CurrPicHIT Then
      PIC.Height = CurrPicHIT
      hp = PIC.Height
      ylonew = 0
      If VS.Min <> 0 Then
         ylonew = ylo / VS.Min '<<<<<<<<<<
      End If
      
      With VS
         .Visible = True
         .Top = PIC.Top
         .Height = PIC.Height
         .Left = PIC.Left + PIC.Width + 4
         ' NB max/min reversed
         .Min = H - hp / Zoom  'maxiVert
         .Max = 0
         
         .Value = .Min  'maxiVert   ' -> VS_Change
         
         If ylonew <> 0 Then
            .Value = ylonew * (VS.Min) ' <<<<<<<<<<
         End If
      
      End With
   Else
      VS.Visible = False
      VS.Value = 0
      If aPicLoaded Then
         PIC.Height = H * Zoom
         hp = PIC.Height
      End If
   End If
   
   If HS.Visible Then
      HS.Top = PIC.Top + PIC.Height + 4
   End If
End Sub

Public Sub MouseMoveCalcs(x As Single, y As Single, HS As HScrollBar, VS As VScrollBar)
' Form1 refs
' X,Y cursor on Picture1
' HS,VS  HScroll1, VScroll1
Dim dx As Long, dy As Long

   If HS.Visible Then
      dx = x - xp1 + 1
      xlo = XTL - dx / Zoom
      If xlo < 0 Then
         xlo = 0
      End If
      xhi = xlo + wp / Zoom
      If xhi > W Then
         xhi = W
         xlo = W - wp / Zoom
      End If
      If xlo > HS.Max Then
         xlo = HS.Max
      End If
      HS.Value = xlo   ' -> HS_Change on Form1
   End If

   If VS.Visible Then
      dy = y - yp1 + 1
      yhi = YTL + dy / Zoom
      If yhi > H Then
         yhi = H
      End If
      ylo = yhi - hp / Zoom
      If ylo < 0 Then
         ylo = 0
         yhi = hp / Zoom
      End If
      If ylo > VS.Min Then
         ylo = VS.Min
      End If
      VS.Value = ylo   ' -> VS_Change on Form1
   End If
End Sub

Public Sub CrossHairs(PIC As PictureBox, HS As HScrollBar, VS As VScrollBar, Lx As Line, LY As Line)
' Form1 refs
' PIC    Picture1
' HS,VS  HScroll1, VScroll1
' LX,LY  LineX, LineY

   If HS.Visible Then
      Lx.Visible = True
      If (HS.Max - HS.Min) <> 0 Then
         xlonew = xlo / (HS.Max - HS.Min)
      Else
         xlonew = 0
      End If
      With Lx
         .X1 = wp * xlonew
         If xlo = HS.Max Then .X1 = .X1 - 1
         .X2 = .X1
         .Y1 = 0
         .Y2 = PIC.Height
      End With
   Else
      Lx.Visible = False
   End If
   If VS.Visible Then
      LY.Visible = True
      If (VS.Min - VS.Max) <> 0 Then
         ylonew = ylo / (VS.Min - VS.Max)
      Else
         ylonew = 0
      End If
      With LY
         .X1 = 0
         .X2 = PIC.Width
         .Y1 = hp - hp * ylonew
         If ylo = 0 Then .Y1 = .Y1 - 1
         .Y2 = .Y1
      End With
   Else
      LY.Visible = False
   End If
End Sub
