Attribute VB_Name = "ModFilters"
' ModFilters.bas

' Some filters modified, adapted from various refs:
' Manuel Santos   PSC CodeId=26303
' Tanner Helland  PSC CodeId=25819
' APaint8         PSC CodeID=53558
' etc


' Note: Input:  Public W & H, image width & height
'               Public PicDataORG(0 To 3, 0 To W - 1, 0 To H - 1)
'               the original image data.
'       Output: Public PicData(0 To 3, 0 To W - 1, 0 To H - 1)
'               which is DISPLAYED in Form1.

Option Explicit

Private Cul As Long
Private ix As Long, iy As Long
Private i As Long, j As Long
Private k As Long
Private r As Long, g As Long, B As Long

Private PicData2() As Byte

Private Intens() As Integer

'<<<< Progress bar >>>>>>>>>>
Private xd As Single, xp As Single

Public Sub Invert()
' NB reversible
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      PicData(0, ix, iy) = Not PicData(0, ix, iy)
      PicData(1, ix, iy) = Not PicData(1, ix, iy)
      PicData(2, ix, iy) = Not PicData(2, ix, iy)
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub Grey()
Dim s As Long
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      B = PicDataORG(0, ix, iy)
      g = PicDataORG(1, ix, iy)
      r = PicDataORG(2, ix, iy)
      s = (r + g + B) \ 3     ' Good enough !?
      PicData(0, ix, iy) = s
      PicData(1, ix, iy) = s
      PicData(2, ix, iy) = s
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub Sepia()
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      B = PicDataORG(0, ix, iy) * 0.114
      g = PicDataORG(1, ix, iy) * 0.587
      r = PicDataORG(2, ix, iy) * 0.299
      k = B + g + r
      B = k
      g = k
      r = k
      PicData(1, ix, iy) = g
      If r < 63 Then
         r = r * 1.1
         B = B * 0.9
      End If
      If r > 62 And r < 192 Then
         r = r * 1.15
         B = B * 0.85
      End If
      If r > 191 Then
         r = r * 1.08
         If r > 255 Then
            r = 255
         End If
         B = B * 0.93
      End If
      PicData(0, ix, iy) = B
      PicData(2, ix, iy) = r
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub VaryRGB(Index As Integer, Frac As Single)
' Index 0,1,2 - B,G,R   ' Frac 0.0 -> 10.0
   'PicData() = PicDataORG()
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      Cul = 1& * PicDataORG(Index, ix, iy) * Frac
      If Cul > 255 Then Cul = 255
      PicData(Index, ix, iy) = Cul
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub Brightness(Frac As Single)
' Frac 0.0 -> 10.0
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      For k = 0 To 3
         Cul = 1& * PicDataORG(k, ix, iy) * Frac
         If Cul > 255 Then Cul = 255
         PicData(k, ix, iy) = Cul
      Next k
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub BlackWhite(BWLim As Long)
' BWLim  1 -> 255
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      Cul = (1& * PicDataORG(0, ix, iy) + _
                  PicDataORG(1, ix, iy) + _
                  PicDataORG(2, ix, iy)) \ 3
      If Cul < BWLim Then
         PicData(0, ix, iy) = 0
         PicData(1, ix, iy) = 0
         PicData(2, ix, iy) = 0
      Else
         PicData(0, ix, iy) = 255
         PicData(1, ix, iy) = 255
         PicData(2, ix, iy) = 255
      End If
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub BlackWhiteDither(zDiv As Single)
' zDiv = 16 -> 48   Floyd-Steinberg B
' Spreader
' 0 7 0
' 3 5 1 /16
Dim greysum As Long
Dim greycount As Long
Dim zMul As Single
Dim zErr As Single
   ReDim Intens(-1 To W + 2, -1 To H + 2)
   zMul = 1 / zDiv
   greysum = 0
   greycount = W * H
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      Intens(ix, iy) = (1& * PicDataORG(0, ix, iy) + _
                             PicDataORG(1, ix, iy) + _
                             PicDataORG(2, ix, iy)) \ 3
      greysum = greysum + Intens(ix, iy)
   Next ix
   Next iy
   greysum = greysum \ greycount
   
   ReDim PicData(0 To 3, 0 To W - 1, 0 To H - 1)
   StartProgress (H - 1)
   For iy = 1 To H - 1
   For ix = 1 To W - 1
      If Intens(ix, iy) > greysum Then
         PicData(0, ix, iy) = 255
         PicData(1, ix, iy) = 255
         PicData(2, ix, iy) = 255
         zErr = (Intens(ix, iy) - 255) * zMul
      Else
         zErr = Intens(ix, iy) * zMul
      End If
      ' Spread error
      Intens(ix - 1, iy + 1) = Intens(ix - 1, iy + 1) + 3 * zErr
      Intens(ix, iy + 1) = Intens(ix, iy + 1) + 5 * zErr
      Intens(ix + 1, iy + 1) = Intens(ix + 1, iy + 1) + zErr
      Intens(ix + 1, iy) = Intens(ix + 1, iy) + 7 * zErr
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
   Erase Intens()
End Sub

Public Sub Contrast(ContrastVal As Long)
' ContrastVal -66 -> 66
' Based on Tanner Helland
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      B = PicDataORG(0, ix, iy)
      g = PicDataORG(1, ix, iy)
      r = PicDataORG(2, ix, iy)
      r = r + (((r - 127) * ContrastVal) \ 100)
      g = g + (((g - 127) * ContrastVal) \ 100)
      B = B + (((B - 127) * ContrastVal) \ 100)
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If g < 0 Then g = 0
      If g > 255 Then g = 255
      If r < 0 Then r = 0
      If r > 255 Then r = 255
      PicData(0, ix, iy) = B
      PicData(1, ix, iy) = g
      PicData(2, ix, iy) = r
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub OutLine(OutLineVal As Long)
' OutLineVal 0 -> 100
' 1 1 1
' 1 0 1
' 1 1 1   8*(i,j)-Sum
   ReDim PicData(0 To 3, 0 To W - 1, 0 To H - 1)
   StartProgress (H - 2)
   For iy = 1 To H - 2
   For ix = 1 To W - 2
      B = 0: g = 0: r = 0
      For j = iy - 1 To iy + 1
      For i = ix - 1 To ix + 1
         If (j <> iy Or i <> ix) Then
            B = B + PicDataORG(0, i, j)
            g = g + PicDataORG(1, i, j)
            r = r + PicDataORG(2, i, j)
         End If
      Next i
      Next j
      B = 8 * PicDataORG(0, ix, iy) - B
      g = 8 * PicDataORG(1, ix, iy) - g
      r = 8 * PicDataORG(2, ix, iy) - r
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If g < 0 Then g = 0
      If g > 255 Then g = 255
      If r < 0 Then r = 0
      If r > 255 Then r = 255
      If (r + g + B) \ 3 > OutLineVal Then
         PicData(0, ix, iy) = 255
         PicData(1, ix, iy) = 255
         PicData(2, ix, iy) = 255
      End If
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub SharpSmooth(SSNum As Long)
Dim n As Long
' SSNum          765-4321
' Sharp-Smooth   123-4321

   If SSNum >= 5 Then  ' SHARP
   ' SHARP
   '-2 -2 -2
   '-2 26 -2
   '-2 -2 -2
      
      SSNum = SSNum - 4 ' 1,2,3
      PicData() = PicDataORG()
      StartProgress (SSNum * (H - 2))
      For n = 1 To SSNum
         For iy = 1 To H - 2
         For ix = 1 To W - 2
            B = 0: g = 0: r = 0
            For j = iy - 1 To iy + 1
            For i = ix - 1 To ix + 1
               If (j <> iy Or i <> ix) Then
                  B = B - 2 * PicData(0, i, j)
                  g = g - 2 * PicData(1, i, j)
                  r = r - 2 * PicData(2, i, j)
               Else
                  B = B + 26 * PicData(0, i, j)
                  g = g + 26 * PicData(1, i, j)
                  r = r + 26 * PicData(2, i, j)
               End If
            Next i
            Next j
            B = B \ 10
            r = r \ 10
            g = g \ 10
            If B < 0 Then B = 0
            If B > 255 Then B = 255
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            PicData(0, ix, iy) = B
            PicData(1, ix, iy) = g
            PicData(2, ix, iy) = r
         Next ix
         DrawProgress
         Next iy
      Next n
      
   
   Else  ' SMOOTH
      ' Blur 1234
      '  1      2      3        4
      '                     1 1 1 1 1
      '0 0 0  0 1 0  1 1 1  1 1 1 1 1
      '1 0 1  1 0 1  1 0 1  1 1 0 1 1
      '0 0 0  0 1 0  1 1 1  1 1 1 1 1
      '                     1 1 1 1 1
      ' Smooths add in (ix,iy)
      
      If SSNum = 5 Then SSNum = 4
      SSNum = 5 - SSNum ' So top of scroll bar is most blurred
      Select Case SSNum
      Case 1   ' H blur
         StartProgress (H - 2)
         For iy = 1 To H - 2
         For ix = 1 To W - 2
            B = 0: g = 0: r = 0
            
            B = B + PicDataORG(0, ix, iy)
            g = g + PicDataORG(1, ix, iy)
            r = r + PicDataORG(2, ix, iy)
            
            B = B + PicDataORG(0, ix - 1, iy)
            g = g + PicDataORG(1, ix - 1, iy)
            r = r + PicDataORG(2, ix - 1, iy)
            
            B = B + PicDataORG(0, ix + 1, iy)
            g = g + PicDataORG(1, ix + 1, iy)
            r = r + PicDataORG(2, ix + 1, iy)
            
            B = B \ 3
            g = g \ 3
            r = r \ 3
            PicData(0, ix, iy) = B
            PicData(1, ix, iy) = g
            PicData(2, ix, iy) = r
         Next ix
         DrawProgress
         Next iy
      Case 2   ' Cross blur
         StartProgress (H - 2)
         For iy = 1 To H - 2
         For ix = 1 To W - 2
            B = 0: g = 0: r = 0
            
            B = B + PicDataORG(0, ix, iy)
            g = g + PicDataORG(1, ix, iy)
            r = r + PicDataORG(2, ix, iy)
            
            B = B + PicDataORG(0, ix, iy - 1)
            g = g + PicDataORG(1, ix, iy - 1)
            r = r + PicDataORG(2, ix, iy - 1)
            
            B = B + PicDataORG(0, ix, iy + 1)
            g = g + PicDataORG(1, ix, iy + 1)
            r = r + PicDataORG(2, ix, iy + 1)
            
            B = B + PicDataORG(0, ix - 1, iy)
            g = g + PicDataORG(1, ix - 1, iy)
            r = r + PicDataORG(2, ix - 1, iy)
            
            B = B + PicDataORG(0, ix + 1, iy)
            g = g + PicDataORG(1, ix + 1, iy)
            r = r + PicDataORG(2, ix + 1, iy)
            
            B = B \ 5
            g = g \ 5
            r = r \ 5
            PicData(0, ix, iy) = B
            PicData(1, ix, iy) = g
            PicData(2, ix, iy) = r
         Next ix
         DrawProgress
         Next iy
      Case 3   ' Surr blur 1
         StartProgress (H - 2)
         For iy = 1 To H - 2
         For ix = 1 To W - 2
            B = 0: g = 0: r = 0
            For j = iy - 1 To iy + 1
            For i = ix - 1 To ix + 1
               'If (j <> iy Or i <> ix) Then
                  B = B + PicDataORG(0, i, j)
                  g = g + PicDataORG(1, i, j)
                  r = r + PicDataORG(2, i, j)
               'End If
            Next i
            Next j
            B = B \ 9
            g = g \ 9
            r = r \ 9
            PicData(0, ix, iy) = B
            PicData(1, ix, iy) = g
            PicData(2, ix, iy) = r
         Next ix
         DrawProgress
         Next iy
      Case 4   ' Surr blur 2
         StartProgress (H - 4)
         For iy = 2 To H - 3
         For ix = 2 To W - 3
            B = 0: g = 0: r = 0
            For j = iy - 2 To iy + 2
            For i = ix - 2 To ix + 2
               'If (j <> iy Or i <> ix) Then
                  B = B + PicDataORG(0, i, j)
                  g = g + PicDataORG(1, i, j)
                  r = r + PicDataORG(2, i, j)
               'End If
            Next i
            Next j
            B = B \ 25
            g = g \ 25
            r = r \ 25
            PicData(0, ix, iy) = B
            PicData(1, ix, iy) = g
            PicData(2, ix, iy) = r
         Next ix
         DrawProgress
         Next iy
      End Select
      ' V Edges
      For iy = 1 To H - 2
         B = 0: g = 0: r = 0  ' Left
         B = B + PicDataORG(0, 0, iy - 1)
         g = g + PicDataORG(1, 0, iy)
         r = r + PicDataORG(2, 0, iy + 1)
         B = B \ 3
         g = g \ 3
         r = r \ 3
         PicData(0, 0, iy) = B
         PicData(1, 0, iy) = g
         PicData(2, 0, iy) = r
         
         B = 0: g = 0: r = 0  ' Right
         B = B + PicDataORG(0, W - 1, iy - 1)
         g = g + PicDataORG(1, W - 1, iy)
         r = r + PicDataORG(2, W - 1, iy + 1)
         B = B \ 3
         g = g \ 3
         r = r \ 3
         PicData(0, W - 1, iy) = B
         PicData(1, W - 1, iy) = g
         PicData(2, W - 1, iy) = r
      Next iy
      ' H Edges
      For ix = 1 To W - 2
         B = 0: g = 0: r = 0  ' Bottom
         B = B + PicDataORG(0, ix - 1, 0)
         g = g + PicDataORG(1, ix, 0)
         r = r + PicDataORG(2, ix + 1, 0)
         B = B \ 3
         g = g \ 3
         r = r \ 3
         PicData(0, ix, 0) = B
         PicData(1, ix, 0) = g
         PicData(2, ix, 0) = r
         
         B = 0: g = 0: r = 0  ' Top
         B = B + PicDataORG(0, ix - 1, H - 1)
         g = g + PicDataORG(1, ix, H - 1)
         r = r + PicDataORG(2, ix + 1, H - 1)
         B = B \ 3
         g = g \ 3
         r = r \ 3
         PicData(0, ix, iy) = B
         PicData(1, ix, iy) = g
         PicData(2, ix, iy) = r
      Next ix
   End If
   Form1.picPB.Cls
End Sub

Public Sub Diffuse(PDIFFUSE As Long)
' PDIFFUSE  1 -> 16
   ReDim PicData(0 To 3, 0 To W - 1, 0 To H - 1)
   StartProgress (H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      j = Rnd * PDIFFUSE - PDIFFUSE \ 2
      i = Rnd * PDIFFUSE - PDIFFUSE \ 2
      If ix + i < 0 Then i = 0
      If ix + i > W - 1 Then i = 0
      If iy + j < 0 Then j = 0
      If iy + j > H - 1 Then j = 0
      PicData(0, ix, iy) = PicDataORG(0, ix + i, iy + j)
      PicData(1, ix, iy) = PicDataORG(1, ix + i, iy + j)
      PicData(2, ix, iy) = PicDataORG(2, ix + i, iy + j)
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub EmbossEngrave(EEPAR As Long)
' EEPAR  -3 -> 3
'+1 0 -1    -1  0  +1
'+1 0 -1    -1  0  +1
'+1 0 -1    -1  0  +1
Dim zEEPAR As Single
   zEEPAR = EEPAR
   If EEPAR = 0 Then zEEPAR = 0.5 'EEPAR = 1
   StartProgress (H - 2)
   ReDim PicData(0 To 3, 0 To W - 1, 0 To H - 1)
   For iy = 1 To H - 2
   For ix = 1 To W - 2
      B = 0: g = 0: r = 0
      For j = iy - 1 To iy + 1
      For i = ix - 1 To ix + 1
         If i <> ix Then
            If i = ix + 1 Then
               B = B - zEEPAR * PicDataORG(0, i, j)
               g = g - zEEPAR * PicDataORG(1, i, j)
               r = r - zEEPAR * PicDataORG(2, i, j)
            Else  ' i=ix-1
               B = B + zEEPAR * PicDataORG(0, i, j)
               g = g + zEEPAR * PicDataORG(1, i, j)
               r = r + zEEPAR * PicDataORG(2, i, j)
            End If
         End If
      Next i
      Next j
      'Backcolor 0 to 255
      B = B + 210  ' Slight blue
      g = g + 200
      r = r + 200
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If g < 0 Then g = 0
      If g > 255 Then g = 255
      If r < 0 Then r = 0
      If r > 255 Then r = 255
      PicData(0, ix, iy) = B
      PicData(1, ix, iy) = g
      PicData(2, ix, iy) = r
   Next ix
   DrawProgress
   Next iy
   Form1.picPB.Cls
End Sub

Public Sub Melt(PMELT As Long)
' PMELT  ' -8 -> 8
Dim n  As Long
Dim NN As Long
Dim iylo As Long, iyhi As Long
Dim ss As Long

   If PMELT = 0 Then Exit Sub
   PicData() = PicDataORG()
   StartProgress (Abs(PMELT) * (H - 1))
   For n = 1 To Abs(PMELT)
      If PMELT > 0 Then ' Melt down
         NN = n   ' 1 to PMELT
         iylo = 0: iyhi = H - 1 - NN
         ss = 1
      Else  ' Melt up
         NN = -n  ' -PMELT to -1
         iyhi = Abs(NN): iylo = H - 1
         ss = -1
      End If
      For iy = iylo To iyhi Step ss
      For ix = 0 To W - 1
         B = PicData(0, ix, iy)
         g = PicData(1, ix, iy)
         r = PicData(2, ix, iy)
         B = B - PicData(0, ix, iy + NN)
         g = g - PicData(1, ix, iy + NN)
         r = r - PicData(2, ix, iy + NN)
         r = B + g + r
         If r < 0 Then
            PicData(0, ix, iy) = PicData(0, ix, iy + NN)
            PicData(1, ix, iy) = PicData(1, ix, iy + NN)
            PicData(2, ix, iy) = PicData(2, ix, iy + NN)
         Else
            PicData(0, ix, iy) = PicData(0, ix, iy)
            PicData(1, ix, iy) = PicData(1, ix, iy)
            PicData(2, ix, iy) = PicData(2, ix, iy)
         End If
      Next ix
      DrawProgress
      Next iy
   Next n
   Form1.picPB.Cls
End Sub

Public Sub Mirrors(Index As Integer)
' Index = 0,1,2,3  - > T,L,R,B
   PicData() = PicDataORG()
   Select Case Index
   Case 0 ' Top
      StartProgress (H \ 2)
      For ix = 0 To W - 1
      For iy = H - 1 To H \ 2 Step -1
         PicData(0, ix, H - 1 - iy) = PicDataORG(0, ix, iy)
         PicData(1, ix, H - 1 - iy) = PicDataORG(1, ix, iy)
         PicData(2, ix, H - 1 - iy) = PicDataORG(2, ix, iy)
      Next iy
      DrawProgress
      Next ix
   Case 1 ' Left
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W \ 2
         PicData(0, W - 1 - ix, iy) = PicDataORG(0, ix, iy)
         PicData(1, W - 1 - ix, iy) = PicDataORG(1, ix, iy)
         PicData(2, W - 1 - ix, iy) = PicDataORG(2, ix, iy)
      Next ix
      DrawProgress
      Next iy
   Case 2 ' Right
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = W - 1 To W \ 2 Step -1
         PicData(0, W - 1 - ix, iy) = PicDataORG(0, ix, iy)
         PicData(1, W - 1 - ix, iy) = PicDataORG(1, ix, iy)
         PicData(2, W - 1 - ix, iy) = PicDataORG(2, ix, iy)
      Next ix
      DrawProgress
      Next iy
   Case 3 ' Bottom
      StartProgress (H \ 2)
      For ix = 0 To W - 1
      For iy = 0 To H \ 2
         PicData(0, ix, H - 1 - iy) = PicDataORG(0, ix, iy)
         PicData(1, ix, H - 1 - iy) = PicDataORG(1, ix, iy)
         PicData(2, ix, H - 1 - iy) = PicDataORG(2, ix, iy)
      Next iy
      DrawProgress
      Next ix
   End Select
   Form1.picPB.Cls
End Sub

Public Sub Flips(Index As Integer)
' Index = 0,1 - > Flip Horz, Vert
   ReDim PicData2(0 To 3, 0 To W - 1, 0 To H - 1)
   PicData2() = PicData()
   Select Case Index
   Case 0 ' Flip Horz
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W - 1
         PicData(0, W - 1 - ix, iy) = PicData2(0, ix, iy)
         PicData(1, W - 1 - ix, iy) = PicData2(1, ix, iy)
         PicData(2, W - 1 - ix, iy) = PicData2(2, ix, iy)
      Next ix
      DrawProgress
      Next iy
   Case 1 ' Flip Vert
      StartProgress (W - 1)
      For ix = 0 To W - 1
      For iy = 0 To H - 1
         PicData(0, ix, H - 1 - iy) = PicData2(0, ix, iy)
         PicData(1, ix, H - 1 - iy) = PicData2(1, ix, iy)
         PicData(2, ix, H - 1 - iy) = PicData2(2, ix, iy)
      Next iy
      DrawProgress
      Next ix
   End Select
   Erase PicData2()
   Form1.picPB.Cls
End Sub


Public Sub Swaps(Index As Integer)
' Index = 0,1,2  - > BG, BR, GR
   PicData() = PicDataORG()
   Select Case Index
   Case 0 ' Swap BG
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W - 1
         PicData(0, ix, iy) = PicDataORG(1, ix, iy)
         PicData(1, ix, iy) = PicDataORG(0, ix, iy)
      Next ix
      DrawProgress
      Next iy
   Case 1 ' Swap BR
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W - 1
         PicData(0, ix, iy) = PicDataORG(2, ix, iy)
         PicData(2, ix, iy) = PicDataORG(0, ix, iy)
      Next ix
      DrawProgress
      Next iy
   Case 2 ' Swap GR
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W - 1
         PicData(1, ix, iy) = PicDataORG(2, ix, iy)
         PicData(2, ix, iy) = PicDataORG(1, ix, iy)
      Next ix
      DrawProgress
      Next iy
   End Select
   Form1.picPB.Cls
End Sub

Public Sub Flute(Index As Integer)
Dim FluteNum As Long
Dim NN As Long
Dim iyv As Long
Dim ixv As Long

   PicData() = PicDataORG()
   FluteNum = 12
   If H < W Then
      NN = H \ FluteNum
   Else
      NN = W \ FluteNum
   End If
   
   If NN < 1 Then NN = 2  ' For images < 12x12 (NB will be > 2x2 from FileOps)
   
   Select Case Index
   Case 0 ' H Flute
      StartProgress (W - 1)
      For ix = 0 To W - 1
      For iy = 0 To H - 1
         iyv = iy + (iy Mod NN) - NN \ 2
         If iyv >= 0 Then
         If iyv <= H - 1 Then
            PicData(0, ix, iy) = PicDataORG(0, ix, iyv)
            PicData(1, ix, iy) = PicDataORG(1, ix, iyv)
            PicData(2, ix, iy) = PicDataORG(2, ix, iyv)
         End If
         End If
      Next iy
      DrawProgress
      Next ix
   Case 1  ' V Flute
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W - 1
         ixv = ix + (ix Mod NN) - NN \ 2
         If ixv >= 0 Then
         If ixv <= W - 1 Then
            PicData(0, ix, iy) = PicDataORG(0, ixv, iy)
            PicData(1, ix, iy) = PicDataORG(1, ixv, iy)
            PicData(2, ix, iy) = PicDataORG(2, ixv, iy)
         End If
         End If
      Next ix
      DrawProgress
      Next iy
   Case 2   ' H & V Flute
      StartProgress (W - 1)
      For ix = 0 To W - 1
      For iy = 0 To H - 1
         iyv = iy + (iy Mod NN) - NN \ 2
         If iyv >= 0 Then
         If iyv <= H - 1 Then
            PicData(0, ix, iy) = PicDataORG(0, ix, iyv)
            PicData(1, ix, iy) = PicDataORG(1, ix, iyv)
            PicData(2, ix, iy) = PicDataORG(2, ix, iyv)
         End If
         End If
      Next iy
      DrawProgress
      Next ix
      StartProgress (H - 1)
      For iy = 0 To H - 1
      For ix = 0 To W - 1
         ixv = ix + (ix Mod NN) - NN \ 2
         If ixv >= 0 Then
         If ixv <= W - 1 Then
            PicData(0, ix, iy) = PicData(0, ixv, iy)
            PicData(1, ix, iy) = PicData(1, ixv, iy)
            PicData(2, ix, iy) = PicData(2, ixv, iy)
         End If
         End If
      Next ix
      DrawProgress
      Next iy
   End Select
   Form1.picPB.Cls
End Sub

Private Sub StartProgress(Par As Long)
' Private xd As Single, xp As Single
' EG: Par = PRELIEF * H
   Form1.picPB.DrawWidth = 3
   Form1.picPB.Cls
   xp = 0
   xd = Form1.picPB.Width / Par
End Sub
Private Sub DrawProgress()
' Private xd As Single, xp As Single
   Form1.picPB.PSet (xp, 2)
   Form1.picPB.Refresh
   xp = xp + xd
End Sub
       
