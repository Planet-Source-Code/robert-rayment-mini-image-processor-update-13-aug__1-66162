Attribute VB_Name = "ModLoadPicture"
' ModLoadPicture.bas

Option Explicit

Public W As Long, H As Long     ' Image width & height

Public Function LoadThePicture(FileSpec$, picARR() As Byte) As Boolean
' Public BHI As BITMAPINFOHEADER   - see ModAPIs.bas
' Public tBitmap As BITMAP
' Could be private: BHI also used in DISPLAY on Form1

' Note: Form1.hdc used. Could Create a DC instead.

Dim ThePic As StdPicture
   On Error GoTo FileError
   Screen.MousePointer = vbHourglass
   Set ThePic = LoadPicture(FileSpec$)
   If GetObject(ThePic, Len(tBitmap), tBitmap) = 0 Then
      Screen.MousePointer = vbDefault
      Set ThePic = Nothing
      MsgBox "FILE ERROR"
      Exit Function
   End If
   W = tBitmap.bmWidth
   H = tBitmap.bmHeight
   ReDim picARR(0 To 3, 0 To W - 1, 0 To H - 1)
   
   ' Copy image colors to picARR()
   With BHI
      .biSize = 40
      .biPlanes = 1
      .biWidth = W
      .biHeight = H
      .biBitCount = 32
   End With
   If GetDIBits(Form1.HDC, ThePic.handle, 0, _
      H, picARR(0, 0, 0), BHI, 0) = 0 Then
      Screen.MousePointer = vbDefault
      Set ThePic = Nothing
      MsgBox "DIB ERROR"
      Exit Function
   End If
   Set ThePic = Nothing
   Screen.MousePointer = vbDefault
   LoadThePicture = True
   Exit Function
'=======
FileError:
      Screen.MousePointer = vbDefault
      Set ThePic = Nothing
      MsgBox "FILE ERROR"
      LoadThePicture = False
End Function


