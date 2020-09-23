Attribute VB_Name = "ModSaveBMP"
' ModSaveBMP.bas

Option Explicit

' Public PicData(0 To 3, 0 To W - 1, 0 To H - 1)

Private Type BITMAPFILEHEADER   ' For 8bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+(3*W+EXB)*H  4  FileSize B
    bReserved1        As Integer  ' 0         2
    bReserved2        As Integer  ' 0         2
    bOffBits          As Long     ' 54        4
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' W         4
    bHeight           As Long     ' H         4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 24        2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' (3*W+EXB)xH     4  Image size B
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 54
End Type

Public ExtraBytes As Long
Private bARRPlus() As Byte  ' <<<<<<<<<<<<<<

Public Function SaveBMP24(FSpec$, bARR() As Byte, bWidth As Long, bHeight As Long) As Boolean
' bArr() = Public PicData(0 To 3, 0 To W - 1, 0 To H - 1)

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Long
Dim BytesPerScanLine As Long
Dim ExtraBytes As Long
Dim ix As Long, iy As Long
Dim ib As Long
On Error GoTo SaveBMPError

   BytesPerScanLine = (3 * bWidth + 3) And &HFFFFFFFC
   ExtraBytes = BytesPerScanLine - 3 * bWidth
   With BFH
      .bType = &H4D42    ' BM
      .bWidth = bWidth
      .bHeight = bHeight
      .bSize = 54 + BytesPerScanLine * Abs(bHeight)
      .bOffBits = 54
      .bHeaderSize = 40
      .bNumPlanes = 1
      .bBPP = 24
      .bBytesInImage = BytesPerScanLine * Abs(bHeight)
   End With
   ' 24bpp image map.  Fill from 32bpp image map.
   ReDim bARRPlus(0 To 3 * bWidth - 1 + ExtraBytes, 0 To Abs(bHeight) - 1)
   For iy = 0 To Abs(bHeight) - 1
      ib = 0
      For ix = 0 To bWidth - 1
         ib = 3 * ix
         bARRPlus(ib, iy) = bARR(0, ix, iy)
         bARRPlus(ib + 1, iy) = bARR(1, ix, iy)
         bARRPlus(ib + 2, iy) = bARR(2, ix, iy)
      Next ix
   Next iy
   
   '-- Kill previous
   On Error Resume Next
   Kill FSpec$
   On Error GoTo 0
   
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Put #fnum, , bARRPlus()  '<<<<<<<<<<<<
   Close #fnum
   Erase bARRPlus()
   SaveBMP24 = True
   On Error GoTo 0
   Exit Function
'=======
SaveBMPError:
   Close
   SaveBMP24 = False
End Function


