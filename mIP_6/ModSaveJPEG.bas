Attribute VB_Name = "ModSaveJPEG"
Option Explicit

Private m_Jpeg As cJPEG

Public Sub SaveDATAJPEG(SaveSpec$, bDATA() As Byte, _
                        ByVal wide As Long, _
                        ByVal hite As Long, _
                        Optional ByVal SrcLeft As Long = 0, _
                        Optional ByVal SrcTop As Long = 0)

Dim Ret As Long
   'Using Default Frequencies & Quality
   ' & No comments
   Set m_Jpeg = New cJPEG
   m_Jpeg.SetSamplingFrequencies 2, 2, 1, 1, 1, 1
   m_Jpeg.Quality = 75
   m_Jpeg.Comment = ""
   m_Jpeg.SampleDATA bDATA(), wide, hite, SrcLeft, SrcTop
   On Error Resume Next
   Kill SaveSpec$
   On Error GoTo 0
   Ret = m_Jpeg.SaveFile(SaveSpec$)
   Set m_Jpeg = Nothing
End Sub

Public Sub SaveHDCJPEG(SaveSpec$, LHDC As Long, ByVal wide As Long, ByVal hite As Long)
Dim Ret As Long
   'Using Default Frequencies & Quality
   ' & No comments
   Set m_Jpeg = New cJPEG
   m_Jpeg.SetSamplingFrequencies 2, 2, 1, 1, 1, 1
   m_Jpeg.Quality = 75
   m_Jpeg.Comment = ""
   m_Jpeg.SampleHDC LHDC, wide, hite
   On Error Resume Next
   Kill SaveSpec$
   On Error GoTo 0
   Ret = m_Jpeg.SaveFile(SaveSpec$)
   Set m_Jpeg = Nothing
End Sub


