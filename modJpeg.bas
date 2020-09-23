Attribute VB_Name = "modJpeg"
Option Explicit

Public m_Image     As cImage
Public m_Jpeg      As cJpeg
Public m_FileName  As String
Public Function FileExists(FileName As String) As Boolean
    If Len(FileName) > 0 Then FileExists = (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function
Public Sub SaveImage(TheImage As cImage, FileName As String)
    Set m_Jpeg = New cJpeg
        Set m_Image = TheImage 'Call this before the form loads to initialize it
    
    m_FileName = FileName
    
    'better than average
    m_Jpeg.Quality = 85
    
           'Sample the cImage by hDC
    m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height

       'Delete file if it exists
    If FileExists(FileName) Then
        SetAttr FileName, vbNormal
        Kill FileName
    End If

       'Save the JPG file
    m_Jpeg.SaveFile m_FileName


    Set m_Image = Nothing
    Set m_Jpeg = Nothing
End Sub
