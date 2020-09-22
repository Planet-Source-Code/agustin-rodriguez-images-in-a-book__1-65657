Attribute VB_Name = "Module1"
Option Explicit
Public Ponteiro As Integer
Public path(-1 To 32768) As String
Public Last_Album As String
Public ultimo_dir As String
Public Passo As Integer
Public Pt As POINTAPI
Public capture As Integer
Public xx As Long
Public yy As Long
Public Ocupado As Integer
Public Arquivo As String
Public Loading_Gif As Integer
Public Listindex_atual As Integer

Public Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "GDI32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE As Long = vbPaletteModeNone

Public Declare Function PlgBlt Lib "GDI32" (ByVal hdcDest As Long, _
                        lpPoint As POINTAPI, _
                        ByVal hdcSrc As Long, _
                        ByVal nXSrc As Long, _
                        ByVal nYSrc As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal hbmMask As Long, _
                        ByVal xMask As Long, _
                        ByVal yMask As Long) As Long

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Pts(2) As POINTAPI
Public Ret As Long

Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Sub Open_img(x As PictureBox, Y As String)

    On Error GoTo erro
    
    x.Picture = LoadPicture(Y)

Exit Sub

erro:
    
    Y = App.path & "\none.err"
        
    Resume

End Sub

Public Function Check_valid(Y As String)

  Dim x(0 To 1) As Byte
  Dim free As Integer
    
    free = FreeFile
    Open Y For Binary As free
        Get #free, FileLen(Y) - 1, x
    Close free

    Select Case UCase$(Right$(Y, 4))
    
      Case ".GIF"
        Check_valid = x(0) = 0 And x(1) = 59
        
      Case ".JPG"
        Check_valid = x(0) = 255 And x(1) = 217
        
      Case ".BMP", ".ERR"
        Check_valid = True
        
    End Select

End Function


