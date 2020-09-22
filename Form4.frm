VERSION 5.00
Begin VB.Form Form4 
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   1230
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   525
      Picture         =   "Form4.frx":164A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrMoveNext 
      Enabled         =   0   'False
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sldFrames_Value As Long
Private sldFrames_Max As Long
Private GIF As cGifReader
Private m_oRenderer             As cBmpRenderer
Private WithEvents m_oReader    As cGifReader
Attribute m_oReader.VB_VarHelpID = -1
Private m_lFrameCount           As Long
Private m_aFrames()             As UcsFrameInfo

Private Type UcsFrameInfo
    oPic        As StdPicture
    nDelay      As Long
End Type

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

    If KeyCode = Asc("C") Then
        Form2.Img_Orig.Width = ScaleWidth
        Form2.Img_Orig.Height = ScaleHeight
        Form2.Img_Orig.Cls

        Call SetStretchBltMode(Form2.Img_Orig.hDC, STRETCHMODE)
        Ret = StretchBlt(Form2.Img_Orig.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)
    
        Clipboard.Clear

        Clipboard.SetData Form2.Img_Orig.Image
    End If

End Sub

Private Sub Form_Load()

    If UCase$(Right$(Arquivo, 4)) = ".GIF" Then
        AutoRedraw = False
    
        Loading_Gif = True
        X_Form_Load
        X_Form_Activate
        Loading_Gif = False
    
      Else
        AutoRedraw = True
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Ocupado Then
        Exit Sub
    End If
        
    xx = x * Screen.TwipsPerPixelX
    yy = Y * Screen.TwipsPerPixelY
    capture = True
    ReleaseCapture
    SetCapture Me.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If capture Then
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - xx, Pt.Y * Screen.TwipsPerPixelY - yy
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    capture = False

End Sub

Private Sub Form_Resize()

    If Right$(UCase$(Arquivo), 4) = ".GIF" Then
        StretchBlt hDC, 0, 0, ScaleWidth, ScaleHeight, picView.hDC, 0, 0, picView.ScaleWidth, picView.ScaleHeight, vbSrcCopy
        Exit Sub
    End If

    Call SetStretchBltMode(hDC, STRETCHMODE)
    Ret = StretchBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, Form2.Img_Orig.hDC, 0, 0, Form2.Img_Orig.ScaleWidth, Form2.Img_Orig.ScaleHeight, vbSrcCopy)
    Refresh

End Sub

Private Sub sldFrames_Change()

  Const FUNC_NAME     As String = "sldFrames_Change"
  Dim lDelay          As Long
    
    On Error GoTo EH
    With m_aFrames(sldFrames_Value)
        lDelay = IIf(.nDelay < 8, 80, .nDelay * 10)
        Set picView.Picture = .oPic
        picView.Refresh
        StretchBlt hDC, 0, 0, ScaleWidth, ScaleHeight, picView.hDC, 0, 0, picView.ScaleWidth, picView.ScaleHeight, vbSrcCopy
       
        If tmrMoveNext.Enabled Then
            tmrMoveNext.Interval = lDelay
            tmrMoveNext.Enabled = False
            tmrMoveNext.Enabled = True
        End If
    End With
        
Exit Sub

EH:
    Resume Next

End Sub

Private Sub tmrMoveNext_Timer()

  Static Gif_step As Integer
  Static n_times As Integer

    If Gif_step = 0 Then
        Gif_step = 1
    End If
    
    If (sldFrames_Value + Gif_step) < 1 Or (sldFrames_Value + Gif_step) > sldFrames_Max Then
        
        Gif_step = -Gif_step
    End If
    sldFrames_Value = sldFrames_Value + Gif_step
    sldFrames_Change
    DoEvents

End Sub

Public Function Init(oRdr As cGifReader) As Boolean

  Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    
    Set m_oReader = oRdr
    If m_oRenderer.Init(oRdr) Then
        Set picView.Picture = Nothing
        If oRdr.MoveLast() Then
            m_lFrameCount = oRdr.FrameIndex + 1
            If m_lFrameCount > 1 Then
                sldFrames_Max = m_lFrameCount
            End If
        End If
    End If

Exit Function

EH:
    Resume Next

End Function

Private Sub X_Form_Load()

  Dim Ret As Long
  
  Dim filenumber As Integer

  Dim sFilename   As String * 260
  Dim lRetval     As Long
  Dim Path_wallpaper As String

    Set m_oRenderer = New cBmpRenderer
    ReDim m_aFrames(-1 To -1)
    
    Set GIF = New cGifReader
    
    If GIF.Init(Arquivo) Then
        GIF.MoveFirst
    End If
    ReDim m_aFrames(-1 To -1)
    Init GIF
      
    Width = GIF.ScreenWidth * Screen.TwipsPerPixelX '+ 1000            'HERE YOU SET THE WIDTH
    Height = GIF.ScreenHeight * Screen.TwipsPerPixelY '+ 1000          'AND THE HEIGHT
    
    picView.Width = Width
    picView.Height = Height
    picView.BackColor = GIF.BackgroundColor
    
End Sub

Private Sub X_Form_Activate()

  Const FUNC_NAME     As String = "Form_Activate"
  Dim lIdx            As Long
  Dim sInfo           As String
 
    On Error GoTo EH
    If UBound(m_aFrames) < 0 And m_lFrameCount > 0 Then
        ReDim m_aFrames(1 To m_lFrameCount)
        If m_oRenderer.MoveFirst() Then
            lIdx = 0
            Do While True
                If Not m_oRenderer.MoveNext Then
                    Exit Do
                End If
                lIdx = lIdx + 1
                With m_aFrames(lIdx)
                    Set .oPic = m_oRenderer.Image
                    .nDelay = m_oRenderer.Reader.DelayTime
                    sldFrames_Value = lIdx
                    If lIdx = 1 Then
                        sldFrames_Change
                    End If
                    DoEvents
                End With
            Loop
        End If
    End If
    sldFrames_Value = 0
    tmrMoveNext_Timer
    tmrMoveNext.Interval = 1
    tmrMoveNext.Enabled = True
    
Exit Sub

EH:
    Resume Next

End Sub


