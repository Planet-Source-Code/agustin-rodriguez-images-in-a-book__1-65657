VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00BC614E&
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   405
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Index           =   3
      Left            =   4005
      TabIndex        =   7
      Top             =   6795
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Index           =   2
      Left            =   4080
      TabIndex        =   6
      Top             =   5835
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Index           =   1
      Left            =   2730
      TabIndex        =   5
      Top             =   6780
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Index           =   0
      Left            =   2805
      TabIndex        =   2
      Top             =   5820
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox PagDIR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   3420
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   285
      Width           =   3000
   End
   Begin VB.PictureBox PagESQ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   420
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   285
      Width           =   3000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   420
      TabIndex        =   8
      Top             =   2580
      Width           =   45
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "§"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6480
      TabIndex        =   4
      Top             =   2370
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6480
      TabIndex        =   3
      Top             =   75
      Width           =   210
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   2820
      Left            =   45
      Picture         =   "Form1.frx":164A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   6750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MakeFileAssociation "Alb", "d:\meus projetos\images-in-a-book", "Images-in-a-Book.exe", "Images Album", ""
Option Explicit

Private Sub Command1_Click(Index As Integer)

  Dim x As Integer
  Dim Y As Integer
  Dim xx As Integer
  Dim yy As Integer

    Select Case Index
      Case Is = 0
        If Image2.Width > 800 Then
            Exit Sub
        End If
        x = 20
        xx = 40
        Y = 0
        yy = 0
      Case 1
        If Image2.Width < 190 Then
            Exit Sub
        End If
        x = -20
        xx = -40
        Y = 0
        yy = 0
      Case 2
        If Image2.Height > 480 Then
            Exit Sub
        End If
        x = 0
        xx = 0
        Y = 20
        yy = 20
      Case 3
        If Image2.Height < 100 Then
            Exit Sub
        End If
        x = 0
        xx = 0
        Y = -20
        yy = -20
    End Select
    Label3.Move Label3.Left, Label3.Top + yy
    Text1.Move Text1.Left, Text1.Top + yy
    
    Width = Width + xx * Screen.TwipsPerPixelX
    Height = Height + yy * Screen.TwipsPerPixelY
    Form2.Width = Form2.Width + xx * Screen.TwipsPerPixelX
    Form2.Height = Form2.Height + yy * Screen.TwipsPerPixelY

    Image2.Width = Image2.Width + xx
    Image2.Height = Image2.Height + yy
    
    Label1.Move Label1.Left + xx
    Label2.Move Label2.Left + xx, Label2.Top + Y

    PagESQ.Width = PagESQ.Width + x
    PagESQ.Height = PagESQ.Height + Y
    
    Form2.picDST(1).Width = Form2.picDST(1).Width + x
    Form2.picDST(1).Height = Form2.picDST(1).Height + Y
    
    PagDIR.Width = PagDIR.Width + x
    PagDIR.Height = PagDIR.Height + Y
    
    PagDIR.Move PagDIR.Left + x
    
    Form2.picDST(0).Width = Form2.picDST(0).Width + x
    Form2.picDST(0).Height = Form2.picDST(0).Height + Y
    
    Form2.picDST(0).Move Form2.picDST(0).Left + x
    
    Form2.Trabalho.Width = Form2.Trabalho.Width + x
    Form2.Trabalho.Height = Form2.Trabalho.Height + Y
    
    Form2.Img_stretch.Width = Form2.Img_stretch.Width + x
    Form2.Img_stretch.Height = Form2.Img_stretch.Height + Y
    
    Listindex_atual = Ponteiro
    Open_img Form2.Img_Orig, path(Ponteiro) & Form3.List1.List(Ponteiro)
    
    Form2.Stretch
    PagDIR.Picture = Form2.Img_stretch.Image
    
    TransparentBlt Form1.PagDIR.hDC, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, Form2.Picture1.hDC, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, vbRed
    
    PagDIR.Refresh
    
    Listindex_atual = Ponteiro - 1
    Open_img Form2.Img_Orig, path(Ponteiro - 1) & Form3.List1.List(Ponteiro - 1)
    
    Form2.Stretch
    PagESQ.Picture = Form2.Img_stretch.Image
    
    TransparentBlt PagESQ.hDC, Form2.Trabalho.Width - Form2.Picture2.Width, 0, Form2.Picture2.Width, Form2.Picture2.Height, Form2.Picture2.hDC, 0, 0, Form2.Picture2.Width, Form2.Picture2.Height, vbRed
    
    PagESQ.Refresh
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      
      Case 33
        Form3.VScroll1.Value = -(Passo - (Passo < 100))
      Case 34
        Form3.VScroll1.Value = -(Passo + (Passo > 1))
        
      Case 39
        If Shift Then
            Command1_Click (0)
            Exit Sub
        End If
        Form2.Command2_Click
      Case 37
        If Shift Then
            Command1_Click (1)
            Exit Sub
        End If
        Form2.Command1_Click
        
      Case 40
        If Shift Then
            Command1_Click (2)
            Exit Sub
        End If
        Form2.Command2_Click
        
      Case 38
        If Shift Then
            Command1_Click (3)
            Exit Sub
        End If
        Form2.Command1_Click
      Case 27
        Unload Form4
    End Select

End Sub

Private Sub Form_Load()

    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hwnd, 12345678, 255, LWA_COLORKEY Or LWA_ALPHA
    Passo = 5

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

    If capture Then
        Form2.Move Left, Top
    End If

    capture = False

End Sub

Private Sub Form_Resize()

  Static vez As Integer

    If vez = 0 Then
        DoEvents
        vez = 1
        Form2.Move Left, Top
        Form2.Show , Me
        
        Dim i As Integer

TransparentBlt PagESQ.hDC, PagESQ.Width - Form2.Picture2.Width, 0, Form2.Picture2.Width, Form2.Picture2.Height, Form2.Picture2.hDC, 0, 0, Form2.Picture2.Width, Form2.Picture2.Height, vbRed
TransparentBlt PagDIR.hDC, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, Form2.Picture1.hDC, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, vbRed

    
PagDIR.Refresh
PagESQ.Refresh

    Form3.Drive1.Drive = Left$(App.path, 2)
    Form3.Dir1.path = App.path
    Last_Album = GetSetting("Images-in-a-Book", "Album", "Last", "")
    
    If Command <> "" Then
        Last_Album = Mid$(Command, 2, Len(Command) - 2)
    End If
    
    If Last_Album <> "" Then
        If Dir$(Last_Album) <> "" Then
            Form3.Cmd_Open (Last_Album)
            Form3.Command3_Click
            'Hide
        End If
    Else
        Label2_Click
    End If
    
    
        'Form3.Show
        Form1.SetFocus
    End If

End Sub

Private Sub Label1_Click()

    End

End Sub

Private Sub Label2_Click()

    If Top < Screen.Height / 2 Then
        Form3.Move Left, Top + 1000, 5940, 4890
      Else
        Form3.Move Left, Top - 3000, 5940, 4890
    End If
    

    
    If Left > Screen.Width / 2 Then
        Form3.Move Left - 200, Form3.Top, 5940, 4890
      Else
        Form3.Move Left + 1000, Form3.Top, 5940, 4890
    End If
    
    Form3.Show 1

End Sub

Private Sub Label3_DblClick()

    Text1 = Form3.List1.ListIndex
    Text1.Visible = True
    Text1.SetFocus

End Sub

Private Sub PagDIR_Click()

    If Loading_Gif Then
        Exit Sub
    End If
    
    Unload Form4

    Arquivo = path(Ponteiro) & Form3.List1.List(Ponteiro)

    Listindex_atual = -1
    Open_img Form2.Img_Orig, Arquivo

    If Form2.Img_Orig.Width * Screen.TwipsPerPixelX > Left Then
        Form4.Move 0, 0, Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Form2.Img_Orig.Height * Screen.TwipsPerPixelY
      Else
        Form4.Move Left - Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Top, Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Form2.Img_Orig.Height * Screen.TwipsPerPixelY
    End If

    Form4.Show
    Me.SetFocus

End Sub

Private Sub PagESQ_Click()

    If Loading_Gif Then
        Exit Sub
    End If
    
    Unload Form4
    Arquivo = path(Ponteiro - 1) & Form3.List1.List(Ponteiro - 1)
    Listindex_atual = -1
    Open_img Form2.Img_Orig, Arquivo

    If Form2.Img_Orig.Width * Screen.TwipsPerPixelX > Left Then
        Form4.Move 0, 0, Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Form2.Img_Orig.Height * Screen.TwipsPerPixelY
      Else
        Form4.Move Left - Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Top, Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Form2.Img_Orig.Height * Screen.TwipsPerPixelY
    End If

    'Form4.Move 0, 0, Form2.Img_Orig.Width * Screen.TwipsPerPixelX, Form2.Img_Orig.Height * Screen.TwipsPerPixelY
    Form4.Show
    Me.SetFocus

End Sub

Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    'Text1.SelText = Text1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If InStr("0123456789" & Chr$(27) & Chr$(46) & Chr$(8) & Chr$(13), Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 And Val(Text1) < Form3.List1.ListCount Then
        KeyAscii = 0
        Ponteiro = Val(Text1) - 2
        Form2.Command1_Click
        Text1.Visible = False
    End If

    If KeyAscii = 27 Then
        Text1.Visible = False
    End If

End Sub


