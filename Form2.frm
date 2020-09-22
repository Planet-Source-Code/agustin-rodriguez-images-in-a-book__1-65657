VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00BC614E&
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   570
      Left            =   585
      TabIndex        =   0
      Top             =   6690
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox Img_stretch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   8640
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   8
      Top             =   5445
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox Img_Orig 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   9990
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   7
      Top             =   5415
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   7860
      Picture         =   "Form2.frx":164A
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   6
      Top             =   405
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   8220
      Picture         =   "Form2.frx":8440
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PREV"
      Height          =   570
      Left            =   2295
      TabIndex        =   4
      Top             =   6750
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox picDST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00BC614E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Index           =   0
      Left            =   3420
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   3
      Top             =   285
      Width           =   3000
   End
   Begin VB.PictureBox Trabalho 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   8625
      Picture         =   "Form2.frx":F236
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox picDST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00BC614E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Index           =   1
      Left            =   420
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   285
      Width           =   3000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Command1_Click()

  Dim i As Long

    On Error GoTo erro

    If Ponteiro + 2 > Form3.List1.ListCount - 1 Then
        Exit Sub
    End If

    Ponteiro = Ponteiro + 2

    Form3.List1.ListIndex = Ponteiro
    
    Trabalho.Picture = Form1.PagDIR
    
    TransparentBlt Trabalho.hDC, 0, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, vbRed
    Form1.Label3 = Form3.List1.ListIndex & "/" & Form3.List1.ListCount - 1
    Form1.Label3.Refresh
    Listindex_atual = Ponteiro
    Open_img Img_Orig, path(Ponteiro) & Form3.List1.List(Ponteiro)
    
    Stretch
    Form1.PagDIR.Picture = Img_stretch.Image
    
    TransparentBlt Form1.PagDIR.hDC, 0, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, vbRed
    
    Form1.PagDIR.Refresh
    
    For i = (Trabalho.ScaleWidth \ Passo) * Passo To 0 Step -Passo
        Pts(0).x = 0
        Pts(0).Y = 0

        Pts(1).x = i
        
        Pts(1).Y = 50 * Cos((i * 90! / Trabalho.ScaleWidth) * 3.14 / 180)
        
        Pts(2).x = 0
        Pts(2).Y = Trabalho.ScaleHeight

        picDST(0).Cls

        Ret = PlgBlt(picDST(0).hDC, Pts(0), Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, 0, 0, 0)
       
        picDST(0).Refresh
        
    Next i
    
    Listindex_atual = Ponteiro - 1
    Open_img Img_Orig, path(Ponteiro - 1) & Form3.List1.List(Ponteiro - 1)
    
    Stretch
    Trabalho.Picture = Img_stretch.Image
    
    TransparentBlt Trabalho.hDC, Trabalho.Width - Picture2.Width, 0, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, vbRed
    
    For i = (Trabalho.ScaleWidth \ Passo) * Passo To 0 Step -Passo
        Pts(0).x = i
        Pts(0).Y = 50 * Sin((i * 90 / Trabalho.ScaleWidth) * 3.14 / 180)

        Pts(1).x = Trabalho.ScaleWidth
        Pts(1).Y = 0

        Pts(2).x = i
        Pts(2).Y = Trabalho.ScaleHeight + 50 * Sin((i * 90 / Trabalho.ScaleWidth) * 3.14 / 180)

        picDST(1).Cls
        Ret = PlgBlt(picDST(1).hDC, Pts(0), Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, 0, 0, 0)

        picDST(1).Refresh
        
    Next i

    Form1.PagESQ.Picture = Trabalho.Image
    Form1.PagESQ.Refresh
    
    picDST(1).Cls
    
Exit Sub
    
erro:
    Img_Orig = LoadPicture(App.path & "\none.err")
    Resume

End Sub

Public Sub Command2_Click()

  Dim i As Integer
    
    If Ponteiro < 2 Then
        Exit Sub
    End If
    
    Ponteiro = Ponteiro - 2
    Form3.List1.ListIndex = Ponteiro

    Trabalho.Picture = Form1.PagESQ
    TransparentBlt Trabalho.hDC, Trabalho.Width - Picture2.Width, 0, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, vbRed
    Form1.Label3 = Form3.List1.ListIndex & "/" & Form3.List1.ListCount - 1
    Form1.Label3.Refresh
    Listindex_atual = Ponteiro - 1
    Open_img Img_Orig, path(Ponteiro - 1) & Form3.List1.List(Ponteiro - 1)
    
    Stretch
    Form1.PagESQ.Picture = Img_stretch.Image
    
    TransparentBlt Form1.PagESQ.hDC, Trabalho.Width - Picture2.Width, 0, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, vbRed
    Form1.PagESQ.Refresh
    
    For i = Trabalho.ScaleWidth Mod Passo To Trabalho.ScaleWidth Step Passo
        Pts(0).x = i
        Pts(0).Y = 50 * Sin((i * 90! / Trabalho.ScaleWidth) * 3.14 / 180)

        Pts(1).x = Trabalho.ScaleWidth
        Pts(1).Y = 0

        Pts(2).x = i
        Pts(2).Y = Trabalho.ScaleHeight + 50 * Sin((i * 90! / Trabalho.ScaleWidth) * 3.14 / 180)

        picDST(1).Cls

        Ret = PlgBlt(picDST(1).hDC, Pts(0), Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, 0, 0, 0)

        picDST(1).Refresh
        
    Next i
        
    Listindex_atual = Ponteiro
    Open_img Img_Orig, path(Ponteiro) & Form3.List1.List(Ponteiro)
    
    Stretch
    Trabalho.Picture = Img_stretch.Image
    
    TransparentBlt Trabalho.hDC, 0, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, vbRed
    
    For i = Trabalho.ScaleWidth Mod Passo To Trabalho.ScaleWidth Step Passo
        Pts(0).x = 0
        Pts(0).Y = 0

        Pts(1).x = i
        Pts(1).Y = 50 * Cos((i * 90! / Trabalho.ScaleWidth) * 3.14 / 180)

        Pts(2).x = 0
        Pts(2).Y = Trabalho.ScaleHeight

        picDST(0).Cls

        Ret = PlgBlt(picDST(0).hDC, Pts(0), Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, 0, 0, 0)

        picDST(0).Refresh
        
    Next i

    Listindex_atual = Ponteiro
    Open_img Img_Orig, path(Ponteiro) & Form3.List1.List(Ponteiro)
    
    Stretch
    Form1.PagDIR.Picture = Img_stretch.Image
    
    TransparentBlt Form1.PagDIR.hDC, 0, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, vbRed
    Form1.PagDIR.Refresh
    
    picDST(0).Cls

End Sub

Private Sub Form_Load()
    
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hwnd, 12345678, 255, LWA_COLORKEY Or LWA_ALPHA

End Sub

Public Sub Stretch()

  Dim th As Single, tw As Single
  Dim xImg As Single
  Dim yImg As Single
  Dim xPic As Single
  Dim yPic As Single
  Dim xRatio As Single
  Dim yRatio As Single
  
    On Error GoTo erro
    
    Img_stretch.Width = Trabalho.Width
    Img_stretch.Height = Trabalho.Height
    Img_stretch.Cls
    
    If Form3.Aspect Then
  
        xImg = Img_Orig.ScaleWidth
        yImg = Img_Orig.ScaleHeight
        xPic = Img_stretch.ScaleWidth
        yPic = Img_stretch.ScaleHeight
  
        xRatio = xImg / xPic
        yRatio = yImg / yPic

        If xRatio >= yRatio Then
            Img_stretch.PaintPicture Img_Orig, (xPic - (Img_Orig.Width / xRatio)) / 2, (yPic - (Img_Orig.Height / xRatio)) / 2, (Img_Orig.Width / xRatio), (Img_Orig.Height / xRatio)
          Else
            Img_stretch.PaintPicture Img_Orig, (xPic - (Img_Orig.Width / yRatio)) / 2, (yPic - (Img_Orig.Height / yRatio)) / 2, (Img_Orig.Width / yRatio), (Img_Orig.Height / yRatio)
        End If
        
        'tw = Img_Orig.ScaleWidth
        'th = Img_Orig.ScaleHeight
        
        'Do While tw > Img_stretch.ScaleWidth Or th > Img_stretch.ScaleHeight
        
        '   tw = tw - tw / 100
        '  th = th - th / 100
        
        'Loop
        
        'Img_stretch.PaintPicture Img_Orig, (Img_stretch.ScaleWidth - tw) / 2, (Img_stretch.ScaleHeight - th) / 2, tw, th, 0, 0, Img_Orig.ScaleWidth, Img_Orig.ScaleHeight, vbSrcCopy
        
      Else
        Img_stretch.PaintPicture Img_Orig, 0, 0, Img_stretch.ScaleWidth, Img_stretch.ScaleHeight, 0, 0, Img_Orig.ScaleWidth, Img_Orig.ScaleHeight, vbSrcCopy
    End If
    
    Img_stretch.Refresh

Exit Sub
    
erro:
    Img_Orig = LoadPicture(App.path & "\none.err")
    Resume

End Sub

