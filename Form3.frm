VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   4230
   ClientLeft      =   -5955
   ClientTop       =   615
   ClientWidth     =   5850
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":164A
   ScaleHeight     =   4230
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   240
      Left            =   5265
      Max             =   1
      Min             =   -100
      TabIndex        =   17
      Top             =   60
      Value           =   -5
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   11
      Top             =   75
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   10
      Top             =   75
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Index           =   2
      Left            =   2595
      TabIndex        =   9
      Top             =   75
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   300
      TabIndex        =   8
      Top             =   480
      Width           =   1830
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00404040&
      Height          =   990
      Left            =   300
      TabIndex        =   7
      Top             =   870
      Width           =   2475
   End
   Begin VB.CommandButton Command4 
      Caption         =   "All"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1995
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   3645
      Width           =   1290
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   2985
      Left            =   2970
      TabIndex        =   4
      Top             =   525
      Width           =   2565
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   2040
      Left            =   285
      MultiSelect     =   1  'Simple
      Pattern         =   "*.bmp;*.jpg;*.gif"
      TabIndex        =   3
      Top             =   1950
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Þ"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2265
      TabIndex        =   2
      Top             =   2985
      Width           =   570
   End
   Begin VB.CheckBox Aspect 
      Caption         =   "Aspect Rate"
      Height          =   225
      Left            =   2370
      TabIndex        =   1
      Top             =   3675
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   390
      Left            =   2280
      TabIndex        =   0
      Top             =   2475
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roll step"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4500
      TabIndex        =   18
      Top             =   75
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   5505
      TabIndex        =   16
      Top             =   45
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aspect Rate"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   2595
      TabIndex        =   15
      Top             =   3690
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GIF"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2895
      TabIndex        =   14
      Top             =   90
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JPG"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1770
      TabIndex        =   13
      Top             =   75
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BMP"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   12
      Top             =   60
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   3840
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   330
      Width           =   5715
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Begin VB.Menu Open_Album 
         Caption         =   "Open Album"
      End
      Begin VB.Menu Save_Album 
         Caption         =   "Save Album"
      End
      Begin VB.Menu Scan_view 
         Caption         =   "Scan"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu About_ind 
         Caption         =   "               Agustin Rodriguez"
         Index           =   0
      End
      Begin VB.Menu About_ind 
         Caption         =   "E-Mail: virtual_guitar_1@hotmail.com"
         Index           =   1
      End
      Begin VB.Menu About_ind 
         Caption         =   "                  Made in Brazil"
         Index           =   2
      End
      Begin VB.Menu About_ind 
         Caption         =   "                     -  2006 -"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click(Index As Integer)

  Dim i As Integer
  Dim arq As String

    For i = 0 To 2
        If Check1(i) Then
            arq = arq + "*." + Label1(i).Caption + ";"
        End If
    Next i
    If arq <> "" Then
        arq = Left$(arq, Len(arq) - 1)
    End If
        
    File1.Pattern = arq
    File1.Refresh

End Sub

Private Sub Command1_Click()

  Dim i As Integer
    
  Dim x As String

    x = File1.path
    If Right$(x, 1) <> "\" Then
        x = x + "\"
    End If
    
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            If Check_valid(x + File1.List(i)) Then
                File1.Selected(i) = False
                List1.AddItem File1.List(i)
                path(List1.NewIndex) = x
            End If
        End If
    Next i

End Sub

Private Sub Command2_Click()

    List1.Clear
    Ponteiro = 0
    
End Sub

Public Sub Command3_Click()

    If Ponteiro And 1 Then
        Ponteiro = Ponteiro + 1
    End If

    If Ponteiro = 0 Then
        Ponteiro = 1
    End If
    
    Listindex_atual = Ponteiro - 1
    Open_img Form2.Img_Orig, path(Ponteiro - 1) & List1.List(Ponteiro - 1)
    
    Form2.Stretch
    Form1.PagESQ.Picture = Form2.Img_stretch.Image
    
    Listindex_atual = Ponteiro
    Open_img Form2.Img_Orig, path(Ponteiro) & List1.List(Ponteiro)
    
    Form2.Stretch
    Form1.PagDIR.Picture = Form2.Img_stretch.Image

    TransparentBlt Form1.PagDIR.hDC, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, Form2.Picture1.hDC, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, vbRed
    TransparentBlt Form1.PagESQ.hDC, Form2.Trabalho.Width - Form2.Picture2.Width, 0, Form2.Picture2.Width, Form2.Picture2.Height, Form2.Picture2.hDC, 0, 0, Form2.Picture2.Width, Form2.Picture2.Height, vbRed
    Hide
    Form1.SetFocus
   
    List1.AddItem ""
       
    Form1.Label3 = Ponteiro & "/" & Form3.List1.ListCount - 1
    Form1.Label3.Refresh

End Sub

Private Sub Command4_Click()

  Static i As Integer
  Dim x As String

    x = File1.path
    If Right$(x, 1) <> "\" Then
        x = x + "\"
    End If
        
    For i = 0 To File1.ListCount - 1
            
        If Check_valid(x + File1.List(i)) Then
            List1.AddItem File1.List(i)
            path(List1.NewIndex) = x
        End If
    Next i

End Sub

Private Sub Dir1_Change()

    File1.path = Dir1.path

End Sub

Private Sub Drive1_Change()

    On Error GoTo erro
    Dir1.path = Drive1.Drive
sair:

Exit Sub

erro:
    Drive1.Drive = "c:"
    Resume sair

End Sub

Private Sub Exit_Click()

    End

End Sub

Private Sub Form_Activate()

    If List1.ListCount Then
        If List1.List(List1.ListCount - 1) = "" Then
            List1.RemoveItem (List1.ListCount - 1)
        End If
    End If
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    Hide
End If

End Sub

Private Sub List1_Click()

    Ponteiro = List1.ListIndex

End Sub

Public Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim i As Integer
  Dim x As Integer

    x = List1.ListIndex

    If KeyCode = 46 And List1.ListIndex <> -1 Then
    
        For i = List1.ListIndex To List1.ListCount - 1
            path(i) = path(i + 1)
        Next i
        path(i) = ""
        List1.RemoveItem List1.ListIndex
        
        If List1.ListCount = 0 Then
            Exit Sub
        End If
            
        If x > List1.ListCount - 1 Then
            x = x - 1
        End If
        
        List1.ListIndex = x
        Ponteiro = List1.ListIndex
    End If

End Sub

Private Sub Open_Album_Click()

  Static ultimo_dir As String

  Dim filename As String
  
    On Error GoTo EHCancel
    
    filename = OpenDialog(Me, "Album of images|*.alb", "Images-in-a-Bock", ultimo_dir)
    
    If filename = "" Then
        Exit Sub
    End If
        
    ultimo_dir = CurDir
    Cmd_Open (filename)
    SaveSetting "Images-in-a-Book", "Album", "Last", filename
    
    On Error GoTo EH
        
EHCancel:

Exit Sub

EH:
    
    Resume Next

End Sub

Private Sub Save_Album_Click()

  Const FUNC_NAME     As String = "cmdSave_Click"
  Dim free As Integer
  Dim filename As String
  Static ultimo_dir As String
  Dim i As Integer
  Dim x As Integer
  
    On Error GoTo EHCancel

    filename = SaveDialog(Me, "Album files|*.alb", ".alb", "Images-in-a-Book", App.path & "\Albums")
    If filename = "" Then
        Exit Sub
    End If

    ultimo_dir = CurDir
    free = FreeFile
    
    Open filename For Binary As free
    x = List1.ListCount - 1
    ReDim nome(x) As String
    For i = 0 To List1.ListCount - 1
        nome(i) = List1.List(i)
    Next i
        
    Put #free, 1, x
    Put #free, , path
    Put #free, , nome
        
    Close free
        
    SaveSetting "Images-in-a-Book", "Album", "Last", filename
    On Error GoTo EH
        
EHCancel:

Exit Sub

EH:
    
    Resume Next

End Sub

Public Sub Cmd_Open(filename)

  Const FUNC_NAME     As String = "cmdOpen_Click"
  Dim free As Integer
  Dim xx As String
  Dim qt_img As Integer
  Dim i As Integer
  
    free = FreeFile
    If filename = "" Then
        Exit Sub
    End If
    ultimo_dir = CurDir
    List1.Clear
    
    Open filename For Binary As free
    Get #free, 1, qt_img
    Get #free, , path
    ReDim nome(qt_img) As String
            
    Get #free, , nome
    
    For i = 0 To qt_img
        List1.AddItem nome(i)
    Next i
    
    Ponteiro = 0
    List1.ListIndex = 0
        
End Sub

Private Sub Scan_view_Click()

    Form5.Move Left + 1000, Top + 1000
    Form5.Show 1

End Sub

Private Sub VScroll1_Change()

    Label2 = Abs(VScroll1.Value)
    Passo = Label2

End Sub


