VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scanner"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":164A
   ScaleHeight     =   4065
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   450
      Left            =   4110
      TabIndex        =   24
      Top             =   3555
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CheckBox Check2 
      Height          =   195
      Left            =   3810
      TabIndex        =   23
      Top             =   2295
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Index           =   2
      Left            =   45
      TabIndex        =   18
      Top             =   3615
      Value           =   1  'Checked
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   3180
      Value           =   1  'Checked
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   16
      Top             =   2745
      Value           =   1  'Checked
      Width           =   180
   End
   Begin VB.CommandButton Scan 
      Caption         =   "Scan"
      Height          =   435
      Left            =   4110
      TabIndex        =   3
      Top             =   3075
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   2505
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "0"
      Top             =   2685
      Width           =   660
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   2490
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "0"
      Top             =   3135
      Width           =   660
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   2490
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "0"
      Top             =   3585
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Append Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   9
      Left            =   4065
      TabIndex        =   22
      Top             =   2190
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   3225
      TabIndex        =   21
      Top             =   3645
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   3240
      TabIndex        =   20
      Top             =   3195
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   3240
      TabIndex        =   19
      Top             =   2715
      Width           =   315
   End
   Begin VB.Label tipo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BMP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   1440
      TabIndex        =   15
      Top             =   2655
      Width           =   615
   End
   Begin VB.Label qtf 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   285
      TabIndex        =   14
      Top             =   2655
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   2265
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "and in all the sub-directories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   3
      Left            =   720
      TabIndex        =   12
      Top             =   1365
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C:\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   900
      Index           =   2
      Left            =   15
      TabIndex        =   11
      Top             =   420
      Width           =   5445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   4095
      TabIndex        =   10
      Top             =   2610
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Looking for Images on the Folder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   0
      Left            =   570
      TabIndex        =   9
      Top             =   0
      Width           =   4425
   End
   Begin VB.Label qtf 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   285
      TabIndex        =   8
      Top             =   3105
      Width           =   960
   End
   Begin VB.Label qtf 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   2
      Left            =   285
      TabIndex        =   7
      Top             =   3540
      Width           =   960
   End
   Begin VB.Label tipo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JPG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   1455
      TabIndex        =   6
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label tipo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GIF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   1485
      TabIndex        =   5
      Top             =   3600
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Length above"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   540
      Index           =   5
      Left            =   2355
      TabIndex        =   4
      Top             =   2100
      Width           =   1035
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

    If Command1.Caption = "Stop" Then
        abortou = True
        Exit Sub
    End If
    
    Unload Me

Exit Sub

End Sub

Private Sub Form_Load()

    Label1(2) = Form3.Dir1.path

End Sub

Private Sub Scan_Click()

  Dim i As Integer
  Dim SearchPath As String, FindStr As String
  Dim FileSize As Long
  Dim NumFiles As Integer, NumDirs As Integer
  
    Scan.Enabled = False
    Command1.Caption = "Stop"
    Command1.Visible = True

    If Check2.Value = 0 Then
        Form3.List1.Clear
        Erase path
        Ponteiro = 0
    End If

    Label1(1).Visible = True
  
    '   Screen.MousePointer = vbHourglass
    
    'List1.Clear
    'SearchPath = Text1.Text
    'FindStr = Text2.Text
    'FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    'Text3.Text = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
    'Text4.Text = "Size of files found under " & SearchPath & " = " & Format$(FileSize, "#,###,###,##0") & " Bytes"

    'Erase path

    qtf(0) = 0
    qtf(1) = 0
    qtf(2) = 0

    DoEvents

    'Make_scan Left$(Disks(Index).Caption, 2) + "\"
    For i = 0 To 2
        Looking = i
        If Check1(i) Then
            If abortou Then
                Exit For '>---> Next
            End If
            qt_files = 0
            FindStr = "*." & tipo(i).Caption
            SearchPath = Form3.Dir1.path
            FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
        End If
    Next i
    
    abortou = False
    Scan.Enabled = True

    Command1.Caption = "OK"
    Label1(1).Visible = False

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Text1_Change(Index As Integer)

    Min_size(Index) = Val(Text1(Index).Text) * 1024

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If InStr("0123456789" & Chr$(27) & Chr$(46) & Chr$(8) & Chr$(13), Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If

End Sub


