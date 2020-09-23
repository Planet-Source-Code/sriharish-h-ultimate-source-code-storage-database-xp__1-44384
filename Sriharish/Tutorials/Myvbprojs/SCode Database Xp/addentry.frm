VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form addentry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Source code Entry"
   ClientHeight    =   6510
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5820
   Icon            =   "addentry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "addentry.frx":08CA
   ScaleHeight     =   6510
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtplatform 
      Height          =   285
      Left            =   2160
      TabIndex        =   33
      Text            =   "VB 6.0"
      Top             =   5640
      Width           =   2775
   End
   Begin Project1.XpBs XpBs7 
      Height          =   495
      Left            =   1800
      TabIndex        =   31
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Clear Fields"
      ButtonStyle     =   3
      Picture         =   "addentry.frx":B5BF
      PictureHover    =   "addentry.frx":BE99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Project1.XpBs cmdadd 
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Add Entry"
      ButtonStyle     =   3
      Picture         =   "addentry.frx":C773
      PictureHover    =   "addentry.frx":D04D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Project1.XpBs XpBs2 
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Browse"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Project1.XpBs XpBs1 
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Browse"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.TextBox Txtcomments 
      Height          =   405
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox Txtdate 
      Height          =   285
      Left            =   2160
      TabIndex        =   23
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Txtinternet 
      Height          =   285
      Left            =   2160
      TabIndex        =   22
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Txtrating 
      Height          =   285
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Txtcategory 
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Txtbackup 
      Height          =   285
      Left            =   2160
      TabIndex        =   18
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Txtproject 
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtscreenshot 
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Txtdescription 
      Height          =   405
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Txtauthor 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Txtcode 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   840
      Width           =   2775
   End
   Begin Project1.XpBs XpBs3 
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Browse"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Project1.XpBs XpBs4 
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Test"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Project1.XpBs XpBs5 
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Insert Todays's Date"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Platform :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Out of 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "VOTE FOR ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   240
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Download Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet code location :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Rating :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Category :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup(ZIP) Location : :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Location :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Screenshot Location :  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Description :(Readme)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Author Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the Code :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "addentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
If Txtcode = "" Then
        MsgBox "You Must Type In Code Name To Add", vbExclamation, ":( VOTE FOR ME- SCODE DB"
Else
'Add entered items to relavent list boxes
'and update them to SCODEDB.dat file
        Form1.lstnames.AddItem Txtcode.Text
        Form1.lstauthor.AddItem Txtauthor.Text
        Form1.lstdescription.AddItem Txtdescription.Text
        Form1.lstplatform.AddItem txtplatform.Text
        Form1.lstscreenshot.AddItem txtscreenshot.Text
        Form1.lstProject.AddItem Txtproject.Text
        Form1.lstbackup.AddItem Txtbackup.Text
        Form1.lstcat.AddItem Txtcategory.Text
        Form1.lstrating.AddItem Txtrating.Text
        Form1.lstinternet.AddItem Txtinternet.Text
        Form1.lstdate.AddItem Txtdate.Text
        Form1.lstComments.AddItem Txtcomments.Text
        Txtcode.Text = ""
        Txtauthor.Text = ""
        Txtdescription.Text = ""
        txtplatform.Text = ""
        txtscreenshot.Text = ""
        Txtproject.Text = ""
        Txtbackup.Text = ""
        Txtcategory.Text = ""
        Txtrating.Text = ""
        Txtinternet.Text = ""
        Txtdate.Text = ""
        Txtcomments.Text = ""
        Close #1
        Open App.Path & "\" & "codedb.dat" For Output As 1
        For i = 0 To Form1.lstnames.ListCount - 1
        Print #1, Form1.lstnames.List(i)
        Print #1, Form1.lstauthor.List(i)
        Print #1, Form1.lstplatform.List(i)
        Print #1, Form1.lstdescription.List(i)
        Print #1, Form1.lstscreenshot.List(i)
        Print #1, Form1.lstProject.List(i)
        Print #1, Form1.lstbackup.List(i)
        Print #1, Form1.lstcat.List(i)
        Print #1, Form1.lstinternet.List(i)
        Print #1, Form1.lstdate.List(i)
        Print #1, Form1.lstrating.List(i)
        Print #1, Form1.lstComments.List(i)
        Next i
Close #1
Form1.totalrec.Caption = Form1.lstnames.ListCount
        MsgBox "Your Entry Has Been Added.PS: VOTE FOR ME", vbInformation, ":( Entry added"
    End If
End Sub
Private Sub XpBs1_Click()
CommonDialog1.Filter = "JPG FILES|*.jpg|Bitmap|*.bmp|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileName) > 0 Then
txtscreenshot.Text = CommonDialog1.FileName
Else
txtscreenshot.Text = ""
End If
End Sub

Private Sub XpBs2_Click()
CommonDialog1.Filter = "Visual Basic Project|*.vbp|Others|*.*|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileName) > 0 Then
Txtproject.Text = CommonDialog1.FileName
Else
Txtproject.Text = ""
End If
End Sub

Private Sub XpBs3_Click()
CommonDialog1.Filter = "ZIP Files|*.Zip|Others|*.*|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileName) > 0 Then
Txtbackup.Text = CommonDialog1.FileName
Else
Txtbackup.Text = ""
End If
End Sub

Private Sub XpBs4_Click()
If Len(Txtinternet.Text) > 0 Then
XpBs4.URL = Txtinternet.Text
Else
MsgBox "A blank Address?.Action canceled!!!", vbInformation, ":(VOTE FOR ME- User Error"
Exit Sub
End If
End Sub

Private Sub XpBs5_Click()
Dim sdate As Date
Txtdate.Text = Format(sdate, Date)
End Sub



Private Sub XpBs7_Click()
Txtcode.Text = ""
Txtauthor.Text = ""
Txtdescription.Text = ""
txtscreenshot.Text = ""
Txtbackup.Text = ""
Txtproject.Text = ""
Txtcategory.Text = ""
Txtrating.Text = ""
Txtinternet.Text = ""
Txtcomments.Text = ""
Txtdate.Text = ""
txtplatform.Text = ""
MsgBox "All clear. But PS: SEE about for more fun", vbOKOnly, ":( VOTE FOR ME_ALL CLEAR"
End Sub
