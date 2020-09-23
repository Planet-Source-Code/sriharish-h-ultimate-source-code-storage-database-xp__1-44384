VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source Code Storage Database Utility-By Sri Harish"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "form1.frx":08CA
   ScaleHeight     =   7185
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstdescription 
      Height          =   255
      Left            =   2280
      TabIndex        =   52
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Index           =   2
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   51
      Top             =   840
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.XpBs XpBs9 
      Height          =   735
      Left            =   7320
      TabIndex        =   50
      Top             =   6450
      Width           =   1215
      _extentx        =   2143
      _extenty        =   1296
      caption         =   "Print"
      pictureposition =   3
      buttonstyle     =   3
      picture         =   "form1.frx":B5BF
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":D2C9
      font            =   "form1.frx":D2F5
      picture         =   "form1.frx":D321
      mousepointer    =   99
   End
   Begin VB.ListBox lstdate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   49
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstinternet 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1320
      TabIndex        =   48
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstrating 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1080
      TabIndex        =   47
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstcat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   46
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstbackup 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   45
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstProject 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   44
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstscreenshot 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   43
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstplatform 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   42
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstauthor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   41
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstComments 
      Height          =   255
      Left            =   2040
      TabIndex        =   40
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin Project1.XpBs cmdclose 
      Height          =   735
      Left            =   8520
      TabIndex        =   39
      Top             =   6450
      Width           =   1095
      _extentx        =   1931
      _extenty        =   1296
      caption         =   "Close"
      pictureposition =   0
      buttonstyle     =   3
      picture         =   "form1.frx":F02B
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":FE7D
      font            =   "form1.frx":FEA9
      picture         =   "form1.frx":FED5
      mousepointer    =   99
   End
   Begin Project1.XpBs cmdsave 
      Height          =   735
      Left            =   6120
      TabIndex        =   38
      Top             =   6450
      Width           =   1215
      _extentx        =   2143
      _extenty        =   1296
      caption         =   "Save Changes"
      pictureposition =   3
      buttonstyle     =   3
      picture         =   "form1.frx":10D27
      picturewidth    =   16
      pictureheight   =   16
      picturesize     =   2
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":12A31
      font            =   "form1.frx":12A5D
      picture         =   "form1.frx":12A89
      mousepointer    =   99
   End
   Begin Project1.XpBs cmddelete 
      Height          =   735
      Left            =   4800
      TabIndex        =   37
      Top             =   6450
      Width           =   1335
      _extentx        =   2355
      _extenty        =   1296
      caption         =   "Delete This Record"
      pictureposition =   3
      buttonstyle     =   3
      picture         =   "form1.frx":14793
      picturewidth    =   16
      pictureheight   =   16
      picturesize     =   2
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":16B75
      font            =   "form1.frx":16BA1
      picture         =   "form1.frx":16BCD
      mousepointer    =   99
   End
   Begin Project1.XpBs cmdnew 
      Height          =   735
      Left            =   3360
      TabIndex        =   36
      Top             =   6450
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1296
      caption         =   "Add New Record"
      pictureposition =   3
      buttonstyle     =   3
      picture         =   "form1.frx":18FAF
      picturewidth    =   16
      pictureheight   =   16
      picturesize     =   2
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":1B391
      font            =   "form1.frx":1B3BD
      picture         =   "form1.frx":1B3E9
      mousepointer    =   99
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   10
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   5760
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   4200
      TabIndex        =   31
      Top             =   5280
      Width           =   3375
   End
   Begin Project1.XpBs XpBs8 
      Height          =   375
      Left            =   7680
      TabIndex        =   29
      Top             =   4755
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      caption         =   "GO"
      buttonstyle     =   3
      picture         =   "form1.frx":1D7CB
      picturewidth    =   16
      pictureheight   =   16
      picturesize     =   0
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":1F4D5
      font            =   "form1.frx":1F501
      picture         =   "form1.frx":1F52D
      mousepointer    =   99
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   8
      Left            =   4200
      TabIndex        =   28
      Top             =   4800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   25
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   4200
      TabIndex        =   23
      Top             =   3840
      Width           =   3375
   End
   Begin Project1.XpBs XpBs7 
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   3360
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Open"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "form1.frx":21237
      font            =   "form1.frx":21263
      mousepointer    =   99
   End
   Begin Project1.XpBs XpBs6 
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      Top             =   3360
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Browse"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "form1.frx":2128F
      font            =   "form1.frx":212BB
      mousepointer    =   99
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   19
      Top             =   3360
      Width           =   3375
   End
   Begin Project1.XpBs XpBs5 
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      Top             =   2880
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Open"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "form1.frx":212E7
      font            =   "form1.frx":21313
      mousepointer    =   99
      url             =   "D:\Sriharish\FUN (partially updated)\flashes\kbc.exe"
   End
   Begin Project1.XpBs XpBs4 
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   2880
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Browse"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "form1.frx":2133F
      font            =   "form1.frx":2136B
      mousepointer    =   99
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   4200
      TabIndex        =   15
      Top             =   2880
      Width           =   3375
   End
   Begin Project1.XpBs XpBs3 
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   2350
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "View"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "form1.frx":21397
      font            =   "form1.frx":213C3
      mousepointer    =   99
   End
   Begin Project1.XpBs XpBs2 
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   2350
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Browse"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "form1.frx":213EF
      font            =   "form1.frx":2141B
      mousepointer    =   99
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   11
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.ListBox lstnames 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   5100
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin Project1.XpBs XpBs1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   300
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Clear"
      buttonstyle     =   3
      picture         =   "form1.frx":21447
      picturewidth    =   16
      pictureheight   =   16
      picturesize     =   0
      originalpicsizew=   16
      originalpicsizeh=   16
      font            =   "form1.frx":23829
      font            =   "form1.frx":23855
      picture         =   "form1.frx":23881
      mousepointer    =   99
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label totalrec 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   30
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Code Location :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Location :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Screenshot :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Author :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2760
      X2              =   9600
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2760
      X2              =   0
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2760
      X2              =   2760
      Y1              =   0
      Y2              =   6360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Code List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================
'This code is written in a way where every user
'can understand its code very easily. Because
'even if its a Database program, my code does
'not use a single line of Database programming,
'Therefore it is easy for everyone to understand
'this code and there is no need for documentation
'Perviously I made this utility for my personal use because
'i had so many codes with unique Zip File names
'I had to find a solution for this.See next line
'*************************************************
'PLS VOTE FOR ME,
'*************************************************
'See "About Dialog" for more
'**************************************************
'

'====================================================
'The database in scodedb.dat file loads in several
'listboxes(Drag the form in design mode below)
'Each list boxes are has a particular name
'For Ex:Lstcomments, lstauthor, lstdate etc

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
'this code removes the selected entry
'and calls the SAVE Code
If lstnames.ListIndex = -1 Then
        If MsgBox("You Do Not Have An Entry Selected", vbExclamation) = vbOK Then Exit Sub
        End If
    If MsgBox("Are You Sure You Want To Delete The Selected Entry?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
    Dim a As Integer
    a = lstnames.ListIndex
    lstnames.RemoveItem a
    lstauthor.RemoveItem a
    lstplatform.RemoveItem a
    lstdescription.RemoveItem a
    lstscreenshot.RemoveItem a
    lstbackup.RemoveItem a
    lstProject.RemoveItem a
    lstcat.RemoveItem a
    lstrating.RemoveItem a
    lstinternet.RemoveItem a
    lstdate.RemoveItem a
    lstComments.RemoveItem a
        If a = lstnames.ListCount Then
        lstnames.ListIndex = a - 1
    Else
        lstnames.ListIndex = a
    End If
    'Switch to next code
    Call lstnames_Click
    'save
    Call Save_It
totalrec.Caption = lstnames.ListCount
End Sub

Private Sub cmdnew_Click()
'Add new enry form
addentry.Show
End Sub

Private Sub cmdsave_Click()
 If MsgBox("Are You Sure You Want to save?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   'At the main form if you have made any
   'changes, then this code is verivied
   'Save code
    Call Save_It
    
End Sub

Private Sub Form_Load()
Close #1
On Error GoTo ErrorHandler
    Dim TempName As String
'open Database
    Open App.Path & "\" & "codedb.dat" For Input As 1
    On Error Resume Next
    'Add items to the relavent lisboxes
    'until end of the file
    Do Until EOF(1)
        Line Input #1, TempName
        lstnames.AddItem TempName
        
        Line Input #1, TempAuthor
        lstauthor.AddItem TempAuthor
        
        Line Input #1, Tempplatform
        lstplatform.AddItem Tempplatform
        
        
        Line Input #1, tempdescription
        lstdescription.AddItem tempdescription
        
        Line Input #1, Tempscreenshot
        lstscreenshot.AddItem Tempscreenshot
        
        Line Input #1, Tempproject
        lstProject.AddItem Tempproject
        
        Line Input #1, Tempbackup
        lstbackup.AddItem Tempbackup
        
        Line Input #1, Tempcat
        lstcat.AddItem Tempcat
        
        Line Input #1, tempinternet
        lstinternet.AddItem tempinternet
        
        Line Input #1, tempdate
        lstdate.AddItem tempdate
        
        Line Input #1, temprating
        lstrating.AddItem temprating
        
        Line Input #1, tempcomments
        lstComments.AddItem tempcomments
        Loop
    Close #1
    'switch list index of names to Zero
    lstnames.ListIndex = 0
    'Update rec no
    totalrec.Caption = lstnames.ListCount
ErrorHandler:
    Select Case Err.Number
    Case 53
        Call Save_It
    End Select
End Sub
Private Sub Save_It()
'This is the code for saving
'first open
Close #1
    Open App.Path & "\" & "codedb.dat" For Output As 1
    'Print items in the list boxes to the end of te list
    For i = 0 To lstnames.ListCount - 1
        Print #1, lstnames.List(i)
        Print #1, lstauthor.List(i)
        Print #1, lstplatform.List(i)
        Print #1, lstdescription.List(i)
        Print #1, lstscreenshot.List(i)
        Print #1, lstProject.List(i)
        Print #1, lstbackup.List(i)
        Print #1, lstcat.List(i)
         Print #1, lstinternet.List(i)
        Print #1, lstdate.List(i)
        Print #1, lstrating.List(i)
        Print #1, lstComments.List(i)
            Next i
    Close #1
End Sub



Private Sub lstnames_Click()
'On Error GoTo lstNamesErr
'This code loads the particular Source code
'when user clicks the Code Name

''Adjust the list boxex
   lstauthor.ListIndex = lstnames.ListIndex
   lstplatform.ListIndex = lstnames.ListIndex
   lstdescription.ListIndex = lstnames.ListIndex
   lstscreenshot.ListIndex = lstnames.ListIndex
   lstProject.ListIndex = lstnames.ListIndex
   lstbackup.ListIndex = lstnames.ListIndex
   lstrating.ListIndex = lstnames.ListIndex
   lstinternet.ListIndex = lstnames.ListIndex
   lstcat.ListIndex = lstnames.ListIndex
   lstdate.ListIndex = lstnames.ListIndex
   lstComments.ListIndex = lstnames.ListIndex
  'Assign the textboxes to print information
  'in the listboxes
  Text1(0).Text = lstauthor.Text
  Text1(1).Text = lstplatform.Text
  Text1(2).Text = lstdescription.Text
  Text1(3).Text = lstscreenshot.Text
  Text1(4).Text = lstProject.Text
  Text1(5).Text = lstbackup.Text
  Text1(6).Text = lstcat.Text
  Text1(7).Text = lstrating.Text
  Text1(8).Text = lstinternet.Text
  Text1(9).Text = lstdate.Text
  Text1(10).Text = lstComments.Text
End Sub


'This following set of code
'verifies whether there is any chamges made to the or database
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If lstnames.ListCount > 0 Then
    lstComments.List(lstnames.ListIndex) = Text1(10).Text
End If
If lstnames.ListCount > 0 Then
    lstdate.List(lstnames.ListIndex) = Text1(9).Text
End If
If lstnames.ListCount > 0 Then
    lstinternet.List(lstnames.ListIndex) = Text1(8).Text
End If
If lstnames.ListCount > 0 Then
    lstrating.List(lstnames.ListIndex) = Text1(7).Text
End If
If lstnames.ListCount > 0 Then
    lstcat.List(lstnames.ListIndex) = Text1(6).Text
End If
If lstnames.ListCount > 0 Then
    lstbackup.List(lstnames.ListIndex) = Text1(5).Text
End If
If lstnames.ListCount > 0 Then
    lstProject.List(lstnames.ListIndex) = Text1(4).Text
End If
If lstnames.ListCount > 0 Then
    lstscreenshot.List(lstnames.ListIndex) = Text1(3).Text
End If
If lstnames.ListCount > 0 Then
    lstdescription.List(lstnames.ListIndex) = Text1(2).Text
End If
If lstnames.ListCount > 0 Then
    lstplatform.List(lstnames.ListIndex) = Text1(1).Text
End If
If lstnames.ListCount > 0 Then
    lstauthor.List(lstnames.ListIndex) = Text1(0).Text
End If
End Sub

Private Sub XpBs1_Click()
'Clear Text search
txtsearch.Text = ""
End Sub

Private Sub XpBs2_Click()
'Screenshot location
'Filter is set to BMP,JPG files
CommonDialog1.Filter = "Bitmap Images|*.bmp|JPG File|*.jpg|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileTitle) > 0 Then
Text1(3).Text = CommonDialog1.FileName
Else
Exit Sub
End If
End Sub

Private Sub XpBs3_Click()

'Internet code location
If Len(Text1(3).Text) > 0 Then
'On Error GoTo error
XpBs3.URL = Text1(3).Text
Else
error: MsgBox "Invalid location.", vbCritical, ":(Invalid Location"
Exit Sub
End If

End Sub

Private Sub XpBs4_Click()
'Project location
'Filter is set to VBP, and others
CommonDialog1.Filter = "Visual Basic Project|*.vbp|Other Projects|*.*|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileTitle) > 0 Then
Text1(4).Text = CommonDialog1.FileName
Else
Exit Sub
End If
End Sub

Private Sub XpBs5_Click()
'open project file
Dim opn
If Len(Text1(4).Text) > 0 Then
'On Error GoTo errors
ShellExecute hwnd, "open", Text1(4).Text, vbNullString, vbNullString, conSwNormal
Else
errors: MsgBox "Type something.", vbCritical, ":(Invalid Location"
Exit Sub
End If
End Sub

Private Sub XpBs6_Click()
'Zip file or backup location
'Filter is set to:ZIP
CommonDialog1.Filter = "Zip Files|*.zip|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileTitle) > 0 Then
Text1(5).Text = CommonDialog1.FileName
Else
Exit Sub
End If
End Sub

Private Sub XpBs7_Click()
'Open Zip file
If Len(Text1(5).Text) > 0 Then
'On Error GoTo errors
ShellExecute hwnd, "open", Text1(5).Text, vbNullString, vbNullString, conSwNormal
Else
errors: MsgBox "Type something.", vbCritical, ":(Invalid Location"
Exit Sub
End If
End Sub

Private Sub XpBs8_Click()
'Open URL
If Len(Text1(8).Text) > 0 Then
XpBs8.URL = Text1(8).Text
Else
MsgBox "Invalid or Blank URL", vbCritical, ":( User Error-Can't Help!!"
Exit Sub
End If
End Sub

Private Sub XpBs9_Click()

On Error GoTo printerror
'Printer code
Printer.Font = "arial"
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Sri Harish's Source Code Database Utility."
    Printer.Print ""
    Printer.Font = "Arial"
    Printer.FontUnderline = False
    Printer.FontSize = 12
    Printer.Print "Code Name: ", ,
    Printer.FontBold = False
    Printer.Print lstnames.Text
    Printer.FontBold = True
    Printer.Print "Author Name: ", ,
    Printer.FontBold = False
    Printer.Print Text1(0).Text
    Printer.FontBold = True
        Printer.Print "Platform: ", ,
    Printer.FontBold = False
    Printer.Print Text1(1).Text
    Printer.FontBold = False
            Printer.Print "Description: ", ,
    Printer.FontBold = False
    Printer.Print Text1(2).Text
    Printer.FontBold = False
            Printer.Print "Screenshot: ", ,
    Printer.FontBold = False
    Printer.Print Text1(3).Text
    Printer.FontBold = False
            Printer.Print "Project Loc: ", ,
    Printer.FontBold = False
    Printer.Print Text1(4).Text
    Printer.FontBold = False
            Printer.Print "Backup Loc: ", ,
    Printer.FontBold = False
    Printer.Print Text1(5).Text
    Printer.FontBold = False
            Printer.Print "Code Category: ", ,
    Printer.FontBold = False
    Printer.Print Text1(6).Text
    Printer.FontBold = False
            Printer.Print "Rating: ", ,
    Printer.FontBold = False
    Printer.Print Text1(7).Text
    Printer.FontBold = False
            Printer.Print "Internet: ", ,
    Printer.FontBold = False
    Printer.Font.Underline = True
    Printer.Print Text1(8).Text
    Printer.FontBold = False
            Printer.Print "Download Date: ", ,
    Printer.FontBold = False
    Printer.Print Text1(9).Text
    Printer.FontBold = False
            Printer.Print "Comments: ", ,
    Printer.FontBold = False
    Printer.Print Text1(10).Text
    Printer.FontBold = False
printerror:
MsgBox "Printer not detected or not connected or not availaable or driver not installed or invalid driver. In the  mean time PLS VOTE FOR ME", vbCritical, ":( Printer Error and VOTE FOR ME"
Exit Sub
End Sub
Private Sub txtSearch_Change()
    
    Dim MatchFound As Boolean
    Dim Last As Integer, J As Integer
    'Code used for search in Lstnames
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text1(7).Text = ""
    Text1(8).Text = ""
    Text1(9).Text = ""
    Text1(10).Text = ""
    
   
    Last = lstnames.ListCount - 1
    J = 0
    MatchFound = False
    Do
        If InStr(1, lstnames.List(J), txtsearch.Text, 1) > 0 Then
            MatchFound = True
            lstnames.ListIndex = J
        End If
        J = J + 1
    Loop Until J > Last Or MatchFound
    If Not MatchFound Then
        lstnames.ListIndex = -1
    End If
    'call lstames click code
    Call lstnames_Click
End Sub
