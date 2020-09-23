VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Sri Harish's Scode Storage Database"
   ClientHeight    =   5535
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5745
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3820.355
   ScaleMode       =   0  'User
   ScaleWidth      =   5394.852
   ShowInTaskbar   =   0   'False
   Begin Project1.XpBs XpBs1 
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   4920
      Width           =   2415
      _extentx        =   4260
      _extenty        =   873
      backcolor       =   12583104
      forecolor       =   16777215
      caption         =   "Email Me"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "frmAbout.frx":1CFA
      font            =   "frmAbout.frx":1D26
      mousepointer    =   99
      url             =   "mailto:sriharish@msn.com"
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "VOTE FOR ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday Reminder 2003 Xp"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Auto Fill ( Like IE Address auto fill)"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Meta Magic Xp"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "License File Registration Xp"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Product Activation *Updated* ( With restore Facility)"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Also Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1D52
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1EFC
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Form_Load()
Dim username As String * 30
Dim returns
returns = GetUserName(username, 30)
username = Left(username, InStr(username, Chr(0)) - 1)
Label2.Caption = username
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Did You VOTE?", vbInformation + vbYesNo, ":( VOTE FOR ME") = vbYes Then
Unload Me
Else
MsgBox "Then Please VOTE", vbInformation + vbOKOnly, ":( VOTE FOR ME"
frmAbout.Show
Exit Sub
End If

End Sub
