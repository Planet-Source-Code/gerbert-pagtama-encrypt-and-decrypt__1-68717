VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypt and Decrypt (by : Gerbert Pagtama >>> E-mail : gerbert_p@yahoo.com"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      FillColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      FillColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   6000
      TabIndex        =   7
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt and Decrypt"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   660
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   5265
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt and Decrypt"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   660
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   5265
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   1680
      TabIndex        =   4
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   3720
      TabIndex        =   6
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   645
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FF0000&
      Height          =   3615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Encrypt and Decrypt by : Gerbert Pagtama
'Email : gerbert_p@yahoo.com

' sample
' text1.text = encrypt(text1.text)

' ENCRYPT
Function Encrypt(p_str As String) As String
    Dim i, strs
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) * 2)
    Next i
        Encrypt = strs
End Function

' DECRYPT
Function Decrypt(p_str As String) As String
    Dim i, strs
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        Decrypt = strs
End Function




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label3.Visible = True
 Label5.Visible = True
 Label7.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.Visible = False
  Label5.Visible = True
  Label7.Visible = True
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.Visible = True
  Label5.Visible = False
  Label7.Visible = True
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.Visible = True
  Label5.Visible = True
  Label7.Visible = False
End Sub

Private Sub Label6_Click()
 Text1.Text = Decrypt(Text1.Text)
End Sub

Private Sub Label4_Click()
  Text1.Text = Encrypt(Text1.Text)
End Sub

Private Sub Label8_Click()
 End
End Sub
