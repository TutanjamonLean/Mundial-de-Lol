VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20205
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   20205
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   5280
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   7560
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   3600
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim equipos(7) As String
Dim kills(9) As Integer
Dim A, B, K As Integer
Private Sub Command1_Click()

If List1.ListCount < 4 And List2.ListCount < 4 Then
    
    For A = 0 To 3
    
        List1.AddItem equipos(A)
    
    Next A

    For B = 4 To 7
    
        List2.AddItem equipos(B)
    
    Next B

End If


    
    
End Sub
Private Sub Form_Activate()

    equipos(0) = "T1"
    equipos(1) = "MAD lions"
    equipos(2) = "Fnatic"
    equipos(3) = "G2"
    equipos(4) = "Gen.g"
    equipos(5) = "Team liquid"
    equipos(6) = "Weibo Gaming"
    equipos(7) = "DAMWON gaming"


End Sub
Private Function Puntaje()




End Function


