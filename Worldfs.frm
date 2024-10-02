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
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   12240
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
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
      Left            =   5160
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   7320
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   2880
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim equiposA(3), equiposB(3) As String
Dim puntajeA(3), puntajeB(3) As Integer
Dim winA, winB As String
Dim A, B As Integer
Dim partidos As Integer

Private Sub Command1_Click()
Dim macht As String
    
partidos = partidos + 1

    If partidos = 1 Then
    
    List1.Clear
    Randomize
    

    For A = 0 To 3
        For B = A + 1 To 3
        macht = equiposA(A) & " vs " & equiposA(B)
        puntajeA(A) = Int(Rnd() * 5)
        puntajeA(B) = 5 - puntajeA(A)
        List1.AddItem macht & ": " & puntajeA(A) & " - " & puntajeA(B)
        Next B
    Next A
    
    List2.Clear
    
    For A = 0 To 3
        For B = A + 1 To 3
        macht = equiposB(A) & " vs " & equiposB(B)
        puntajeB(A) = Int(Rnd() * 5)
        puntajeB(B) = 5 - puntajeB(A)
        List2.AddItem macht & ": " & puntajeB(A) & " - " & puntajeB(B)
        Next B
    Next A
  
  ElseIf partidos = 2 Then
    
    
    
  End If
        
End Sub

Private Sub Command2_Click()

    List3.AddItem (Verificar(List3.List(0), List3.List(0)))
    
    
    
End Sub

Private Sub Form_Activate()
    
    equiposA(0) = "T1"
    equiposA(1) = "Mad lions"
    equiposA(2) = "G2"
    equiposA(3) = "Gen.g"
    equiposB(0) = "R7"
    equiposB(1) = "TES"
    equiposB(2) = "FNC"
    equiposB(3) = "C9"
    
End Sub
Private Function Verificar(maxPuntajeA As Integer, maxPuntajeB As Integer) As String
Dim T As Integer
Dim resultado As String
    
    maxPuntajeA = puntajeA(0)
    winA = equiposA(0)
    maxPuntajeB = puntajeB(0)
    winB = equiposB(0)


    For T = 0 To List1.ListCount
        If puntajeA(T) > maxPuntajeA Then
            maxPuntajeA = puntajeA(T)
            winA = equiposA(T)
        End If


    
    resultado = "ganador de la semifinal: " & winA & " pasando con " & maxPuntajeA & " puntos"

    Verificar = resultado

End Function
