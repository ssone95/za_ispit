VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRacunaj 
      Caption         =   "Racunaj"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtRezultat 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtZbir 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtProizvod 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtZBI 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRacunaj_Click()
    Dim A(1000) As Integer, ZBI As Integer, PetA As Single, ZbirA As Integer, Rezultat As Single, N As Integer, i As Integer
    
    ZBI = Int(txtZBI.Text)
    If ZBI = 0 Then
        ZBI = 8
    End If
    
    N = 5 + ZBI
    
    For i = 0 To N - 1
        A(i) = Val(InputBox("Unesi broj:"))
    Next i
    
    
    PetA = 1
    For i = 0 To 4
        PetA = PetA * A(i)
    Next i
    
    
    ZbirA = 0
    For i = 0 To N - 1
        ZbirA = ZbirA + A(i)
    Next i
    
    Rezultat = PetA / ZbirA
    
    txtProizvod.Text = PetA
    txtZbir.Text = ZbirA
    txtRezultat.Text = Rezultat

End Sub
