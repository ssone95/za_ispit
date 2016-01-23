VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVreme 
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer tVreme 
      Interval        =   1000
      Left            =   8040
      Top             =   120
   End
   Begin VB.TextBox txtSlajd 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.HScrollBar HSSlajd 
      Height          =   375
      Left            =   240
      Max             =   1000
      Min             =   -500
      TabIndex        =   8
      Top             =   3840
      Width           =   7935
   End
   Begin VB.TextBox txtRezultat 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCelobrojno 
      Caption         =   "Celobrojno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdPodeljeno 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdPuta 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtBroj2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtBroj1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCelobrojno_Click()
    Dim Broj1 As Double, Broj2 As Double, Rezultat As Double
    
    Broj1 = txtBroj1.Text
    Broj2 = txtBroj2.Text
    
    If Broj2 = 0 Then
        MsgBox "Delilac ne sme biti nula!"
        Exit Sub
    End If
    
    Rezultat = Broj1 \ Broj2
    
    txtRezultat.Text = Rezultat
End Sub

Private Sub cmdMinus_Click()
    Dim Broj1 As Double, Broj2 As Double, Rezultat As Double
    
    Broj1 = txtBroj1.Text
    Broj2 = txtBroj2.Text
    
    Rezultat = Broj1 - Broj2
    
    txtRezultat.Text = Rezultat
End Sub

Private Sub cmdPlus_Click()
    Dim Broj1 As Double, Broj2 As Double, Rezultat As Double
    
    Broj1 = txtBroj1.Text
    Broj2 = txtBroj2.Text
    
    Rezultat = Broj1 + Broj2
    
    txtRezultat.Text = Rezultat
End Sub

Private Sub cmdPodeljeno_Click()
    Dim Broj1 As Double, Broj2 As Double, Rezultat As Double
    
    Broj1 = txtBroj1.Text
    Broj2 = txtBroj2.Text
    
    If Broj2 = 0 Then
        MsgBox "Delilac ne sme biti nula!"
        Exit Sub
    End If
    
    Rezultat = Broj1 / Broj2
    
    txtRezultat.Text = Rezultat
End Sub

Private Sub cmdPuta_Click()
    Dim Broj1 As Double, Broj2 As Double, Rezultat As Double
    
    Broj1 = txtBroj1.Text
    Broj2 = txtBroj2.Text
    
    Rezultat = Broj1 * Broj2
    
    txtRezultat.Text = Rezultat
End Sub

Private Sub Form_Load()
    tVreme.Enabled = True
End Sub

Private Sub HSSlajd_Scroll()
    txtSlajd.Text = HSSlajd.Value
End Sub

Private Sub tVreme_Timer()
    txtVreme.Text = Now
End Sub
