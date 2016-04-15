VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdProveri 
      Caption         =   "Proveri"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox TekstDeljivo 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox TekstBroj 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdProveri_Click()
    Dim broj As Double
    broj = TekstBroj.Text
    
    If broj Mod 2 = 0 Then
        TekstDeljivo.Text = "Ja"
    Else
        TekstDeljivo.Text = "Nein"
    End If
End Sub
