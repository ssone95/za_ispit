VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdIzracunaj 
      Caption         =   "Izracunaj"
      Height          =   915
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   2355
   End
   Begin VB.TextBox TekstIspis 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox TekstZBI 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox TekstX 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdIzracunaj_Click()
    Dim X As Double
    Dim ZBI As Integer
    Dim Y As Double
    Dim privremeno As Double
   
    X = TekstX.Text
    ZBI = TekstZBI.Text
   
    If ZBI < 1 Then
        ZBI = 10
    End If
    
    
    If X <> -10 Then
        privremeno = 2 * (X * X) - 5 * X - 4
   
        Y = privremeno / (X + 5)
    
        TekstIspis.Text = Y
    End If
End Sub
