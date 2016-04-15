VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdIzracunaj 
      Caption         =   "Izracunaj"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox TekstIspis 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox TekstZBI 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox TekstX 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Ispis"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ZBI"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label X 
      Caption         =   "X"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
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
    
    privremeno = (X - ZBI)
    
    If privremeno >= 0 Then
        Y = Sqr(privremeno * X)
    
        TekstIspis.Text = Y
    End If
    ' ne ide dalje jer je broj manji od 0
End Sub
