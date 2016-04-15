VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdIzracunaj 
      Caption         =   "Izracunaj"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox TekstZaokruzeno 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox TekstReal 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox TekstUnos 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Zaokruzeno"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Realni deo"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Unos"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdIzracunaj_Click()
    Dim unos As Double
    Dim ceobroj As Integer
    Dim realanDeo As Double
    
    unos = TekstUnos.Text
    
    ceobroj = Int(unos)
    
    realanDeo = unos - ceobroj
    
    TekstReal.Text = realanDeo
    
    unos = Round(unos, 2)
    
    TekstZaokruzeno.Text = unos
End Sub
