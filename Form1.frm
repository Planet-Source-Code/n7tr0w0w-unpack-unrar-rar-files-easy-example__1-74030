VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unrar vb6 example by NiTrOwow"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtapp 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Unrar"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to unrar:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   870
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
RARExecute OP_EXTRACT, txtapp.Text
End Sub

Private Sub Form_Load()
txtapp.Text = App.Path
End Sub
