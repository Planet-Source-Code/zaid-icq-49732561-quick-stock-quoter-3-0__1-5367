VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "View Graph"
   ClientHeight    =   1230
   ClientLeft      =   4440
   ClientTop       =   4290
   ClientWidth     =   1575
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   1575
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Resize()
Form2.Height = Picture1.Height + 420
Form2.Width = Picture1.Width + 120
End Sub
