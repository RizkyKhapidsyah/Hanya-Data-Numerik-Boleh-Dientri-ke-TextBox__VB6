VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hanya Data Numerik Boleh Dientri ke TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Created by Rizky Khapidsyah
'Source Code Dimulai Dari sini

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub

'Hanya karakter 0 sampai dengan 9 saja.
Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

'Cara di atas hanya menerima karakter 0 sampai dengan 9 'saja. Agar tombol lainnya seperti Delete, BackSpace, 'dan SpaceBar juga bisa diterima, Anda bisa menggunakan 'tips di bawah ini:

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
End Sub


