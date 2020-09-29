VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memindahkan String dari VB ke Lingkungan DOS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Created by Rizky Khapidsyah
'Tips ini menggunakan clipboard untuk melewatkan String 'ari Visual Basic ke lingkungan DOS, seperti jika user 'mengetik suatu string di Command Prompt dan menekan 'Enter...
'Klik tombol pertama untuk membuka sebuah jendela DOS.
'Klik tombol kedua untuk melewatkan string ke 'lingkungan DOS.
'Source Code Program Dimulai Dari Sini

Private Sub Command1_Click()
  'Membuka sebuah jendela DOS secara maximized
  Shell ("cmd.exe"), vbMaximizedFocus
End Sub

Private Sub Command2_Click()
   'Bersihkan clipboard
   Clipboard.Clear
   'Menyalin string yang Anda lewatkan ke dalam
   'clipboard, termasuk code dari Enter (Chr$(13) =
      'Enter Key)
   'Ganti "Dir *.*" di bawah dengan string yang Ingin
   'Anda lewatkan.
   Clipboard.SetText "Dir *.*" + Chr$(13)
   'Fokuskan ke jendela DOS. "MS-DOS Prompt" adalah
   'judul dari
   'jendela DOS. Jika jendela DOS Anda mempunya judul
   'yang lain (Judul adalah yang tertulis di bar atas
   'jendela DOS), ganti "MS-DOS Prompt" di bawah
   'dengan judul jendela DOS Anda.
   AppActivate "MS-DOS Prompt"
   'Gunakan fungsi "SendKeys" untuk mengirim string
   'dari clipboard ke lingkungan DOS.
   SendKeys "% ep", 1
End Sub

