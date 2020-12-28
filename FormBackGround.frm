VERSION 4.00
Begin VB.Form BackGround 
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   4050
   ClientTop       =   4200
   ClientWidth     =   7575
   ControlBox      =   0   'False
   Height          =   4410
   Left            =   3990
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   ShowInTaskbar   =   0   'False
   Top             =   3855
   Width           =   7695
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   240
   End
   Begin VB.Image ImageBoxBackGround 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   0
      Picture         =   "FormBackGround.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "BackGround"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub


Private Sub Timer1_Timer()

    Let ImageBoxBackGround.Width = FormBackGround.Width / 15
    Let ImageBoxBackGround.Height = FormBackGround.Height / 15
    
End Sub


