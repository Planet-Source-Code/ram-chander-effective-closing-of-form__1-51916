VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Its an special form event which acts specially
'and closes it.
' This I developed for my friend who asked me something special in
'closing the form
'If you like it plz comment to me @ chandervv@yahoo.com
Private Sub Command1_Click()
GotoVal = Me.Height / 2

For Gointo = 1 To GotoVal
    'NEW ADDITION NEXT LINE

    DoEvents
        Me.Height = Me.Height - 10
        'Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then GoTo horiz
    Next Gointo
    'This is the width part of the same sequence above
horiz:
    Me.Height = 30
    GotoVal = Me.Width / 2

    For Gointo = 1 To GotoVal
        'NEW ADDITION NEXT LINE
        DoEvents
            Me.Width = Me.Width - 10
            'Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
        Next Gointo
        
End
End Sub
