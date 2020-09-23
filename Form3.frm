VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form3 
   Caption         =   "Code Window"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3885
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "See Code"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser code 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Note: You can also learn resizing code with this window"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
code.Navigate (App.Path & "\resize.html")
End Sub

Private Sub Form_Load()
code.Navigate "about:<center><font color=red>CODE</font> <font color=blue>WINDOW</font></center>"
Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
code.Width = Me.ScaleWidth
code.Height = Me.ScaleHeight - Label1.Height - Command1.Height
Label1.Width = Me.ScaleWidth
Label1.Top = Me.ScaleHeight - Label1.Height - Command1.Height
Command1.Top = Me.ScaleHeight - Command1.Height
End Sub
