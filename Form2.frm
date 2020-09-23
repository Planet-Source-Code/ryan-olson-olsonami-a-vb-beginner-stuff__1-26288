VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "It Worked!"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Insert Text in Textbox"
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command4 
         Caption         =   "See Code"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear Text"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Insert It!"
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Center Form"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Center Form"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   3615
      Begin VB.CommandButton Command5 
         Caption         =   "See Code"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub



Private Sub Command2_Click()
Text1.SelText = Text2.Text
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Form3.code.Navigate (App.Path & "\textbox.html")
Form3.SetFocus
End Sub

Private Sub Command5_Click()
Form3.code.Navigate (App.Path & "\center.html")
Form3.SetFocus
End Sub
