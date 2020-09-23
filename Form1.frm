VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beginner Stuff"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Load Forms"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   4695
      Begin VB.CommandButton Command9 
         Caption         =   "See Code"
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "See Code"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Form.Visible = True Method"
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Form.Show Method"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Note: These both do the same thing. (the new window has more stuff)"
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   520
         Width           =   4455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "OnMouseOver"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   4695
      Begin VB.CommandButton Command10 
         Caption         =   "See Code"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Move Mouse Here"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add Items to Listbox"
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   4815
      Begin VB.CommandButton Command7 
         Caption         =   "See Code"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add This"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fake Progress Bar With Timers"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
      Begin VB.CommandButton Command6 
         Caption         =   "See Code"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4200
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3840
         Top             =   120
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Form's Title"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "See Code"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change It!"
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Caption = Text1.Text 'sets the caption (title)
                        'to whatever is typed in the text in the text box
End Sub

Private Sub Command10_Click()
Form3.code.Navigate (App.Path & "\mouseover.html")
End Sub

Private Sub Command2_Click()
List1.AddItem Text2.Text
End Sub
Private Sub Command3_Click()
Form2.Show
End Sub
Private Sub Command4_Click()
Form2.Visible = True
Form2.SetFocus
End Sub
Private Sub Command5_Click()
Form3.code.Navigate (App.Path & "\title.html")
Form3.SetFocus
End Sub

Private Sub Command6_Click()
Form3.code.Navigate (App.Path & "\bar.html")
Form3.Show
End Sub

Private Sub Command7_Click()
Form3.code.Navigate (App.Path & "\list.html")
Form3.SetFocus
End Sub

Private Sub Command8_Click()
Form3.code.Navigate (App.Path & "\formshow.html")
Form3.SetFocus
End Sub

Private Sub Command9_Click()
Form3.code.Navigate (App.Path & "\formvis.html")
Form3.SetFocus
End Sub

Private Sub Form_Load()
Form3.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)


Do
If Form2.Visible = True Then
Form2.Top = Form2.Top + 60
End If
Me.Top = Me.Top + 30
Form3.Top = Form3.Top + 30
Me.Left = Me.Left + 30
Form3.Left = Form3.Left + 30
Me.Width = Me.Width - 60
Form3.Width = Form3.Width - 60
Me.Height = Me.Height - 60
Form3.Height = Form3.Height - 60
Loop Until Me.Top > Screen.Width
End

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = vbWhite
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = vbRed
'look at Frame4_mousemove to see how to change it back
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If ProgressBar1.Value < 100 Then
ProgressBar1.Value = ProgressBar1.Value + 2
Else
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If ProgressBar1.Value > 1 Then
ProgressBar1.Value = ProgressBar1.Value - 2
Else
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub
