VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Using the mouse wheel in vb - Form2"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   195
      TabIndex        =   3
      Text            =   "Setting the focus here won't work"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   195
      TabIndex        =   1
      Text            =   "Place your mouse here and use the mouse wheel see the value change"
      Top             =   540
      Width           =   7485
   End
   Begin VB.Label lblmouseval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   45
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Form1    ' unload the forms
    Unload Form2
End Sub

Private Sub Form_Activate()
    Text1.SetFocus 'the mouse wheel won't do anything if the right tex box isn't the focus
End Sub

Private Sub Form_Load()
    oldproc2 = GetWindowLong(txt.hWnd, GWL_WNDPROC)
    SetWindowLong txt.hWnd, GWL_WNDPROC, AddressOf TWndProc2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong txt.hWnd, GWL_WNDPROC, oldproc2
    Set Form2 = Nothing
End Sub


