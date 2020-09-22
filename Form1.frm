VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Using the mouse wheel in vb - Form1"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   195
      TabIndex        =   3
      Text            =   "Placing the focus here will work"
      Top             =   960
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
      Text            =   "Place your mouse in either text box and use the mouse wheel see the value change"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mouse wheel by Ben Jones
' Email dreamvb@yahoo.com
' I hope you find this code helpfull
' Modification for multiple forms and entire form subclassing by Al Klein 4/17/2003
' Email rukbat@optonline.net
'
' Thanks for the idea, Ben.  I needed to be able to do this.

Private Sub Command1_Click()
    Unload Form2    ' unload the forms
    Unload Form1
End Sub

Private Sub Form_Activate()
' this form will see the mouse wheel regardless of which control is the focus
End Sub

Private Sub Form_Load()
    Form2.Show
    Form2.Top = 8800
    OldProc1 = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf TWndProc1 ' Subclass the entire form
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong Me.hWnd, GWL_WNDPROC, OldProc1
    Set Form1 = Nothing ' Release the form from memory
End Sub


