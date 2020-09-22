Attribute VB_Name = "Module1"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_MOUSEWHEEL = &H20A ' window message for mouse wheel
Private MouseWheelUp As Boolean     ' true if mouse wheel up, false if down
Private I As Long, J As Long        ' used to hold our counter value
Public Const GWL_WNDPROC = (-4)
Public OldProc1 As Long             ' Holds the old TWndProc for form 1
Public oldproc2 As Long             ' Holds the old TWndProc for form 2

Public Function TWndProc1(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'  If hWnd = Form1.txt.hWnd Then
    If wMsg = WM_MOUSEWHEEL Then ' Have we got a mouse wheel message
    
    If wParam > 0 Then MouseWheelUp = True Else MouseWheelUp = False
    
    Select Case MouseWheelUp
        Case True ' mouse up value is found
            Form1.lblmouseval.Caption = "The value of the mouse wheel is set to Up" ' update the label caption
            I = I + 1 ' update our counter
        Case False ' mouse value down is found
            Form1.lblmouseval.Caption = "The value of the mouse wheel is set to Down" ' update the label caption
            I = I - 1 ' update our counter
            'If I <= 0 Then I = 0 'reset our counter if below zero
    End Select
        Form1.txt.Text = "The value of the mouse wheel is " & I ' update the text in the textbox
    End If
'  End If
  TWndProc1 = CallWindowProc(OldProc1, hWnd, wMsg, wParam, lParam)

End Function
Public Function TWndProc2(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'If hWnd = Form2.txt.hWnd Then
    If wMsg = WM_MOUSEWHEEL Then
    
    If wParam > 0 Then MouseWheelUp = True Else MouseWheelUp = False
    
    Select Case MouseWheelUp
        Case True
            Form2.lblmouseval.Caption = "The value of the mouse wheel is set to Up"
            J = J + 1
        Case False
            Form2.lblmouseval.Caption = "The value of the mouse wheel is set to Down"
            J = J - 1
            'If J <= 0 Then J = 0 'reset our counter if below zero
    End Select
        Form2.txt.Text = "The value of the mouse wheel is " & J
    End If
'  End If
  TWndProc2 = CallWindowProc(oldproc2, hWnd, wMsg, wParam, lParam)

End Function

