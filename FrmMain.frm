VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   1920
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   2580
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

' Get information
    Open "c:\Windows\Control.ini" For Input As #1
    
    INIFound = False
    'This is the same as form_load on frmconfig
    Do While Not EOF(1)
        Line Input #1, temp 'Read Data
        If temp = "[Screen Saver.CarlsCredits]" Then 'If settings found
            INIFound = True
            
            '10 lines to read
            For i = 0 To 6
                Line Input #1, CreditText(i) 'Get text
            Next i
            Line Input #1, temp 'Get colors
            Colo.R = Val(temp)
            Line Input #1, temp
            Colo.G = Val(temp)
            Line Input #1, temp
            Colo.B = Val(temp)
            
            Exit Do 'Exit DO-LOOP when settings are found
        End If
    Loop
    Close #1

' If settings isn't found
If INIFound = False Then
    CreditText(0) = "Carls" 'Use Default Values
    CreditText(1) = "Scrolling Credits" 'Use Default Values
    CreditText(2) = "Screen Saver" 'Use Default Values
    CreditText(3) = "For a free game made in VB visit" 'Use Default Values
    CreditText(4) = "http://www.parkstonemot.freeserve.co.uk" 'Use Default Values
    CreditText(5) = "For Motocross visit" 'Use Default Values
    CreditText(6) = "http://www.dirtrider.net/freeride" 'Use Default Values
    
    Colo.R = 255 'Use Default Values
    Colo.G = 125 'Use Default Values
    Colo.B = 80 'Use Default Values
End If

' Set background to black
FrmMain.BackColor = RGB(0, 0, 0)

lblCredits(0).Top = FrmMain.ScaleHeight 'Set first label just out of sight
lblCredits(0).Left = (FrmMain.ScaleWidth / 2) - (lblCredits(0).Width / 2) 'Centre first label

For i = 1 To NumLines
    Load lblCredits(i) 'Create new labels
    lblCredits(i).Top = lblCredits(i - 1).Top + lblCredits(i - 1).Height        'Set labels vertical position
    lblCredits(i).Left = (FrmMain.ScaleWidth / 2) - (lblCredits(i).Width / 2)   'Centre label
    lblCredits(i).Visible = True    'Show label
Next i

For i = 0 To NumLines
    lblCredits(i).Caption = CreditText(i)   'Set labels Text
Next i
End Sub

'End if mouse or key is pressed
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
End
End Sub

'Not used in the screen saver, but sets the labels to their
'correct positions if screen is resized.
Private Sub Form_Resize()
lblCredits(0).Top = FrmMain.ScaleHeight
lblCredits(0).Left = (FrmMain.ScaleWidth / 2) - (lblCredits(0).Width / 2)

For i = 1 To NumLines
    lblCredits(i).Top = lblCredits(i - 1).Top + lblCredits(i - 1).Height
    lblCredits(i).Left = (FrmMain.ScaleWidth / 2) - (lblCredits(i).Width / 2)
Next i

End Sub

Private Sub Timer1_Timer()
For i = 0 To NumLines
    'Check if label is out of the top
    If lblCredits(i).Top <= -lblCredits(i).Height Then lblCredits(i).Top = FrmMain.ScaleHeight

    'Sets the color for the text
    If lblCredits(i).Top <= FrmMain.ScaleHeight And lblCredits(i).Top >= FrmMain.ScaleHeight / 2 Then
        lblCredits(i).ForeColor = RGB(((FrmMain.ScaleHeight - lblCredits(i).Top) / (FrmMain.ScaleHeight / 2)) * Colo.R, ((FrmMain.ScaleHeight - lblCredits(i).Top) / (FrmMain.ScaleHeight / 2)) * Colo.G, ((FrmMain.ScaleHeight - lblCredits(i).Top) / (FrmMain.ScaleHeight / 2)) * Colo.B)
    End If
    If lblCredits(i).Top <= FrmMain.ScaleHeight / 2 And lblCredits(i).Top >= 0 Then
        lblCredits(i).ForeColor = RGB((lblCredits(i).Top) / (FrmMain.ScaleHeight / 2) * Colo.R, (lblCredits(i).Top) / (FrmMain.ScaleHeight / 2) * Colo.G, (lblCredits(i).Top) / (FrmMain.ScaleHeight / 2) * Colo.B)
    End If
    
    'Moves the label up
    lblCredits(i).Top = lblCredits(i).Top - 3
Next i
End Sub
