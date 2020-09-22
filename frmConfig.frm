VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Scrolling Credits Setup"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Colour"
      Height          =   2295
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      Begin VB.HScrollBar ColBar 
         Height          =   255
         Index           =   2
         LargeChange     =   30
         Left            =   120
         Max             =   255
         TabIndex        =   13
         Top             =   1680
         Value           =   255
         Width           =   1575
      End
      Begin VB.HScrollBar ColBar 
         Height          =   255
         Index           =   1
         LargeChange     =   30
         Left            =   120
         Max             =   255
         TabIndex        =   12
         Top             =   1080
         Value           =   255
         Width           =   1575
      End
      Begin VB.HScrollBar ColBar 
         Height          =   255
         Index           =   0
         LargeChange     =   30
         Left            =   120
         Max             =   255
         TabIndex        =   11
         Top             =   480
         Value           =   255
         Width           =   1575
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Blue"
         Height          =   195
         Left            =   660
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Green"
         Height          =   195
         Left            =   660
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Red"
         Height          =   195
         Left            =   660
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Text            =   "http://www.dirtrider.net/freeride"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Text            =   "For Motocross visit"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Text            =   "http://www.parkstonemot.freeserve.co.uk"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Text            =   "For a free game made in VB visit"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Text            =   "Screen Saver"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Text            =   "Scrolling Credits"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Text            =   "Carls"
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
' Unload the form if cancel is pressed
    Unload Me
End Sub

Private Sub cmdOK_Click()
' If OK is pressed then set the values
For i = 0 To NumLines
    CreditText(i) = frmConfig.txtText(i).Text   ' Set the text used
Next i
Colo.R = ColBar(0).Value ' Set the RGB colours
Colo.G = ColBar(1).Value
Colo.B = ColBar(2).Value

'Read File then change it
    Open "c:\Windows\Control.ini" For Input As #1
        Open "c:\Windows\ControlTemp.tmp" For Output As #2  'Open temporary File
        
        INIFound = False
        
        Do While Not EOF(1)     'Loop until End of file
            Line Input #1, temp 'Store data in temp
            
                If temp = "[Screen Saver.CarlsCredits]" Then    'First Line of Screen Saver settings
                    For i = 0 To 10
                    Line Input #1, temp 'Just read the lines to stop them being read later
                    Next i
                    
                    Print #2, "[Screen Saver.CarlsCredits]" 'Write Data to temporary file
                    
                    For i = 0 To 6
                    Print #2, CreditText(i) 'Write Text Data to temporary file
                    Next i
                    
                    Print #2, Format(Colo.R) 'Write Color Data to temporary file
                    Print #2, Format(Colo.G)
                    Print #2, Format(Colo.B)
                    INIFound = True 'Settings were found
                End If
                
                
            Print #2, temp  'Copy data from temp (Original data) into temporary file
        Loop
        
        ' If Settings isn't found
        If INIFound = False Then
            Print #2, ""
            Print #2, "[Screen Saver.CarlsCredits]" 'Write data to temporary file
                    
            For i = 0 To 6
                Print #2, CreditText(i)
            Next i
                    
            Print #2, Format(Colo.R)
            Print #2, Format(Colo.G)
            Print #2, Format(Colo.B)
            Print #2, ""
        End If
        
        Close #2    'Close files
    Close #1
    
    Kill "c:\Windows\Control.ini"   'Delete original File
    Name "c:\Windows\ControlTemp.tmp" As "c:\Windows\Control.ini"   'Rename temporary file to original name

    'Unload the form
    Unload Me
End Sub

'If the Scroll bars change then update colour
Private Sub ColBar_Change(Index As Integer)
    lblColor.BackColor = RGB(ColBar(0).Value, ColBar(1).Value, ColBar(2).Value)
End Sub
Private Sub ColBar_Scroll(Index As Integer)
    lblColor.BackColor = RGB(ColBar(0).Value, ColBar(1).Value, ColBar(2).Value)
End Sub

Private Sub Form_Load()

' Get information
    Open "c:\Windows\Control.ini" For Input As #1
    
    INIFound = False
    
    Do While Not EOF(1)
        Line Input #1, temp 'Put data into temp
        If temp = "[Screen Saver.CarlsCredits]" Then    'Start of settings found
            INIFound = True
            
            '10 lines to read
            For i = 0 To 6
                Line Input #1, CreditText(i)    'Read the text
            Next i
            Line Input #1, temp 'Read colours
            Colo.R = Val(temp)  'Put colours into variable
            Line Input #1, temp
            Colo.G = Val(temp)
            Line Input #1, temp
            Colo.B = Val(temp)
            
            Exit Do 'Exit DO LOOP if settings are found
        End If
    Loop
    Close #1    'Close file

' If settings isn't found
If INIFound = False Then
    CreditText(0) = "Carls" 'Set Default Values
    CreditText(1) = "Scrolling Credits" 'Set Default Values
    CreditText(2) = "Screen Saver" 'Set Default Values
    CreditText(3) = "For a free game made in VB visit" 'Set Default Values
    CreditText(4) = "http://www.parkstonemot.freeserve.co.uk" 'Set Default Values
    CreditText(5) = "For Motocross visit" 'Set Default Values
    CreditText(6) = "http://www.dirtrider.net/freeride" 'Set Default Values
    
    Colo.R = 255 'Set Default Values
    Colo.G = 125 'Set Default Values
    Colo.B = 80 'Set Default Values
End If

'Set text boxes with correct text
For i = 0 To 6
    frmConfig.txtText(i).Text = CreditText(i)
Next i

ColBar(0).Value = Colo.R 'Set scroll bars with correct value
ColBar(1).Value = Colo.G
ColBar(2).Value = Colo.B

End Sub
