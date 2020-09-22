Attribute VB_Name = "modScreenSaver"
'//// If you are using Notepad make sure Word wrap is on ////

'This program was created by Carl Warwick 1999 using VB5
'professional edition.
'You can use this program for any purpose as long as no
'profit is made.

'For a free game created in VB go to http://www.parkstonemot.freeserve.co.uk/indexfw.htm
'If your into Motocross then go to http://www.dirtrider.net/freeride

'You are free to alter this program or use it in your own
'programs as long as I'm given credit for it.

'This program demonstrates how to create a screen saver in
'Visual Basic, it also shows how to make scrolling credits
'that fade in and out of the screen.

'The Fading, Scrolling credits:-
' Useful for end credits for a game etc.
' Easy to adapt for you own programs.
' Only one label and one timer are required at design time.

'Creating your own Screen Saver:-
' In the module you'll find the code used to make the screen
'   saver, without this code your screen saver will just run
'   straight away without you even selecting it.
' It also shows you how to use a settings screen, and saves
'   these settings to control.ini in windows which is where
'   the information for all screen savers is saved.

' Once you have made your screen saver make it into an exe file.
' Change the file extension to .scr
' Right click it and click install
' and there you have it, your very own screen saver...
 
'E-Mail me at freeride@dirtrider.net

Type RGB
    R As Integer
    G As Integer
    B As Integer
End Type

Public CreditText(20) As String, NumLines As Integer, i As Integer
Public Colo As RGB, INIFound As Boolean


Sub Main()

' This is the number of lines of text used
NumLines = 6

'These are the strings returned for use below
'/a = password
'/s = run
'/p = end
'/c = config

'This is the code needed for a screen saver...
    Dim strCmdLine As String
    strCmdLine = Left(Command, 2)
        
    If strCmdLine = "/p" Then
        End
        'on select of your screen saver
    ElseIf strCmdLine = "/c" Then
        Call Load(frmConfig)
        frmConfig.Show
        'Function to call when
        'Settings'
        'button is pushed
    Else
        Call Load(FrmMain)
        FrmMain.Show
        'your screen saver has been loaded
    End If


End Sub

