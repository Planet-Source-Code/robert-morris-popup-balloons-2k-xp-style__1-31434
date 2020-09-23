VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Balloons Form"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd3 
      Caption         =   "Balloon 3"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Balloon 2"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Balloon 1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtInformation 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmSample.frx":0000
      Top             =   1920
      Width           =   4935
   End
   Begin VB.CommandButton cmdExample2 
      Caption         =   "&Another Example"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdPopIt 
      Caption         =   "&Pop It Up!"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblMore 
      Caption         =   "The buttons below will also display some more styles of pop-up balloons. Click them!"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lblExample2 
      Caption         =   "For another, more practical example, click:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblInformation 
      Caption         =   "Click the button to pop up the sample balloon to see what it looks like:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd1_Click()
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd1.hwnd, WinRect)
BalloonXY.X = (WinRect.Left * Screen.TwipsPerPixelX) + (cmd1.Width / 2)
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY - (cmd1.Height / 2)

Dim frmBalloon1 As New frmBalloon

frmBalloon1.SetBalloon "Balloon One", "This is a balloon! My properties are set " & _
    "so that I have a close button, no icon, and will not automatically close " & _
    "after a certain amount of time. My coordinates are also set so that I " & _
    "display in the middle of the ""Balloon 1"" button.", BalloonXY.X, _
    BalloonXY.Y, , True
    
frmBalloon1.Show , Me
Me.SetFocus
End Sub
Private Sub cmd2_Click()
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd2.hwnd, WinRect)
BalloonXY.X = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmBalloon2 As New frmBalloon

frmBalloon2.SetBalloon "Balloon Two", "I am Balloon 2. My properties are set " & _
    "so that I do not display a close (X) button, will auto-close after ten " & _
    "seconds, have a custom height and width, appear next to the ""Balloon 2"" " & _
    "button, and display a 9x-style ""!"" icon", _
    BalloonXY.X, BalloonXY.Y, "!9", , 10000, 2500, 2100
    
frmBalloon2.Show , Me
Me.SetFocus
End Sub

Private Sub cmd3_Click()

Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd1.hwnd, WinRect)
BalloonXY.X = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmBalloon3 As New frmBalloon

frmBalloon3.SetBalloon "Balloon Three", "I am Balloon 3. I am set to auto-" & _
    "close after fifteen seconds, display a close (X) button, show an (XP-style" & _
    ") ""i"" icon, appear lined up with the first button in this row (but be " & _
    "about centered on this form), and show using a custom font, Times New Roman.", _
    BalloonXY.X, BalloonXY.Y, "i", True, 15000, , 5000, "Times New Roman", 9
 
frmBalloon3.Show , Me
Me.SetFocus

End Sub

Private Sub cmdExample2_Click()
frmSample2.Show
Unload Me
End Sub
Private Sub cmdPopIt_Click()
Dim WinRect As RECT         'These are used to hold some values we
Dim WinPoint As POINTAPI    'get during the API calls, and for
Dim BalloonXY As BalloonCoords 'storing the X and Y coordinates
                            'of the balloon that we pass when showing it

'This code is used to determine the position of the balloon. We (usually)
'want it to be displayed near some type of object, and since we need to
'set the balloon's coordinates relative to the screen, not the form, we
'need to determine the screen position of the control by which we want to
'place the balloon so it will show in the right spot.

'Get coordinates of the object on it that we want to
'display the ballooon by
Call GetWindowRect(cmdPopIt.hwnd, WinRect) 'When you use this code, replace
                                    'cmdPopIt with whatever control you
                                    'want to place the balloon by.
                                    
'This is multiplied by TwipsPerPixel because VB works
'with twips by default, but the API works in pixels. We'll be assigning
'the X and Y coordinates we get (which will be the coordiate for the lower
'left-hand corner of the control we chose above) to a BalloonXY (with .X
'and .Y properties) type object so we can easily use these coordinates
'later when we call to show the balloon.
'You can just assign them to two variables or whatever if you like, or
'just use the "formula" for figuring them directly when you call SetBalloon()
'instead of calculating them and then holding them in a variable.

BalloonXY.X = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmPopUpBalloon As New frmBalloon 'Make a new form based on frmBalloon

frmPopUpBalloon.SetBalloon "Sample Balloon", "This is a sample balloon to " & _
    "demonstrate the capabilities of the pop-up balloon/tooltips that you can " & _
    "use in your programs! They can include a title, multi-line text, an optional " & _
    "close button, automatiically close after a certain amount of time, " & _
    "an icon, and show in any font!", BalloonXY.X, BalloonXY.Y, "i", True, _
    10000               'These preceeding lines set the properties
                        '(text, etc.) for the balloon. For the list of arguments
                        ' and what their values can be and what they do, see
                        ' frmBalloon's SetBalloon() procedure.
    
   
frmPopUpBalloon.Show , Me 'Show the balloon, with me as the owner

Me.SetFocus 'Since the balloon is a window (a form), showing it will
            'take focus away from this form, which it's called from,
            'and we don't want that to happen. We're working around it
            'by giving me focus after showing it. There IS away to show
            'a window without giving it focus via API, but I haven't
            'gotten that to work yet.
End Sub

Private Sub txtInformation_KeyPress(KeyAscii As Integer)
'Show a balloon if you try to type in the textbox (its Locked property
'is true, so you can't edit it anyway, but this will tell you if you try)

End Sub
