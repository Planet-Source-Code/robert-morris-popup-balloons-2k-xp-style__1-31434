VERSION 5.00
Begin VB.Form frmSample2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample 2: Calculator"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmSample2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "&Divide!"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1170
      Width           =   1095
   End
   Begin VB.TextBox txtSecondNumber 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFirstNumber 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   $"frmSample2.frx":0442
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Closes this and returns to the previous example."
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label lblQuotient 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2415
      TabIndex        =   7
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label lblEqualsSign 
      Caption         =   "="
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label lblDivisionSign 
      Caption         =   "/"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Enter two numbers to divide below, and press the ""Divide!"" button:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmSample2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
frmSample.Show
End Sub
Private Sub cmdDivide_Click()
'Variables used if balloon is shown:
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

' Check for division by zero before we start to divide, because
' it would cause an error if we didn't
If txtSecondNumber.Text = "0" Then 'Show the error balloon

    Call GetWindowRect(txtSecondNumber.hwnd, WinRect)

    BalloonXY.X = WinRect.Left * Screen.TwipsPerPixelX
    BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY
    Dim frmDivisionByZero As New frmBalloon

    frmDivisionByZero.SetBalloon "Division by Zero", "You have attempted a " & _
        "division by zero, which is not possible. Please check your second number " & _
        "and try again.", BalloonXY.X, BalloonXY.Y, _
        "x", True, 10000
    
    frmDivisionByZero.Show , Me
    
    Me.SetFocus
    Exit Sub
End If
  
' Divide the two numbers
On Error GoTo DivisionError
lblQuotient.Caption = Me.txtFirstNumber.Text / Me.txtSecondNumber.Text
Exit Sub

DivisionError:

Call GetWindowRect(cmdDivide.hwnd, WinRect)
BalloonXY.X = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmDivisionError As New frmBalloon

frmDivisionError.SetBalloon "Division Error", "An error occured trying to " & _
    "divide the numbers. Make sure they are both valid numbers, don't contain " & _
    "letters, etc., and try again." & vbCrLf & vbCrLf & "This also " & _
    "demonstrates a balloon with a different font!", BalloonXY.X, BalloonXY.Y, _
    "x", True, 10000, 1720, , "Arial"
    
frmDivisionError.Show , Me
Me.SetFocus

End Sub

Private Sub txtFirstNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtSecondNumber.SetFocus
End If
End Sub

Private Sub txtSecondNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdDivide_Click
End If
End Sub
