VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fishing Game - 0 Fish Caught"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5400
      Top             =   5520
   End
   Begin VB.Frame frameDepth 
      Caption         =   "Depth"
      Height          =   1815
      Left            =   2520
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
      Begin VB.OptionButton opt5FT 
         Caption         =   "5 Feet"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton opt10FT 
         Caption         =   "10 Feet"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton opt3FT 
         Caption         =   "3 Feet"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton opt1FT 
         Caption         =   "1 Foot"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame frameBait 
      Caption         =   "Bait"
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   2055
      Begin VB.OptionButton optSquid 
         Caption         =   "Squid"
         Height          =   855
         Left            =   1080
         Picture         =   "Form1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optBread 
         Caption         =   "Bread"
         Height          =   855
         Left            =   120
         Picture         =   "Form1.frx":06CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optWorm 
         Caption         =   "Worm"
         Height          =   495
         Left            =   1080
         Picture         =   "Form1.frx":0A79
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optHotDog 
         Caption         =   "Hot Dog"
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":0D83
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   3600
      Top             =   1200
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   3720
      TabIndex        =   11
      Top             =   5160
      Width           =   3210
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3840
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image imgFish 
      Height          =   480
      Left            =   4320
      Picture         =   "Form1.frx":108D
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgReel3 
      Height          =   135
      Left            =   720
      Picture         =   "Form1.frx":1397
      Top             =   3600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgReel2 
      Height          =   135
      Left            =   720
      Picture         =   "Form1.frx":16F2
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgReel1 
      Height          =   165
      Left            =   720
      Picture         =   "Form1.frx":1A48
      Top             =   3120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgReel 
      Height          =   165
      Left            =   75
      Picture         =   "Form1.frx":1DA7
      Top             =   6075
      Width           =   180
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   1920
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Image imgBobber 
      Height          =   150
      Left            =   1920
      Picture         =   "Form1.frx":2106
      Top             =   2600
      Width           =   135
   End
   Begin VB.Image imgRod 
      Height          =   3825
      Left            =   0
      Picture         =   "Form1.frx":245C
      Top             =   3120
      Width           =   135
   End
   Begin VB.Image imgBackground 
      Height          =   4980
      Left            =   0
      Picture         =   "Form1.frx":28DB
      Top             =   0
      Width           =   7005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FishType  'Creates a type called 'FishType'
LocX As Integer        'which makes it much easier
LocY As Integer        'to have the fish 'swimming'
SeaDepth As Integer    'around and much easier to
Bait As String         'set the properties
IsCaught As Boolean    '(Bait,Location,Depth,etc.)
End Type

Dim Fish(20) As FishType
Dim CanGo2 As Boolean
Dim CanGo As Boolean
Dim Casting As Boolean    'All the variables, of course
Dim Reeling As Boolean
Dim CurX As Single
Dim CurY As Single
Dim INum As Integer
Dim FishCaught As Integer

'My notes ;)

'x1=imgbobber.left+20
'y1=imgbobber.top+60
'6840=max left
'2910=max top

Sub Pause(Duration)       'A little 'Pausing' routine
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub


Sub ResetFish()       'Sub to reset all of the fish to
For i = 0 To 20       'not caught
Fish(i).IsCaught = False 'Actually, the IsCaught
Next i                   'property is almost useless
imgFish.Visible = False  'and was implemented barely
End Sub                  'and never removed.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode       'This is all just for
Case vbKeyQ               'the Debug label.
Label1.Visible = True     'When I was making the game,
Timer3.Enabled = True     'I was having many problems...
lblStatus.Caption = "Debug code activated."
Case vbKeyE               'So I put a label in so I
Label1.Visible = False    'could look at what was going on.
Timer3.Enabled = False
lblStatus.Caption = "Debug code deactivated."
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Reeling = False
Casting = True
imgBobber.Left = 1920
imgBobber.Top = 3100
Line1.X1 = 120      'Nothing special, just sets up some
Line1.X2 = 1920     'stuff in the Load event.
Line1.Y1 = 3120
Line1.Y2 = 3100
End Sub

Private Sub imgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Ok, this is confusing to me!
'Let's see if I can get this right!

If Casting = True And Reeling = False Then
Timer2.Enabled = False  'This makes sure that a fish
                        'is not being reeled in
                        'at the moment

If Y <= 2600 Or Y >= 4500 Then
CanGo = False                 'This will keep it
imgBobber.Top = imgBobber.Top 'within the boundaries
Else                          'of the water,
CanGo = True                  'so you're not fishing
End If                        'on the rocks!

If X >= 6700 Then
CanGo2 = False                  'Same with this, only
imgBobber.Left = imgBobber.Left 'this time it keeps the
Else                            'Left within the boundaries.
CanGo2 = True
End If

If CanGo2 = True Then
imgBobber.Left = X + 150  'These two will make it attach
End If                    'to the mouse pointer
                          'if it is within the boundaries.
If CanGo = True Then
imgBobber.Top = Y + 300
End If

CurY = Y 'I was having trouble with something, so I made
CurX = X 'it put the X and Y into variables, but I don't
         'think that they are used anywhere important.
If CanGo2 = True Then
Line1.X1 = X + 210 'This attaches it to the bobber.
End If

Line1.X2 = imgRod.Left + 80 'Keeps the Line (the pun was
Line1.Y2 = imgRod.Top + 10  'not intentional) on the rod.

If CanGo = True Then
Line1.Y1 = Y + 320 'This also attaches it to the bobber.
End If

ElseIf Casting = False And Reeling = True Then
Timer1.Enabled = True  'I believe this is when a fish
Timer2.Enabled = False 'has been caught.

ElseIf Casting = False And Reeling = False Then
Timer2.Enabled = True 'This is when the user has clicked
End If                'and 'casted' the line.

End Sub

Private Sub imgBackground_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton And Reeling = False And Casting = False Then
lblStatus.Caption = "Try another spot?"
Casting = True 'This is when the user right-clicks
Reeling = Fals 'and can cast to another location.
Timer1.Enabled = False
Timer2.Enabled = False
frameBait.Enabled = True
frameDepth.Enabled = True
ElseIf Button = vbLeftButton Then

If Casting = False And Reeling = True Then
Timer2.Enabled = False
imgFish.Left = imgFish.Left - 20
'This is when the person has caught a fish and is reeling
'it in. Note that if the user clicks on the fish, nothing
'will happen... that would make the code even more
'confusing! :)

If imgFish.Left <= 0 Then

'All of this stuff below is when the user has brought
'the fish all the way over to the left and therefore has
'caught it. It resets the line, adds a fish to the score,
'and re-enables the controls on the bottom.



Casting = True
Reeling = False
imgFish.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Line1.Visible = True
imgBobber.Visible = True
FishCaught = FishCaught + 1
frameBait.Enabled = True
lblStatus.Caption = "Congratulations! You have successfully caught a fish! You now have " & FishCaught & " total fish."
Form1.Caption = "Fishing Game - " & FishCaught & " Fish Caught"
Timer1.Enabled = False
Timer2.Enabled = False
Casting = True
Reeling = False
frameBait.Enabled = True
frameDepth.Enabled = True
Line1.Visible = True
imgBobber.Visible = True
imgFish.Visible = False

'Pretty annoying code, huh? :)

ResetFish
Exit Sub
End If

ElseIf Casting = True And Reeling = False Then
lblStatus.Caption = "Line casted!"
Casting = False
Reeling = False
frameBait.Enabled = False
frameDepth.Enabled = False
Timer2.Enabled = False

'This is when the user has clicked to cast the line,
'since he is Casting and not Reeling therefore he
'is not reeling in a fish and is not not Reeling and
'he is not not Casting either, so he has not just
'casted a line and is not idle and...

'*Gasp*

'Goodness I talk too much and use too many double
'negatives!

'Anyways...


End If

End If
End Sub

Private Sub Timer1_Timer()

'I have almost NO idea what is going on in this timer!
'All I did was change a bunch of things until it worked!
'I'll try to specify what is going on as best as I can...

If Fish(Index).IsCaught = True Then

'I don't think that property is even used!

'??

'FishCaught = FishCaught + 1
'lblStatus.Caption = "Congratulations! You have successfully caught a fish! You now have " & FishCaught & " total fish."

'??

Timer1.Enabled = False
Timer2.Enabled = False
Casting = True
Reeling = False
frameBait.Enabled = True
frameDepth.Enabled = True
Line1.Visible = True
imgBobber.Visible = True
imgFish.Visible = False
Else

If imgFish.Visible = False Then imgFish.Visible = True
'Well this is obviously what happens when you catch a fish...
lblStatus.Caption = "Quick! Click as fast as you can on the background to reel in the fish!"
Randomize
RndN = Int((Rnd * 20) + 1)
If imgFish.Left >= 7000 Then
'Oh no, it got away! :(
Timer1.Enabled = False
Timer2.Enabled = False
Casting = True
Reeling = False
frameBait.Enabled = True
frameDepth.Enabled = True
Line1.Visible = True
imgBobber.Visible = True
imgFish.Visible = False
lblStatus.Caption = "Tough luck, it got away!"
Exit Sub
Else
'Notice how it goes away in random intervals?
'Just another boring detail in Burbble's programming :)
imgFish.Left = imgFish.Left + RndN
End If
imgReel.Picture = imgReel1.Picture
Pause 0.01 + INum
imgReel.Picture = imgReel2.Picture
Pause 0.01 + INum
imgReel.Picture = imgReel3.Picture
Pause 0.01 + INum
imgReel.Picture = imgReel2.Picture
Pause 0.01 + INum
INum = INum + 0.001
'Reeling animation... Badly done, but hey, it works!
End If
End Sub

Private Sub Timer2_Timer()

'Everything in this timer generates the fish
'and checks if the user's line and the fish are matched.


If Casting = False And Reeling = False Then

Randomize

For i = 0 To 20
RNum1 = Int((Rnd * 2) + 1)
RNum2 = Int((Rnd * 2) + 1)
If RNum1 = 1 Then RNum1 = 0
If RNum1 = 2 Then RNum1 = 5
If RNum2 = 1 Then RNum2 = 0
If RNum2 = 2 Then RNum2 = 5

'Below it generates the locations, bait, and sea depth
'of the fish.

Fish(i).LocX = Int((Rnd * 690) + 0) & RNum1
Fish(i).LocY = Int((Rnd * 480) + 0) & RNum2
rndnum1 = Int((Rnd * 4) + 1)
If rndnum1 = 1 Then
Fish(i).Bait = "hotdog"
End If
If rndnum1 = 2 Then
Fish(i).Bait = "worm"
End If
If rndnum1 = 3 Then
Fish(i).Bait = "bread"
End If
If rndnum1 = 4 Then
Fish(i).Bait = "squid"
End If

rndnum1 = Int((Rnd * 4) + 1)

If rndnum1 = 1 Then
Fish(i).SeaDepth = 1
End If
If rndnum1 = 2 Then
Fish(i).SeaDepth = 3
End If
If rndnum1 = 3 Then
Fish(i).SeaDepth = 5
End If
If rndnum1 = 4 Then
Fish(i).SeaDepth = 10
End If

Next i

'Now it checks for any matches...

'I have NO idea why I put it in Abs(), but it works
'so I'm not complaining or changing it :)

For i = 0 To 20
If Line1.X1 = Abs(Fish(i).LocX) Or Line1.Y1 = Abs(Fish(i).LocY) Then
If Fish(i).Bait = "hotdog" And optHotDog.Value = True Or Fish(i).Bait = "worm" And optWorm.Value = True Or Fish(i).Bait = "bread" And optBread.Value = True Or Fish(i).Bait = "squid" And optSquid.Value = True Then
If Fish(i).SeaDepth = 1 And opt1FT.Value = True Or Fish(i).SeaDepth = 3 And opt3FT.Value = True Or Fish(i).SeaDepth = 5 And opt5FT.Value = True Or Fish(i).SeaDepth = 10 And opt10FT.Value = True Then
imgFish.Top = imgBobber.Top
imgFish.Left = imgBobber.Left
imgBobber.Visible = False
Line1.Visible = False
Reeling = True
Casting = False
'Fish(i).IsCaught = True
'See? Told ya that feature wasn't being used!
Timer2.Enabled = False
Timer1.Enabled = True
lblStatus.Caption = "You caught a fish!"
'You caught a fish. Woo.
Exit Sub
End If
End If
End If
Next i

'What on earth?

'I suppose it checks, once the fish has been caught,
'if it can disable the detection timer and return
'to normal.
ElseIf Casting = False And Reeling = True Or Casting = True And Reeling = False Then
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
'!!!!!
'Tell Burbble you found the ancient debug timer and
'to give you super-duper extra fish points!
'!!!!!

For i = 0 To 20
Label1.Caption = Fish(i).LocX & "," & Fish(i).LocY & " " & Fish(i).IsCaught & "," & Fish(i).Bait & "," & Fish(i).SeaDepth & " " & Line1.X1 & "," & Line1.Y1 & " " & Abs(imgFish.Top) & "," & Abs(imgFish.Left) & " " & Casting & "," & Reeling
'Pause 1

'Only enable Pause 1 if you want it to pause each time
'it updates the debug label so you can actually read it,
'but it will be more out-of-date.

Next i
End Sub
