VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conspiracy Generator v1.0.0"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10710
   Icon            =   "RANDOM~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10710
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10650
      TabIndex        =   12
      Top             =   5400
      Width           =   10710
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Reset Top List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtcounter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtmove 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4200
      Width           =   10455
   End
   Begin VB.ListBox lstrandom 
      Height          =   2535
      ItemData        =   "RANDOM~1.frx":0442
      Left            =   0
      List            =   "RANDOM~1.frx":0444
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   120
      Width           =   10575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Conspiracy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   2880
      TabIndex        =   14
      Top             =   2760
      Width           =   7455
   End
   Begin VB.Label Label7 
      Caption         =   "Conspiracy News Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "Conspiracies Generated"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Label4"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label3"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label2"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About & Disclaimer"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim totalnumber As Integer
Dim start As String
Dim names As String
Dim action As String
Dim counter As Integer
Dim output
Dim i
Dim ending As String


Private Sub Command1_Click()
'This picks a number for each area (label array 0 to 3)
Dim Number(4) As Integer
Dim Lowest, Highest As Integer
Lowest = 1
Highest = 22
Randomize
Number(0) = Int((Highest - Lowest + 1) * Rnd + Lowest)
Number(1) = Int((Highest - Lowest + 1) * Rnd + Lowest)
Number(2) = Int((Highest - Lowest + 1) * Rnd + Lowest)
Number(3) = Int((Highest - Lowest + 1) * Rnd + Lowest)

Label2.Caption = Number(1)
Label3.Caption = Number(2)
Label4.Caption = Number(3)
totalnumber = Number(0) + Number(1) + Number(2) + Number(3)
Label5.Caption = totalnumber
 
'start conspiracy generator

start = Number(0)
Label1.Caption = start
Select Case start
Case 1
Label1.Caption = "It was about 20 years ago that, "
Case 2
Label1.Caption = "It was reported last week that, "
Case 3
Label1.Caption = "During the early 90's it was reported that,  "
Case 4
Label1.Caption = "The networks are reporting that  "
Case 5
Label1.Caption = "The News is now reporting that about " & totalnumber & "  years ago,  "
Case 6
Label1.Caption = "Various News Reporters, who've been silenced, have spoken out and said that  "
Case 7
Label1.Caption = "About " & totalnumber & " years ago, "
Case 8
Label1.Caption = "British News sources revealed that "
Case 9
Label1.Caption = "At least " & totalnumber & " News Sources are now reporting that "
Case 10
Label1.Caption = "About " & totalnumber & " days ago, it was reported by good sources that  "
Case 11
Label1.Caption = "Reports show that "
Case 12
Label1.Caption = "Underground News Services are reporting that " & totalnumber & " days ago, "
Case 13
Label1.Caption = "Reuters reported today that "
Case 14
Label1.Caption = "AP News and Wire Service, in a small news bite, reported that "
Case 15
Label1.Caption = "The BBC reported yesterday, in a quick news story, that "
Case 16
Label1.Caption = "News Sources worldwide reported on Tuesday that, "
Case 17
Label1.Caption = "Europe Newswires are reporting that "
Case 18
Label1.Caption = "The White House is reporting that "
Case 19
Label1.Caption = "Right Wing Conservatives are saying that "
Case 20
Label1.Caption = "Moscow admits they know that "
Case 21
Label1.Caption = "Sources throughout Asia are reporting that "
Case 22
Label1.Caption = "Hidden Microphones throughout Russia revealed that "
End Select


names = Number(1)
Label2.Caption = names
Select Case names
Case 1
Label2.Caption = "Bill Clinton "
Case 2
Label2.Caption = "George Bush, Sr "
Case 3
Label2.Caption = "CIA Agents "
Case 4
Label2.Caption = "Tony Blair "
Case 5
Label2.Caption = "Al Gore "
Case 6
Label2.Caption = "North Korea agents "
Case 7
Label2.Caption = "The Commitee of 300 (a Secret Society of Elite Families) "
Case 8
Label2.Caption = "Green Peace "
Case 9
Label2.Caption = "Soviet KGB Agents "
Case 10
Label2.Caption = "Hillary Clinton "
Case 11
Label2.Caption = "various Ambassador's to the United Nations "
Case 12
Label2.Caption = "United Nations Leaders, "
Case 13
Label2.Caption = "David Rockefeller "
Case 14
Label2.Caption = "Colin Powell "
Case 15
Label2.Caption = "Norman Cousins "
Case 16
Label2.Caption = "Pope John Paul II "
Case 17
Label2.Caption = "Osama Bin Laden "
Case 18
Label2.Caption = "Dick Cheney "
Case 19
Label2.Caption = "George W Bush "
Case 20
Label2.Caption = "the Federal Reserve "
Case 21
Label2.Caption = "Tipper Gore "
Case 22
Label2.Caption = "Gray Davis (D-Ca-Gov)  "
End Select

action = Number(2)
Label3.Caption = action
Select Case action
Case 1
Label3.Caption = "planted listening devices in all hotel rooms around Washington DC and Moscow."
Case 2
Label3.Caption = "implemented Global Location Satellites and Transponders, worldwide! "
Case 3
Label3.Caption = "insisted on micro-chip implants in all humans, by the year 2005. "
Case 4
Label3.Caption = "lobbied for all U.S. cities to impliment roadway cameras for population control."
Case 5
Label3.Caption = "spyed on China, Russia, North Korea, and possibly the United States."
Case 6
Label3.Caption = "controlled the weather and every disaster worldwide."
Case 7
Label3.Caption = "demanded population control, wanting the worlds population reduced to 500,000 by 2010!"
Case 8
Label3.Caption = "used psychics to read the minds of criminal suspects while in prisons."
Case 9
Label3.Caption = "assisted double agents during the Vietnam war."
Case 10
Label3.Caption = "pushed for Thumbprint Identification for consumer purchases and new job applications."
Case 11
Label3.Caption = "pushed drugs into various countries to suppress the will of the people."
Case 12
Label3.Caption = "implemented International Biosphere Reserves in the U.S., giving U.S. OWNED Land away to the United Nations."
Case 13
Label3.Caption = "claimed that One World Government and the NEW WORLD ORDER is COMING."
Case 14
Label3.Caption = "set up the F.E.M.A. agency to control the United States and its people."
Case 15
Label3.Caption = "denied claims of black helicopters and foreign troops on American Soil."
Case 16
Label3.Caption = "oversaw the placement of UN concentration camps now in place throughout America."
Case 17
Label3.Caption = "had inside information on the bombing of the Oklahoma Federal Building, in Oklahoma City!"
Case 18
Label3.Caption = "had numerous Double Agents spy on U.S. TOP SECRET SITES and collect data."
Case 19
Label3.Caption = "covered up all documents that proved a One World Government was starting in Europe."
Case 20
Label3.Caption = "set up the Gorbhachav Foundation at the Presidio in San Franciso to oversee Environmental issues."
Case 21
Label3.Caption = "oversaw the movement of Foreign Troops in the United States."
Case 22
Label3.Caption = "introduced bio-chemical warfare against the American People during the 90's."
End Select

ending = Number(3)
Label4.Caption = ending

Select Case ending
Case 1
Label4.Caption = "  (News Service India)"
Case 2
Label4.Caption = " (News Service Bejing)"
Case 3
Label4.Caption = " (News from the South Pacific)"
Case 4
Label4.Caption = " (Conservative News Media)"
Case 5
Label4.Caption = " (BBC , GBC, FBC and FCC News Media)"
Case 6
Label4.Caption = " (Foreign News)"
Case 7
Label4.Caption = " (Northern News Views)"
Case 8
Label4.Caption = " (Liberal News Blockade)"
Case 9
Label4.Caption = " (News and Views, Canada)"
Case 10
Label4.Caption = " (Hawaii News International)"
Case 11
Label4.Caption = "(Underworld Daily)"
Case Else
Label4.Caption = " (Conspiracy Theory Generator)"
End Select

'counts by 1's in count box
counter = counter + 1
lstrandom.AddItem Label1.Caption & Label2.Caption & Label3.Caption & Label4.Caption
txtcounter.Text = counter
End Sub

Private Sub Command2_Click()
lstrandom.Clear
counter = 0
txtcounter.Text = ""
txtmove.Text = ""
End Sub

Private Sub Command3_Click()
lstrandom.Clear
counter = 0
txtcounter.Text = ""
End Sub



Private Sub lstrandom_Click()
'this puts checked conspiracy theory into the bottom form

output = lstrandom.ListCount
For i = 0 To output - 1
If lstrandom.Selected(i) Then txtmove.Text = lstrandom.Text
Next i
End Sub


Private Sub mnuabout_Click()
'first is the company then the disclaimer
MsgBox "8Ball Software (c) 2002 - Best in Hard to Find Software!!", 64, "About This Program"
'then this pops up
MsgBox "This Program is for ENTERTAINMENT PURPOSES ONLY", 16, "Disclaimer - No Truth To Any Generated Conspiracy"
End Sub

Private Sub mnuprint_Click()
'very simple printform feature
Form1.PrintForm
End Sub

Private Sub mnuquit_Click()
End
End Sub
