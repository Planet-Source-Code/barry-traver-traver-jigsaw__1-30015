VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJigsaw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traver Jigsaw"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6690
   Icon            =   "Jigsaw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4170
      Left            =   480
      Picture         =   "Jigsaw.frx":0442
      ScaleHeight     =   4110
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   940
      Visible         =   0   'False
      Width           =   6060
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      Height          =   4170
      Left            =   180
      ScaleHeight     =   4110
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   700
      Width           =   6060
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Visible         =   0   'False
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9578
            MinWidth        =   1058
            Text            =   "Congratulations!  Do another puzzle?"
            TextSave        =   "Congratulations!  Do another puzzle?"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Yes"
            TextSave        =   "Yes"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "No"
            TextSave        =   "No"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   873
      SimpleText      =   "Here's your puzzle!"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11757
            MinWidth        =   3528
            Text            =   "Here's your puzzle!"
            TextSave        =   "Here's your puzzle!"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadPicture 
         Caption         =   "&Load Picture"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScale 
      Caption         =   "&Scale"
      Begin VB.Menu mnuWhichScale 
         Caption         =   "&Normal"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuWhichScale 
         Caption         =   "&Fit Screen"
         Index           =   2
      End
   End
   Begin VB.Menu mnuCondition 
      Caption         =   "&Condition"
      Begin VB.Menu mnuScramblePicture 
         Caption         =   "&Scrambled "
      End
      Begin VB.Menu mnuUnscramblePicture 
         Caption         =   "&Unscrambled "
      End
   End
   Begin VB.Menu mnuSize 
      Caption         =   "&Pieces"
      Begin VB.Menu mnuSetSOS 
         Caption         =   "2 x 2  (4 pieces)"
         Index           =   2
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "3 x 3  (9 pieces)"
         Index           =   3
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "4 x 4  (16 pieces)"
         Index           =   4
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "5 x 5  (25 pieces)"
         Index           =   5
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "6 x 6  (36 pieces)"
         Index           =   6
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "7 x 7  (49 pieces)"
         Index           =   7
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "8 x 8  (64 pieces)"
         Index           =   8
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "9 x 9  (81 pieces)"
         Index           =   9
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "10 x 10  (100 pieces)"
         Index           =   10
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "11 x 11  (121 pieces)"
         Index           =   11
      End
      Begin VB.Menu mnuSetSOS 
         Caption         =   "12 x 12  (144 pieces)"
         Index           =   12
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "&Mode"
      Begin VB.Menu mnuDragMode 
         Caption         =   "&Drag Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSwapMode 
         Caption         =   "&Swap Mode"
      End
   End
   Begin VB.Menu mnuStatusLine 
      Caption         =   "S&tatus"
      Begin VB.Menu mnuReportProgress 
         Caption         =   "&Show"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Traver Jigsaw"
      End
      Begin VB.Menu mnuHowToPlay 
         Caption         =   "&How to Play"
      End
   End
End
Attribute VB_Name = "frmJigsaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************'
'                       Traver Jigsaw                           '
'                        written by                             '
'                      Barry A. Traver                          '
'                 ( email: btraver@traver.org )                 '
'                                                               '
'  You are free to use this source code in your own private,    '
'  non-commercial programs without permission, provided that    '
'  proper credit is given in your source code and compiled      '
'  program.  If you want to use this code in programs that      '
'  you intend to distribute to others, explicit permission      '
'  from the author should be obtained.                          '
'                                                               '
'  Traver Jigsaw is a simple jigsaw puzzle program intended     '
'  for children of all ages.  The author (a non-professional    '
'  Visual Basic programmer at this point) may be contacted by   '
'  email at btraver@traver.org (as above), and he would love    '
'  to receive comments (or questions) related to this program.  '
'                                                               '
'  The squirrel photograph is public domain.  I downloaded it   '
'  from http://www.pdimages.com/03702.html-ssi on the Public    '
'  Domain Images site.  (I had wanted to use a Garfield image,  '
'  but I was unable to confirm that it was public domain.)      '
'                                                               '
'  In the code, the squares are numbered as you might read a    '
'  book:  first you read the first line from left to right,     '
'  then the second line from left to right, and so on.          '
'                                                               '
'  1 2 3     Here, for example, is how a 3 x 3 puzzle would     '
'  4 5 6     be numbered.                                       '
'  7 8 9                                                        '
'                                                               '
'  Contrary to appearances, Traver Jigsaw makes use of only     '
'  three Picture controls (one always invisible, one always     '
'  not, and one sometimes visible and sometimes not).  No       '
'  use has been made of any Image or ImageList controls.  The   '
'  three Picture controls are used in such a way to create      '
'  the illusion of the existence of separate puzzle pieces.     '
'  In reality, it is always the same small Picture control      '
'  (repainted as appropriate) that is getting moved.  Even      '
'  that Picture box is only visibly moved when Drag mode is     '
'  chosen rather than Swap mode.                                '
'                                                               '
'  The PaintPicture method is what does most of the work in     '
'  the program.  The simple but powerful approach taken here    '
'  should be useful in other situations (such as boardgame      '
'  programs) where the movement of rectangular pieces or of     '
'  pieces of other shapes may be involved.  (For irregular      '
'  shapes, masking of the Picture control will be required.)    '
'                                                               '
'  Although Traver Jigsaw has many features, the source code    '
'  is rather short.  Program code need not be especially long   '
'  or complicated to accomplish many important tasks.  That is  '
'  one thing I learned when I programmed long ago on the old    '
'  Texas Instruments TI-99/4A, a computer that was well in      '
'  advance of its time.                                         '
'                                                               '
'  Back then, when I was a big frog in a little pond, I was     '
'  (without exaggeration) an internationally-known computer     '
'  guru (e.g., I regularly wrote the TI FORUM monthly column    '
'  for Computer Shopper).  Having moved from the TI to the      '
'  IBM, I'm now a little frog in a big pond, but one thing has  '
'  not changed:  I program for fun, and I enjoy the demands for '
'  invention and creativity placed on a person who is called to '
'  respond to programming challenges (especially when he is     '
'  largely self-taught, and everyone knows that a person who    '
'  tries to teach himself has a fool for a student!).           '
'                                                               '
'  At any rate, I enjoy interacting with other people who like  '
'  to program for fun.  I hope that you have fun with Traver    '
'  Jigsaw and that you find the source code helpful.  Let me    '
'  hear from you.  And ... Enjoy!                               '
'                                                               '
'               Copyright © 2001 by Barry Traver                '
'                                                               '
'***************************************************************'

' Caution:  Some of the code does not follow generally recommended
' practices.  For example, I use symbols (%, &, !, #, $) when I
' define data types.  For another example, I make frequent use of
' global variables, which can be a somewhat risky practice at times
' (definitely not recommended for programming teams!), but for a
' programmer working entirely by himself, the time it can save can
' be worth the risk.  (Yes, I've occasionally run into problems as
' a result of not "encapsulating" enough to protect one part of my
' program from another, but Visual Basic has good debugging tools!
' <grin>)
'
' Some of the non-recommended practices are due to conscious choice
' and can be regarded as perhaps different but still defensible.
' Other non-recommended practices you may see here, however, are
' due to ignorance on my part.  Thus I welcome your comments on the
' program, including suggestions on how the code might be improved.
' (The program would be even worse if it were not for help from a
' number of Usenet newsgroup friends, including Rocky Clark and Mike
' Williams, who supplied code for moving the one Picture box.)  Send
' your comments to me at btraver@traver.org (as mentioned earlier).
' Thanks!

Option Explicit

Private PictureTitle$, PictureFilename$

Private NormalScale As Boolean, BestScaleRatio!
Private NormalLeft!, NormalTop!

Private SOS%, TotalPlaces%  ' "SOS" = number of "Squares On Side"
Private PicturePieceIn%() ' tells what picture piece is in what place
Private NumberWrong%

Private TimeToChoosePiece As Boolean
Private FirstPlace%, FirstPiece%
Private SecondPlace%, SecondPiece%

Private SourceX&, SourceY&, Source2X&, Source2Y&
Private SourceWidth&, SourceHeight&
Private DestinationX&, DestinationY&
Private DestinationWidth&, DestinationHeight&

Private Dragging As Boolean
Private Xmouse As Single, Ymouse As Single
Private Xoffset As Single, Yoffset As Single

Private Sub Form_Load()
  ' We'll start out with a built-in puzzle and set it for 3 x 3.
  SOS = 3
  TotalPlaces = SOS ^ 2
  mnuSetSOS_Click SOS
  NormalScale = True: BestScaleRatio! = 1
  AdjustFormAndPictureSize
  PictureTitle = "Squirrel"
  PictureFilename = App.Path & "\Squirrel.jpg"
  Call ShowScrambledPicture
End Sub

Private Sub mnuAbout_Click()
  MsgBox "This program is" & vbCrLf & vbCrLf _
      & "Copyright (C) 2001" & vbCrLf & vbCrLf & "by Barry A. Traver" _
      & vbCrLf & vbCrLf & "All Rights Reserved." & vbCrLf & vbCrLf _
      & vbCrLf & "I would love to hear" & vbCrLf & vbCrLf _
      & "comments from you" & vbCrLf & vbCrLf & "about this program.  " _
      & vbCrLf & vbCrLf & "Please write to me at " & vbCrLf & vbCrLf _
      & "this email address:" & vbCrLf & vbCrLf & "btraver@traver.org" _
      & vbCrLf & vbCrLf & "Thank you."
End Sub

Private Sub mnuDragMode_Click()
  mnuDragMode.Checked = Not mnuDragMode.Checked
  mnuSwapMode.Checked = Not mnuDragMode.Checked
  ShowWindowCaption
  Picture3.Visible = False
End Sub

Private Sub mnuExit_Click()
  Unload Me
  End
End Sub

Private Sub mnuHowToPlay_Click()
  ' This routine gives "short and sweet" instructions.
  MsgBox "TRAVER JIGSAW is puzzle fun for children" & vbCrLf _
    & vbCrLf & "of all ages and is very simple to play.  Just" _
    & vbCrLf & vbCrLf & "put the pieces in the proper places " _
    & "to make" & vbCrLf & vbCrLf & "a perfect picture." & vbCrLf _
    & vbCrLf & vbCrLf & "IF YOU USE SWAP MODE:" & vbCrLf & vbCrLf _
    & "Click on a piece you want to move," & vbCrLf _
    & vbCrLf & "and then click on where you want it" & vbCrLf _
    & vbCrLf & "put.  The two pieces will swap places." & vbCrLf _
    & vbCrLf & vbCrLf & "IF YOU USE DRAG MODE:" & vbCrLf & vbCrLf _
    & "Click on a piece you want to move, and" & vbCrLf _
    & vbCrLf & "then drag it to where you want to put it." _
    & vbCrLf & vbCrLf & vbCrLf & "ENJOY!"
End Sub

Private Sub mnuLoadPicture_Click()
  ' This routine allows the user to load in any standard picture
  ' file (JPEG, GIF, etc.) from disk or other appropropiate media.
  Dim Pos%
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist _
      Or cdlOFNFileMustExist
  CommonDialog1.Filter = "Picture Files (*.jpg, *gif, and " & _
      "others)|*.jpg;*.jpeg;*.gif;*.bmp;*.cur;*.emf;*.ico;" _
      & "*.rle;*.wmf"
  CommonDialog1.FilterIndex = 1
  CommonDialog1.ShowOpen
  ' Suggested by Mike Williams to correct *.wmf problem:
  With Picture1
    .BackColor = vbWhite
    .Cls
    .Picture = LoadPicture(CommonDialog1.Filename)
    .PaintPicture .Picture, 0, 0, .ScaleWidth, .ScaleHeight
    .Picture = .Image
  End With
  PictureTitle = CommonDialog1.FileTitle
  Pos = InStrRev(CommonDialog1.FileTitle, ".")
  PictureTitle = Left$(CommonDialog1.FileTitle, Pos - 1)
  PictureFilename = CommonDialog1.Filename
  AdjustFormAndPictureSize
  Call ShowScrambledPicture
  frmJigsaw.BackColor = &H8000000F ' Normal gray/grey
  TimeToChoosePiece = True ' Used for Swap Mode.
ErrHandler:
  Exit Sub
End Sub

Private Sub mnuPrint_Click()
  ' Print out as large a picture as possible with minimal distortion
  ' (only what's needed for minimum margin of one-half inch).
  Dim HeightRatio!, WidthRatio!, BestRatio!
  Dim NewHeight!, NewWidth!
  Dim DestX!, DestY!, DestWidth!, DestHeight!
  If Picture1.Height >= Picture1.Width Then
    HeightRatio = Printer.ScaleHeight / Picture1.ScaleHeight
    WidthRatio = Printer.ScaleWidth / Picture1.ScaleWidth
    BestRatio = Min(HeightRatio, WidthRatio)
    NewHeight = Picture1.ScaleHeight * BestRatio
    NewWidth = Picture1.ScaleWidth * BestRatio
    Printer.Orientation = vbPRORPortrait
  Else
    HeightRatio = Printer.ScaleHeight / Picture1.ScaleWidth
    WidthRatio = Printer.ScaleWidth / Picture1.ScaleHeight
    BestRatio = Min(HeightRatio, WidthRatio)
    NewHeight = Picture1.ScaleHeight * BestRatio
    NewWidth = Picture1.ScaleWidth * BestRatio
    Printer.Orientation = vbPRORLandscape
  End If
  DestX = Max(0, (Printer.ScaleWidth - NewWidth) / 2)
  DestY = Max(0, (Printer.ScaleHeight - NewHeight) / 2)
  DestWidth = NewWidth
  DestHeight = NewHeight
  If DestX < 720 Then
    DestX = 720
    DestWidth = DestWidth - (2 * (720 - DestX))
  End If
  If DestY < 720 Then
    DestY = 720
    DestHeight = DestHeight - (2 * (720 - DestY))
  End If
  Printer.PaintPicture Picture1.Picture, _
      DestX, DestY, DestWidth, DestHeight
  Printer.EndDoc
End Sub

Private Sub mnuReportProgress_Click()
  mnuReportProgress.Checked = Not mnuReportProgress.Checked
  Call CheckProgress
End Sub

Private Sub mnuRestore_Click()
  ' Restores puzzle and puzzle state previously saved.
  Dim Filenum%, Filename$, I%
  On Error GoTo Errortrap
  Filenum = FreeFile
  Open App.Path & "\Jigsaw.cfg" For Input As #Filenum
  Line Input #Filenum, PictureFilename
  Line Input #Filenum, PictureTitle
  Picture1.Picture = LoadPicture(PictureFilename)
  AdjustFormAndPictureSize
  Input #Filenum, SOS
  mnuSetSOS_Click SOS
  Picture2.Cls ' Restore black background for grid.
  For I = 1 To TotalPlaces
    Input #Filenum, PicturePieceIn(I)
    Call MovePiece(PicturePieceIn(I), I)
  Next I
  Close #Filenum
  Call CheckProgress
  Call TestIt
  TimeToChoosePiece = True ' Used with Swap Mode
  Exit Sub
Errortrap:
  MsgBox "Error in Restoring Information."
End Sub

Private Sub mnuSave_Click()
  ' Saves current puzzle and puzzle state (to be restored later).
  Dim Filenum%, I%
  On Error GoTo Errortrap
  Filenum = FreeFile
  Open App.Path & "\Jigsaw.cfg" For Output As #Filenum
  Print #Filenum, PictureFilename
  Print #Filenum, PictureTitle
  Print #Filenum, SOS
  For I = 1 To TotalPlaces
    Print #Filenum, PicturePieceIn(I)
  Next I
  Close #Filenum
  Exit Sub
Errortrap:
  MsgBox "Error in Saving Information."
End Sub

Private Sub mnuScramblePicture_Click()
  ' This routine places puzzle pieces in a "random" arrangement.
  ShowScrambledPicture
End Sub

Private Sub mnuSetSOS_Click(Index%)
  ' This routine sets up a puzzle for 2 x 2, 3 x 3, etc.
  Dim I%
  SOS = Index
  TotalPlaces = SOS ^ 2
  For I = 2 To 12
    mnuSetSOS(I).Checked = (I = SOS)
  Next I
  ' Set size of movable piece
  Picture3.Width = Picture2.ScaleWidth / SOS
  Picture3.Height = Picture2.ScaleHeight / SOS
  ShowScrambledPicture
End Sub

Private Sub mnuUnscramblePicture_Click()
  ' This routine copies the entire picture from the invisible
  ' Picture1 to the visible Picture2.
  mnuUnscramblePicture.Checked = True
  mnuScramblePicture.Checked = False
  Call ShowUnscrambledPicture
  Call AdjustStatusBar(1, "Here's your puzzle put together!")
End Sub

Private Sub mnuSwapMode_Click()
  mnuSwapMode.Checked = Not mnuSwapMode.Checked
  mnuDragMode.Checked = Not mnuSwapMode.Checked
  ShowWindowCaption
  Picture3.Visible = False
End Sub

Private Sub mnuWhichScale_Click(Index As Integer)
  ' Adjust for normal size or optimal size for screen.
  If mnuWhichScale(1).Checked = True Then
    NormalLeft = frmJigsaw.Left
    NormalTop = frmJigsaw.Top
  End If
  mnuWhichScale(1).Checked = Not mnuWhichScale(1).Checked
  mnuWhichScale(2).Checked = Not mnuWhichScale(2).Checked
  If mnuWhichScale(1).Checked Then
    NormalScale = True
    frmJigsaw.Left = NormalLeft
    frmJigsaw.Top = NormalTop
  Else
    NormalScale = False
    frmJigsaw.Move 0, 0
  End If
  Call AdjustFormAndPictureSize
  If NumberWrong = 0 Then
    ShowUnscrambledPicture
  Else
    Call RefreshPuzzleDisplay
  End If
End Sub

Private Sub Picture2_MouseDown(Button%, Shift%, X!, Y!)
  ' Exit if puzzle is solved.
  If NumberWrong = 0 Then Exit Sub
  ' Provide hidden help if right mouse button is clicked.
  If Button = 2 And (Shift > 0 And Shift < 4) Then
    Call ProvideHiddenHelp(Shift%, X!, Y!)
  ElseIf Button = 1 And Shift = 0 Then ' Plain left-click of mouse
    If mnuDragMode.Checked Then ' Handle mouse left-click in Drag Mode
      Dragging = True
      FirstPlace = Val(Int(X / (Picture2.ScaleWidth / SOS))) _
          + SOS * Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
      FirstPiece = PicturePieceIn(FirstPlace)
      ' "Paint" movable puzzle piece
      Source2X = ((FirstPiece - 1) Mod SOS) * Picture1.ScaleWidth / SOS
      Source2Y = Int((FirstPiece - 1) / SOS) * Picture1.ScaleHeight / SOS
      Picture3.PaintPicture _
          Picture1.Picture, _
          0, 0, DestinationWidth, DestinationHeight, _
          Source2X, Source2Y, SourceWidth, SourceHeight
      ' Change location of moveable puzzle piece to the place clicked
      DestinationX = ((FirstPlace - 1) Mod SOS) * _
          Picture2.ScaleWidth / SOS
      DestinationY = Int((FirstPlace - 1) / SOS) * _
          Picture2.ScaleHeight / SOS
      Picture3.Move DestinationX + Picture2.Left, _
          DestinationY + Picture2.Top
      ' Make movable puzzle piece visible
      Picture3.Visible = True
      Xoffset = Picture3.Left - Picture2.Left
      Yoffset = Picture3.Top - Picture2.Top
      Xmouse = X
      Ymouse = Y
      ' Put black box on Picture 2 under movable puzzle piece
      Picture2.Line (DestinationX, DestinationY) _
          -(DestinationX + DestinationWidth, DestinationY + DestinationHeight), _
          , BF
    Else ' Handle mouse left-click in Swap Mode
      If TimeToChoosePiece = True Then
        FirstPlace = Val(Int(X / (Picture2.ScaleWidth / SOS))) _
            + SOS * Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
        FirstPiece = PicturePieceIn(FirstPlace)
        ' Visually emphasize what piece was clicked.
        ' "Paint" movable puzzle piece
        Source2X = ((FirstPiece - 1) Mod SOS) * Picture1.ScaleWidth / SOS
        Source2Y = Int((FirstPiece - 1) / SOS) * Picture1.ScaleHeight / SOS
        Picture3.PaintPicture _
            Picture1.Picture, _
            0, 0, DestinationWidth, DestinationHeight, _
            Source2X, Source2Y, SourceWidth, SourceHeight
        ' Change location of moveable puzzle piece to the place clicked
        DestinationX = ((FirstPlace - 1) Mod SOS) * _
            Picture2.ScaleWidth / SOS
        DestinationY = Int((FirstPlace - 1) / SOS) * _
            Picture2.ScaleHeight / SOS
        Picture3.Move DestinationX + Picture2.Left, _
            DestinationY + Picture2.Top
        ' Make movable puzzle piece visible
        Picture3.Visible = True
        TimeToChoosePiece = False
      Else ' Time to click on place to move chosen piece
        SecondPlace = Val(Int(X / (Picture2.ScaleWidth / SOS))) _
            + SOS * Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
        If SecondPlace = FirstPlace Then
          Picture3.Visible = False
          TimeToChoosePiece = True
          Exit Sub
        End If
        SecondPiece = PicturePieceIn(SecondPlace)
        ' The pieces swap places (or at least it looks that way).
        Call MovePiece(FirstPiece, SecondPlace)
        Call MovePiece(SecondPiece, FirstPlace)
        Picture3.Visible = False
        Call CheckProgress
        Call TestIt
      End If
    End If
  End If
End Sub

Private Sub Picture2_MouseMove(Button%, Shift%, X!, Y!)
  Dim X1!, Y1!
  ' Exit if puzzle is solved.
  If NumberWrong = 0 Then Exit Sub
  ' Exit if right mouse button is clicked.
  If Button = 2 Then Exit Sub
  ' Exit if Mode = Swap Mode, because ...
  If mnuSwapMode.Checked = True Then Exit Sub
  ' this routine is useful only for Dragging in Drag Mode.
  If Dragging Then
    X1 = Picture2.Left + X - Xmouse
    Y1 = Picture2.Top + Y - Ymouse
    If (Picture2.Left <> X1) Or (Picture2.Top <> Y1) Then
      Picture3.Move X1 + Xoffset, Y1 + Yoffset
    End If
  End If
End Sub

Private Sub Picture2_MouseUp(Button%, Shift%, X!, Y!)
  Dim Column%, Row%, Position%
  ' Exit if puzzle is solved.
  If NumberWrong = 0 Then Exit Sub
  ' Exit if right mouse button is clicked.
  If Button = 2 Then Exit Sub
  ' Exit if Mode = Swap Mode
  If mnuSwapMode.Checked = True Then Exit Sub
  If Button = 1 And Shift = 0 And Dragging = False Then Exit Sub
  X = Picture3.Left - Picture2.Left + DestinationWidth / 2
  Y = Picture3.Top - Picture2.Top + DestinationHeight / 2
  Column = Val(Int(X / (Picture2.ScaleWidth / SOS))) + 1
  Row = Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
  ' Don't go outside the limits of the picture!
  If Column < 1 Then Column = 1
  If Column > SOS Then Column = SOS
  If Row < 1 Then Row = 1
  If Row > SOS Then Row = SOS
  Position = (Column - 1) + (SOS * (Row - 1)) + 1
  SecondPlace = Position
  SecondPiece = PicturePieceIn(SecondPlace)
  ' The pieces swap places (or at least it looks that way).
  Call MovePiece(FirstPiece, SecondPlace)
  Call MovePiece(SecondPiece, FirstPlace)
  Call CheckProgress
  Call TestIt
  Picture3.Visible = False
  Dragging = False
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mnuSwapMode.Checked = True Then
    ' Remove emphasis upon piece clicked earlier
    Picture3.Visible = False
    TimeToChoosePiece = True
  End If
End Sub

Private Sub StatusBar1_PanelClick(Index As Integer, ByVal Panel As MSComctlLib.Panel)
  If Index = 1 Then Exit Sub
  If Left$(StatusBar1(2).Panels(1).Text, 16) = "Congratulations!" Then
    ' "Congratulations!  Do another puzzle?"
    If Panel = "Yes" Then
      Call AdjustStatusBar(2, "Do the same picture again?")
      Beep
      frmJigsaw.BackColor = vbBlue
    Else
      End
    End If
  Else ' Do the same picture again?
    If Panel = "Yes" Then
      ShowScrambledPicture
    Else
      Call AdjustStatusBar(1, "Choose another picture.")
      mnuLoadPicture_Click
    End If
  End If
End Sub

' The preceding are control event procedures.  The following
' are other procedures used in the program.

Private Sub AdjustFormAndPictureSize()
  Dim Ratio1!, Ratio2!
  Dim TempFormWidth!, TempFormHeight!
  ' This routine resizes the form according to the picture loaded.
  ' and also adjusts dimensions for Picture2 and Picture 3
  ' Don't adjust size if minimizing the program.
  If frmJigsaw.WindowState = 1 Then Exit Sub
  If NormalScale = True Then
    BestScaleRatio! = 1
  Else
    Ratio1! = Screen.Width / (Picture1.Width + 400)
    Ratio2! = Screen.Height / (Picture1.Height + 1600)
    BestScaleRatio! = Min(Ratio1!, Ratio2!)
    ' If form height > .95 sceen height, leave room for task bar.
    If (BestScaleRatio! * (Picture1.Height + 1600)) > 0.95 _
        * Screen.Height Then
      BestScaleRatio! = 0.95 * BestScaleRatio!
    End If
  End If
  ' Set size of visible picture.
  Picture2.Width = (BestScaleRatio! * (Picture1.Width + 400)) _
      - 400
  Picture2.Height = (BestScaleRatio! * (Picture1.Height + 1600)) _
      - 1600
  ' Set size of form
  frmJigsaw.Width = Picture2.Width + 400
  frmJigsaw.Height = Picture2.Height + 1600
  ' Set size of movable piece
  Picture3.Width = Picture2.ScaleWidth / SOS
  Picture3.Height = Picture2.ScaleHeight / SOS
  ' Adust if not wide enough to show top menu on one line.
  If frmJigsaw.Width < 4500 Then
    frmJigsaw.Width = 4500
    Picture2.Left = (frmJigsaw.Width - Picture2.Width) / 2
  Else
    Picture2.Left = 180
  End If
End Sub

Private Sub AdjustStatusBar(WhichOne%, Message$)
  StatusBar1(3 - WhichOne).Visible = False
  StatusBar1(WhichOne).Panels(1).Text = Message$
  StatusBar1(WhichOne).Visible = True
End Sub

Private Sub CheckNumberWrong()
  Dim I%
  NumberWrong = 0
  For I = 1 To TotalPlaces
    If PicturePieceIn(I) <> I Then
      NumberWrong = NumberWrong + 1
    End If
  Next I
End Sub

Private Sub CheckProgress()
  Call CheckNumberWrong
  If mnuReportProgress.Checked = True Then
    Call AdjustStatusBar(1, Str$(NumberWrong) _
        & " pieces are out of place.")
  Else
    Call AdjustStatusBar(1, "Here's your puzzle to solve!")
  End If
End Sub

Function Max!(Number1!, Number2!)
  Max = Number1
  If Number2 > Number1 Then Max = Number2
End Function

Function Min!(Number1!, Number2!)
  Min = Number1
  If Number2 < Number1 Then Min = Number2
End Function

Private Sub MovePiece(PuzzlePiece%, PuzzlePlace%)
  ' Get proper puzzle piece from Picture1 and place it in Picture2.
  Call SetSourceParameters(PuzzlePiece, SourceX, SourceY, _
      SourceWidth, SourceHeight)
  Call SetDestinationParameters(PuzzlePlace, DestinationX, _
      DestinationY, DestinationWidth, DestinationHeight)
  ' Slightly reduce rectangle size to create appearance of grid.
  SourceX = SourceX + 10
  SourceY = SourceY + 10
  SourceWidth = SourceWidth - 20
  SourceHeight = SourceHeight - 20
  DestinationX = DestinationX + 10
  DestinationY = DestinationY + 10
  DestinationWidth = DestinationWidth - 20
  DestinationHeight = DestinationHeight - 20
  Picture2.PaintPicture _
      Picture1.Picture, _
      DestinationX, DestinationY, DestinationWidth, DestinationHeight, _
      SourceX, SourceY, SourceWidth, SourceHeight
  PicturePieceIn(PuzzlePlace) = PuzzlePiece
  TimeToChoosePiece = True ' Used with Swap Mode
End Sub

Private Sub ProvideHiddenHelp(Shift%, X!, Y!)
  ' Certain keypress combinations involving a right mouse-click
  ' will put one or more pieces in their proper places.
  Dim I%
  Select Case Shift
   Case 1 ' Shift-right click was pressed.
      ' Shift-Right click provides "hidden help."  The piece clicked
      ' will be put in its proper place (if it is not already there).
      FirstPlace = Val(Int(X / (Picture2.ScaleWidth / SOS))) _
          + SOS * Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
      FirstPiece = PicturePieceIn(FirstPlace)
      SecondPlace = PicturePieceIn(FirstPlace)
      SecondPiece = PicturePieceIn(SecondPlace)
    Case 2 ' Ctrl-right arrow was pressed
    ' Ctrl-Right click provides "hidden help."  If place clicked
    ' does not contain proper piece, it will be put in that place.
    FirstPlace = Val(Int(X / (Picture2.ScaleWidth / SOS))) _
        + SOS * Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
    FirstPiece = PicturePieceIn(FirstPlace)
    For I = 1 To TotalPlaces
      If PicturePieceIn(I) = FirstPlace Then Exit For
    Next I
    SecondPlace = I
    SecondPiece = PicturePieceIn(I)
  Case 3 ' Ctrl-Shift-right click was pressed
    ' Put the correct piece in this place....
    FirstPlace = Val(Int(X / (Picture2.ScaleWidth / SOS))) _
        + SOS * Val(Int(Y / (Picture2.ScaleHeight / SOS))) + 1
    FirstPiece = PicturePieceIn(FirstPlace)
    For I = 1 To TotalPlaces
      If PicturePieceIn(I) = FirstPlace Then Exit For
    Next I
    SecondPlace = I
    SecondPiece = PicturePieceIn(I)
    Call MovePiece(FirstPiece, SecondPlace)
    Call MovePiece(SecondPiece, FirstPlace)
    ' ... and put the piece that was here earlier in its proper place.
    FirstPlace = SecondPlace
    FirstPiece = PicturePieceIn(SecondPlace)
    SecondPlace = PicturePieceIn(FirstPlace)
    SecondPiece = PicturePieceIn(SecondPlace)
  End Select
  Call MovePiece(FirstPiece, SecondPlace)
  Call MovePiece(SecondPiece, FirstPlace)
  Call CheckProgress
  Call TestIt
End Sub

Private Sub SetDestinationParameters(PuzzlePlace%, DestinationX&, _
    DestinationY&, DestinationWidth&, DestinationHeight&)
  ' The "destination" is Picture2, the visible puzzle seen by the
  ' user.  The "source" is Picture1, the invisible completed puzzle
  ' used as a pattern and answer key behind the visible puzzle.
  DestinationX = (((PuzzlePlace - 1) Mod SOS) * Picture2.ScaleWidth / SOS)
  DestinationY = Int((PuzzlePlace - 1) / SOS) * Picture2.ScaleHeight / SOS
  DestinationWidth = Picture2.ScaleWidth / SOS
  DestinationHeight = Picture2.ScaleHeight / SOS
End Sub

Private Sub SetSourceParameters(PuzzlePiece%, SourceX&, SourceY&, _
    SourceWidth&, SourceHeight&)
  ' The 'source' is Picture1.  To be technically consistent with
  ' SetDestinationParameter, perhaps we should speak of PuzzlePlace
  ' here also (rather than PuzzlePiece), but in Picture1 all pieces
  ' are in their proper place, i.e., for Picture1, PuzzlePlace% =
  ' PuzzlePiece%, so it works the same either way, and using two
  ' different words may make the code a bit easier to follow.
  SourceX = ((PuzzlePiece - 1) Mod SOS) * Picture1.ScaleWidth / SOS
  SourceY = Int((PuzzlePiece - 1) / SOS) * Picture1.ScaleHeight / SOS
  SourceWidth = Picture1.ScaleWidth / SOS
  SourceHeight = Picture1.ScaleHeight / SOS
End Sub

Private Sub ShowScrambledPicture()
  ' This routine places puzzle pieces in a "random" arrangement.
  mnuScramblePicture.Checked = True
  mnuUnscramblePicture.Checked = False
  Dim Choices%(), I%, J%, LeftToChoose, OKToExit As Boolean
  TotalPlaces = SOS ^ 2
  ReDim Choices(TotalPlaces), PicturePieceIn(TotalPlaces)
  For I = 1 To TotalPlaces
    Choices(I) = I
  Next I
  frmJigsaw.BackColor = &H8000000F ' Normal gray/grey for form.
  Picture2.Cls ' Restore black background for picture grid.
  ' This loop chooses a "random" piece for each place
  For I = 1 To TotalPlaces
    LeftToChoose = (TotalPlaces) + 1 - I
    Do ' This code loop maximizes the disorganization.
      OKToExit = True
      Randomize
      J = Int(((LeftToChoose) * Rnd) + 1)
      If I >= TotalPlaces - 2 Then Exit Do
      ' Avoid putting correct piece in correct place.
      If I = Choices(J) Then OKToExit = False
      ' Avoid side-by-side match if possible.
      If I > 1 Then
        If (PicturePieceIn(I - 1) + 1 = Choices(J)) Then OKToExit = False
      End If
      ' Avoid vertical match if possible.
      If I > SOS Then
        If (PicturePieceIn(I - SOS) + SOS = Choices(J)) _
          Then OKToExit = False
      End If
      If OKToExit = True Then Exit Do
    Loop
    PicturePieceIn(I) = Choices(J)
    Call MovePiece(PicturePieceIn(I), I)
    Choices(J) = Choices(LeftToChoose)
  Next I
  ShowWindowCaption
  Call CheckProgress
End Sub

Private Sub RefreshPuzzleDisplay()
  ' Useful when changing scale; scrambled picture keeps same position
  Dim I%
  Picture2.Cls
  For I = 1 To TotalPlaces
    Call MovePiece(PicturePieceIn(I), I)
  Next I
End Sub

Private Sub ShowUnscrambledPicture()
  Dim I%
  mnuUnscramblePicture.Checked = True
  mnuScramblePicture.Checked = False
  Picture2.PaintPicture _
      Picture1.Picture, _
      0, 0, Picture2.Width, Picture2.Height, _
      0, 0, Picture1.Width, Picture1.Height
    For I = 1 To TotalPlaces
      PicturePieceIn(I) = I
    Next I
  NumberWrong = 0
End Sub

Private Sub ShowWindowCaption()
  Dim Mode$
  If mnuDragMode.Checked = True Then
    Mode$ = "Drag Mode"
  Else
    Mode$ = "Swap Mode"
  End If
  frmJigsaw.Caption = "Traver Jigsaw - " & PictureTitle _
      & " - " & SOS & " x " & SOS & " - " & Mode
End Sub

Private Sub TestIt()
  ' This routine checks to see whether pieces are all in their
  ' proper places and responds accordingly.
  Call CheckProgress
  If NumberWrong > 0 Then Exit Sub
  ' The following gets rid of any lines between puzzle pieces:
  Call ShowUnscrambledPicture
  ' Put a black frame around the picture for attention.
  frmJigsaw.BackColor = vbBlack
  Call AdjustStatusBar(2, "Congratulations!  Do another puzzle?")
  Beep
End Sub
