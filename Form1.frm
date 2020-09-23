VERSION 5.00
Begin VB.Form frmDungeon 
   Caption         =   "Use the arrow Keys to move, Esc to exit"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You Standing in a vast wasteland that streaches as far as you can see in all directions. There is a hut near here"
      Height          =   1695
      Left            =   4920
      TabIndex        =   3
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Current Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":030A
      Height          =   2415
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmDungeon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is just an idea for a rougelike game written in VB, the map and
'Charachters are all drawn using ascii chars. The program needs fixing

'Time get time returns the time, getasynckeystate lets us know what keys are
'Pressed and getactivewindow is the currently active window handle
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private MapArray(1 To 40, 1 To 20) As String * 1    'The actual map
Private EventArray() As EventData       'Map Events
Private CurrentR As Integer     'Current row pos of player
Private CurrentC As Integer     'Current column position of player
Private CurrentMap As Integer   'Map currently in use
Private LineCount As Integer    'Used for redimming Events array
Private NextFrame As Long        'Used as a timer to stop the game going too fast
Private GameOver As Boolean      'To exit the main loop

'Constants for the map numbers
Private Const Level_1 As Integer = 0
Private Const Level_2 As Integer = 1
Private Const Level_3 As Integer = 2
Private Const Level_4 As Integer = 3

Private Type EventData                  'Used for The events Textfile
    EventX As Integer
    EventY As Integer
    Description As String
End Type


Private Sub LoadMap(ByVal map As Integer)
'This populates Maparray() with the correct map
Dim i As Integer        'Counter Var (for populating array)
Dim j As Integer        'Counter Var (for populating array)
Dim k As Integer        'Counter Var (For file position)
Dim EmptyString As String * 1   'For storing file data
Dim MapPath As String           'Address to the file
Dim FileHandle As Integer       'To store the file number (one not in use)
Dim EventFile As String

'More can be added here as more maps are written, I should probably get these
'Values from a file later
If map = Level_1 Then MapPath = "\Data\StartMap.txt"
If map = Level_2 Then MapPath = "\Data\Map2.txt"
If map = Level_3 Then MapPath = "\Data\map3.txt"
If map = Level_4 Then MapPath = "\Data\Map4.txt"

CurrentMap = map                'Used for map events
EmptyString = Space(1)          'This is so only one char is read from the file
k = 1                           'k is used for the current possition in the file
FileHandle = FreeFile           'Assigns a number representing the next available
                                'File handle to the var filehandle

Open App.Path & MapPath For Binary As #FileHandle
    For i = 1 To 20
        For j = 1 To 40
            Get #1, k, EmptyString
            k = k + 1
            MapArray(j, i) = EmptyString
            'This reads from the text file (one char at a time) and
            'Sticks that info into the array
            If MapArray(j, i) = "ö" Then        'Get the location of the player
                CurrentR = j                    'symbol, and take note of its
                CurrentC = i                    'Map coords
            End If
        Next j
    Next i
Close #FileHandle

FileHandle = FreeFile           'Assigns a number representing the next available
                                'File handle to the var filehandle
                                
                                
If map = Level_1 Then EventFile = "\Data\Map1Events.txt"
If map = Level_2 Then EventFile = "\Data\Map2Events.txt"
If map = Level_3 Then EventFile = "\Data\Map3Events.txt"
If map = Level_4 Then EventFile = "\Data\Map4Events.txt"
LineCount = 0

'The following block Pulls the data from the events file and stores it in an
'(of a user defined type(See above))
Open App.Path & EventFile For Input As #FileHandle
    Do While Not EOF(FileHandle)        'While there is text left
        LineCount = LineCount + 1
        ReDim Preserve EventArray(LineCount)
        Input #FileHandle, _
            EventArray(LineCount).EventX, _
            EventArray(LineCount).EventY, _
            EventArray(LineCount).Description
    Loop
Close #FileHandle
    
'Default messages for the maps. I might assign these to vars later on
If map = Level_1 Then lblDisplay.Caption = _
    "You Standing in a vast wasteland that streaches as far as you can see in all directions. There is a hut to the north west and a cave to the east"
If map = Level_2 Then lblDisplay.Caption = _
    "You are in a small cave, the only light comes from your small lantern, the exit is to the west"
If map = Level_3 Then lblDisplay.Caption = _
    "You are in the Rat King layer, As you step in he roars, causing parts of the cave system to come crashing down, trapping you here *CRASH*"
If map = Level_4 Then lblDisplay.Caption = _
    "You plummet through the collapsing floor with the demise of the rat king. You are in a large room containing an underground lake, there is a shelter of some kind on the island"

'Draw the Map and start the game
Call DrawMap
Call MasterLoop
End Sub

Private Sub DrawMap()
'This sub just cycles through maparray() and prints the data to the screen
Dim i As Integer
Dim j As Integer

picDisplay.Cls          'To stop the information going below the old
For i = 1 To 20
    For j = 1 To 40
    picDisplay.Print MapArray(j, i);
    Next j
    picDisplay.Print " "        'This is just to give a new line
Next i
End Sub

Private Sub Form_Activate()
Call LoadMap(Level_1)
'If this is done in form load the Screen blanks (picdisplay) the 0 can of course
'Be switched if you want to start with a diffrent map
End Sub

Private Sub MasterLoop()
'The master loop will keep on going untill the bool (GameOver) = true
Do
    NextFrame = timeGetTime + 200
    'GetAsyncKeyState(n) is a boolean value representing whether the
    'Key (n) is pressed or not (value is 0 if its not)
        
    'The following 4 if blocks check the arrow keys, if they are pressed and
    'There is no wall in that direction it moves the player and redraws the
    'Map, it then checks if that square raises an event. I used to have the
    'Event check at the end of the loop, but I think that it works better
    'if its only called when you move. Imactive is a test to see if
    'the window is active
    
    If (ImActive = True) And (GetAsyncKeyState(37) < 0) Then    'Left Arrow
        If MapArray(CurrentR - 1, CurrentC) = " " Then
            MapArray(CurrentR, CurrentC) = " "
            MapArray(CurrentR - 1, CurrentC) = "ö"
            CurrentR = CurrentR - 1
            Call DrawMap
            Call CheckForEvent(CurrentR, CurrentC)
        End If
    End If
        
    If (ImActive = True) And (GetAsyncKeyState(39) < 0) Then    'Right Arrow
        If MapArray(CurrentR + 1, CurrentC) = " " Then
            MapArray(CurrentR, CurrentC) = " "
            MapArray(CurrentR + 1, CurrentC) = "ö"
            CurrentR = CurrentR + 1
            Call DrawMap
            Call CheckForEvent(CurrentR, CurrentC)
        End If
    End If
        
    If (ImActive = True) And (GetAsyncKeyState(38) < 0) Then    'Up Arrow
        If MapArray(CurrentR, CurrentC - 1) = " " Then
            MapArray(CurrentR, CurrentC) = " "
            MapArray(CurrentR, CurrentC - 1) = "ö"
            CurrentC = CurrentC - 1
            Call DrawMap
            Call CheckForEvent(CurrentR, CurrentC)
        End If
    End If
        
    If (ImActive = True) And (GetAsyncKeyState(40) < 0) Then    'Down Arrow
        If MapArray(CurrentR, CurrentC + 1) = " " Then
            MapArray(CurrentR, CurrentC) = " "
            MapArray(CurrentR, CurrentC + 1) = "ö"
            CurrentC = CurrentC + 1
            Call DrawMap
            Call CheckForEvent(CurrentR, CurrentC)
        End If
    End If
    
    'If the Escape key is pressed, Exit the loop and close the program
    If (ImActive = True) And (GetAsyncKeyState(27) < 0) Then GameOver = True: End

    Me.Caption = "Current Possition: " & CurrentR & ":" & CurrentC
    'This was used for design purpose (designing events)
    
    Do                      'If the frame is not finished, just wait
        DoEvents
    Loop Until timeGetTime >= NextFrame

Loop Until GameOver         'Game over can also be changed from subs to end the
                            'game, Will include an ending message later
End Sub


Private Sub CheckForEvent(ByVal Xpos As Integer, ByVal Ypos As Integer)
Dim i As Integer        'Counter Variable

'This sub just checks if the square Raises an event


'Checks against the contents of the EventArray array, if it finds a match,
'Display it
For i = 1 To LineCount
    If (EventArray(i).EventX = Xpos) And _
       (EventArray(i).EventY = Ypos) Then
       lblDisplay.Caption = EventArray(i).Description
       Exit For
    End If
Next i

'This is for the entry and exit points of the map, it just loads
'the requested map
If (CurrentMap = Level_1) And (Xpos = 40) And (Ypos = 13) Then Call LoadMap(Level_2)
If (CurrentMap = Level_2) And (Xpos = 1) And (Ypos = 13) Then Call LoadMap(Level_1)
If (CurrentMap = Level_2) And (Xpos = 29) And (Ypos = 6) Then Call LoadMap(Level_3)
If (CurrentMap = Level_3) And (Xpos = 34) And (Ypos = 11) Then Call LoadMap(Level_4)

If (CurrentMap = Level_3) And (Xpos = 29) And (Ypos = 7) Then
    MapArray(29, 6) = "@"
    Call DrawMap
End If

End Sub

Private Function ImActive() As Boolean
'Get active window returns a long representing the window handle of the
'Currently active window, Me.hWnd returns a long representing the handle of
'The form window, these are compared to find out who has focus !

If Me.hWnd = GetActiveWindow Then
    ImActive = True
    Else
        ImActive = False
End If
End Function

Private Sub Label3_Click()

End Sub
