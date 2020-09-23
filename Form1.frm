VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Fire"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   600
      Top             =   360
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4005
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Left-click to place more plants, right-click to quit"
      Top             =   1455
      Width           =   3465
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3225
      Top             =   1965
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BgBuffer As clsMemoryBitmap  'The offscreen buffer all pictures are assembled in
Private Backdrop As clsMemoryBitmap  'The background picture that erases old pictures
Private Sprites() As clsMemoryBitmap 'The moving pictures that are drawn with transparency.
Private FrameCount As Long, StartTime As Long, NextTime As Long 'FPS timer variables
Private LastX As Long, LastY As Long
Private QuitGame As Boolean
Private MoveByYourself As Boolean

Private Sub Form_Load()

    Dim RetVal As Long
    
    'Set the max size of the screen, and resize the form
    ScreenRect.Left = 0
    ScreenRect.Top = 0
    ScreenRect.Right = 640
    ScreenRect.Bottom = 468
    
    'Set up the form
    Me.ScaleMode = vbPixels
    Me.Move 0, 0, (ScreenRect.Right * Screen.TwipsPerPixelX), (ScreenRect.Bottom * Screen.TwipsPerPixelY)
    Text1.Move 0, Me.ScaleHeight - Text1.Height
    Text2.Move Me.ScaleWidth - Text2.Width, Me.ScaleHeight - Text2.Height
    'Me.Palette = App.Path & "\Background.bmp"
    Me.PaletteMode = 2
    Me.Show
    Me.SetFocus
    
    'Load the backdrop picture - REQUIRED
    Set Backdrop = New clsMemoryBitmap
    Backdrop.Init App.Path & "\Background.bmp"
    
    'Load the background buffer - REQUIRED
    Set BgBuffer = New clsMemoryBitmap
    BgBuffer.Init App.Path & "\Background.bmp"
    
    'Load the sprite - REQUIRED
    ReDim Sprites(1 To 1)
    Set Sprites(1) = New clsMemoryBitmap
    Sprites(1).Init App.Path & "\firemap2.bmp"
    
    'Set the bounding rectangle and destination surfaces
    BgBuffer.SetDC Form1.HDC    'background buffer draws to the screen
    Backdrop.SetDC BgBuffer.HDC 'backdrop draws to the background buffer
    Sprites(1).SetDC BgBuffer.HDC   'sprite draws to the background buffer
    
    'Add the first dirty rectangle.
    ReDim DirtyRects(0 To 0)
    AddDirtyRect ScreenRect
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set the sprite location
    If X <> LastX Or Y <> LastY Then
        LastX = X: LastY = Y
        Sprites(1).SetXY LastX, LastY
        MoveByYourself = False
        Timer2.Enabled = True
    End If
End Sub

Private Sub RenderLoop()
    Const TheInterval As Long = 3
    Dim xInterval As Long, yInterval As Long
    xInterval = TheInterval: yInterval = TheInterval / 3
    Do
        If QuitGame Then
            Exit Do
        End If
        'Check for demo mode (10 seconds of inactivity)
        If MoveByYourself Then
            Sprites(1).SetXY LastX, LastY
            LastX = LastX + xInterval
            LastY = LastY + yInterval
            If LastX > ScreenRect.Right Then
                LastX = ScreenRect.Right: xInterval = 0 - TheInterval
            End If
            If LastY > ScreenRect.Bottom Then
                LastY = ScreenRect.Bottom: yInterval = (0 - TheInterval) / 3
            End If
            If LastX < 0 Then
                LastX = 0: xInterval = TheInterval
            End If
            If LastY < 0 Then
                LastY = 0: yInterval = TheInterval / 3
            End If
        End If
        'Display the pictures
        Render
        'Increment the FPS counter
        FrameCount = FrameCount + 1
        If timeGetTime > NextTime Then
            'Display the FPS counter
            Text1.Text = UBound(Sprites, 1) & " plants  - " & Format(FrameCount * 1000 / (NextTime - StartTime), "0.00") & " fps"
            StartTime = timeGetTime
            NextTime = StartTime + 1000
            FrameCount = 0
        End If
        DoEvents
    Loop
End Sub

Private Sub Render()
    Dim RetVal As Long, i As Long, j As Long
    
    MergeRects
    
    For i = 1 To UBound(DirtyRects, 1)
    
        If DirtyRects(i).Right = 0 Then
            'it's an emptied rectangle - skip it
        Else
            'Draw the backdrop, erasing all that went before
            Backdrop.Draw DirtyRects(i), False
            
            'Draw the sprite with transparency
            For j = 2 To UBound(Sprites, 1)
                Sprites(j).Draw DirtyRects(i), True
            Next j
            Sprites(1).Draw DirtyRects(i), True
            
            'Copy the background buffer to the form
            BgBuffer.Draw DirtyRects(i), False
        End If
    Next i
        
    'Erase the dirty rectangles
    ReDim DirtyRects(0 To 0)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    If Button = 2 Then
        QuitGame = True
    Else
        i = UBound(Sprites, 1) + 1
        ReDim Preserve Sprites(1 To i)
        Set Sprites(i) = New clsMemoryBitmap
        Sprites(i).Init App.Path & "\firemap2.bmp"
        Sprites(i).SetXY X, Y
        Sprites(i).SetDC BgBuffer.HDC   'sprite draws to the background buffer
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QuitGame = True
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    Timer1.Enabled = False
    
    'Start the FPS timer
    StartTime = timeGetTime
    NextTime = StartTime + 1000
    
    'Start the main rendering loop
    RenderLoop
    
    'Quit
    For i = 1 To UBound(Sprites, 1)
        Set Sprites(i) = Nothing
    Next i
    Set Backdrop = Nothing
    Set BgBuffer = Nothing
    End
End Sub

Private Sub Timer2_Timer()
    MoveByYourself = True
    Timer2.Enabled = False
End Sub
