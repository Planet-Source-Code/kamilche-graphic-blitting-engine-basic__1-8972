VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variables
Private PicDC As Long, PicPrevBmp As Long, PicBmp As Long
Private picWidth As Long, picHeight As Long
Private PicRect As RECT
Private DestDC As Long ', BoundingRect As RECT

Friend Sub Init(ByVal FileName As String)
'Loads a filename into a picture in memory.
    Dim FileNo As Integer
    Dim ScreenDC As Long
    Dim FileHead As BITMAPFILEHEADER
    Dim bmpInfo_24 As BITMAPINFO_24
    Dim bmpInfo_8 As BITMAPINFO_8
    Dim bmpInfoHead As BITMAPINFOHEADER
    Dim PicBytes() As Byte
    Dim RetVal As Long
    Dim i As Long, j As Long, PicOffset As Long, PicLen As Long
    
    On Error GoTo Err_Init
    
    'Make sure the file exists
    If Dir(FileName, vbArchive Or vbNormal) = "" Then
        MsgBox "You must pass a valid BMP file to this routine!"
        Exit Sub
    End If
    
    PicOffset = 1
    PicLen = FileLen(FileName)
    
    'Read in the file data
    FileNo = FreeFile
    Open FileName For Binary Access Read As #FileNo
    
    'Retrieve the bitmap information
    Get #FileNo, PicOffset, FileHead
    Get #FileNo, PicOffset + 14, bmpInfoHead
    
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case 8
            Get #FileNo, PicOffset + 14, bmpInfo_8
        Case 24
            Get #FileNo, PicOffset + 14, bmpInfo_24
    End Select
    
    ReDim PicBytes(0 To PicLen - 44)
    Get #FileNo, , PicBytes
    
    'Close the file
    Close #FileNo

    'Create the memory bitmap
    ScreenDC = GetDC(0)
    PicDC = CreateCompatibleDC(ScreenDC)
    picWidth = bmpInfoHead.biWidth
    picHeight = bmpInfoHead.biHeight
    PicBmp = CreateCompatibleBitmap(ScreenDC, picWidth, picHeight)
    PicPrevBmp = SelectObject(PicDC, PicBmp)
    
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case 8
            RetVal = SetDIBits_8(PicDC, PicBmp, 0, picWidth, PicBytes(0), bmpInfo_8, DIB_RGB_COLORS)
        Case 24
            RetVal = SetDIBits_24(PicDC, PicBmp, 0, picHeight, PicBytes(0), bmpInfo_24, DIB_RGB_COLORS)
    End Select
    
    'Clean up the bitmap
    RetVal = ReleaseDC(0, ScreenDC)
    
    'Set the initial location
    PicRect.Left = 0
    PicRect.Top = 0
    PicRect.Right = picWidth
    PicRect.Bottom = picHeight

    Exit Sub
    
Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Sub

Friend Sub SetXY(ByVal DestX As Long, ByVal DestY As Long)
    AddDirtyRect PicRect
    With PicRect
        .Left = (DestX - picWidth / 2)
        .Top = (DestY - picHeight / 2)
        .Right = .Left + picWidth - 1
        .Bottom = .Top + picHeight - 1
    End With
    AddDirtyRect PicRect
End Sub

Friend Sub Draw(BoundingRect As RECT, Transparent As Boolean)
'Draws the actor on the backdrop, CLIPPING the actor as necessary

    Dim dstRect As RECT, srcRect As RECT, TempRect As RECT
    Dim X As Long
    Dim Y As Long
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim x_off As Long, y_off As Long
    Dim dx As Long, dy As Long
    Dim BoundingRectWidth As Long, BoundingRectHeight As Long
    Dim RetVal As Long
    
    On Error GoTo ErrHandler
    

    'first do trivial rejections of bitmap, is it totally invisible?
    RetVal = IntersectRect(TempRect, PicRect, BoundingRect)
    If RetVal = 0 Then
        Exit Sub
    End If
    
    'Save the current xy of the picture
    X = PicRect.Left
    Y = PicRect.Top
    
'    'first do trivial rejections of bitmap, is it totally invisible?
'    If X >= BoundingRect.Right Or Y >= BoundingRect.Bottom Or (X + picWidth) <= 0 Or (Y + picHeight) <= 0 Then
'        Exit Sub
'    End If
'
    'Calculate width and height of bounding rect
    BoundingRectWidth = BoundingRect.Right - BoundingRect.Left
    BoundingRectHeight = BoundingRect.Bottom - BoundingRect.Top
    
    'clip destination rectangle
    'pre-compute the bounding rect to make life easy
    x1 = PicRect.Left
    y1 = PicRect.Top
    x2 = PicRect.Right 'x1 + PicWidth - 1
    y2 = PicRect.Bottom 'y1 + PicHeight - 1
    
    'upper left hand corner first
    If x1 < BoundingRect.Left Then x1 = BoundingRect.Left
    If y1 < BoundingRect.Top Then y1 = BoundingRect.Top
    
    'now lower left hand corner
    If x2 >= BoundingRect.Right Then x2 = BoundingRect.Right
    If y2 >= BoundingRect.Bottom Then y2 = BoundingRect.Bottom
    
    'fill in the destination rectangle
    With dstRect
        .Left = x1: .Right = x2
        .Top = y1: .Bottom = y2
    End With
    
    'clip source rectangle
    'now we know to draw only the portions of the bitmap from (x1,y1) to (x2,y2).
    'compute offsets into bitmap on x,y axes,
    'we need to compute starting point to blit from.
    x_off = x1 - X
    y_off = y1 - Y
    
    'compute number of rows and columns to blit
    dx = x2 - x1 + 1
    dy = y2 - y1 + 1
    
    'fill in the source rectangle
    With srcRect
        .Left = x_off: .Right = .Left + dx
        .Top = y_off: .Bottom = .Top + dy
    End With
    
    'Draw the graphic
    If Transparent Then
        TransparentBlit dstRect.Left, dstRect.Top, srcRect.Left, srcRect.Top, dx, dy ', vbWhite
    Else
        RetVal = BitBlt(DestDC, dstRect.Left, dstRect.Top, dx, dy, PicDC, srcRect.Left, srcRect.Top, vbSrcCopy)
    End If
    
    Exit Sub
    
ErrHandler:

    MsgBox Err.Number & " - " & Err.Description
    
End Sub

Friend Sub SetDC(BgHDC As Long)
    DestDC = BgHDC
End Sub

Private Sub Class_Terminate()
    Dim PrevBmp As Long, RetVal As Long
    PrevBmp = SelectObject(PicDC, PicPrevBmp)
    RetVal = DeleteObject(PicBmp)
    RetVal = DeleteDC(PicDC)
End Sub

Private Sub TransparentBlit(ByVal DestX As Long, ByVal DestY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal ClipWidth As Long, ByVal ClipHeight As Long)
'IMPORTANT! BACKGROUND COLOR OF PICTURE SHOULD BE WHITE

    Dim hPrevBmp As Long
    Dim hMaskBmp As Long, hMaskPrevBmp As Long, maskDC As Long
    Dim hWorkBmp As Long, hWorkPrevBmp As Long, workDC As Long
    Dim di As Long, OrigColor As Long
    Dim TransColor As Long
    
    TransColor = vbWhite
 
    '------------------------------------------------------------------------
    ' Prepare memory bitmaps for use by this routine.
    '------------------------------------------------------------------------
    'Create mask bitmap
    maskDC = CreateCompatibleDC(DestDC)
    hMaskBmp = CreateBitmap(ClipWidth, ClipHeight, 1, 1, ByVal 0&)
    hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
 
    'Create work bitmap
    workDC = CreateCompatibleDC(DestDC)
    hWorkBmp = CreateCompatibleBitmap(DestDC, ClipWidth, ClipHeight)
    hWorkPrevBmp = SelectObject(workDC, hWorkBmp)
 
    '------------------------------------------------------------------------
    ' Create the mask, so there's a black object on a white background.
    '------------------------------------------------------------------------
    'Set the transparent color
    OrigColor = SetBkColor(PicDC, TransColor)
    'Copy the original picture to the black/white bitmap
    di = BitBlt(maskDC, 0, 0, ClipWidth, ClipHeight, PicDC, SrcX, SrcY, SRCCOPY)
    'Set the background color back to what it was
    TransColor = SetBkColor(PicDC, OrigColor)
 
    '------------------------------------------------------------------------
    ' Invert the mask, so there's a white object on a black background.
    '------------------------------------------------------------------------
    di = BitBlt(maskDC, 0, 0, ClipWidth, ClipHeight, 0&, 0, 0, DSTINVERT)
 
    '------------------------------------------------------------------------
    ' Bring the piece of the affected destination picture into the work area
    '------------------------------------------------------------------------
    di = BitBlt(workDC, 0, 0, ClipWidth, ClipHeight, DestDC, DestX, DestY, SRCCOPY)
 
    '------------------------------------------------------------------------
    ' Now OR in the inverted mask
    ' After this, we have a white hole where our picture will be.
    '------------------------------------------------------------------------
    di = BitBlt(workDC, 0, 0, ClipWidth, ClipHeight, maskDC, 0, 0, SRCPAINT)
 
    '------------------------------------------------------------------------
    ' Now AND in the source picture
    ' After this, we have a new, modified picture piece.
    '------------------------------------------------------------------------
    di = BitBlt(workDC, 0, 0, ClipWidth, ClipHeight, PicDC, SrcX, SrcY, SRCAND)
 
    ' And finally,  Copy the bitmap back to the original spread
    di = BitBlt(DestDC, DestX, DestY, ClipWidth, ClipHeight, workDC, 0, 0, SRCCOPY)
 
    '------------------------------------------------------------------------
    ' Clean up after ourselves
    '------------------------------------------------------------------------
    'Remove mask bitmap
    hPrevBmp = SelectObject(maskDC, hMaskPrevBmp)
    di = DeleteObject(hMaskBmp)
    di = DeleteDC(maskDC)
 
    'Remove work bitmap
    hPrevBmp = SelectObject(workDC, hWorkPrevBmp)
    di = DeleteObject(hWorkBmp)
    di = DeleteDC(workDC)
 
 
End Sub

Friend Property Get Width() As Long
    Width = picWidth
End Property

Friend Property Get Height() As Long
    Height = picHeight
End Property

Friend Property Get HDC() As Long
    HDC = PicDC
End Property
