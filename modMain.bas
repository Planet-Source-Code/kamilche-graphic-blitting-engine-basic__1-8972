Attribute VB_Name = "modMain"
Option Explicit

'Types
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BITMAPFILEHEADER    '14 bytes
   bfType As Integer
   bfSize As Long
   bfReserved1 As Integer
   bfReserved2 As Integer
   bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER   '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Public Type BITMAPINFO_8
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 255) As RGBQUAD
End Type

Public Type BITMAPINFO_24
   bmiHeader As BITMAPINFOHEADER
End Type

'Declares
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal HDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, _
  ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, _
  ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetDIBits_8 Lib "gdi32" Alias "SetDIBits" _
  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
  ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_8, _
  ByVal wUsage As Long) As Long
Public Declare Function SetDIBits_24 Lib "gdi32" Alias "SetDIBits" _
  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
  ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_24, _
  ByVal wUsage As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

'Constants
Public Const DIB_RGB_COLORS As Long = 0
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const DSTINVERT = &H550009

'Variables
Public DirtyRects() As RECT 'The list of rectangles to be drawn.
Public ScreenRect As RECT   'The screen size.

Public Sub AddDirtyRect(TheRect As RECT)
    'Add a rectangle to the list of rectangles to be drawn.
    Dim i As Long
    i = UBound(DirtyRects, 1) + 1
    ReDim Preserve DirtyRects(0 To i)
    DirtyRects(i) = TheRect
End Sub

Public Sub MergeRects()
    Dim Max As Long, i As Long, RetVal As Long, j As Long, TempRect As RECT
    Max = UBound(DirtyRects, 1)
    If Max > 2 Then
        If Max > 10 Then
            'bah. just empty out and start over, there's too many.
            ReDim DirtyRects(0 To 1)
            DirtyRects(1) = ScreenRect
        Else
            For i = 2 To Max
                'dirtyrect(i) is the rectangle to be compared against
                If DirtyRects(i).Right = 0 Then
                    'skip it - emptied rectangle
                Else
                    For j = 2 To Max
                        If i = j Then
                            'same rectangle
                        ElseIf DirtyRects(j).Right = 0 Then
                            'emptied rectangle
                        Else
                            RetVal = IntersectRect(TempRect, DirtyRects(i), DirtyRects(j))
                            If RetVal > 0 Then
                                'they intersect - make the first rectangle larger.
                                With DirtyRects(i)
                                    .Left = IIf(.Left < DirtyRects(j).Left, .Left, DirtyRects(j).Left)
                                    .Right = IIf(.Right > DirtyRects(j).Right, .Right, DirtyRects(j).Right)
                                    .Top = IIf(.Top < DirtyRects(j).Top, .Top, DirtyRects(j).Top)
                                    .Bottom = IIf(.Bottom > DirtyRects(j).Bottom, .Bottom, DirtyRects(j).Bottom)
                                End With
                                'empty out the second rectangle
                                With DirtyRects(j)
                                    .Left = 0: .Right = 0: .Top = 0: .Bottom = 0
                                End With
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
    End If
End Sub

