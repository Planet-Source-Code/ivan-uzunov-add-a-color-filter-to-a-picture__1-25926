Attribute VB_Name = "Filter"
'***********************************************************************************
' Author    : Ivan Uzunov
' E- mail   : kicheto@goatrance.com
' Purpose   : Add color filter to bitmap
' Made with : Visual Basic 6.00
' Date      : 08-07-2001
'***********************************************************************************
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Enum FilterEnum
       BlackWhite = 0
       Red = 1
       Blue = 2
       Green = 3
       Yellow = 4
End Enum

Public Sub AddFilter(picSource As PictureBox, ByVal Filter As FilterEnum)
Dim X As Long
Dim Y As Long
Dim R As Long
Dim NewColor As Long
Dim MaxValue As Long
    'First get the max RGB value
    MaxValue = RGB(255, 255, 255)
    Select Case Filter
           'All filters have a similar code
           Case BlackWhite:
                    'Start moving in the picturebox
                    For X = 0 To picSource.ScaleWidth
                        For Y = 0 To picSource.ScaleHeight
                            'get the color in a definite coordinates
                            R = GetPixel(picSource.hdc, X, Y)
                            'calculate the new color for Red, Blue and Green
                            NewColor = (R / MaxValue) * 255
                            'get the RGB value of the new color
                            NewColor = RGB(NewColor, NewColor, NewColor)
                            'set the new color in the same coordinates
                            Call SetPixel(picSource.hdc, X, Y, NewColor)
                        Next Y
                    Next X
            Case Red:
                    For X = 0 To picSource.ScaleWidth
                        For Y = 0 To picSource.ScaleHeight
                            R = GetPixel(picSource.hdc, X, Y)
                            NewColor = (R / MaxValue) * 255
                            If NewColor <= 255 / 2 Then
                               NewColor = RGB(NewColor, 0, 0)
                            Else
                               NewColor = RGB(255, NewColor, NewColor)
                            End If
                            Call SetPixel(picSource.hdc, X, Y, NewColor)
                        Next Y
                    Next X
            Case Blue:
                    For X = 0 To picSource.ScaleWidth
                        For Y = 0 To picSource.ScaleHeight
                            R = GetPixel(picSource.hdc, X, Y)
                            NewColor = (R / MaxValue) * 255
                            If NewColor <= 255 / 2 Then
                               NewColor = RGB(0, 0, NewColor)
                            Else
                               NewColor = RGB(NewColor, NewColor, 255)
                            End If
                            Call SetPixel(picSource.hdc, X, Y, NewColor)
                        Next Y
                    Next X
            Case Green:
                    For X = 0 To picSource.ScaleWidth
                        For Y = 0 To picSource.ScaleHeight
                            R = GetPixel(picSource.hdc, X, Y)
                            NewColor = (R / MaxValue) * 255
                            If NewColor <= 255 / 2 Then
                               NewColor = RGB(0, NewColor, 0)
                            Else
                               NewColor = RGB(NewColor, 255, NewColor)
                            End If
                            Call SetPixel(picSource.hdc, X, Y, NewColor)
                        Next Y
                    Next X
            Case Yellow:
                    For X = 0 To picSource.ScaleWidth
                        For Y = 0 To picSource.ScaleHeight
                            R = GetPixel(picSource.hdc, X, Y)
                            NewColor = (R / MaxValue) * 255
                            If NewColor <= 255 / 2 Then
                               NewColor = RGB(NewColor, NewColor, 0)
                            Else
                               NewColor = RGB(255, 255, NewColor)
                            End If
                            Call SetPixel(picSource.hdc, X, Y, NewColor)
                        Next Y
                    Next X
     End Select
    DoEvents
    picSource.Refresh
End Sub
