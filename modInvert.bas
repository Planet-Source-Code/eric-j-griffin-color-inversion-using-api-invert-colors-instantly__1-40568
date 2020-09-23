Attribute VB_Name = "modInvert"
Option Explicit

'## Api Declarations
Public Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'## Type Declaration
Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
