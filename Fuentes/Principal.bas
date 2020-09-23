Attribute VB_Name = "Principal"
Option Explicit

Public gLeftMargin As Integer
Public gRightMargin As Integer
Public gTopMargin As Integer
Public gBottomMargin As Integer

Public gprint As Boolean


Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type



Sub Main()
    gLeftMargin = 1                   ' Initialize
    gRightMargin = 1
    gTopMargin = 1
    gBottomMargin = 1
    gprint = False
    Ventana.Show
End Sub

