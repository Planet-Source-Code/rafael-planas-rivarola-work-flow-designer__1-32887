VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior = 0   'vbNone
'  MTSTransactionMode = 0   'NotAnMTSObject
'End
'   ###########################################################
'   #
'   #   Filename    : clsMouseOver.cls
'   #   Author      : Albert Peck (albert.peck@inter.net.uk)
'   #   Comment     : New MouseOver Class
'   #
'   #   To Use
'   #   ------
'   #
'   #   Create and then setup this class in the form_load event,
'   #   using SetBorderStyle(Me.BorderStyle).
'   #
'   #   Then use IsMouseOver(Me, [AnyGraphicalObject]) at any
'   #   point to recieve a boolean value back if the mouse is
'   #   over the geven object.
'   #
'   ###########################################################

Option Explicit

Private intBorder As Integer    ' To hold the borderstyle of the current form
Private RelativeObjectX         ' Hold the X pos of the object
Private RelativeObjectY         ' Hold the Y pos of the object
Private MousePos As POINTAPI    ' To gather the mouse position

Private Const OffsetX = 6       ' The offset when there is a border round the form
Private Const OffsetY = 42      ' The offset when there is a title bar on the form

Private Type POINTAPI           ' Type to retieve mouse position
    X As Long
    Y As Long
End Type

' Declare API call to get mouse position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' To setup the correct border style
Public Sub SetBorderStyle(intBorderStyle As Integer)

    intBorder = intBorderStyle

End Sub

' Detect if the mouse if over the given object on the give form
Public Function IsMouseOver(frmHost As Form, objSubject As Object, Optional Pixel As Boolean = True)

    ' Work out the relative position of the object in question
    Select Case intBorder

        Case 0
        
            ' Plain form
            RelativeObjectX = ((frmHost.Left) / IIf(Pixel, Screen.TwipsPerPixelX, 1)) + objSubject.Left + 6
            RelativeObjectY = ((frmHost.Top) / IIf(Pixel, Screen.TwipsPerPixelY, 1)) + objSubject.Top + 25
        
        Case Else
            
            ' Has title bar and border
            RelativeObjectX = ((frmHost.Left) / IIf(Pixel, Screen.TwipsPerPixelX, 1)) + OffsetX + objSubject.Left + 6
            RelativeObjectY = ((frmHost.Top) / IIf(Pixel, Screen.TwipsPerPixelY, 1)) + OffsetY + objSubject.Top + 25
    
    End Select

    ' Call API to get mouse position
    GetCursorPos MousePos
    
    ' See if mouse lays with the boundaries of the object
'    Ventana.StatusBar1.Panels.Item(3).Text = Str(MousePos.x) + "," + Str(MousePos.y)
'    Ventana.StatusBar1.Panels.Item(2).Text = Str(RelativeObjectX) + "," + Str(RelativeObjectY)
    
    If (MousePos.X >= RelativeObjectX And MousePos.X <= RelativeObjectX + objSubject.Width) And (MousePos.Y >= RelativeObjectY And MousePos.Y <= RelativeObjectY + objSubject.Height) Then
    
        IsMouseOver = True
        
    Else
    
        IsMouseOver = False
    
    End If

    
End Function
    
