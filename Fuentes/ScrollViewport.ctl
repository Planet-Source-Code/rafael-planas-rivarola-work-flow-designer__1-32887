VERSION 5.00
Begin VB.UserControl ScrollViewport 
   Alignable       =   -1  'True
   BackColor       =   &H8000000C&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ControlContainer=   -1  'True
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   50
   ToolboxBitmap   =   "ScrollViewport.ctx":0000
   Begin VB.HScrollBar sclHorizontal 
      Height          =   240
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   480
   End
   Begin VB.VScrollBar sclVertical 
      Height          =   480
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox imgNook 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "ScrollViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// CONTROL ACTIVEX USR_ScrollViewport
'// SCROLLABLE VIEWPORTS
'// Por: Rafael Planas Rivarola, creado Nov-1999, revizado Dic-1999
Option Explicit

'//ENUMS
Public Enum enm_ViewScroll
    svpBoth
    svpVertical
    svpHorizontal
End Enum

Private cnth As Byte
Private cntv As Byte

'//CONSTANTES
Private Const NookSize = 240 / 15         '//Twips
Private Const MinWidth = NookSize * 3  '//Control
Private Const MinHeight = NookSize * 3 '//Control

'//MIEMBROS
Private m_ViewPort      As PictureBox
Private m_ViewScroll    As enm_ViewScroll
Private m_ViewContainer As String
Private m_HValue        As Long
Private m_VValue        As Long

'//VARIABLES
Private mblnViewPort As Boolean
Private CurNookSize  As Long

'//EVENTOS
Event HorizontalScroll(Stat As Byte)
Event VerticalScroll(Stat As Byte)
Event MouseMove()


Private Sub sclHorizontal_Scroll()
    If mblnViewPort Then
       m_ViewPort.Left = -sclHorizontal.Value
       RaiseEvent HorizontalScroll(0)
    Else
       If m_ViewPort.Left > 0 Then
          RaiseEvent HorizontalScroll(1)
       End If
    End If
End Sub

Private Sub sclVertical_Scroll()
    If mblnViewPort Then
       m_ViewPort.Top = -sclVertical.Value
       RaiseEvent VerticalScroll(0)
    End If
End Sub

Private Sub UserControl_Initialize()
    '//Inicializar los scroll bars.
    With sclVertical
        .Top = 0
        .Width = NookSize
        .SmallChange = 90
        .LargeChange = 180
    End With
    With sclHorizontal
        .Left = 0
        .Height = NookSize
        .SmallChange = 90
        .LargeChange = 180
    End With
    CurNookSize = NookSize
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub UserControl_Resize()
    '//Filtra tmaño minimo
    If Width < MinWidth Then
       Width = MinWidth
       Exit Sub
    End If
    If Height < MinHeight Then
       Height = MinHeight
       Exit Sub
    End If
    '//Reubicar controles scroll
    sclVertical.Left = ScaleWidth - NookSize
    sclVertical.Height = ScaleHeight - CurNookSize
    sclHorizontal.Top = ScaleHeight - NookSize
    sclHorizontal.Width = ScaleWidth - CurNookSize
    '//Mueve nook image
    If ViewScroll = svpBoth Then
       imgNook.Move sclHorizontal.Width, sclVertical.Height
    End If
    '//Scroll
    If mblnViewPort Then
       Call ConfigScrollBars
    End If
End Sub

Private Sub sclHorizontal_Change()
    If mblnViewPort Then
       m_ViewPort.Left = -sclHorizontal.Value
       If sclHorizontal.Value < sclHorizontal.Max Then
          RaiseEvent HorizontalScroll(0)
       Else
          cnth = cnth + 1
          sclHorizontal.Max = sclHorizontal.Max + 1
          Debug.Print cnth
          If cnth > 3 Then
             RaiseEvent HorizontalScroll(1)
             cnth = 0
          End If
       End If
    End If
End Sub

Private Sub sclVertical_Change()
    If mblnViewPort Then
       m_ViewPort.Top = -sclVertical.Value
       
       If sclVertical.Value < sclVertical.Max Then
          RaiseEvent VerticalScroll(0)
       Else
          cntv = cntv + 1
          sclVertical.Max = sclVertical.Max + 1
          Debug.Print cntv
          If cntv > 3 Then
             RaiseEvent VerticalScroll(1)
             cntv = 0
          End If
       End If
    End If
End Sub

Private Sub ConfigScrollBars()
    '//Posicion del scroll bar horizontal .
    Dim CurScaleMode As ScaleModeConstants
    CurScaleMode = Extender.Container.ScaleMode
    Extender.Container.ScaleMode = vbTwips
    If sclHorizontal.Visible Then
       With sclHorizontal
           If Extender.Width < m_ViewPort.Width Then
             .Max = m_ViewPort.Width - Extender.Width
             .Enabled = True
           Else
              m_ViewPort.Left = 0
             .Value = 0
             .Enabled = False
           End If
       End With
    End If
    '//Posicion del scroll bar vertical.
    If sclVertical.Visible Then
       With sclVertical
           If Extender.Height < m_ViewPort.Height Then
             .Max = m_ViewPort.Height - Extender.Height
             .Enabled = True
           Else
              m_ViewPort.Top = 0
             .Value = 0
             .Enabled = False
           End If
       End With
    End If
    Extender.Container.ScaleMode = CurScaleMode
End Sub

'//PROPERTY: BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Get HValue() As Long
    HValue = sclHorizontal.Value
End Property

Public Property Let HValue(ByVal m_HVal As Long)
    If m_HVal < 1 Then m_HVal = 0
    If m_HVal > sclHorizontal.Max Then m_HVal = sclHorizontal.Max
    sclHorizontal.Value = m_HVal
'    sclHorizontal.Max = m_HVal
    PropertyChanged "HValue"
End Property

Public Property Get VValue() As Long
    VValue = sclVertical.Value
End Property

Public Property Get VMaxValue() As Long
    VMaxValue = sclVertical.Max
End Property

Public Property Get HMaxValue() As Long
    HMaxValue = sclHorizontal.Max
End Property

Public Property Let VValue(ByVal m_VVal As Long)
    If m_VVal < 1 Then m_VVal = 0
    If m_VVal > sclVertical.Max Then m_VVal = sclVertical.Max
    sclVertical.Value = m_VVal
    PropertyChanged "VValue"
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'//PROPERTY: ViewScroll
Public Property Get ViewScroll() As enm_ViewScroll
    ViewScroll = m_ViewScroll
End Property

Public Property Let ViewScroll(ByVal New_ViewScroll As enm_ViewScroll)
    m_ViewScroll = New_ViewScroll
    Call LetViewScroll
    PropertyChanged "ViewScroll"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H808080)
    m_ViewScroll = PropBag.ReadProperty("ViewScroll", svpBoth)
    m_ViewContainer = PropBag.ReadProperty("ViewContainer", "")
    m_HValue = PropBag.ReadProperty("HValue", 0)
    m_VValue = PropBag.ReadProperty("VValue", 0)
    '//Acciones
    Call LetViewScroll
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H808080)
    Call PropBag.WriteProperty("ViewScroll", m_ViewScroll, svpBoth)
    Call PropBag.WriteProperty("ViewContainer", m_ViewContainer, "")
    Call PropBag.WriteProperty("HValue", m_HValue, 0)
    Call PropBag.WriteProperty("VValue", m_VValue, 0)
End Sub

Private Sub LetViewScroll()
    sclHorizontal.Visible = (m_ViewScroll = svpBoth Or m_ViewScroll = svpHorizontal)
    sclVertical.Visible = (m_ViewScroll = svpBoth Or m_ViewScroll = svpVertical)
    CurNookSize = IIf(m_ViewScroll = svpBoth, NookSize, 0)
    imgNook.Visible = (m_ViewScroll = svpBoth)
    '//Reloc scroll controls
    sclVertical.Height = ScaleHeight - CurNookSize
    sclHorizontal.Width = ScaleWidth - CurNookSize
    '//Move nook image
    If ViewScroll = svpBoth Then
       imgNook.Move sclHorizontal.Width, sclVertical.Height
    End If
End Sub

'//PROPERTY: ViewContainer
Public Property Get ViewContainer() As String
    ViewContainer = m_ViewContainer
End Property

Public Property Let ViewContainer(ByVal New_ViewContainer As String)
    m_ViewContainer = New_ViewContainer
    PropertyChanged "ViewContainer"
End Property

Public Sub Refresh()
    Dim ctl As Control
    
    On Error GoTo ErrorHandler
    
    '//cambia tamaño del contenedor
    If mblnViewPort Then
       ConfigScrollBars
       Exit Sub
    End If
    
    '//verifica parametros
    mblnViewPort = False
    
    For Each ctl In Extender.Parent.Controls
        If TypeName(ctl) = "PictureBox" Then
           If LCase(ctl.Name) = LCase(m_ViewContainer) Then
              Set m_ViewPort = ctl
              With m_ViewPort
                  Set .Container = Extender
                 .BorderStyle = vbBSNone
                 .Move 0, 0, .Width + NookSize, .Height + NookSize
                 .ZOrder 1
              End With
              Call ConfigScrollBars
              mblnViewPort = True
              Exit For
           End If
        End If
    Next
    Exit Sub
    
ErrorHandler:
    MsgBox "Ha fallado la asignacion del control contenedor.", vbInformation
    Exit Sub
    UserControl.Enabled = False
End Sub


