Attribute VB_Name = "modPropiedades"
Option Explicit

Public Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long

Public objSession As Object
Public adoConexion As New ADODB.Connection
Public adoRecordset As New ADODB.Recordset
Public intNumeroTarea As Integer
Public intArregloGlobal As Integer
Public strUsuarios As String

Public ListaUsuarios() As Usuarios

Type Usuarios
    Codigo As String * 3
    Nombre() As String
    Usuario() As String
    Modificado As String * 1
End Type

Sub sConexion()
    adoConexion.Open "Sistema de Reclamos", "", ""
End Sub

Sub sSession()
    Dim objRecipientes As Object
    Set objSession = CreateObject("MAPI.Session")
    objSession.Logon newSession:=False
    'Set objRecipientes = objSession.AddressBook(Title:="Lista de Usuarios - Wiese Sudameris", recipLists:=1, forceResolution:=True, parentWindow:=frmPropiedades.hWnd)
End Sub

'Sub Main()
'    Call sConexion
'    Call sSession
'    'Call sCargarFormulario("FYI")
'    Call sCargarFormulario(1)
'    'frmPropiedades.Show (1)
'
'End Sub

Sub sCargarFormulario(intArreglo As Integer)
    'ReDim EventoTarea(1)
    'ReDim EventoPropiedades(1)
    Select Case GEventoTarea.TareaTipo
        Case 6
            With frmPropiedades
                .txtMensaje.Height = 3495
                .cmdCheck.Visible = False
                .cmdCancel.Visible = False
                .txtRespuesta.Visible = False
                .lvwRespuestas.Visible = False
                .Show 1
            End With
        Case 5
            With frmPropiedades
                .txtMensaje.Height = 1455
                .cmdCheck.Visible = True
                .cmdCancel.Visible = True
                .txtRespuesta.Visible = True
                .lvwRespuestas.Visible = True
                .Show 1
            End With
        Case Else
            
            'MsgBox "No definido aun"
    End Select
'    frmPropiedades.txtNombreTarea = Trim(GEventoTarea.Definicion)
'    frmPropiedades.txtMensaje = Trim(GEventoPropiedades.Mensaje)
    frmPropiedades.Show 1
    
End Sub
