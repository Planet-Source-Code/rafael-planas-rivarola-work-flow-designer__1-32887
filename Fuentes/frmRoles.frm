VERSION 5.00
Begin VB.Form frmRoles 
   Caption         =   "Usuarios"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstRoles 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstUsuarios 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   2640
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4740
      TabIndex        =   2
      Top             =   2640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1395
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdAgregarRol 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Picture         =   "frmRoles.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdEliminarRol 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   360
         Picture         =   "frmRoles.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   3420
      TabIndex        =   3
      Top             =   360
      Width           =   3480
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Picture         =   "frmRoles.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   360
         Picture         =   "frmRoles.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label lblEtiquetas 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEtiquetas 
      BackStyle       =   0  'Transparent
      Caption         =   "Roles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIr_Click()

End Sub

Private Sub cmdAceptar_Click()

    Dim i As Integer, a As Integer
    
    For i = 1 To UBound(ListaUsuarios)
        If ListaUsuarios(i - 1).Modificado = "1" Then
            adoConexion.BeginTrans
            adoConexion.Execute "Delete from Usuarios Where CodigoRol = '" & ListaUsuarios(i - 1).Codigo & "'"
            For a = 1 To UBound(ListaUsuarios(i - 1).Nombre)
                adoConexion.Execute "Insert Into Usuarios (CodigoRol, NombreUsuario) Values ('" & ListaUsuarios(i - 1).Codigo & "', '" & ListaUsuarios(i - 1).Usuario(a) & "')"
            Next
            adoConexion.CommitTrans
        End If
    Next
    frmPropiedades.txtPara = strUsuarios
    Unload Me
End Sub

Private Sub cmdAgregar_Click()

    'On Error Resume Next
    Dim objRecipientes As Object
    Dim strUsuarios As String, i As Integer, a As Integer
    Dim strCodigo As String
    
    Set objRecipientes = objSession.AddressBook(Title:="Lista de Usuarios - Wiese Sudameris", recipLists:=1, forceResolution:=True, parentWindow:=frmPropiedades.hwnd)
    For i = 1 To objRecipientes.Count
        frmRoles.lstUsuarios.AddItem objRecipientes(i).Name
        strCodigo = ListaUsuarios(lstRoles.ListIndex).Codigo
        ReDim ListaUsuarios(lstRoles.ListIndex).Nombre(UBound(ListaUsuarios(lstRoles.ListIndex).Nombre) + 1)
        ReDim ListaUsuarios(lstRoles.ListIndex).Usuario(UBound(ListaUsuarios(lstRoles.ListIndex).Usuario) + 1)
        For a = 1 To lstUsuarios.ListCount
            ListaUsuarios(lstRoles.ListIndex).Usuario(a) = lstUsuarios.List(a - 1)
        Next
    Next
    ListaUsuarios(lstRoles.ListIndex).Modificado = "1"
End Sub



Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim a As Integer, i As Integer
    Dim intElementosSeleccionados As Integer
    intElementosSeleccionados = lstUsuarios.SelCount
    ReDim ListaUsuarios(lstRoles.ListIndex).Nombre(UBound(ListaUsuarios(lstRoles.ListIndex).Nombre) - intElementosSeleccionados)
    ReDim ListaUsuarios(lstRoles.ListIndex).Usuario(UBound(ListaUsuarios(lstRoles.ListIndex).Usuario) - intElementosSeleccionados)
    For i = lstUsuarios.ListCount To 1 Step -1
        If lstUsuarios.Selected(i - 1) Then
            lstUsuarios.RemoveItem (i - 1)
        End If
    Next
    For a = 1 To lstUsuarios.ListCount
        ListaUsuarios(lstRoles.ListIndex).Usuario(a) = lstUsuarios.List(a - 1)
    Next
    ListaUsuarios(lstRoles.ListIndex).Modificado = "1"
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Dim i As Integer, a As Integer
    Dim adoTemporal As New ADODB.Recordset
    
    adoRecordset.CursorLocation = adUseServer
    adoRecordset.CursorType = adOpenStatic
    adoRecordset.Open "Select * from Roles", adoConexion
    ReDim ListaUsuarios(adoRecordset.RecordCount)
    
    For i = 1 To adoRecordset.RecordCount
        adoTemporal.CursorLocation = adUseServer
        adoTemporal.CursorType = adOpenStatic
        adoTemporal.Open "Select * from Usuarios Where CodigoRol = '" & adoRecordset("Codigo") & "'", adoConexion
        ReDim ListaUsuarios(i - 1).Nombre(adoTemporal.RecordCount)
        ReDim ListaUsuarios(i - 1).Usuario(adoTemporal.RecordCount)
        ListaUsuarios(i - 1).Codigo = adoRecordset("Codigo")
        lstRoles.AddItem adoRecordset("Descripcion")
        For a = 1 To adoTemporal.RecordCount
            ListaUsuarios(i - 1).Nombre(a) = adoRecordset("Descripcion")
            ListaUsuarios(i - 1).Usuario(a) = adoTemporal("NombreUsuario")
            adoTemporal.MoveNext
        Next
        adoRecordset.MoveNext
        adoTemporal.Close
    Next
    adoRecordset.Close
    lstRoles.ListIndex = 0
End Sub

Private Sub lstRoles_Click()
    Dim i As Integer
    lstUsuarios.Clear
    For i = 1 To UBound(ListaUsuarios(lstRoles.ListIndex).Nombre)
        lstUsuarios.AddItem ListaUsuarios(lstRoles.ListIndex).Usuario(i)
    Next
End Sub

