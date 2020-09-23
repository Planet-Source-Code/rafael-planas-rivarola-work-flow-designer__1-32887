VERSION 5.00
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roles"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfigurar 
      Caption         =   "&Configurar"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
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
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
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
      Height          =   1860
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lblEtiquetas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Roles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Dim i As Integer
    strUsuarios = ""
    For i = 1 To lstRoles.ListCount
        If lstRoles.Selected(i - 1) Then
            strUsuarios = strUsuarios & lstRoles.List(i - 1)
            strUsuarios = strUsuarios & "; "
        End If
    Next
    If Len(Trim(strUsuarios)) > 0 Then
       strUsuarios = Left(strUsuarios, Len(strUsuarios) - 2)
    End If
    
    frmPropiedades.txtPara = strUsuarios
    Unload Me
End Sub

Private Sub cmdConfigurar_Click()
    frmRoles.Show 1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    adoRecordset.CursorLocation = adUseServer
    adoRecordset.CursorType = adOpenStatic
    adoRecordset.Open "Select * from Roles", adoConexion
    
    For i = 1 To adoRecordset.RecordCount
        lstRoles.AddItem adoRecordset("Descripcion")
        adoRecordset.MoveNext
    Next
    adoRecordset.Close
    lstRoles.ListIndex = 0
End Sub
