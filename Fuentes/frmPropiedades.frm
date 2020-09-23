VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPropiedades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades de la Tarea"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmPropiedades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   1455
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
      Left            =   4440
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtNombreTarea 
      Appearance      =   0  'Flat
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
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   4335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mensaje"
      TabPicture(0)   =   "frmPropiedades.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAsunto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtMensaje"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPara"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPara"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAsunto"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbNombreLetra"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbSize"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkFormato(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkFormato(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkFormato(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optAlineacion(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optAlineacion(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optAlineacion(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lvwRespuestas"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdCancel"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdCheck"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtRespuesta"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.TextBox txtRespuesta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CommandButton cmdCheck 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         Picture         =   "frmPropiedades.frx":0028
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3360
         Width           =   255
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   480
         Picture         =   "frmPropiedades.frx":012A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3360
         Width           =   255
      End
      Begin ComctlLib.ListView lvwRespuestas 
         Height          =   1575
         Left            =   240
         TabIndex        =   14
         Top             =   3720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.OptionButton optAlineacion 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   6120
         Picture         =   "frmPropiedades.frx":022C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton optAlineacion 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   5760
         Picture         =   "frmPropiedades.frx":032E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton optAlineacion 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   5400
         Picture         =   "frmPropiedades.frx":0430
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox chkFormato 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   4920
         Picture         =   "frmPropiedades.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox chkFormato 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   4560
         Picture         =   "frmPropiedades.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox chkFormato 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   4200
         Picture         =   "frmPropiedades.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox cmbSize 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmPropiedades.frx":0838
         Left            =   3240
         List            =   "frmPropiedades.frx":0851
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNombreLetra 
         Appearance      =   0  'Flat
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
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "cmbNombreLetra"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtAsunto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   6135
      End
      Begin VB.TextBox txtPara 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   6135
      End
      Begin VB.CommandButton cmdPara 
         Appearance      =   0  'Flat
         Caption         =   "&Para..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin RichTextLib.RichTextBox txtMensaje 
         Height          =   1455
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2566
         _Version        =   393217
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmPropiedades.frx":0870
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   240
         X2              =   7080
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblAsunto 
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPropiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkFormato_Click(Index As Integer)
    Select Case Index
        Case 0
            txtMensaje.SelBold = Not txtMensaje.SelBold
        Case 1
            txtMensaje.SelItalic = Not txtMensaje.SelItalic
        Case 2
            txtMensaje.SelUnderline = Not txtMensaje.SelUnderline
    End Select
End Sub

Private Sub cmbNombreLetra_Click()
    txtMensaje.SelFontName = cmbNombreLetra.Text
End Sub

Private Sub cmbSize_Click()
    txtMensaje.SelFontSize = cmbSize.Text
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    GEventoPropiedades.Para = txtPara
    GEventoPropiedades.Asunto = txtAsunto
    GEventoPropiedades.Mensaje = txtMensaje
    GEventoPropiedades.Definicion = txtNombreTarea
    GEventoTarea.Definicion = txtNombreTarea
    
    GEventoPropiedades.NoCondiciones = lvwRespuestas.ListItems.Count
    ReDim GEventoPropiedades.Condicion(lvwRespuestas.ListItems.Count)
    For i = 1 To lvwRespuestas.ListItems.Count
        GEventoPropiedades.Condicion(i).CondicionActiva = True
        GEventoPropiedades.Condicion(i).NoCondicion = i
        GEventoPropiedades.Condicion(i).Definicion = frmPropiedades.lvwRespuestas.ListItems.Item(i).Text
    Next
    Unload Me
    Load frmPropiedades
End Sub

Private Sub cmdCancel_Click()
    Dim i As Byte
    If lvwRespuestas.SelectedItem.Selected Then
        i = lvwRespuestas.SelectedItem.Index
        lvwRespuestas.ListItems.Remove (i)
    End If
End Sub

Private Sub cmdCancelar_Click()
'    adoConexion.Close
'    Unload Me
End Sub

Private Sub cmdCheck_Click()
    If Trim(txtRespuesta) = "" Then
        Exit Sub
    End If
    With frmPropiedades.lvwRespuestas
        .ListItems.Add , , txtRespuesta
        txtRespuesta.Text = ""
        txtRespuesta.SetFocus
    End With
End Sub

Private Sub cmdPara_Click()
    frmUsuarios.Show 1
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    Me.Refresh
'    For i = 1 To Screen.FontCount
'        If Trim(Screen.Fonts(i)) <> "" Then
'            cmbNombreLetra.AddItem Screen.Fonts(i)
'        End If
'    Next
    
    cmbNombreLetra.Text = txtMensaje.Font.Name
    cmbSize.ListIndex = 0
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    lvwRespuestas.ColumnHeaders.Add , , "", lvwRespuestas.Width - 280
    lvwRespuestas.View = lvwReport
    txtNombreTarea = GEventoTarea.Definicion
    For i = 1 To GEventoPropiedades.NoCondiciones
         frmPropiedades.lvwRespuestas.ListItems.Add = GEventoPropiedades.Condicion(i).Definicion
    Next
End Sub

Private Sub optAlineacion_Click(Index As Integer)
    Select Case Index
        Case 0
            txtMensaje.SelAlignment = 0
        Case 1
            txtMensaje.SelAlignment = 2
        Case 2
            txtMensaje.SelAlignment = 1
    End Select
End Sub


Private Sub txtRespuesta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdCheck = True
    End If
End Sub

