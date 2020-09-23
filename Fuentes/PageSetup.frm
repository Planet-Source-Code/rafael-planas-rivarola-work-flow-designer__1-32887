VERSION 5.00
Begin VB.Form FrmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Impresión"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "PageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   60
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2490
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2490
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3390
      Begin VB.ComboBox cboLeftMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   1140
      End
      Begin VB.ComboBox cboRightMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   795
         Width           =   1140
      End
      Begin VB.ComboBox cboTopMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1245
         Width           =   1140
      End
      Begin VB.ComboBox cboBottomMargin 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1695
         Width           =   1140
      End
      Begin VB.Label lblLeftMargin 
         AutoSize        =   -1  'True
         Caption         =   "Izquierda (Pulg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lblRightMargin 
         AutoSize        =   -1  'True
         Caption         =   "Derecha  (Pulg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label lblTopMargin 
         AutoSize        =   -1  'True
         Caption         =   "Superior  (Pulg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1275
         Width           =   1335
      End
      Begin VB.Label lblBottomMargin 
         AutoSize        =   -1  'True
         Caption         =   "Inferior   (Pulg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   250
         TabIndex        =   6
         Top             =   1740
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2490
      Width           =   1035
   End
End
Attribute VB_Name = "FrmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    gprint = False
    Dim i, m
    m = 4
    For i = 0 To m Step 0.05
        FrmPageSetup.cboLeftMargin.AddItem Format(i, "0.00")
    Next
    For i = 0 To m Step 0.05
        FrmPageSetup.cboRightMargin.AddItem Format(i, "0.00")
    Next
    For i = 0 To m Step 0.05
        FrmPageSetup.cboTopMargin.AddItem Format(i, "0.00")
    Next
    For i = 0 To m Step 0.05
        FrmPageSetup.cboBottomMargin.AddItem Format(i, "0.00")
    Next
    FrmPageSetup.cboLeftMargin.Text = cboLeftMargin.List(gLeftMargin / 0.05)
    FrmPageSetup.cboRightMargin.Text = cboRightMargin.List(gRightMargin / 0.05)
    FrmPageSetup.cboTopMargin.Text = cboTopMargin.List(gTopMargin / 0.05)
    FrmPageSetup.cboBottomMargin.Text = cboBottomMargin.List(gBottomMargin / 0.05)
End Sub


Private Sub cmdOK_Click()
    gLeftMargin = cboLeftMargin.Text
    gRightMargin = cboRightMargin.Text
    gTopMargin = cboTopMargin.Text
    gBottomMargin = cboBottomMargin.Text
    Unload Me
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdPrint_Click()
    gprint = True
    Unload Me
End Sub

