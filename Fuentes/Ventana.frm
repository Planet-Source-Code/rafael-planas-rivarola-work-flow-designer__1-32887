VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.MDIForm Ventana 
   BackColor       =   &H8000000C&
   Caption         =   "Modelador de Procesos"
   ClientHeight    =   4470
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9030
   Icon            =   "Ventana.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4185
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4948
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2716
            MinWidth        =   2716
            Text            =   "DESP "
            TextSave        =   "DESP "
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2011
            MinWidth        =   2011
            Text            =   "INACTIVO"
            TextSave        =   "INACTIVO"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2187
            Picture         =   "Ventana.frx":0582
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2011
            MinWidth        =   2011
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "03:27 p.m."
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1170
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Nuevo 
         Caption         =   "&Nuevo Proceso"
         Index           =   0
      End
      Begin VB.Menu Nuevo 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Nuevo 
         Caption         =   "&Salir"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Ventana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
    Me.Move 0, 0, Screen.Width, Screen.Height
    ArchivoACargar = Trim(Command)
    ReDim Flujo(0)
    Dim frmD As PizarraFlujos
    Set frmD = New PizarraFlujos
    Load frmD
End Sub


Private Sub Nuevo_Click(Index As Integer)
    Select Case Index
           Case 2
                End
           Case 0
                Dim frmD As PizarraFlujos
                Set frmD = New PizarraFlujos
                Load frmD
   End Select
End Sub

Private Sub Timer1_Timer()
    Exit Sub
    Dim lData As Long, lType As Long, lSize As Long
    Dim hKey As Long, Statbar As Double
                   
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
                    
    'If there's a problem accessing the registry
    If Qry <> 0 Then
         MsgBox "Can't Open Statistics Key"
         End
    End If
                    
    lType = REG_DWORD
    lSize = 4
                    
    'Querying the registry for CPUUsage
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", _
        0, lType, lData, lSize)
                                    
                                    
    'statbar is the variable that holds the CPU
    'usage divided by 10.
    '(ex. if 79% of the CPU is being used then
    ' statbar will hold the int(7.9) = 8)
    Statbar = lData / 10
    If Statbar >= 1 Then Statbar = Statbar - 1
    
    'used to fill the SSPanel with the color green
    'beginning with 0 and ending with the value of
    'statbar.
'    ProgressBar1.Max = 10
'    ProgressBar1.Value = Statbar
                    
    Ventana.StatusBar1.Panels.Item(5).Text = " CPU = " + Format$(lData, "##0") + "%"
    Qry = RegCloseKey(hKey)

End Sub



