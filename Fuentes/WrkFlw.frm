VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6FBB0048-0642-11D2-89AB-448504C10000}#1.0#0"; "PopNote.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid.ocx"
Begin VB.Form PizarraFlujos 
   ClientHeight    =   7485
   ClientLeft      =   615
   ClientTop       =   795
   ClientWidth     =   8655
   Icon            =   "WrkFlw.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   Begin ComctlLib.Toolbar HerramientasMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   19
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NUEVO"
            Description     =   "Permite Cargar o Crear un Nuevo Flujo de Proceso"
            Object.ToolTipText     =   "Cargar Nuevo Proceso"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LIMPIAR"
            Description     =   "Inicializara la Pizarra de Diseño a Fin de Recrear el Proceso Activo."
            Object.ToolTipText     =   "Limpiar Proceso"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ABRIR"
            Description     =   "Cargara un Proceso Previamente Grabado."
            Object.ToolTipText     =   "Abrir Proceso Existente"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GRABAR"
            Description     =   "Grabara Cualquir Cambio Realizado en la Pizarra de Diseño de Procesos."
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            Object.Width           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "AUMENTAR"
            Description     =   "Incrementara la Capacidad de Graficación de la Pizarra de Diseño."
            Object.ToolTipText     =   "Aumentar Tamaño de Pizarra"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "REDUCIR"
            Description     =   "Disminuira la Capacidad de Graficación de la Pizarra de Diseño."
            Object.ToolTipText     =   "Reducir Tamaño de Pizarra"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "HERRAMIENTA"
            Description     =   "Muestra u Oculta la Barra de Herraminetas de Diseño de Procesos."
            Object.ToolTipText     =   "Mostrar Herramientas"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DATOS"
            Description     =   "Muestra u Oculta el Panel de Datos e Informacion del Diseño de Procesos."
            Object.ToolTipText     =   "Mostrar Datos e Información del Proceso"
            Object.Tag             =   ""
            ImageIndex      =   14
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PUBLICAR"
            Description     =   "Publicara el Proceso Depurado Declarandolo Vigente."
            Object.ToolTipText     =   "Publicar Proceso"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "APAGAR"
            Description     =   "Terminara la Vigencia del Proceso Publicado."
            Object.ToolTipText     =   "Apagar Proceso"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PROPIEDADES"
            Description     =   "Muestra las Propiedades del Proceso Diseñado."
            Object.ToolTipText     =   "Propiedades del Proceso"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "AYUDA"
            Description     =   "Muestra Acerca de Modelador de Procesos"
            Object.ToolTipText     =   "Ayuda"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "DESHACER"
            Object.ToolTipText     =   "DesHacer"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar Herramientas 
      Height          =   5010
      Left            =   7680
      TabIndex        =   2
      Top             =   420
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   8837
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NORMAL"
            Object.ToolTipText     =   "Selección"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NOTIF"
            Description     =   "Adiciona una tarea o paso 'Informar a : ... de : ...'"
            Object.ToolTipText     =   "Informar"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NOTIF_RESP"
            Description     =   "Adiciona una tarea o paso 'Toma de decisión'"
            Object.ToolTipText     =   "Toma de Decisión"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BIFURC"
            Description     =   "Adiciona una tarea o paso 'Retorno' con registro del evento."
            Object.ToolTipText     =   "Retorno"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RETORNO"
            Description     =   "Adiciona una tarea o paso 'Bifurcación'"
            Object.ToolTipText     =   "Bifurcación "
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "META"
            Description     =   "Adiciona una tarea o paso 'Meta o Hito', con registro del evento."
            Object.ToolTipText     =   "Hito"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SUBPROCESO"
            Description     =   "Adiciona una tarea o paso 'SubProceso', ejecutandolo y continuando con el proceso."
            Object.ToolTipText     =   "Sub Proceso"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "FINAL"
            Description     =   "Adiciona una tarea o paso 'Fin de Proceso' registrando este evento."
            Object.ToolTipText     =   "Fin de Proceso"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "INTERCAMBIO"
            Object.ToolTipText     =   "Intercambio de Información"
            Object.Tag             =   ""
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "RELACION"
            Description     =   "Relaciona tareas y define el flujo de proceso."
            Object.ToolTipText     =   "Relacionar Tareas"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "ELIMINAR"
            Description     =   "Elimina tareas, relaciones o notas del diseño de proceso."
            Object.ToolTipText     =   "Elimina"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "INSERT"
            Description     =   "Inserta una tarea o paso entro dos tareas relacionadas."
            Object.ToolTipText     =   "Insertar Tarea"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   2
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NOTAS"
            Description     =   "Adiciona una Nota o Recordatorio."
            Object.ToolTipText     =   "Insertar Anotaciones"
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ARRASTRE"
            Description     =   "Enciende o Apaga el desplazamiento por arrastre."
            Object.ToolTipText     =   "Desplazamiento de Pantalla"
            Object.Tag             =   ""
            ImageIndex      =   13
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7320
      Top             =   4800
   End
   Begin VB.PictureBox Pizarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2580
      MousePointer    =   99  'Custom
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   246
      TabIndex        =   0
      Top             =   540
      Width           =   3720
      Begin VB.PictureBox ObjEvento 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   0
         Left            =   2385
         ScaleHeight     =   390
         ScaleWidth      =   2190
         TabIndex        =   8
         Top             =   4860
         Visible         =   0   'False
         Width           =   2190
      End
      Begin POPNOTE.PopUpNote Note 
         Index           =   0
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         BeginProperty NoteFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label DefNotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Esta es una prueba de ..."
         BeginProperty Font 
            Name            =   "Mirror"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   660
         TabIndex        =   18
         Top             =   4920
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image Notas 
         Height          =   375
         Index           =   0
         Left            =   1320
         Picture         =   "WrkFlw.frx":030A
         Stretch         =   -1  'True
         Top             =   4500
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape ShapeStat 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   390
         Index           =   0
         Left            =   2400
         Top             =   4410
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Shape ShapeMov 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Height          =   390
         Left            =   2400
         Top             =   3960
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label LabelIns 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Insertar"
         BeginProperty Font 
            Name            =   "Mirror"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   5280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape ShapeDel 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Height          =   390
         Left            =   2400
         Top             =   3480
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label ObjRelT1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ninguno"
         BeginProperty Font 
            Name            =   "Mirror"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   150
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   5520
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Line ObjRelS3 
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   219
         X2              =   300
         Y1              =   356
         Y2              =   356
      End
      Begin VB.Line ObjRelS2 
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   219
         X2              =   300
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line ObjRelS1 
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   219
         X2              =   300
         Y1              =   366
         Y2              =   366
      End
      Begin VB.Shape ObjRelP1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2370
         Shape           =   3  'Circle
         Top             =   5445
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape ObjRelP4 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2595
         Shape           =   3  'Circle
         Top             =   5445
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape ObjRelP2 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   2820
         Top             =   5445
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape ObjRelP3 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   3000
         Top             =   5445
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   0
      Left            =   225
      Picture         =   "WrkFlw.frx":074C
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   16
      Top             =   3465
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   1
      Left            =   225
      Picture         =   "WrkFlw.frx":343E
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   15
      Top             =   3105
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   2
      Left            =   225
      Picture         =   "WrkFlw.frx":6130
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   14
      Top             =   2745
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   3
      Left            =   225
      Picture         =   "WrkFlw.frx":8E22
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   13
      Top             =   2385
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   4
      Left            =   225
      Picture         =   "WrkFlw.frx":BB14
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   12
      Top             =   2025
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   5
      Left            =   225
      Picture         =   "WrkFlw.frx":E806
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   11
      Top             =   1665
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   6
      Left            =   225
      Picture         =   "WrkFlw.frx":114F8
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   10
      Top             =   1305
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   7
      Left            =   225
      Picture         =   "WrkFlw.frx":141EA
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   9
      Top             =   945
      Visible         =   0   'False
      Width           =   2190
   End
   Begin TabDlg.SSTab ContenedorDatos 
      Height          =   2085
      Left            =   300
      TabIndex        =   7
      Top             =   4560
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   3678
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   529
      BackColor       =   -2147483644
      TabCaption(0)   =   "Adjuntos a Proceso"
      TabPicture(0)   =   "WrkFlw.frx":16EDC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "HerramientasAdjuntos"
      Tab(0).Control(1)=   "Adjuntos"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Accesos del Proceso"
      TabPicture(1)   =   "WrkFlw.frx":16EF8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HerramientasAccesos"
      Tab(1).Control(1)=   "Accesos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Datos de Proceso"
      TabPicture(2)   =   "WrkFlw.frx":16F14
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "DatosAplicMtx"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Errores de Diseño"
      TabPicture(3)   =   "WrkFlw.frx":16F30
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "PizarraDatos(0)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Secuencia de Proceso"
      TabPicture(4)   =   "WrkFlw.frx":16F4C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "PizarraDatos(1)"
      Tab(4).ControlCount=   1
      Begin ComctlLib.Toolbar HerramientasAdjuntos 
         Height          =   420
         Left            =   -74880
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   741
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageListAdjuntos"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   6
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "ADJUNTAR"
               Description     =   "Adjunta un archivo para ser anexado al inicio del proceso."
               Object.ToolTipText     =   "Adjunta Archivo"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "ELIMINAR"
               Description     =   "Elimina el archivo adjunto seleccionado de la lista."
               Object.ToolTipText     =   "Eliminar archivo adjunto seleccionado."
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "ELIMINARTODOS"
               Description     =   "Elimina todos los archivos adjuntos definidos en la lista."
               Object.ToolTipText     =   "Eliminar todos los archivos adjuntos."
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "ABRIR"
               Description     =   "Abre el archivo adjunto seleccionado a fin de ser revisado."
               Object.ToolTipText     =   "Abrir archivo adjunto seleccionado."
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin ComctlLib.Toolbar HerramientasAccesos 
         Height          =   420
         Left            =   -74910
         TabIndex        =   24
         Top             =   135
         Visible         =   0   'False
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   741
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageListAdjuntos"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   6
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "ADJUNTAR"
               Description     =   "Adjunta un archivo para ser anexado al inicio del proceso."
               Object.ToolTipText     =   "Adjunta Archivo"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "ELIMINAR"
               Description     =   "Elimina el archivo adjunto seleccionado de la lista."
               Object.ToolTipText     =   "Eliminar archivo adjunto seleccionado."
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "ELIMINARTODOS"
               Description     =   "Elimina todos los archivos adjuntos definidos en la lista."
               Object.ToolTipText     =   "Eliminar todos los archivos adjuntos."
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "ABRIR"
               Description     =   "Abre el archivo adjunto seleccionado a fin de ser revisado."
               Object.ToolTipText     =   "Abrir archivo adjunto seleccionado."
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin vbAcceleratorGrid6.vbalGrid DatosAplicMtx 
         Height          =   1575
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2778
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
         Begin VB.TextBox TxtDatos 
            Height          =   285
            Left            =   1920
            TabIndex        =   28
            Top             =   600
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   600
            Visible         =   0   'False
            Width           =   2355
         End
      End
      Begin ComctlLib.ListView Adjuntos 
         Height          =   1110
         Left            =   -74880
         TabIndex        =   20
         Top             =   540
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1958
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "prueba"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin RichTextLib.RichTextBox PizarraDatos 
         Height          =   1455
         Index           =   1
         Left            =   -74775
         TabIndex        =   21
         Top             =   135
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         RightMargin     =   65535
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"WrkFlw.frx":16F68
      End
      Begin RichTextLib.RichTextBox PizarraDatos 
         Height          =   1455
         Index           =   0
         Left            =   -74820
         TabIndex        =   22
         Top             =   135
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         RightMargin     =   65535
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"WrkFlw.frx":16FEE
      End
      Begin ComctlLib.ListView Accesos 
         Height          =   1110
         Left            =   -74910
         TabIndex        =   23
         Top             =   540
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1958
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "prueba"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox Contenedor 
      Height          =   1185
      Left            =   270
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   5
      Top             =   3825
      Width           =   6540
      Begin WrkFlw.ScrollViewport VentanaMovil 
         Height          =   1155
         Left            =   60
         TabIndex        =   17
         Top             =   0
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   2037
         BackColor       =   -2147483636
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   840
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicVSplit 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   360
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   6
      Top             =   6960
      Width           =   2175
   End
   Begin VB.PictureBox TareaImagen 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   8
      Left            =   225
      Picture         =   "WrkFlw.frx":17074
      ScaleHeight     =   390
      ScaleWidth      =   2190
      TabIndex        =   25
      Top             =   585
      Visible         =   0   'False
      Width           =   2190
   End
   Begin ComctlLib.ImageList ImageListAdjuntos 
      Left            =   6900
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":19D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1A080
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1A39A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1A6B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList prueba 
      Left            =   6960
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1A9CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1ACE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1B002
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1B31C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1B636
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1B950
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1BC6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6705
      Top             =   2745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1BF84
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1C29E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1C5B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1C8D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1CBEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1CDC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1CFA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1D17A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1D354
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1D66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1D848
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1DA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1DD3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1E056
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1E370
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1E68A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1E9A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1ECBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1EFD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1F2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1F60C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageListCur 
      Left            =   6705
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1F7E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1FB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":1FE1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":20134
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":2044E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":20768
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":20A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":20D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":210B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":213D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6705
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":216EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":218C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":21A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":21C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":21E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":2202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":22206
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":223E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":225BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":22794
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":2296E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":22B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":22E62
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WrkFlw.frx":2317C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu Archivo 
         Caption         =   "&Cargar Nuevo Proceso"
         Index           =   0
      End
      Begin VB.Menu Archivo 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Archivo 
         Caption         =   "&Limpiar Proceso"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu Archivo 
         Caption         =   "&Abrir"
         Index           =   3
         Shortcut        =   ^A
      End
      Begin VB.Menu Archivo 
         Caption         =   "&Grabar"
         Index           =   4
         Shortcut        =   ^G
      End
      Begin VB.Menu Archivo 
         Caption         =   "Grabar &Como"
         Index           =   5
         Shortcut        =   ^R
      End
      Begin VB.Menu Archivo 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu Mnu_Print 
         Caption         =   "&Configurar Impresión"
         Index           =   0
      End
      Begin VB.Menu Mnu_Print 
         Caption         =   "&Previsualizar Impresion"
         Index           =   1
         Begin VB.Menu Prv1 
            Caption         =   "&Errores de Diseño"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu Prv1 
            Caption         =   "&Secuencia de Proceso"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu Mnu_Print 
         Caption         =   "&Imprimir"
         Index           =   2
         Begin VB.Menu prn1 
            Caption         =   "&Diseño de Proceso"
            Index           =   0
         End
         Begin VB.Menu prn1 
            Caption         =   "&Errores de Diseño"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu prn1 
            Caption         =   "&Secuencia de Proceso"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu Mnu_Print 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Mnu_Salir 
         Caption         =   "&Salir"
         Index           =   0
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MnuEditar 
      Caption         =   "&Editar"
      Visible         =   0   'False
      Begin VB.Menu TamañoPizarra 
         Caption         =   "Incrementar Tamaño de Pizarra"
         Index           =   0
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu TamañoPizarra 
         Caption         =   "Disminuir Tamaño de Pizarra"
         Index           =   1
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "&Ver"
      Visible         =   0   'False
      Begin VB.Menu Ver 
         Caption         =   "&Herramientas"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu MnuVentana 
      Caption         =   "&Ventana"
      Begin VB.Menu Mnu_vreo 
         Caption         =   "&Reordenar"
         Index           =   0
      End
      Begin VB.Menu Mnu_vreo 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFlujos 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu Mnu_Ayuda 
         Caption         =   "&Ayuda del Modelador de Procesos"
         Index           =   0
      End
      Begin VB.Menu Mnu_Ayuda 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Mnu_Ayuda 
         Caption         =   "Acerca de &Modelador de Procesos"
         Index           =   2
      End
   End
End
Attribute VB_Name = "PizarraFlujos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private FORM_NUM As Integer


Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type ApTx
    X As Long
    Y As Long
    TareaEncontrada As Integer
    Palabra As String
End Type

Private Type ANTERIOR
    Accion As Integer
    Elemento As Integer
    Activo  As Boolean
End Type

Dim Deshacer As ANTERIOR

Dim ApuntadorTexto As ApTx

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private clsMouser As New clsMouseOver


Dim EventoNota()        As ElementoNota
Dim EventoTarea()       As ElementoTarea
Dim EventoPropiedades() As PropiedadesTarea
Dim Relacion()          As ElementoRelacion
Dim DatosProceso()      As DatosProc
Dim r                   As RelacionCordenadas
Dim Plantilla           As WorkArea

Dim AplicMtx            As Grilla


Dim HerramientaSeleccionada As Integer
Dim EventoTareaSeleccionado As Integer
     

Dim NombreDeArchivo$
Dim RequiereGrabar%
Dim RequiereReproc%
Dim RequiereReproc1%
Dim RequiereReproc2%


Dim PDCancel%
Dim SecObj$
Dim NivelProc%
Dim DiasProc As Single
Dim Circular As Boolean

Dim DiasMax  As Single
Dim DiasMin  As Single
Dim ProcMax  As Single
Dim ProcMin  As Single
Dim Procesos As Integer

Dim cVS As New cSplitDDC

Dim DefTareaTipo(9) As String
Dim NoTareasTipo(9) As Integer

Dim ObjEnMovimiento As Integer
Dim EsPosibleRelacion As Boolean

Dim MaxProcObj As Integer
Dim MaxRelaObj As Integer
Dim MaxNotaObj As Integer

Dim TeclaStat As Boolean


Dim DespY As Integer
Dim DespX As Integer

Dim TareaTipo As Integer

' Que se esta Arrastrando?
Dim Arrastrando         As Boolean
Dim ArrastrandoRelacion As Boolean
Dim ArrastrandoNota     As Boolean
Dim ArrastrandoLinea    As Boolean

Dim NuevaLocRelacion    As Integer
Dim NuevaLocObjeto      As Integer

Dim RCInicial As RelacionCordenadas
Dim RCActual  As RelacionCordenadas
Dim RAnterior As Integer
Dim PAnterior As Integer

' coordenadas de arrastre.
Dim PosInicialX As Single
Dim PosInicialY As Single
Dim PosActualX  As Single
Dim PosActualY  As Single

Dim Splt    As Boolean
Dim SpltDif As Integer

Dim Btn As Integer


Private Sub Adjuntos_DblClick()
    If Adjuntos.SelectedItem.Index > 0 Then
       MsgBox "Hay que editar el adjunto '" + Adjuntos.SelectedItem.Text + "'", vbInformation, "Adjuntos al Proceso"
    End If
End Sub

Private Sub Archivo_Click(Index As Integer)
    Dim ret%
    Select Case Index
           Case 0
              Dim frmD As PizarraFlujos
              Set frmD = New PizarraFlujos
              Load frmD
           Case 2
              ret = GrabarCambios("Nuevo")
              If ret = IDNOSAVE Or ret = IDYES Or ret = IDNO Then
                LimpiarPizarra
              End If
           Case 3
              Call AbrirArchivo
           Case 4
              ret = Grabar()
           Case 5
              ret = GrabarComo()
    End Select
End Sub


Private Sub cboTipo_Click()
   If cboTipo.Visible Then
      Dim i As Long
      i = AplicMtx.Fila
      DatosAplicMtx.CellText(i, 3) = cboTipo.List(cboTipo.ListIndex)
      DatosProceso(i).Tipo = cboTipo.List(cboTipo.ListIndex)
      
      Select Case DatosAplicMtx.CellText(i, 3)
             Case "Caracter"
                  DatosAplicMtx.CellText(i, 4) = "10"
                  DatosAplicMtx.CellText(i, 5) = "Nulo"
                  DatosProceso(i).Longitud = "10"
                  DatosProceso(i).VDefecto = "Nulo"
             Case "Numerico"
                  DatosAplicMtx.CellText(i, 4) = "10.2"
                  DatosAplicMtx.CellText(i, 5) = "0.00"
                  DatosProceso(i).Longitud = "10.2"
                  DatosProceso(i).VDefecto = "0.00"
             Case "Moneda"
                  DatosAplicMtx.CellText(i, 4) = "12.2"
                  DatosAplicMtx.CellText(i, 5) = "0.00"
                  DatosProceso(i).Longitud = "12.2"
                  DatosProceso(i).VDefecto = "0.00"
             Case "Fecha - Hora"
                  DatosAplicMtx.CellText(i, 4) = "19"
                  DatosAplicMtx.CellText(i, 5) = ""
                  DatosProceso(i).Longitud = "19"
                  DatosProceso(i).VDefecto = ""
             Case "Logico"
                  DatosAplicMtx.CellText(i, 4) = "2"
                  DatosAplicMtx.CellText(i, 5) = "No"
                  DatosProceso(i).Longitud = "2"
                  DatosProceso(i).VDefecto = "No"
      End Select
      
      cboTipo.Visible = False
      DatosAplicMtx.SetFocus
      SendKeys "{RIGHT}"
      DatosAplicMtx.SelectedCol = AplicMtx.Columna
      
      RequiereGrabar = True
      RequiereReproc = True
      Call MuestraProceso
      
   End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      KeyAscii = 0
      cboTipo.Visible = False
      DatosAplicMtx.SetFocus
      Exit Sub
   End If
End Sub

Private Sub cboTipo_LostFocus()
   cboTipo.Visible = False
   DatosAplicMtx.CancelEdit
   DatosAplicMtx.SetFocus
End Sub

Private Sub Contenedor_DragDrop(Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos
End Sub

Private Sub ContenedorDatos_Click(PreviousTab As Integer)
    HerramientasAdjuntos.Visible = False
    Select Case ContenedorDatos.Tab
           Case 0
                HerramientasAdjuntos.ZOrder 0
                Adjuntos.ZOrder 0
                HerramientasAdjuntos.Visible = True
           Case 1
                HerramientasAccesos.ZOrder 0
                Accesos.ZOrder 0
                HerramientasAccesos.Visible = True
           Case 2
                DatosAplicMtx.ZOrder 0
           Case 3
                PizarraDatos(0).ZOrder 0
           Case 4
                PizarraDatos(1).ZOrder 0
    End Select
    PDCancel = True
    
    Call MuestraProceso(False)
    Call Form_Resize
End Sub

Private Sub ContenedorDatos_DragDrop(Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos
End Sub

Private Sub ContenedorDatos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     LabelIns.Visible = False
     ShapeMov.Visible = False
End Sub

Private Sub DatosAplicMtx_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
   Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
   Dim i As Long

   With AplicMtx
       .Fila = lRow
       .Columna = lCol
   End With
   
   Select Case lCol
          Case 1, 2, 4, 5
              If lCol > 1 And DatosAplicMtx.CellText(lRow, 1) = "(ninguno)" Then Exit Sub
              If lCol > 2 And DatosAplicMtx.CellText(lRow, 2) = "(ninguno)" Then Exit Sub
              If lCol <> 4 Or (lCol = 4 And DatosAplicMtx.CellText(lRow, 3) = "Caracter") Then
                 DatosAplicMtx.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
                 With TxtDatos
                    If lCol = 4 And DatosAplicMtx.CellText(lRow, 3) = "Caracter" Then
                        If (iKeyAscii >= Asc("0") And iKeyAscii <= Asc("9")) Or iKeyAscii = 13 Or iKeyAscii = 8 Then
                        Else
                           iKeyAscii = 0
                           Exit Sub
                        End If
                    End If
                    .Text = IIf(iKeyAscii > Asc(" "), Chr(iKeyAscii), DatosAplicMtx.CellText(lRow, lCol))
                    .SelStart = IIf(iKeyAscii > Asc(" "), 1, 0)
                    .SelLength = 65535
                    .Move lLeft, lTop, lWidth
                    .Visible = True
                    .ZOrder
                    .SetFocus
                 End With
               End If
          Case 3
               If DatosAplicMtx.CellText(lRow, 1) = "(ninguno)" Then Exit Sub
               If DatosAplicMtx.CellText(lRow, 2) = "(ninguno)" Then Exit Sub
               
               AplicMtx.Fila = lRow
               AplicMtx.Columna = lCol
                
               DatosAplicMtx.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
               With cboTipo
                    .Move lLeft, lTop, lWidth
                    For i = 0 To .ListCount - 1
                        If .List(i) = DatosAplicMtx.CellText(lRow, lCol) Then
                           .ListIndex = i
                           Exit For
                        End If
                    Next i
                    .Tag = lRow
                    .Visible = True
                    .ZOrder
                    .SetFocus
               End With
   End Select
End Sub

Private Sub DefNotas_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
     Dim X1 As Single, Y1 As Single
     X1 = Notas(Index).Left + (X / Screen.TwipsPerPixelX) - Notas(Index).Width
     Y1 = Notas(Index).Top + (Y / Screen.TwipsPerPixelY) + Notas(Index).Height
     Call Pizarra_DragDrop(Source, Int(X1), Int(Y1))
End Sub

Private Sub DefNotas_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
     Dim X1 As Single, Y1 As Single
     X1 = Notas(Index).Left + (X / Screen.TwipsPerPixelX) - Notas(Index).Width
     Y1 = Notas(Index).Top + (Y / Screen.TwipsPerPixelY) + Notas(Index).Height
     Call Pizarra_DragOver(Source, Int(X1), Int(Y1), State)
End Sub

Private Sub DefNotas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim X1 As Single, Y1 As Single
     X1 = Notas(Index).Left + (X / Screen.TwipsPerPixelX) - Notas(Index).Width
     Y1 = Notas(Index).Top + (Y / Screen.TwipsPerPixelY) + Notas(Index).Height
     Call Pizarra_MouseMove(Button, Shift, Int(X1), Int(Y1))
     DefNotas(Index).ZOrder 0
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cVS.SplitterFormMouseUp 0, 0
    Ventana.StatusBar1.Panels.Item(2).Text = IIf(TeclaStat, "DESP ENCENDIDO", "DESP APAGADO")
    Form_Resize
    
    If ObjEnMovimiento > 0 Then
       ObjEvento(ObjEnMovimiento).Drag 0: ObjEvento(ObjEnMovimiento).DragMode = 0: ObjEvento(ObjEnMovimiento).Visible = True
       ShapeDel.Visible = False
    End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos
End Sub

Private Sub Form_Initialize()
    MaxProcObj = 0
    MaxNotaObj = 0
    Splt = False
    SpltDif = 120
    
    ReDim Preserve DatosProceso(1)
    Dim Nt%
    DefTareaTipo(9) = "Intercambio de Información"
    DefTareaTipo(8) = "Inicio De Proceso"
    DefTareaTipo(7) = "Informar"
    DefTareaTipo(6) = "Toma De Decisión"
    DefTareaTipo(5) = "Retorno"
    DefTareaTipo(4) = "Bifurcacion"
    DefTareaTipo(3) = "Hito"
    DefTareaTipo(2) = "Sub Proceso"
    DefTareaTipo(1) = "Fin De Proceso"
    
    Call InicializaMtxDatosAplic: Nt = 0
    For Nt = 0 To 9: NoTareasTipo(Nt) = 0: Next
    If Len(Trim(ArchivoACargar)) > 0 Then
       If LeeArchivo(Trim(ArchivoACargar)) Then
       Else
          NombreDeArchivo = GetFileName(ArchivoACargar)
          Me.Caption = "Proceso - " + GetFileName(ArchivoACargar)
          Flujo(FORM_NUM).Caption = GetFileName(NombreDeArchivo)
          mnuFlujos(FORM_NUM - 1).Caption = Flujo(FORM_NUM).Caption
          
       End If
       ArchivoACargar = ""
    End If
    
End Sub


Private Sub Form_Load()

    Dim i%
    FORM_NUM = UBound(Flujo) + 1
    ReDim Preserve Flujo(FORM_NUM)
    
    
    NombreDeArchivo = "Intitulado.prc"
    Me.Caption = "Proceso - " + GetFileName(NombreDeArchivo)
    
    Dim Nuevo As Boolean
    clsMouser.SetBorderStyle Me.BorderStyle
    
    Flujo(FORM_NUM).Caption = GetFileName(NombreDeArchivo)
    Flujo(FORM_NUM).Handle = Me.hdc
    
    i = mnuFlujos.Count
    If i < FORM_NUM Then
       Load mnuFlujos(FORM_NUM - 1)
    End If
    For i = 1 To UBound(Flujo)
        mnuFlujos(i - 1).Caption = Flujo(i).Caption
    Next
    With cVS
        .Orientation = espHorizontal
        .Border(espbBottom) = 82
        .Border(espbTop) = 104
        .Border(espbLeft) = 2
        .Border(espbRight) = 2
        .SplitObject = PicVSplit
    End With
    
    With VentanaMovil
        .Align = vbAlignTop
        .BackColor = Pizarra.BackColor
        .ViewContainer = "Pizarra"
        .Refresh
    End With
    
        
    
    Call Form_Resize
    
    Call InitCPU
    
    If Len(Trim(ArchivoACargar)) > 0 Then
       PizarraDatos(0).Text = ""
       ContenedorDatos.Tab = 3
       PizarraDatos(0).ZOrder 0
    Else
       Call LimpiarPizarra
       ContenedorDatos.Tab = 3
       PizarraDatos(0).ZOrder 0
    End If
'    Ventana.StatusBar1.Panels.Item(2).Text = IIf(TeclaStat, "DESP ENCENDIDO", "DESP APAGADO")
    
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LabelIns.Visible = False
    cVS.SplitterFormMouseMove X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (cVS.SplitterFormMouseUp(X, Y)) Then
        Splt = True
        Form_Resize
    End If
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    If (Me.Height / Screen.TwipsPerPixelY) < 300 Then
       Me.Height = 300 * Screen.TwipsPerPixelY
       Exit Sub
    End If
    If (Me.Width / Screen.TwipsPerPixelX) < 420 Then
       Me.Width = 420 * Screen.TwipsPerPixelX
       Exit Sub
    End If
    
    Dim lH As Long, i%
    Contenedor.ScaleMode = 3
    
    If Not Splt Then
       If ContenedorDatos.Visible Then
          PicVSplit.Top = IIf(((Me.Height / Screen.TwipsPerPixelY) - SpltDif) < 104, 104, ((Me.Height / Screen.TwipsPerPixelY) - SpltDif))
       Else
          PicVSplit.Top = Me.ScaleHeight - 5 '  Me.Height / Screen.TwipsPerPixelY - 3
       End If
    Else
       SpltDif = (Me.Height / Screen.TwipsPerPixelY) - PicVSplit.Top
       Splt = False
    End If
    
    lH = PicVSplit.Top + 2 - HerramientasMenu.Height
    Herramientas.Align = 0
    Herramientas.Height = Me.ScaleHeight
    Herramientas.Left = (Me.ScaleWidth - Herramientas.ButtonWidth)
    With Contenedor
            .Move 0, _
            HerramientasMenu.Top + HerramientasMenu.Height + 2, _
            Me.ScaleWidth - IIf(Me.ScaleWidth > Herramientas.Width + 1, IIf(Herramientas.Visible, Herramientas.Width + 1, 1), 1), _
            lH
    End With
    cVS.Border(espbRight) = IIf(Me.ScaleWidth > Herramientas.Width + 1, IIf(Herramientas.Visible, Herramientas.Width + 1, 1), 1)
    PicVSplit.Width = Contenedor.Width
    
    VentanaMovil.Move 0, 0, Contenedor.ScaleWidth, Contenedor.ScaleHeight
    Pizarra.Height = IIf(VentanaMovil.Height * 15 + 30 * 15 > Pizarra.Height, VentanaMovil.Height * 15 + 30 * 15, Pizarra.Height)
    Pizarra.Width = IIf(VentanaMovil.Width * 15 + 30 * 15 > Pizarra.Width, VentanaMovil.Width * 15 + 30 * 15, Pizarra.Width)
    VentanaMovil.Refresh
    
    If ContenedorDatos.Visible Then
        With ContenedorDatos
            .Move -5, PicVSplit.Top + PicVSplit.Height - 2, _
             Contenedor.Width + 5, _
             Me.ScaleHeight - (PicVSplit.Top + PicVSplit.Height)
        End With
               
        With HerramientasAdjuntos
             .Move 75, 0, _
             (Contenedor.Width + 2) * 15, _
             .Height
        End With
               
        With HerramientasAccesos
             .Move 75, 0, _
             (Contenedor.Width + 2) * 15, _
             .Height
        End With
               
        With Adjuntos
             .Move 45, HerramientasAdjuntos.Top + HerramientasAdjuntos.Height - 15, _
             (Contenedor.Width + 2) * 15, _
             ContenedorDatos.Height * 15 - ContenedorDatos.TabHeight * 15 - HerramientasAdjuntos.Height
        End With
         
        With Accesos
             .Move 45, HerramientasAccesos.Top + HerramientasAccesos.Height - 15, _
             (Contenedor.Width + 2) * 15, _
             ContenedorDatos.Height * 15 - ContenedorDatos.TabHeight * 15 - HerramientasAccesos.Height
        End With
         
         
        With DatosAplicMtx
             .Move 60, 0, _
             (Contenedor.Width + 2) * 15, _
             ContenedorDatos.Height * 15 - ContenedorDatos.TabHeight * 15
        End With
        
        For i = 1 To 0 Step -1
            With PizarraDatos(i)
                .Move 60, 0, _
                (Contenedor.Width + 2) * 15, _
                ContenedorDatos.Height * 15 - ContenedorDatos.TabHeight * 15
            End With
        Next
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = RequiereSalir
End Sub


Private Sub HerramientasAccesos_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim a%
    Dim Fnd As Boolean
    Fnd = False
    Select Case Button.Key
    Case Is = "ADJUNTAR"
         Call AbrirAcceso
         If Accesos.ListItems.Count > 0 Then
            HerramientasAccesos.Buttons.Item(3).Enabled = True
            HerramientasAccesos.Buttons.Item(4).Enabled = True
            HerramientasAccesos.Buttons.Item(6).Enabled = True
         End If
    Case Is = "ELIMINAR"
         If Accesos.ListItems.Count > 0 Then
            For a = Accesos.ListItems.Count To 1 Step -1
                If Accesos.ListItems.Item(a).Selected Then
                   Accesos.ListItems.Remove a
                   Fnd = True
                End If
            Next
            If Fnd Then
               RequiereGrabar = True
               RequiereReproc = True
               MuestraProceso
            Else
               MsgBox "No existe ningun Acceso adjunto seleccionado para ser eliminado." + vbCrLf + "Debe seleccionar un Accesos adjunto para proceder con la eliminación", vbInformation, "Eliminar Archivo Adjunto Seleccionado."
            End If
         End If
         If Accesos.ListItems.Count = 0 Then
            HerramientasAccesos.Buttons.Item(3).Enabled = False
            HerramientasAccesos.Buttons.Item(4).Enabled = False
            HerramientasAccesos.Buttons.Item(6).Enabled = False
         End If
    Case Is = "ELIMINARTODOS"
         If Accesos.ListItems.Count > 0 Then
            Dim Rsp%
            Rsp = MsgBox("Desea realmente eliminar todos los Accesos adjuntos de este proceso?", vbYesNo + vbQuestion, "Eliminar Todos Los Archivo Adjuntos.")
            If Rsp = 6 Then
               For a = Accesos.ListItems.Count To 1 Step -1
                   Accesos.ListItems.Remove a
               Next
               RequiereGrabar = True
               RequiereReproc = True
               MuestraProceso
            End If
         Else
            MsgBox "No existe ningun Acceso adjunto para ser eliminado.", vbInformation, "Eliminar Todos Los Accesos Adjuntos."
         End If
         If Accesos.ListItems.Count = 0 Then
            HerramientasAccesos.Buttons.Item(3).Enabled = False
            HerramientasAccesos.Buttons.Item(4).Enabled = False
            HerramientasAccesos.Buttons.Item(6).Enabled = False
         End If
    Case Is = "ABRIR"
         If Accesos.ListItems.Count > 0 Then
            For a = Accesos.ListItems.Count To 1 Step -1
                If Accesos.ListItems.Item(a).Selected Then
                   Fnd = True
                   Exit For
                End If
            Next
            If Fnd Then
               MsgBox "Hay que editar el Acceso '" + Accesos.SelectedItem.Text + "'", vbInformation, "Abrir Accesos Adjunto Seleccionado."
            Else
               MsgBox "No existe ningun Acceso adjunto seleccionado para ser Abierto." + vbCrLf + "Debe seleccionar un Acceso adjunto para proceder con abrir el Archivo.", vbInformation, "Abrir Acceso Adjunto Seleccionado."
            End If
         Else
            MsgBox "Lo siento, No existe ningun Acceso adjunto para ser Abierto.", vbInformation, "Abrir Accesos Adjunto Seleccionado."
         End If
    End Select

End Sub

Private Sub HerramientasAccesos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     LabelIns.Visible = False
     ShapeMov.Visible = False
     Dim C%
     Dim Xp%
     Xp = X / Screen.TwipsPerPixelX '+ HerramientasAdjuntos.Left
     For C = 1 To HerramientasAccesos.Buttons.Count
         If (Xp >= HerramientasAccesos.Buttons.Item(C).Left / Screen.TwipsPerPixelX And Xp <= (HerramientasAccesos.Buttons.Item(C).Left / Screen.TwipsPerPixelX + HerramientasAccesos.Buttons.Item(C).Width / Screen.TwipsPerPixelX)) Then
             Ventana.StatusBar1.Panels.Item(1).Text = HerramientasAccesos.Buttons.Item(C).Description
             Exit For
         End If
     Next

End Sub

Private Sub HerramientasAdjuntos_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim a%
    Dim Fnd As Boolean
    Fnd = False
    Select Case Button.Key
    Case Is = "ADJUNTAR"
         Call AbrirAdjunto
         If Adjuntos.ListItems.Count > 0 Then
            HerramientasAdjuntos.Buttons.Item(3).Enabled = True
            HerramientasAdjuntos.Buttons.Item(4).Enabled = True
            HerramientasAdjuntos.Buttons.Item(6).Enabled = True
         End If
    Case Is = "ELIMINAR"
         If Adjuntos.ListItems.Count > 0 Then
            For a = Adjuntos.ListItems.Count To 1 Step -1
                If Adjuntos.ListItems.Item(a).Selected Then
                   Adjuntos.ListItems.Remove a
                   Fnd = True
                End If
            Next
            If Fnd Then
               RequiereGrabar = True
               RequiereReproc = True
               MuestraProceso
            Else
               MsgBox "No existe ningun archivo adjunto seleccionado para ser eliminado." + vbCrLf + "Debe seleccionar un archivo adjunto para proceder con la eliminación", vbInformation, "Eliminar Archivo Adjunto Seleccionado."
            End If
         End If
         If Adjuntos.ListItems.Count = 0 Then
            HerramientasAdjuntos.Buttons.Item(3).Enabled = False
            HerramientasAdjuntos.Buttons.Item(4).Enabled = False
            HerramientasAdjuntos.Buttons.Item(6).Enabled = False
         End If
    Case Is = "ELIMINARTODOS"
         If Adjuntos.ListItems.Count > 0 Then
            Dim Rsp%
            Rsp = MsgBox("Desea realmente eliminar todos los archivos adjuntos de este proceso?", vbYesNo + vbQuestion, "Eliminar Todos Los Archivo Adjuntos.")
            If Rsp = 6 Then
               For a = Adjuntos.ListItems.Count To 1 Step -1
                   Adjuntos.ListItems.Remove a
               Next
               RequiereGrabar = True
               RequiereReproc = True
               MuestraProceso
            End If
         Else
            MsgBox "No existe ningun archivo adjunto para ser eliminado.", vbInformation, "Eliminar Todos Los Archivos Adjuntos."
         End If
         If Adjuntos.ListItems.Count = 0 Then
            HerramientasAdjuntos.Buttons.Item(3).Enabled = False
            HerramientasAdjuntos.Buttons.Item(4).Enabled = False
            HerramientasAdjuntos.Buttons.Item(6).Enabled = False
         End If
    Case Is = "ABRIR"
         If Adjuntos.ListItems.Count > 0 Then
            For a = Adjuntos.ListItems.Count To 1 Step -1
                If Adjuntos.ListItems.Item(a).Selected Then
                   Fnd = True
                   Exit For
                End If
            Next
            If Fnd Then
               MsgBox "Hay que editar el adjunto '" + Adjuntos.SelectedItem.Text + "'", vbInformation, "Abrir Archivo Adjunto Seleccionado."
            Else
               MsgBox "No existe ningun archivo adjunto seleccionado para ser Abierto." + vbCrLf + "Debe seleccionar un archivo adjunto para proceder con abrir el Archivo.", vbInformation, "Abrir Archivo Adjunto Seleccionado."
            End If
         Else
            MsgBox "Lo siento, No existe ningun archivo adjunto para ser Abierto.", vbInformation, "Abrir Archivo Adjunto Seleccionado."
         End If
    End Select
End Sub

Private Sub HerramientasAdjuntos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     LabelIns.Visible = False
     ShapeMov.Visible = False
     Dim C%
     Dim Xp%
     Xp = X / Screen.TwipsPerPixelX '+ HerramientasAdjuntos.Left
     For C = 1 To HerramientasAdjuntos.Buttons.Count
         If (Xp >= HerramientasAdjuntos.Buttons.Item(C).Left / Screen.TwipsPerPixelX And Xp <= (HerramientasAdjuntos.Buttons.Item(C).Left / Screen.TwipsPerPixelX + HerramientasAdjuntos.Buttons.Item(C).Width / Screen.TwipsPerPixelX)) Then
             Ventana.StatusBar1.Panels.Item(1).Text = HerramientasAdjuntos.Buttons.Item(C).Description
             Exit For
         End If
     Next

End Sub

Private Sub Mnu_Ayuda_Click(Index As Integer)
  Select Case Index
         Case 2
            Call ShellAbout(hwnd, "Acerca de Modelador de Procesos# ....Adaptado a Microsoft Exchange WORKFLOW", _
                                              "Copyright © Rafael Planas 1999 - 2000" & vbCrLf & _
                                              "Banco Wiese Sudameris.", 0)
 End Select
End Sub

Private Sub Mnu_Print_Click(Index As Integer)
    If Index = 0 Then
       FrmPageSetup.Show vbModal
       If gprint = True Then
          frmDocPreview.DocPrintProc
       End If
    End If
End Sub

Private Sub Mnu_Salir_Click(Index As Integer)
    If RequiereSalir = 0 Then
       End
    End If

End Sub

Private Sub Notas_DblClick(Index As Integer)
     Note(Index).NoteCaption = EventoNota(Index).Titulo
     Note(Index).NoteText = EventoNota(Index).Definicion
     Note(Index).NoteInfo = EventoNota(Index).Fecha
     Note(Index).ShowNote

End Sub

Private Sub Notas_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Not Herramientas.Visible Then Exit Sub
     
     If ArrastrandoNota Then
        Notas(ObjEnMovimiento).Drag 0: Notas(ObjEnMovimiento).DragMode = 0: Notas(ObjEnMovimiento).Visible = True
        ShapeDel.Visible = False
        ShapeMov.Visible = False
        ArrastrandoNota = False
     Else
        Call RemarcaRelaciones(ObjEnMovimiento, 0)
        Call OcultaElementos
     End If

End Sub

Private Sub Notas_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    If Not Herramientas.Visible Then Exit Sub
    
    Btn = Button
    If (Button = 1 And (HerramientaSeleccionada = 0 Or HerramientaSeleccionada = 5)) Then
        ShapeDel.Move Notas(Index).Left - 2, Notas(Index).Top - 2, Notas(Index).Width + 4, Notas(Index).Height + 4
        ShapeDel.Visible = True
        
        Notas(Index).DragIcon = ImageListCur.ListImages.Item(7).Picture
        
        ObjEnMovimiento = Index
        Notas(Index).Drag 1
        Notas(Index).DragMode = 1
        DespX = X / Screen.TwipsPerPixelX
        DespY = Y / Screen.TwipsPerPixelY
        ArrastrandoNota = True
   End If
 
   If HerramientaSeleccionada = 3 Then
      If Button = 1 Then
         Call EliminaNota(Index)
      Else
         Notas(Index).DragIcon = ImageListCur.ListImages.Item(7).Picture
         DoEvents
      End If
   End If

End Sub

Private Sub Notas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Not Herramientas.Visible Then Exit Sub
   
   Select Case HerramientaSeleccionada
        Case 0
             Pizarra.MouseIcon = ImageListCur.ListImages.Item(7).Picture
             Pizarra.MousePointer = 99
        Case 2
             LabelIns.Move ObjEvento(Index).Left + (X / Screen.TwipsPerPixelX) + 32, ObjEvento(Index).Top + (Y / Screen.TwipsPerPixelY) - 27
             LabelIns.Caption = "Relacionar "
             LabelIns.Visible = True
             LabelIns.ZOrder 0
        Case 3
             LabelIns.Move Notas(Index).Left + (X / Screen.TwipsPerPixelX) + 32, Notas(Index).Top + (Y / Screen.TwipsPerPixelY) - 27
             LabelIns.Caption = "Eliminar Nota "
             LabelIns.Visible = True
             LabelIns.ZOrder 0
             
             ShapeDel.Move Notas(Index).Left - 2, Notas(Index).Top - 2, Notas(Index).Width + 4, Notas(Index).Height + 4
             ShapeDel.Visible = True
             PAnterior = Index
   End Select
End Sub

Private Sub ObjEvento_DblClick(Index As Integer)
    'Invoca pantalla de atributos de la tarea
    If Not Herramientas.Visible Then Exit Sub
    EventoTareaSeleccionado = Index
    
    If EventoTarea(Index).TareaTipo = 1 Then
       ArchivoACargar = "d:\proyectos\workflow\subp1.prc"
       Dim frmD As PizarraFlujos
       Set frmD = New PizarraFlujos
       Load frmD
       
    Else
    '    Call sSession
        'Call sCargarFormulario("FYI")
    '    Call sConexion
        
        GEventoTarea = EventoTarea(Index)
        GEventoPropiedades = EventoPropiedades(Index)
        
        Call sCargarFormulario(Index)
        
        EventoPropiedades(Index) = GEventoPropiedades
        EventoTarea(Index) = GEventoTarea
        EventoPropiedades(Index).Personalizada = True
        
        ObjEvento(Index).Cls
        ObjEvento(Index).FontSize = 8
        ObjEvento(Index).Print
        ObjEvento(Index).Font = "Mirror"
        ObjEvento(Index).FontSize = 7
        ObjEvento(Index).Print Space(10) + Trim(EventoTarea(Index).Definicion)
        Call MuestraProceso
    End If
End Sub

Private Sub ObjEvento_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
     If Not Herramientas.Visible Then Exit Sub
     If ArrastrandoRelacion Then
        If EventoTarea(Index).TareaTipo <> 7 Then
            
            ArrastrandoRelacion = False: DibujaLinea RCInicial
            If ObjEnMovimiento <> Index Then Pizarra.MouseIcon = ImageListCur.ListImages.Item(6).Picture
            If ObjEnMovimiento <> Index Then Call CreaRelacion(ObjEnMovimiento, Index)
            Call RemarcaRelaciones(ObjEnMovimiento, 0)
        Else
            Call RemarcaRelaciones(ObjEnMovimiento, 0)
            Call OcultaElementos
            MsgBox "No es posible efectuar este tipo de relación," + Chr(10) + "con una tarea de INICIO DE PROCESO", vbExclamation, "Relacionando Tareas"
        
        End If
     End If
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos
End Sub

Private Sub ObjEvento_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Select Case HerramientaSeleccionada
           Case 0, 1, 3, 4
               If Not ArrastrandoNota Then
                  ObjEvento(ObjEnMovimiento).Visible = False
               End If
           Case Else
               If Not ArrastrandoNota Then
                  If Btn = 2 Then ObjEvento(ObjEnMovimiento).Visible = False
                  LabelIns.Move ObjEvento(Index).Left + (X / Screen.TwipsPerPixelX) + 30, ObjEvento(Index).Top + (Y / Screen.TwipsPerPixelY) - 25
                  LabelIns.Caption = IIf(Btn = 1, "Relacionando ", "")
                  LabelIns.Visible = True
                  LabelIns.ZOrder 0
                  ShapeMov.Move ObjEvento(Index).Left + (X / Screen.TwipsPerPixelX) - DespX + 1, ObjEvento(Index).Top + (Y / Screen.TwipsPerPixelY) - DespY + 1, ObjEvento(0).Width, ObjEvento(0).Height
                  ShapeMov.Visible = True
                    
                  If Not ArrastrandoRelacion Then Exit Sub
                  If ObjEnMovimiento = Index Then
                     DibujaLinea RCInicial
                     ObjRelS1(0).Visible = False: ObjRelS2(0).Visible = False: ObjRelS3(0).Visible = False
                  Else
                     If EventoTarea(Index).TareaTipo <> 7 Then
                        DibujaLinea RCInicial
                        Call CalculaCoordenadasRelacion(ObjEvento(ObjEnMovimiento), ObjEvento(Index), ObjEnMovimiento, Index)
                        RCActual = r:  DibujaLinea RCActual, True:  RCInicial = r
                     End If
                  End If
               End If
    End Select
End Sub



Private Sub ObjEvento_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Not Herramientas.Visible Then Exit Sub
   
   Select Case HerramientaSeleccionada
        Case 0
             Pizarra.MouseIcon = ImageListCur.ListImages.Item(7).Picture
             Pizarra.MousePointer = 99
        Case 2
             LabelIns.Move ObjEvento(Index).Left + (X / Screen.TwipsPerPixelX) + 32, ObjEvento(Index).Top + (Y / Screen.TwipsPerPixelY) - 27
             LabelIns.Caption = "Relacionar "
             LabelIns.Visible = True
             LabelIns.ZOrder 0
        Case 3
             LabelIns.Move ObjEvento(Index).Left + (X / Screen.TwipsPerPixelX) + 32, ObjEvento(Index).Top + (Y / Screen.TwipsPerPixelY) - 27
             LabelIns.Caption = "Eliminar Tarea "
             LabelIns.Visible = True
             LabelIns.ZOrder 0
             
             ShapeDel.Move ObjEvento(Index).Left - 2, ObjEvento(Index).Top - 2, ObjEvento(Index).Width + 4, ObjEvento(Index).Height + 4
             ShapeDel.Visible = True
             Call RemarcaRelaciones(Index, 1)
             PAnterior = Index
        Case 1, 4
             If HerramientaSeleccionada = 1 Then
                ShapeMov.Visible = False
                Pizarra.MousePointer = 0
             End If
             LabelIns.Visible = False
             If HerramientaSeleccionada = 4 Then
                Pizarra.MousePointer = vbDefault
             End If
   End Select
End Sub

Private Sub ObjRelT1_DblClick(Index As Integer)
  If Herramientas.Visible Then
     
     Select Case HerramientaSeleccionada
            Case 0, 1, 2, 4
                Dim Txo%, Rxo%, Txd%, Rxd%, Nc%, i%, S%
                Txo = Relacion(Index).TareaOrigen
                Rxo = Relacion(Index).ConsecuentesOrigen
                Txd = Relacion(Index).TareaDestin
                Rxd = Relacion(Index).PrescedentesDestin
                Nc = EventoPropiedades(Txo).NoCondiciones
                S = EventoTarea(Txo).Consecuente(Rxo).RelacionInfinita
                If S Then
                   ObjRelT1(Index).Caption = Mid(ObjRelT1(Index).Caption, 1, InStr(1, ObjRelT1(Index).Caption, Chr(10)) - 1)
                End If
                For i = 1 To Nc
                   If Trim(ObjRelT1(Index).Caption) = Trim(EventoPropiedades(Txo).Condicion(i).Definicion) Then
                      Exit For
                   End If
                Next
                i = i + 1
                If i > Nc Then i = 1
                If EventoPropiedades(Txo).Condicion(i).CondicionActiva Then
                   ObjRelT1(Index).Caption = Trim(EventoPropiedades(Txo).Condicion(i).Definicion) + IIf(S, Chr(10) + "(Posible Flujo Circular)", "")
                   EventoTarea(Txo).Consecuente(Rxo).RelacionTipo = EventoPropiedades(Txo).Condicion(i).NoCondicion
                   EventoTarea(Txd).Prescedente(Rxd).RelacionTipo = EventoPropiedades(Txo).Condicion(i).NoCondicion
                ElseIf i <= Nc Then
                       i = i + 1
                       If i > Nc + 1 Then i = 1
                       ObjRelT1(Index).Caption = Trim(EventoPropiedades(Txo).Condicion(i).Definicion) + IIf(S, Chr(10) + "(Posible Flujo Circular)", "")
                       EventoTarea(Txo).Consecuente(Rxo).RelacionTipo = EventoPropiedades(Txo).Condicion(i).NoCondicion
                       EventoTarea(Txd).Prescedente(Rxd).RelacionTipo = EventoPropiedades(Txo).Condicion(i).NoCondicion
                End If
                RequiereGrabar = True
                RequiereReproc = True
                Call MuestraProceso
     End Select
  End If
End Sub

Private Sub ObjRelT1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos
End Sub

Private Sub ObjRelT1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Herramientas.Visible Then
   Select Case HerramientaSeleccionada
        Case 0, 1, 2, 4
             Dim Txo%, Rxo%, Txd%, Rxd%, Nc%, i%
                
             Txo = Relacion(Index).TareaOrigen
             Rxo = Relacion(Index).ConsecuentesOrigen
             Nc = EventoPropiedades(Txo).NoCondiciones
             If Nc > 1 Then
                Pizarra.MouseIcon = ImageListCur.ListImages.Item(10).Picture
                Pizarra.MousePointer = 99
                LabelIns.Move ObjRelT1(Index).Left + (X / Screen.TwipsPerPixelX) + 32, ObjRelT1(Index).Top + (Y / Screen.TwipsPerPixelY) - 27
                LabelIns.Caption = "Cambiar Condición "
                LabelIns.Visible = True
                LabelIns.ZOrder 0
             Else
                LabelIns.Visible = False
             End If
        Case 3
             LabelIns.Visible = False
   End Select
End If
End Sub

Private Sub PicVSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cVS.SplitterMouseDown Me.hwnd, X, Y
End Sub

Private Sub Pizarra_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    If X - DespX <= 0 Then X = DespX
    If Y - DespY <= 0 Then Y = DespY
    
    Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - DespX) + "," + Str(Y - DespY) + " "
    If ArrastrandoNota Then
       ShapeMov.Move X - DespX + 1, Y - DespY + 1, Notas(0).Width, Notas(0).Height
       Notas(0).Move X - DespX + 1, Y - DespY + 1, Notas(0).Width, Notas(0).Height
    Else
       ShapeMov.Move X - DespX + 1, Y - DespY + 1, ObjEvento(0).Width, ObjEvento(0).Height
       ObjEvento(0).Move X - DespX + 1, Y - DespY + 1, ObjEvento(0).Width, ObjEvento(0).Height
    End If
    ShapeMov.Visible = True
    
    If HerramientaSeleccionada = 2 Then
       LabelIns.Move X + 30, Y - 25
       LabelIns.Caption = IIf(Btn = 1, "Relacionando ", "Moviendo ")
       LabelIns.Visible = True
       LabelIns.ZOrder 0
    End If
    
    
    Pizarra.MouseIcon = ImageListCur.ListImages.Item(6).Picture
    Pizarra.MousePointer = 99
    
    If TeclaStat Then
       Call VerificaDimiensionPizarra(X, Y, State)
    End If
    If Not ArrastrandoNota And Not ArrastrandoRelacion Then
       Call MueveRelaciones(X, Y)
    End If
    
    If Not ArrastrandoRelacion Then Exit Sub
    DibujaLinea RCInicial
    ObjEvento(0).Move X - (ObjEvento(0).Width / 2), Y - (ObjEvento(0).Height / 2), ObjEvento(0).Width, ObjEvento(0).Height
    Call CalculaCoordenadasRelacion(Source, ObjEvento(0), ObjEnMovimiento, 0)
    RCActual = r
    DibujaLinea RCActual
    RCInicial = r

End Sub


Private Sub Pizarra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Herramientas.Visible Then
       Dim Xv%, Yv%, NroSegmentoRelacion%, NCT%
       Dim PO%, pd%, C%
       NCT = 0
       Xv = Int(X): Yv = Int(Y) 'en caso de no ser los pixeles
       
       NroSegmentoRelacion = DetectaRelacion(Xv, Yv, NCT)
       If NroSegmentoRelacion > -1 And Button = 1 Then
          Select Case HerramientaSeleccionada
                 Case 0, 2
                     If NCT > 4 Then
                        ObjRelP2(NroSegmentoRelacion).Visible = Not (ObjRelP2(NroSegmentoRelacion).Visible)
                        ObjRelP3(NroSegmentoRelacion).Visible = Not (ObjRelP3(NroSegmentoRelacion).Visible)
                        If ObjRelP2(NroSegmentoRelacion).Visible Then
                           For C = 1 To MaxRelaObj - 1
                               If C <> NroSegmentoRelacion And ObjRelP2(C).Visible Then
                                  ObjRelP2(C).Visible = False: ObjRelP3(C).Visible = False
                               End If
                           Next C
                           ObjRelP2(NroSegmentoRelacion).ZOrder 0
                           ObjRelP3(NroSegmentoRelacion).ZOrder 0
                           ObjRelP1(NroSegmentoRelacion).ZOrder 0
                           ObjRelP4(NroSegmentoRelacion).ZOrder 0
                        End If
                     End If
                     ArrastrandoLinea = True
                     NuevaLocRelacion = NroSegmentoRelacion
                     NuevaLocObjeto = NCT
                     Exit Sub
                Case 3
                     Call EliminaRelacion(NroSegmentoRelacion)
                     RequiereGrabar = True
                     RequiereReproc = True
                     Call MuestraProceso
                Case 4
                     Dim Nt%, Rt%, Ro%, Rd%
                     PO = Relacion(NroSegmentoRelacion).TareaOrigen
                     Ro = Relacion(NroSegmentoRelacion).ConsecuentesOrigen
                     pd = Relacion(NroSegmentoRelacion).TareaDestin
                     Rd = Relacion(NroSegmentoRelacion).ConsecuentesOrigen
                     Rt = EventoTarea(PO).Consecuente(Ro).RelacionTipo
'                     EventoPropiedades(PO).Condicion(1).Tipo
                     
                     Call EliminaRelacion(NroSegmentoRelacion)
                     Call CreaTarea(X, Y, TareaTipo, Nt)
                     Call CreaRelacion(PO, Nt, Rt)
                     Call CreaRelacion(Nt, pd)
                     EventoTarea(Nt).Prescedente(1).RelacionTipo = Rt
          End Select
       End If
       
       Select Case HerramientaSeleccionada
       Case 0
       Case 2
       Case 1, 5
            ' calcelar arrastre si es presionado el boton derecho
            If Button = vbRightButton Then
                If Arrastrando Then
                    Arrastrando = False
                    DibujaCaja PosInicialX, PosInicialY   ' borrar la caja antigua
                End If
                
            End If
            
            If Not Arrastrando Then
               Arrastrando = True
               PosInicialX = X: PosActualX = PosInicialX
               PosInicialY = Y: PosActualY = PosInicialY
               DibujaCaja PosActualX, PosActualY
            End If
       End Select
    End If
End Sub

Private Sub Pizarra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Herramientas.Visible Then
       Dim NCT%, RN%, p%
       Select Case HerramientaSeleccionada
                              
              Case 1
                   LabelIns.Move X + 30, Y - 25
                   Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - (ObjEvento(0).Width / 2)) + "," + Str(Y - (ObjEvento(0).Height / 2)) + " "
                   ShapeMov.Move X - (ObjEvento(0).Width / 2), Y - (ObjEvento(0).Height / 2), ObjEvento(0).Width, ObjEvento(0).Height
                   ShapeMov.Visible = True
                   
                   LabelIns.Caption = "Nueva Tarea (" + DefTareaTipo(TareaTipo + 1) + ") "
                   LabelIns.Visible = True
                   LabelIns.ZOrder 0
                   
                   Pizarra.MouseIcon = ImageListCur.ListImages.Item(8).Picture
                   Pizarra.MousePointer = 99
                   If Not Arrastrando Then Exit Sub
                    
                   DibujaCaja PosInicialX, PosInicialY   ' Borrar caja antigua
                   PosActualX = X: PosActualY = Y
                   DibujaCaja PosActualX, PosActualY    ' Dibujar nueva Caja
                   PosInicialX = X: PosInicialY = Y
              
              Case 0, 2
                    LabelIns.Visible = False
                    
                    Dim DR%
                    DR = DetectaRelacion(X, Y, NCT)
                    If HerramientaSeleccionada = 2 Then
                       Pizarra.MouseIcon = ImageListCur.ListImages.Item(5).Picture
                    Else
                       Pizarra.MouseIcon = ImageListCur.ListImages.Item(2).Picture
                    End If
                    Pizarra.MousePointer = IIf(DR > -1, 99, IIf(HerramientaSeleccionada = 0, 0, 99))
                    LabelIns.Move X + 30, Y - 25
                    If RAnterior > 0 Then
                       If Relacion(RAnterior).RelacionActiva Then
                          ObjRelS1(RAnterior).BorderColor = &H80000008
                          ObjRelS2(RAnterior).BorderColor = &H80000008
                          ObjRelS3(RAnterior).BorderColor = &H80000008
                       End If
                    End If
                    
                    If DR < 0 Then
                       LabelIns.Caption = "Relacionar "
                       
                    Else
                       ObjRelS1(DR).BorderColor = &H8000000B
                       ObjRelS2(DR).BorderColor = &H8000000B
                       ObjRelS3(DR).BorderColor = &H8000000B
                       ObjRelS1(DR).ZOrder 0: ObjRelS2(DR).ZOrder 0:  ObjRelS3(DR).ZOrder 0
                       RAnterior = DR
                    
                       Select Case NCT
                              Case 1, 2
                                   If ObjRelP2(DR).Visible Then
                                      LabelIns.Caption = "Desplazar Linea Relacion "
                                      Pizarra.MousePointer = 5
                                   Else
                                      LabelIns.Caption = "Relacionar "
                                   End If
                              Case 3, 4
                                   If ObjRelP2(DR).Visible Then
                                      LabelIns.Caption = "Desplazar Linea Relacion "
                                      If ObjRelP2(DR).Top = ObjRelP1(DR).Top Then
                                         Pizarra.MousePointer = 7
                                      Else
                                         Pizarra.MousePointer = 9
                                      End If
                                   Else
                                      LabelIns.Caption = "Relacionar "
                                   End If
                              Case 5, 6, 7
                                   If ObjRelP2(DR).Visible Then
                                      LabelIns.Caption = "Deselecionar Linea Relacion "
                                   Else
                                      LabelIns.Caption = "Selecionar Linea Relacion "
                                   End If
                        End Select
                        LabelIns.Visible = True
                        LabelIns.ZOrder 0
                   End If
                   
                   If ArrastrandoLinea = True Then
                       Dim Xv%, Yv%
                       Dim Pad%, NumRP%
                       Dim Hij%, NumRH%
                       
                       Dim Pi As ElementoCordenadas
                       Dim PF As ElementoCordenadas
                       Dim VH As String * 2
                       Dim lH As Boolean
                       Dim LV As Boolean
                       
                       Hij% = Relacion(NuevaLocRelacion).TareaOrigen
                       NumRH = Relacion(NuevaLocRelacion).ConsecuentesOrigen
                       Pad% = Relacion(NuevaLocRelacion).TareaDestin
                       NumRP = Relacion(NuevaLocRelacion).PrescedentesDestin
                        
                       With Pi
                           .XCordLeft = Int(ObjEvento(Hij).Left + 0.5)
                           .XCordMed = Int(ObjEvento(Hij).Left + (ObjEvento(Hij).Width / 2) + 0.5)
                           .XCordRigth = Int(ObjEvento(Hij).Left + ObjEvento(Hij).Width + 0.5)
                           .YCordTop = Int(ObjEvento(Hij).Top + 0.5)
                           .YCordMed = Int(ObjEvento(Hij).Top + (ObjEvento(Hij).Height / 2) + 0.5)
                           .YCordBottom = Int(ObjEvento(Hij).Top + ObjEvento(Hij).Height + 0.5)
                       End With
                            
                       With PF
                           .XCordLeft = Int(ObjEvento(Pad).Left + 0.5)
                           .XCordMed = Int(ObjEvento(Pad).Left + (ObjEvento(Pad).Width / 2) + 0.5)
                           .XCordRigth = Int(ObjEvento(Pad).Left + ObjEvento(Pad).Width + 0.5)
                           .YCordTop = Int(ObjEvento(Pad).Top + 0.5)
                           .YCordMed = Int(ObjEvento(Pad).Top + (ObjEvento(Pad).Height / 2) + 0.5)
                           .YCordBottom = Int(ObjEvento(Pad).Top + ObjEvento(Pad).Height + 0.5)
                       End With
                       
                       EventoTarea(Pad).Prescedente(NumRP).PosicionXYPersonal = True
                       EventoTarea(Hij).Consecuente(NumRH).PosicionXYPersonal = True
                       
                       VH = EventoTarea(Hij).Consecuente(NumRH).VH
                       
                       Xv = Int(X): Yv = Int(Y)
                       r = EventoTarea(Pad).Prescedente(NumRP)
                       
                       Select Case VH
                            Case "UL", "UN", "UR", "DL", "DN", "DR"
                                 Select Case NuevaLocObjeto
                                        Case 1
                                            LV = IIf((r.Y1 > r.Y4), (Yv <= r.Y1) And (Yv >= r.Y4), (Yv >= r.Y1) And (Yv <= r.Y4))
                                            If LV Then r.Y2 = Yv: r.Y3 = r.Y2
                                            'If Yv < PI.YCordTop Then R.Y2 = PI.YCordTop: R.Y3 = R.Y2
                                            'If Yv > PI.YCordBottom Then R.Y2 = PI.YCordBottom: R.Y3 = R.Y2
                                            
                                            lH = Xv >= Pi.XCordLeft And Xv <= Pi.XCordRigth
                                            If lH Then r.X2 = Xv
                                            If Xv < Pi.XCordLeft Then r.X2 = Pi.XCordLeft
                                            If Xv > Pi.XCordRigth Then r.X2 = Pi.XCordRigth
                                            r.X1 = r.X2
                                        Case 2
                                            LV = IIf((r.Y1 > r.Y4), (Yv <= r.Y1) And (Yv >= r.Y4), (Yv >= r.Y1) And (Yv <= r.Y4))
                                            If LV Then r.Y3 = Yv: r.Y2 = r.Y3
                                            'If Yv < PF.YCordTop Then R.Y3 = PF.YCordTop
                                            'If Yv > PF.YCordBottom Then R.Y3 = PF.YCordBottom
                                            
                                            lH = Xv >= PF.XCordLeft And Xv <= PF.XCordRigth
                                            If lH Then r.X3 = Xv
                                            If Xv < PF.XCordLeft Then r.X3 = PF.XCordLeft
                                            If Xv > PF.XCordRigth Then r.X3 = PF.XCordRigth
                                            r.X4 = r.X3
                                        Case 3
                                            If Int(r.Y1 + 0.5) = Pi.YCordTop And Mid(VH, 1, 1) = "U" Or Int(r.Y1 + 0.5) = Pi.YCordBottom And Mid(VH, 1, 1) = "D" Then
                                               lH = Xv >= Pi.XCordLeft And Xv <= Pi.XCordRigth
                                               If lH Then r.X1 = Xv
                                               If Xv < Pi.XCordLeft Then r.X1 = Pi.XCordLeft
                                               If Xv > Pi.XCordRigth Then r.X1 = Pi.XCordRigth
                                               If r.Y1 = r.Y2 Then r.Y2 = r.Y3
                                               r.X2 = r.X1
                                            End If
                                            If r.X1 = Pi.XCordLeft And Mid(VH, 2, 1) = "L" Or r.X1 = Pi.XCordRigth And Mid(VH, 2, 1) = "R" Then
                                               LV = Yv >= Pi.YCordTop And Yv <= Pi.YCordBottom
                                               If LV Then r.Y1 = Yv
                                               If Yv < Pi.YCordTop Then r.Y1 = Pi.YCordTop
                                               If Yv > Pi.YCordBottom Then r.Y1 = Pi.YCordBottom
                                               r.Y2 = r.Y1: r.X2 = r.X4
                                            End If
                                        Case 4
                                            If Int(r.Y4 + 0.5) = PF.YCordTop And Mid(VH, 1, 1) = "D" Or Int(r.Y4 + 0.5) = PF.YCordBottom And Mid(VH, 1, 1) = "U" Then
                                               lH = Xv >= PF.XCordLeft And Xv <= PF.XCordRigth
                                               If lH Then r.X4 = Xv
                                               If Xv < PF.XCordLeft Then r.X4 = PF.XCordLeft
                                               If Xv > PF.XCordRigth Then r.X4 = PF.XCordRigth
                                               If r.Y4 = r.Y3 Then r.Y3 = r.Y2
                                               r.X3 = r.X4
                                            End If
                                            If r.X4 = PF.XCordLeft And Mid(VH, 2, 1) = "R" Or r.X4 = PF.XCordRigth And Mid(VH, 2, 1) = "L" Then
                                               LV = Yv >= PF.YCordTop And Yv <= PF.YCordBottom
                                               If LV Then r.Y4 = Yv ': R.Y3 = R.Y4
                                               If Yv < PF.YCordTop Then r.Y4 = PF.YCordTop
                                               If Yv > PF.YCordBottom Then r.Y4 = PF.YCordBottom
                                               r.Y3 = r.Y4: r.X3 = r.X1
                                            End If
                                 End Select
                            Case "NL", "NR"
                                 Select Case NuevaLocObjeto
                                        Case 1
                                            lH = IIf((r.X1 > r.X4), (Xv <= r.X1) And (Xv >= r.X4), (Xv >= r.X1) And (Xv <= r.X4))
                                            If lH Then r.X2 = Xv: r.X3 = r.X2
                                            LV = Yv >= Pi.YCordTop And Yv <= Pi.YCordBottom
                                            If LV Then r.Y2 = Yv
                                            If Yv < Pi.YCordTop Then r.Y2 = Pi.YCordTop
                                            If Yv > Pi.YCordBottom Then r.Y2 = Pi.YCordBottom
                                            r.Y1 = r.Y2
                                        Case 2
                                            lH = IIf((r.X1 > r.X4), (Xv <= r.X1) And (Xv >= r.X4), (Xv >= r.X1) And (Xv <= r.X4))
                                            If lH Then r.X3 = Xv: r.X2 = r.X3
                                            LV = Yv >= PF.YCordTop And Yv <= PF.YCordBottom
                                            If LV Then r.Y3 = Yv
                                            If Yv < PF.YCordTop Then r.Y3 = PF.YCordTop
                                            If Yv > PF.YCordBottom Then r.Y3 = PF.YCordBottom
                                            r.Y4 = r.Y3
                                        Case 3
                                            LV = Yv >= Pi.YCordTop And Yv <= Pi.YCordBottom
                                            If LV Then r.Y2 = Yv
                                            If Yv < Pi.YCordTop Then r.Y2 = Pi.YCordTop
                                            If Yv > Pi.YCordBottom Then r.Y2 = Pi.YCordBottom
                                            r.Y1 = r.Y2
                                        Case 4
                                            LV = Yv >= PF.YCordTop And Yv <= PF.YCordBottom
                                            If LV Then r.Y3 = Yv
                                            If Yv < Pi.YCordTop Then r.Y3 = Pi.YCordTop
                                            If Yv > Pi.YCordBottom Then r.Y3 = Pi.YCordBottom
                                            r.Y4 = r.Y3
                                 End Select
                       End Select
                       EventoTarea(Pad).Prescedente(NumRP).X1 = r.X1
                       EventoTarea(Pad).Prescedente(NumRP).X2 = r.X2
                       EventoTarea(Pad).Prescedente(NumRP).X3 = r.X3
                       EventoTarea(Pad).Prescedente(NumRP).X4 = r.X4
                       EventoTarea(Pad).Prescedente(NumRP).Y1 = r.Y1
                       EventoTarea(Pad).Prescedente(NumRP).Y2 = r.Y2
                       EventoTarea(Pad).Prescedente(NumRP).Y3 = r.Y3
                       EventoTarea(Pad).Prescedente(NumRP).Y4 = r.Y4
                       
                       EventoTarea(Hij).Consecuente(NumRH).X1 = r.X1
                       EventoTarea(Hij).Consecuente(NumRH).X2 = r.X2
                       EventoTarea(Hij).Consecuente(NumRH).X3 = r.X3
                       EventoTarea(Hij).Consecuente(NumRH).X4 = r.X4
                       EventoTarea(Hij).Consecuente(NumRH).Y1 = r.Y1
                       EventoTarea(Hij).Consecuente(NumRH).Y2 = r.Y2
                       EventoTarea(Hij).Consecuente(NumRH).Y3 = r.Y3
                       EventoTarea(Hij).Consecuente(NumRH).Y4 = r.Y4
                       Call DibujaRelacion(NuevaLocRelacion, Hij, NumRH, Pad, NumRP)
                    End If
              
              
              Case 3, 4
                   If HerramientaSeleccionada = 4 Then
                      LabelIns.Caption = "Insertar (" + DefTareaTipo(TareaTipo + 1) + ") "
                      Pizarra.MouseIcon = ImageListCur.ListImages.Item(2).Picture
                   Else
                      LabelIns.Caption = "Eliminar "
                      Pizarra.MouseIcon = ImageListCur.ListImages.Item(1).Picture
                   End If
                   LabelIns.Visible = True
                   LabelIns.ZOrder 0
                   ShapeDel.Visible = False
                   RN = DetectaRelacion(X, Y, NCT)
                   If RAnterior > 0 Then
                      If Relacion(RAnterior).RelacionActiva Then
                         ObjRelS1(RAnterior).BorderColor = &H80000008
                         ObjRelS2(RAnterior).BorderColor = &H80000008
                         ObjRelS3(RAnterior).BorderColor = &H80000008
                      End If
                   End If
                   Pizarra.MousePointer = 99 ' IIf(Rn > -1 And HerramientaSeleccionada = 4, vbCrosshair, IIf(HerramientaSeleccionada = 4, 6, 99))
                   LabelIns.Move X + 30, Y - 25
                   
                   If RN > -1 Then
                      If HerramientaSeleccionada = 3 Then
                         LabelIns.Caption = "Eliminar Relacion "
                      End If
                      ObjRelS1(RN).BorderColor = &H8000000B
                      ObjRelS2(RN).BorderColor = &H8000000B
                      ObjRelS3(RN).BorderColor = &H8000000B
                      ObjRelS1(RN).ZOrder 0: ObjRelS2(RN).ZOrder 0:  ObjRelS3(RN).ZOrder 0
                      RAnterior = RN
                   End If
                   p = PAnterior
                   PAnterior = 0
                   Call RemarcaRelaciones(p, 0)
              Case 5
                   LabelIns.Move X + 30, Y - 25
                   Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - (ObjEvento(0).Width / 2)) + "," + Str(Y - (ObjEvento(0).Height / 2)) + " "
                   ShapeMov.Move X - (Notas(0).Width / 2), Y - (Notas(0).Height / 2), Notas(0).Width, Notas(0).Height
                   ShapeMov.Visible = True
                   
                   LabelIns.Caption = "Anotación"
                   LabelIns.Visible = True
                   LabelIns.ZOrder 0
                   
                   Pizarra.MouseIcon = ImageListCur.ListImages.Item(2).Picture
                   Pizarra.MousePointer = 99
                   If Not Arrastrando Then Exit Sub
                    
                   DibujaCaja PosInicialX, PosInicialY   ' Borrar caja antigua
                   PosActualX = X: PosActualY = Y
                   DibujaCaja PosActualX, PosActualY    ' Dibujar nueva Caja
                   PosInicialX = X: PosInicialY = Y
       End Select
    End If

End Sub

Private Sub Pizarra_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Herramientas.Visible Then
        
        Dim Xv%, Yv%, NroSegmentoRelacion%, NCT%
        NCT = 0
        Xv = Int(X): Yv = Int(Y) 'en caso de no ser los pixeles
        NroSegmentoRelacion = DetectaRelacion(Xv, Yv, NCT)
        If NroSegmentoRelacion > -1 And Button = 1 Then ' En caso de querer multiples puntos
           If ObjRelS1(NroSegmentoRelacion).BorderColor = &H8000000B Then
              ObjRelS1(NroSegmentoRelacion).BorderColor = &H80000008
              ObjRelS2(NroSegmentoRelacion).BorderColor = &H80000008
              ObjRelS3(NroSegmentoRelacion).BorderColor = &H80000008
           End If
        End If
        
        If ArrastrandoLinea = True Then
           ArrastrandoLinea = False
        End If
        
        If Not Arrastrando Then Exit Sub
        Arrastrando = False
        
        
        DibujaCaja PosInicialX, PosInicialY   ' Borrar caja antigua
        PosActualX = X: PosActualY = Y
        
        Dim Nt%
        Select Case HerramientaSeleccionada
             Case 1
                    Pizarra.MouseIcon = ImageListCur.ListImages.Item(9).Picture
                    Call CreaTarea(PosActualX, PosActualY, TareaTipo, Nt)
                    ShapeMov.Visible = False
             Case 5
                    Pizarra.MouseIcon = ImageListCur.ListImages.Item(9).Picture
                    Call CreaNotas(PosActualX, PosActualY, TareaTipo, Nt)
                    ShapeMov.Visible = False
        End Select
    End If
End Sub



Private Sub PizarraDatos_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos
     ShapeMov.Visible = False
End Sub

Private Sub PizarraDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub PizarraDatos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ApuntadorTexto.X = X
    ApuntadorTexto.Y = Y
    Dim ObjSrch As String
    Dim C%
    If PizarraDatos(Index).SelLength = 0 And ApuntadorTexto.TareaEncontrada > 0 Then
       If EventoTarea(ApuntadorTexto.TareaEncontrada).Señalado Then
          For C = ApuntadorTexto.TareaEncontrada To 1 Step -1
              If EventoTarea(C).ProcesoActivo Then
                 If UCase(Trim(EventoTarea(C).Definicion)) = UCase(Trim(ApuntadorTexto.Palabra)) Then
                    Call SeñalaTarea((C))
                 End If
              End If
          Next
       End If
    End If
End Sub

Private Sub PizarraDatos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSrch As String
    Dim C%
    If PizarraDatos(Index).SelLength > 0 Then
       ObjSrch = Trim(PizarraDatos(Index).SelText) '  RichWordOver(PizarraDatos(Index), (ApuntadorTexto.x), (ApuntadorTexto.y))
       For C = 1 To MaxProcObj
           If EventoTarea(C).ProcesoActivo Then
               If UCase(Trim(EventoTarea(C).Definicion)) = UCase(Trim(ObjSrch)) Then
                  ApuntadorTexto.TareaEncontrada = C
                  ApuntadorTexto.Palabra = ObjSrch
                  Call SeñalaTarea(C, True, vbYellow)
               End If
           End If
       Next
    End If

End Sub

Private Sub prn1_Click(Index As Integer)
    'http://www.angelfire.com/wv/Visualbasic05
    Select Case Index
           Case 1
                PizarraDatos(0).SetFocus
                frmDocPreview.DocPrintProc
           Case 2
                PizarraDatos(1).SetFocus
                frmDocPreview.DocPrintProc
    End Select
    

End Sub

Private Sub Prv1_Click(Index As Integer)
    'http://www.angelfire.com/wv/Visualbasic05
    Select Case Index
           Case 0
                PizarraDatos(0).SetFocus
                frmDocPreview.Show vbModal
                If gprint = True Then
                     frmDocPreview.DocPrintProc
                End If
           Case 1
                PizarraDatos(1).SetFocus
                frmDocPreview.Show vbModal
                If gprint = True Then
                     frmDocPreview.DocPrintProc
                End If
'                frmDocPreview.DocPrintProc
    End Select

End Sub

Private Sub TamañoPizarra_Click(Index As Integer)
    Pizarra.SetFocus
    Select Case Index
        Case 0
            Pizarra.Height = IIf(Pizarra.Height < ((VentanaMovil.Height * 15)) * 2, Pizarra.Height * 1.2, Pizarra.Height)
            Pizarra.Width = IIf(Pizarra.Width < ((VentanaMovil.Width * 15) * 2), Pizarra.Width * 1.2, Pizarra.Width)
            Plantilla.XMax = Pizarra.Width
            Plantilla.YMax = Pizarra.Height
        Case 1
            Pizarra.Height = IIf((Pizarra.Height / Screen.TwipsPerPixelY) / 1.2 > Plantilla.YObjMax, Pizarra.Height / 1.2, Pizarra.Height)
            Pizarra.Width = IIf((Pizarra.Width / Screen.TwipsPerPixelX) / 1.2 > Plantilla.XObjMax, Pizarra.Width / 1.2, Pizarra.Width)
            Plantilla.XMax = Pizarra.Width
            Plantilla.YMax = Pizarra.Height
    
    End Select
    VentanaMovil.Refresh
End Sub


Private Sub ObjEvento_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Herramientas.Visible Then Exit Sub
    
    Btn = Button
    If (Button = 1 And (HerramientaSeleccionada = 0 Or HerramientaSeleccionada = 1)) Or _
       (Button = 2 And (HerramientaSeleccionada = 2 Or HerramientaSeleccionada = 3 Or HerramientaSeleccionada = 4)) Then
        ShapeDel.Move ObjEvento(Index).Left - 2, ObjEvento(Index).Top - 2, ObjEvento(Index).Width + 4, ObjEvento(Index).Height + 4
        ShapeDel.Visible = True
        
        ObjEvento(Index).DragIcon = ImageListCur.ListImages.Item(7).Picture
        
        ObjEnMovimiento = Index
        ObjEvento(Index).Drag 1
        ObjEvento(Index).DragMode = 1
        DespX = X / Screen.TwipsPerPixelX
        DespY = Y / Screen.TwipsPerPixelY
        Call RemarcaRelaciones(Index, 1)
     
                
        
   End If
   
   If Button = 1 And HerramientaSeleccionada = 2 Then
      If EventoTarea(Index).TareaTipo = 0 Then Exit Sub
      ObjEnMovimiento = Index
      ObjEvento(Index).Drag 1
      ObjEvento(Index).DragMode = 1
      DespX = X / Screen.TwipsPerPixelX
      DespY = Y / Screen.TwipsPerPixelY
      ObjEvento(Index).DragIcon = ImageListCur.ListImages.Item(4).Picture
      
      If Button = 1 Then
         If ArrastrandoRelacion Then
            ArrastrandoRelacion = False
            DibujaLinea RCInicial ' Borrar caja antigua
         End If
      End If
            
      If Not ArrastrandoRelacion Then
         ArrastrandoRelacion = True
         ObjEvento(0).Move ObjEvento(Index).Left, ObjEvento(Index).Top, ObjEvento(0).Width, ObjEvento(0).Height
         Call CalculaCoordenadasRelacion(ObjEvento(Index), ObjEvento(0), (Index), 0)
         RCInicial = r: RCActual = RCInicial
         DibujaLinea RCActual
      End If
   End If
   If HerramientaSeleccionada = 3 Then
      If Button = 1 Then
         Call EliminaTarea(Index)
      Else
       ObjEvento(Index).DragIcon = ImageListCur.ListImages.Item(7).Picture
       DoEvents
      
      End If
   End If
End Sub

Private Sub Pizarra_DragDrop(Source As Control, X As Single, Y As Single)
  Dim NRC%
  
  If HerramientaSeleccionada <> 2 Or (HerramientaSeleccionada = 2 And Not ArrastrandoRelacion) Then
        ShapeDel.Visible = False
        ShapeMov.Visible = False
        
        If X - DespX < 0 Then
           X = DespX
        End If
        If Y - DespY < 0 Then
           Y = DespY
        End If
        Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - DespX) + "," + Str(Y - DespY) + " "
           
        If ArrastrandoNota Then
            Notas(ObjEnMovimiento).Drag 0
            Notas(ObjEnMovimiento).DragMode = 0
            
            If EventoNota(ObjEnMovimiento).posx <> (X - DespX) + (Notas(0).Width / 2) Or EventoNota(ObjEnMovimiento).posy <> (Y - DespY) + (Notas(0).Height / 2) Then
                Notas(ObjEnMovimiento).Move X - DespX, Y - DespY
                With DefNotas(ObjEnMovimiento)
                    .ZOrder 0
                    .Move Notas(ObjEnMovimiento).Left - (.Width / 2) + DespX, Notas(ObjEnMovimiento).Top + Notas(ObjEnMovimiento).Height + 2, .Width, .Height
                End With
                
                ShapeStat(ObjEnMovimiento).Move Notas(ObjEnMovimiento).Left - 2, Notas(ObjEnMovimiento).Top - 2, Notas(ObjEnMovimiento).Width + 4, Notas(ObjEnMovimiento).Height + 4
                
                RequiereGrabar = True
                EventoNota(ObjEnMovimiento).posx = (X - DespX) + (Notas(0).Width / 2)
                EventoNota(ObjEnMovimiento).posy = (Y - DespY) + (Notas(0).Height / 2)
                Plantilla.XObjMax = 0
                Plantilla.YObjMax = 0
                
                
                For NRC = 1 To MaxNotaObj
                    Plantilla.XObjMax = IIf(EventoNota(NRC).posx + Notas(0).Width > Plantilla.XObjMax, EventoNota(NRC).posx + Notas(0).Width, Plantilla.XObjMax)
                    Plantilla.YObjMax = IIf(EventoNota(NRC).posy + Notas(0).Height > Plantilla.YObjMax, EventoNota(NRC).posy + Notas(0).Height, Plantilla.YObjMax)
                Next NRC
            End If
            Notas(ObjEnMovimiento).Visible = True
            ArrastrandoNota = False
        Else
            ObjEvento(ObjEnMovimiento).Drag 0
            ObjEvento(ObjEnMovimiento).DragMode = 0
            
            If EventoTarea(ObjEnMovimiento).posx <> (X - DespX) + (ObjEvento(0).Width / 2) Or EventoTarea(ObjEnMovimiento).posy <> (Y - DespY) + (ObjEvento(0).Height / 2) Then
                Call MueveRelaciones(X, Y)
            Else
                Call RemarcaRelaciones(ObjEnMovimiento, 0)
            End If
            ObjEvento(ObjEnMovimiento).Visible = True
        End If
  Else
        ShapeDel.Visible = False
        ShapeMov.Visible = False
        ObjEvento(ObjEnMovimiento).Drag 0
        ObjEvento(ObjEnMovimiento).DragMode = 0
        ObjEvento(ObjEnMovimiento).Visible = True
        If Not ArrastrandoRelacion Then Exit Sub
        ArrastrandoRelacion = False
        DibujaLinea RCInicial
        ObjRelS1(0).Visible = False: ObjRelS2(0).Visible = False: ObjRelS3(0).Visible = False
        ShapeMov.Visible = False
  End If
  
End Sub

' ********************************************
'  Crea nuevo control con cordenadas
' PosActualX, PosActualY.
' ********************************************
Private Sub CreaTarea(X As Single, Y As Single, Tipo As Integer, ByRef Nt As Integer)
    ' Cargar Nuevo Control
    
        
    Dim Mx%, Proc%
    
    If MaxProcObj > 0 Then
       For Mx = 1 To MaxProcObj
           If Not EventoTarea(Mx).ProcesoActivo Then Exit For
       Next Mx
    Else
       Mx = 1
    End If
    If Mx > MaxProcObj Then
       MaxProcObj = MaxProcObj + 1
       Select Case MaxProcObj
              Case 30
                   MsgBox "Se recomienda no emplear mas de treinta tareas dentro de un proceso;" + Chr(10) + Chr(13) + _
                          "Precaución!, El hacerlo podria volver muy confuso el diseño del mismo.", vbInformation, "Creación de Procesos"
              Case 40
                   MsgBox "Ud. esta excediendo el limite de tareas permitido por los estandares" + Chr(10) + Chr(13) + _
                          "Internacionales de modelamiento de procesos.", vbExclamation, "Creación de Procesos"
              Case Is > 50
                   MsgBox "Ud. ha roto con la norma de modelamiento de Procesos." + Chr(10) + Chr(13) + _
                          "Se sugiere revizar su diseño y reevaluarlo.", vbExclamation, "Creación de Procesos"
                   MaxProcObj = 50
                   Exit Sub
       End Select
       Proc = MaxProcObj
       ReDim Preserve EventoTarea(MaxProcObj)
       ReDim Preserve EventoPropiedades(MaxProcObj)
    Else
       Proc = Mx
    End If
    
        
    
    Herramientas.Buttons.Item(9).Enabled = IIf(Proc > 1, True, False)
    Herramientas.Buttons.Item(10).Enabled = IIf(Proc > 1, True, False)

    NoTareasTipo(TareaTipo) = NoTareasTipo(TareaTipo) + 1
    
    ' Guarda coordenadas del control creado
    EventoTarea(Proc).IdProceso = "PXXXXX"
    EventoTarea(Proc).IdTarea = "T" + Format(MaxProcObj, "00000")
    EventoTarea(Proc).NoTarea = Proc
    EventoTarea(Proc).TareaTipo = TareaTipo
    EventoTarea(Proc).Definicion = DefTareaTipo(TareaTipo + 1) + " " + Trim(Str(NoTareasTipo(TareaTipo)))
    EventoTarea(Proc).posx = X
    EventoTarea(Proc).posy = Y
    EventoTarea(Proc).NroPrescedentes = 0
    EventoTarea(Proc).NroConsecuentes = 0
    EventoTarea(Proc).ProcesoActivo = True
    EventoTarea(Proc).Terminal = True
    
    Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - (ObjEvento(0).Width / 2)) + "," + Str(Y - (ObjEvento(0).Height / 2)) + " "
    
    
    EventoPropiedades(Proc).IdProceso = EventoTarea(Proc).IdProceso
    EventoPropiedades(Proc).NoTarea = EventoTarea(Proc).NoTarea
    EventoPropiedades(Proc).TareaTipo = EventoTarea(Proc).TareaTipo
    EventoPropiedades(Proc).Definicion = EventoTarea(Proc).Definicion
    EventoPropiedades(Proc).IdTarea = EventoTarea(Proc).IdTarea
    
    EventoPropiedades(Proc).Para = "(ninguno)"
    EventoPropiedades(Proc).Mensaje = "(ninguno)"
    EventoPropiedades(Proc).Asunto = "(ninguno)"
    EventoPropiedades(Proc).DiasMaximo = 1  'ojo por defecto
    EventoPropiedades(Proc).Personalizada = False
    
    EventoPropiedades(Proc).NoCondiciones = 1
    ReDim EventoPropiedades(Proc).Condicion(1)
    EventoPropiedades(Proc).Condicion(1).CondicionActiva = True
    EventoPropiedades(Proc).Condicion(1).NoCondicion = 1
    EventoPropiedades(Proc).Condicion(1).Tipo = 0
    EventoPropiedades(Proc).Condicion(1).Definicion = "Concluido"

    DibujaTarea Proc, X, Y, EventoTarea(Proc).Definicion, TareaTipo
    If Proc >= 1 Then
       RequiereGrabar = True
       RequiereReproc = True
    End If
       
    Nt = Proc
    Call MuestraProceso
       
    With Deshacer
     .Accion = 1
     .Activo = True
     .Elemento = Proc
     HerramientasMenu.Buttons.Item(19).Enabled = .Activo
    End With
    Call VerificaHerramientas
    
    
    
End Sub


' ********************************************
' Dibuja elemento tarea (Referencial)
' ********************************************
Private Sub DibujaCaja(ByVal X As Integer, ByVal Y As Integer)
Dim old_mode As Integer
     If HerramientaSeleccionada = 5 Then
        ShapeMov.Move X - (Notas(0).Width / 2), Y - (Notas(0).Height / 2), Notas(0).Width, Notas(0).Height
     Else
        ShapeMov.Move X - (ObjEvento(0).Width / 2), Y - (ObjEvento(0).Height / 2), ObjEvento(0).Width, ObjEvento(0).Height
     End If
     ShapeMov.Visible = True
End Sub

' ********************************************
' Dibuja relacion entre dos tareas (referencial)
' ********************************************
Private Sub DibujaLinea(RC As RelacionCordenadas, Optional Posible As Boolean = False)
    Dim old_mode As Integer
    If Not EsPosibleRelacion Then Exit Sub


    ObjRelS1(0).BorderColor = IIf(Not Posible, &H8000000B, &HC00000)
    ObjRelS1(0).X1 = IIf(Mid(RC.VH, 1, 1) = "N", IIf(Mid(RC.VH, 2, 1) = "R", RC.X1, IIf(Mid(RC.VH, 2, 1) = "L", RC.X1, RC.X1)), RC.X1)
    ObjRelS1(0).X2 = RC.X2
    ObjRelS1(0).Y1 = IIf(Mid(RC.VH, 1, 1) = "D", RC.Y1, IIf(Mid(RC.VH, 1, 1) = "U", RC.Y1, RC.Y1))
    ObjRelS1(0).Y2 = RC.Y2: ObjRelS1(0).Visible = True ' IIf(ObjRelS1(0).Visible = True, False, True)
    
    
    ObjRelS2(0).BorderColor = IIf(Not Posible, &H8000000B, &HC00000)
    ObjRelS2(0).X1 = RC.X2: ObjRelS2(0).X2 = RC.X3
    ObjRelS2(0).Y1 = RC.Y2: ObjRelS2(0).Y2 = RC.Y3: ObjRelS2(0).Visible = True 'IIf(ObjRelS2(0).Visible = True, False, True)
    
    
    ObjRelS3(0).BorderColor = IIf(Not Posible, &H8000000B, &HC00000)
    ObjRelS3(0).X1 = RC.X3
    ObjRelS3(0).X2 = IIf(Mid(RC.VH, 1, 1) = "N", IIf(Mid(RC.VH, 2, 1) = "L", RC.X4, IIf(Mid(RC.VH, 2, 1) = "R", RC.X4, RC.X4)), RC.X4)
    ObjRelS3(0).Y1 = RC.Y3
    ObjRelS3(0).Y2 = IIf(Mid(RC.VH, 1, 1) = "U", RC.Y4, IIf(Mid(RC.VH, 1, 1) = "D", RC.Y4, RC.Y4))
    ObjRelS3(0).Visible = True

End Sub





 Private Sub Timer1_Timer()

    clsMouser.SetBorderStyle 1
    If clsMouser.IsMouseOver(Me, HerramientasMenu) Then
    ElseIf clsMouser.IsMouseOver(Me, Herramientas) Then
    Else
       Ventana.StatusBar1.Panels.Item(1).Text = ""
    End If

    If clsMouser.IsMouseOver(Ventana.ActiveForm, Contenedor) Then
    
    
    Else
        
        If TxtDatos.Visible Or cboTipo.Visible Then
           Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
           DatosAplicMtx.CellBoundary AplicMtx.Fila, AplicMtx.Columna, lLeft, lTop, lWidth, lHeight
           If TxtDatos.Visible Then TxtDatos.Move lLeft, lTop, lWidth
'           If cboTipo.Visible Then cboTipo.Move lLeft, lTop, lWidth
        End If
        
        If Len(Ventana.StatusBar1.Panels.Item(4).Text) > 0 Then
           Ventana.StatusBar1.Panels.Item(4).Text = "0, 0"
        End If
        
        LabelIns.Visible = False
        ShapeDel.Visible = False
        ShapeMov.Visible = False
    
    
    End If

End Sub

Private Sub Herramientas_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case Is = "NORMAL"
        HerramientaSeleccionada = 0
        ShapeMov.Visible = False
    Case Is = "NOTIF"
        HerramientaSeleccionada = 1
        TareaTipo = 6
    Case Is = "NOTIF_RESP"
        HerramientaSeleccionada = 1
        TareaTipo = 5
    Case Is = "BIFURC"
        HerramientaSeleccionada = 1
        TareaTipo = 4
    Case Is = "RETORNO"
        HerramientaSeleccionada = 1
        TareaTipo = 3
    Case Is = "META"
        HerramientaSeleccionada = 1
        TareaTipo = 2
    Case Is = "SUBPROCESO"
        HerramientaSeleccionada = 1
        TareaTipo = 1
    Case Is = "FINAL"
        HerramientaSeleccionada = 1
        TareaTipo = 0
    Case Is = "INTERCAMBIO"
        HerramientaSeleccionada = 1
        TareaTipo = 8
    Case Is = "RELACION"
        ShapeMov.Visible = False
        HerramientaSeleccionada = 2
        Pizarra.MouseIcon = ImageListCur.ListImages.Item(5).Picture
    Case Is = "ELIMINAR"
        ShapeMov.Visible = False
        HerramientaSeleccionada = 3
        Pizarra.MouseIcon = ImageListCur.ListImages.Item(1).Picture
    Case Is = "INSERT"
        ShapeMov.Visible = False
        HerramientaSeleccionada = 4
        Pizarra.MouseIcon = ImageListCur.ListImages.Item(2).Picture
    Case Is = "NOTAS"
        ShapeMov.Visible = False
        HerramientaSeleccionada = 5
        Pizarra.MouseIcon = ImageListCur.ListImages.Item(2).Picture
    Case Is = "ARRASTRE"
        TeclaStat = Not TeclaStat
        Herramientas.Buttons.Item(14).Value = IIf(TeclaStat, tbrPressed, tbrUnpressed)
        Ventana.StatusBar1.Panels.Item(2).Text = IIf(TeclaStat, "DESP ENCENDIDO", "DESP APAGADO")
    End Select
    If HerramientaSeleccionada = 1 Then
       Pizarra.MouseIcon = ImageListCur.ListImages.Item(8).Picture
    End If

End Sub

Private Sub Herramientas_DragDrop(Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos

End Sub

Private Sub Herramientas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     LabelIns.Visible = False
     ShapeMov.Visible = False
     Dim C%
     Dim Xp%
     Xp = Y / Screen.TwipsPerPixelY
     For C = 1 To Herramientas.Buttons.Count
         If (Xp >= Herramientas.Buttons.Item(C).Top And Xp <= (Herramientas.Buttons.Item(C).Top + Herramientas.Buttons.Item(C).Height)) Then
             Ventana.StatusBar1.Panels.Item(1).Text = Herramientas.Buttons.Item(C).Description
             Exit For
         End If
     Next

End Sub

Private Sub HerramientasMenu_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case Is = "NUEVO"
         Call Archivo_Click(0)
         HerramientasMenu.Buttons.Item(1).Value = tbrUnpressed
    Case Is = "LIMPIAR"
         Call Archivo_Click(2)
         HerramientasMenu.Buttons.Item(2).Value = tbrUnpressed
    Case Is = "ABRIR"
         Call Archivo_Click(3)
         HerramientasMenu.Buttons.Item(3).Value = tbrUnpressed
    Case Is = "GRABAR"
         Call Archivo_Click(4)
         HerramientasMenu.Buttons.Item(4).Value = tbrUnpressed
    Case Is = "AUMENTAR"
         Call TamañoPizarra_Click(0)
    Case Is = "REDUCIR"
         Call TamañoPizarra_Click(1)
    Case Is = "HERRAMIENTA"
         Herramientas.Visible = Not Herramientas.Visible
         HerramientasMenu.Buttons.Item(9).Value = IIf(Herramientas.Visible, tbrPressed, tbrUnpressed)
         If Not Herramientas.Visible Then
            Herramientas.Buttons.Item(1).Value = tbrPressed
            HerramientaSeleccionada = 0
            Pizarra.MousePointer = 0
         End If
         Form_Resize
    Case Is = "DATOS"
         ContenedorDatos.Visible = Not ContenedorDatos.Visible
         HerramientasMenu.Buttons.Item(10).Value = IIf(ContenedorDatos.Visible, tbrPressed, tbrUnpressed)
         Form_Resize
    
    Case Is = "AYUDA"
         Call Mnu_Ayuda_Click(2)
    Case Is = "DESHACER"
         With Deshacer
             Select Case .Accion
                    Case 1
                        Call EliminaTarea(.Elemento)
                    Case 2
                        Call EliminaRelacion(.Elemento)
             End Select
         .Activo = False
         HerramientasMenu.Buttons.Item(19).Enabled = .Activo
         End With
         
    End Select


End Sub

Private Sub HerramientasMenu_DragDrop(Source As Control, X As Single, Y As Single)
     Call RemarcaRelaciones(ObjEnMovimiento, 0)
     Call OcultaElementos

End Sub



Private Sub HerramientasMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     LabelIns.Visible = False
     ShapeMov.Visible = False
     Dim C%
     Dim Xp%
     Xp = X / Screen.TwipsPerPixelX
     For C = 1 To HerramientasMenu.Buttons.Count
         If (Xp >= HerramientasMenu.Buttons.Item(C).Left And Xp <= (HerramientasMenu.Buttons.Item(C).Left + HerramientasMenu.Buttons.Item(C).Width)) Then
             Ventana.StatusBar1.Panels.Item(1).Text = HerramientasMenu.Buttons.Item(C).Description
             Exit For
         End If
     Next
     
End Sub


Private Sub TxtDatos_KeyPress(KeyAscii As Integer)
   Dim i As Long
   If KeyAscii = 27 Then
      TxtDatos.Text = DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna)
      KeyAscii = 0
      TxtDatos.Visible = False
      DatosAplicMtx.SetFocus
      Exit Sub
   End If
   If AplicMtx.Columna = 4 And DatosAplicMtx.CellText(AplicMtx.Fila, 3) = "Caracter" Then
      If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 13 Or KeyAscii = 8 Then
      Else
      
         KeyAscii = 0
         Exit Sub
      End If
      If Val(TxtDatos.Text + Chr(KeyAscii)) > 255 And KeyAscii <> 13 And KeyAscii <> 8 Then
         KeyAscii = 0
         Exit Sub
      End If
   End If
   If TxtDatos.Visible And KeyAscii = 13 Then
      If TxtDatos.Text <> "(ninguno)" Then
         If AplicMtx.Columna = 1 Or AplicMtx.Columna = 2 Then
            For i = 1 To DatosAplicMtx.Rows
               If i <> AplicMtx.Fila Then
                  If DatosAplicMtx.RowVisible(i) Then
                     If UCase(TxtDatos.Text) = UCase(DatosAplicMtx.CellText(i, AplicMtx.Columna)) Then
                        MsgBox "El Nombre de Campo Ingresado Ya Existe!", vbExclamation, "Datos De Proceso"
                        TxtDatos.Text = ""
                        DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna) = TxtDatos.Text
                        Exit Sub
                     End If
                  End If
               End If
            Next i
         End If
      
         Select Case AplicMtx.Columna
           Case 1
           
               i = CLng(AplicMtx.Fila)
             With DatosAplicMtx
               .CellText(i, 1) = TxtDatos.Text
               .CellText(i, 2) = "(ninguno)"
               .CellText(i, 3) = "Caracter"
               .CellText(i, 4) = "10"
               .CellText(i, 5) = "Nulo"
               .CellText(i, 6) = LCase(Trim(TxtDatos.Text))
               .CellForeColor(i, 1) = vbWindowText
               .CellForeColor(i, 2) = vbWindowText
               .CellForeColor(i, 3) = vbWindowText
               .CellForeColor(i, 4) = vbWindowText
               .CellForeColor(i, 5) = vbWindowText
               
               If i = DatosAplicMtx.Rows Then
                  i = i + 1
                  .AddRow , , (i = 1)
                  .CellText(i, 1) = "(ninguno)"
                  .CellText(i, 2) = "(ninguno)"
                  .CellText(i, 3) = "Caracter"
                  .CellText(i, 4) = "10"
                  .CellText(i, 5) = "Nulo"
                  .CellForeColor(i, 1) = vbButtonFace
                  .CellForeColor(i, 2) = vbButtonFace
                  .CellForeColor(i, 3) = vbButtonFace
                  .CellForeColor(i, 4) = vbButtonFace
                  .CellForeColor(i, 5) = vbButtonFace
                  
                  If Not .RowVisible(i) Then
                     ReDim Preserve DatosProceso(i)
                     .RowVisible(i) = True
                  End If
               End If
             End With
           Case 4
               TxtDatos.Text = Val(TxtDatos.Text)
         End Select
         DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna) = TxtDatos.Text
      Else
         TxtDatos.Text = ""
         DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna) = TxtDatos.Text
      End If
      Select Case AplicMtx.Columna
             Case 1
                 DatosProceso(AplicMtx.Fila).Campo = TxtDatos.Text
                 DatosProceso(AplicMtx.Fila).Clave = LCase(Trim(TxtDatos.Text))
             Case 2
                 DatosProceso(AplicMtx.Fila).Definicion = TxtDatos.Text
             Case 4
                 DatosProceso(AplicMtx.Fila).Longitud = TxtDatos.Text
             Case 5
                 DatosProceso(AplicMtx.Fila).VDefecto = TxtDatos.Text
      End Select
      DatosProceso(AplicMtx.Fila).Activo = True
      TxtDatos.Visible = False
      DatosAplicMtx.SetFocus
      SendKeys "{RIGHT}"
      DatosAplicMtx.SelectedCol = AplicMtx.Columna
      
      RequiereGrabar = True
      RequiereReproc = True
      Call MuestraProceso
   End If
End Sub

Private Sub TxtDatos_LostFocus()
   Dim i As Long
   Dim lLastRow As Long
   Dim lNC As Long
   Select Case AplicMtx.Columna
          Case 1, 2
               If Len(Trim(TxtDatos)) = 0 Then
                  TxtDatos.Text = "(ninguno)"
                  DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna) = TxtDatos.Text
                  If AplicMtx.Columna = 1 Then
                     For i = AplicMtx.Fila + 1 To DatosAplicMtx.Rows
                         If DatosAplicMtx.RowVisible(i) Then
                            DatosAplicMtx.CellText(i - 1, 1) = DatosAplicMtx.CellText(i, 1)
                            DatosAplicMtx.CellText(i - 1, 2) = DatosAplicMtx.CellText(i, 2)
                            DatosAplicMtx.CellText(i - 1, 3) = DatosAplicMtx.CellText(i, 3)
                            DatosAplicMtx.CellText(i - 1, 4) = DatosAplicMtx.CellText(i, 4)
                            DatosAplicMtx.CellText(i - 1, 5) = DatosAplicMtx.CellText(i, 5)
                            DatosAplicMtx.CellText(i - 1, 6) = DatosAplicMtx.CellText(i, 6)
                            DatosProceso(i - 1).Activo = DatosProceso(i).Activo
                            DatosProceso(i - 1).Campo = DatosProceso(i).Campo
                            DatosProceso(i - 1).Definicion = DatosProceso(i).Definicion
                            DatosProceso(i - 1).Tipo = DatosProceso(i).Tipo
                            DatosProceso(i - 1).Longitud = DatosProceso(i).Longitud
                            DatosProceso(i - 1).VDefecto = DatosProceso(i).VDefecto
                            DatosProceso(i - 1).Clave = DatosProceso(i).Clave
                            lLastRow = i
                         End If
                     Next i
                     If (lLastRow = 0) Then lLastRow = DatosAplicMtx.Rows
                     DatosAplicMtx.CellText(lLastRow, 1) = "(ninguno)"
                     DatosAplicMtx.CellForeColor(lLastRow, 1) = vbButtonFace
                     For i = 1 To DatosAplicMtx.Rows
                        If DatosAplicMtx.CellText(i, 1) = "(ninguno)" Then
                           lNC = lNC + 1
                           DatosAplicMtx.CellForeColor(i, 1) = vbButtonFace
                        End If
                        If lNC > 1 Then
                           DatosAplicMtx.RowVisible(i) = False
                           DatosAplicMtx.RemoveRow i
                           ReDim Preserve DatosProceso(i - 1)
                           RequiereGrabar = True
                           RequiereReproc = True
                           Call MuestraProceso
                        End If
                     Next i
                  End If
               End If
          Case 4
               If Len(Trim(TxtDatos)) = 0 Then
                  Select Case DatosAplicMtx.CellText(AplicMtx.Fila, 3)
                         Case "Caracter"
                              TxtDatos.Text = "10"
                         Case "Numerico"
                              TxtDatos.Text = "10.2"
                         Case "Moneda"
                              TxtDatos.Text = "12.2"
                         Case "Fecha - Hora"
                              '09/09/2000 23:45:00"
                              TxtDatos.Text = "19"
                         Case "Logico"
                              TxtDatos.Text = "2"
                  End Select
                  DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna) = TxtDatos.Text
                  DatosProceso(AplicMtx.Fila).Longitud = TxtDatos.Text
                  RequiereGrabar = True
                  RequiereReproc = True
                  Call MuestraProceso
               End If
          Case 5
               If Len(Trim(TxtDatos)) = 0 Then
                  Select Case DatosAplicMtx.CellText(AplicMtx.Fila, 3)
                         Case "Caracter"
                              TxtDatos.Text = "Nulo"
                         Case "Numerico"
                              TxtDatos.Text = "0.00"
                         Case "Moneda"
                              TxtDatos.Text = "0.00"
                         Case "Fecha - Hora"
                              TxtDatos.Text = ""
                         Case "Logico"
                              TxtDatos.Text = "No"
                  End Select
                  DatosAplicMtx.CellText(AplicMtx.Fila, AplicMtx.Columna) = TxtDatos.Text
                  DatosProceso(AplicMtx.Fila).VDefecto = TxtDatos.Text
                  RequiereGrabar = True
                  RequiereReproc = True
                  Call MuestraProceso
              End If
   End Select
               
   TxtDatos.Visible = False
   DatosAplicMtx.CancelEdit
   DatosAplicMtx.SetFocus
   
End Sub

Private Sub VentanaMovil_DragDrop(Source As Control, X As Single, Y As Single)
    Call RemarcaRelaciones(ObjEnMovimiento, 0)
    Call OcultaElementos

End Sub


Private Sub VentanaMovil_HorizontalScroll(Stat As Byte)
   If Not Herramientas.Visible Then Exit Sub
   
   If Stat = 1 Then
      Pizarra.Width = IIf(Pizarra.Width < ((VentanaMovil.Width * 15) * 2), Pizarra.Width * 1.2, Pizarra.Width)
      Plantilla.XMax = Pizarra.Width
      VentanaMovil.Refresh
   End If
End Sub

Private Sub VentanaMovil_MouseMove()
   If HerramientaSeleccionada = 1 Then ShapeMov.Visible = False
      

End Sub

Private Sub VentanaMovil_VerticalScroll(Stat As Byte)
   If Not Herramientas.Visible Then Exit Sub
   
   If Stat = 1 Then
      Pizarra.Height = IIf(Pizarra.Height < ((VentanaMovil.Height * 15) * 2), Pizarra.Height * 1.2, Pizarra.Height)
      Plantilla.YMax = Pizarra.Height
      VentanaMovil.Refresh
   End If

End Sub

Private Sub Ver_Click(Index As Integer)
    Select Case Index
        Case 0
            'conmuta el estado de la barra de herramientas (visible/no visible)
            If Not Herramientas.Visible Then
                Herramientas.Visible = True
            Else
                Herramientas.Visible = False
            End If
            Form_Resize
    End Select

End Sub


Private Sub CreaRelacion(ProcOri As Integer, ProcDst As Integer, Optional TipoRela As Integer)
    
    Dim CreadaPrevia As Boolean
    CreadaPrevia = False
    Dim Nc%, i%
    Nc = EventoTarea(ProcOri).NroConsecuentes
    If Nc > 0 Then
       For i = 1 To Nc
           If EventoTarea(ProcOri).Consecuente(i).TareaConsecuente = ProcDst Then Exit For
       Next
       If i <= Nc Then
          If EventoPropiedades(ProcOri).NoCondiciones > 1 Then
             CreadaPrevia = True
          Else
             MsgBox "Lo siento, ya fue creada previamente esta relacion!", vbInformation, "Relacionador de Tareas"
             Exit Sub
          End If
       End If
    End If
    Nc = EventoTarea(ProcOri).NroPrescedentes
    If Nc > 0 Then
       For i = 1 To Nc
           If EventoTarea(ProcOri).Prescedente(i).TareaPrescedente = ProcDst Then Exit For
       Next
       If i <= Nc Then
          MsgBox "ATENCION!..., Ud. no puede crear una relacion de naturaleza circular" + Chr(10) + Chr(13) + "Para establecer ese tipo de relacion debe emplear la tarea de RETORNO!", vbInformation, "Relacionador de Tareas"
          Exit Sub
       End If
    End If
    
    Call CalculaCoordenadasRelacion(ObjEvento(ProcOri), ObjEvento(ProcDst), ProcOri, ProcDst, CreadaPrevia)
       
    
    Dim Mx%, Relac%
    
    If MaxRelaObj > 0 Then
       For Mx = 1 To MaxRelaObj
           If Not Relacion(Mx).RelacionActiva Then Exit For
       Next Mx
    Else
       Mx = 1
    End If
    If Mx > MaxRelaObj Then
       MaxRelaObj = MaxRelaObj + 1: Relac = MaxRelaObj
       ReDim Preserve Relacion(MaxRelaObj)
    Else
       Relac = Mx
    End If
    
    Herramientas.Buttons.Item(12).Enabled = IIf(Relac > 0, True, False)

    Load ObjRelP1(Relac): Load ObjRelP4(Relac)
    Load ObjRelP2(Relac): Load ObjRelP3(Relac)
    
    Load ObjRelS1(Relac): Load ObjRelS2(Relac): Load ObjRelS3(Relac)
    Load ObjRelT1(Relac)
    ObjRelT1(Relac).Caption = "Relacion " + Str(ProcOri) + " a " + Str(ProcDst) + " " + Str(Relac)
        
    Dim MxPrescedentes As Integer, MxConsecuentes As Integer
        
    MxConsecuentes = EventoTarea(ProcOri).NroConsecuentes + 1
    MxPrescedentes = EventoTarea(ProcDst).NroPrescedentes + 1
    
    Relacion(Relac).TareaOrigen = ProcOri
    Relacion(Relac).ConsecuentesOrigen = MxConsecuentes
    Relacion(Relac).TareaDestin = ProcDst
    Relacion(Relac).PrescedentesDestin = MxPrescedentes
    Relacion(Relac).RelacionActiva = True
    
    
    ReDim Preserve EventoTarea(ProcOri).Consecuente(MxConsecuentes): EventoTarea(ProcOri).NroConsecuentes = MxConsecuentes
    ReDim Preserve EventoTarea(ProcDst).Prescedente(MxPrescedentes): EventoTarea(ProcDst).NroPrescedentes = MxPrescedentes
        
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).TareaPrescedente = ProcOri
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).TareaConsecuente = ProcDst
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).RelacionTipo = EventoPropiedades(ProcOri).Condicion(1).NoCondicion
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).PosicionXYPersonal = False
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).RelacionActiva = True
       
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).NoElementoRelacion = Relac
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).VH = r.VH
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).Ruta = r.Ruta
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).X1 = r.X1
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).X2 = r.X2
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).X3 = r.X3
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).X4 = r.X4
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).Y1 = r.Y1
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).Y2 = r.Y2
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).Y3 = r.Y3
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).Y4 = r.Y4
    Circular = False
    Call SecuenciaReversaObjeto(ProcOri, ProcDst)
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).RelacionInfinita = Circular
'    EventoTarea(ProcOri).Consecuente(MxConsecuentes).RelacionInfinita = Circular
    
    If TipoRela > 0 Then EventoTarea(ProcDst).Prescedente(MxPrescedentes).RelacionTipo = TipoRela
    
    EventoTarea(ProcOri).Consecuente(MxConsecuentes) = EventoTarea(ProcDst).Prescedente(MxPrescedentes)
    
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).NumeroRPrescedente = MxConsecuentes
    EventoTarea(ProcDst).Prescedente(MxPrescedentes).TareaPrescedente = ProcOri
    EventoTarea(ProcOri).Consecuente(MxConsecuentes).NumeroRConsecuente = MxPrescedentes
    EventoTarea(ProcOri).Consecuente(MxConsecuentes).TareaConsecuente = ProcDst
        
    EventoTarea(ProcDst).Terminal = IIf(EventoTarea(ProcDst).NroConsecuentes > 0, False, True)
    
    Call DibujaRelacion(Relac, ProcOri, MxConsecuentes, ProcDst, MxPrescedentes)
    
    RequiereGrabar = True
    RequiereReproc = True
    Call MuestraProceso

    With Deshacer
     .Accion = 2
     .Activo = True
     .Elemento = Relac
     HerramientasMenu.Buttons.Item(19).Enabled = .Activo
    End With


End Sub



Private Function RequiereSalir%()
  Dim ret%, i%
  ret = GrabarCambios("Salir")
  If ret <> IDCANCEL Then
     If FORM_NUM < UBound(Flujo) Then
        For i% = FORM_NUM + 1 To UBound(Flujo)
           Flujo(i% - 1).Caption = Flujo(i%).Caption
           Flujo(i% - 1).Handle = Flujo(i%).Handle
        Next i%
     End If
     If UBound(Flujo) > 0 Then
        ReDim Preserve Flujo(UBound(Flujo) - 1)
     Else
        ReDim Flujo(0)
     End If
     RequiereSalir = 0
  Else
     RequiereSalir = 1
  End If
End Function

Private Function Grabar%()
  Dim ret%

  If NombreDeArchivo$ = "" Then
    ret = GrabarComo()
    Grabar = ret
  Else
    ret = EscribeArchivo(NombreDeArchivo)
    Grabar = IDOK
  End If

End Function

Private Function GrabarCambios%(ByVal Caption$)
  Dim Message$
  Dim MsgBoxType%
  Dim Response%

  If Not RequiereGrabar Then
    GrabarCambios = IDNOSAVE
    Exit Function
  End If

  Message$ = "         Grabar los Cambios?"
  MsgBoxType = MB_YESNOCANCEL + MB_ICONEXCLAMATION + MB_DEFBUTTON1
  Response = MsgBox(Message$, MsgBoxType, Caption)

  If Response = IDYES Then
    If NombreDeArchivo = "" Then
      Response = GrabarComo()
    Else
      Response = Grabar()
      If Response = IDOK Then
        Response = IDYES
      End If
    End If
  End If

  GrabarCambios = Response

End Function


Private Function GrabarComo%()
  Dim Message$
  Dim MsgBoxType%
  Dim Response%

  GrabarComo = IDOK

  On Error GoTo grabarError

  CMDialog1.CancelError = True
  CMDialog1.Filter = "Todos los Archivos (*.*)|*.*|Archivos de Procesos (*.PRC)|*.prc"
  CMDialog1.FilterIndex = 2
  CMDialog1.Action = 2

  If Dir$(CMDialog1.FileName, 0) <> "" Then
     Message$ = "Archivo ya existe. Desea reemplazarlo?"
     MsgBoxType = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2
     Response = MsgBox(Message$, MsgBoxType, "Grabar Como")
     GrabarComo = Response
     If Response = IDNO Then
        GrabarComo = GrabarComo()
        Exit Function
     End If
  End If

  Dim Fname$
  Fname = CMDialog1.FileName
  If EscribeArchivo(Fname) Then
    GoTo grabarError
  End If

  NombreDeArchivo = CMDialog1.FileName
  Flujo(FORM_NUM).Caption = GetFileName(NombreDeArchivo)
  mnuFlujos(FORM_NUM - 1).Caption = Flujo(FORM_NUM).Caption
  Me.Caption = "Proceso - " + GetFileName(NombreDeArchivo)

  Exit Function

grabarError:
  If Err = CDERR_CANCEL Then
    GrabarComo = IDCANCEL
  End If
  Exit Function

End Function



Private Function EscribeArchivo%(ByVal Fname$)


  Dim FileNum%
  Dim i%, p%, C%
  Dim da%
  On Error GoTo EscribeError

  'Si el archivo existe , eliminarlo
  If Dir$(Fname, 0) <> "" Then
     Kill Fname
  End If

  FileNum = FreeFile

  Open Fname For Binary As FileNum
   
  Put #FileNum, , "WSWORKFLOW"
  Put #FileNum, , Pizarra.Width      'As Single
  Put #FileNum, , Pizarra.Height     'As Single
  Put #FileNum, , Plantilla.XObjMax   'As Single
  Put #FileNum, , Plantilla.YObjMax   'As Single
  Put #FileNum, , MaxProcObj
  Put #FileNum, , MaxRelaObj
  For i = 1 To MaxProcObj
        Put #FileNum, , EventoTarea(i).ProcesoActivo 'As Boolean
        Put #FileNum, , EventoTarea(i).IdProceso     'As String * 6
        Put #FileNum, , EventoTarea(i).IdTarea       'As String * 6
        Put #FileNum, , EventoTarea(i).NoTarea       'As Integer
        Put #FileNum, , EventoTarea(i).TareaTipo       'As Byte
        Put #FileNum, , EventoTarea(i).posx            'As Single
        Put #FileNum, , EventoTarea(i).posy            'As Single
        Put #FileNum, , EventoTarea(i).Definicion      'As String * 25
        Put #FileNum, , EventoTarea(i).NroPrescedentes 'As Byte
        Put #FileNum, , EventoTarea(i).NroConsecuentes 'As Byte
        Put #FileNum, , EventoTarea(i).Terminal        'As Boolean
        If EventoTarea(i).NroPrescedentes > 0 Then
           For p = 1 To EventoTarea(i).NroPrescedentes
                Put #FileNum, , EventoTarea(i).Prescedente(p).RelacionActiva      'As Boolean
                Put #FileNum, , EventoTarea(i).Prescedente(p).RelacionInfinita    'As Boolean
                Put #FileNum, , EventoTarea(i).Prescedente(p).TareaPrescedente    'As Integer
                Put #FileNum, , EventoTarea(i).Prescedente(p).NumeroRPrescedente  'As Integer
                Put #FileNum, , EventoTarea(i).Prescedente(p).TareaConsecuente    'As Integer
                Put #FileNum, , EventoTarea(i).Prescedente(p).NumeroRConsecuente  'As Integer
                Put #FileNum, , EventoTarea(i).Prescedente(p).NoElementoRelacion  'As Integer
                Put #FileNum, , EventoTarea(i).Prescedente(p).RelacionTipo        'As Integer
                Put #FileNum, , EventoTarea(i).Prescedente(p).PosicionXYPersonal  'As Boolean
                Put #FileNum, , EventoTarea(i).Prescedente(p).VH                  'As String * 2
                Put #FileNum, , EventoTarea(i).Prescedente(p).Ruta                'As String * 1
                Put #FileNum, , EventoTarea(i).Prescedente(p).X1                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).Y1                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).X2                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).Y2                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).X3                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).Y3                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).X4                  'As Single
                Put #FileNum, , EventoTarea(i).Prescedente(p).Y4                  'As Single
           Next p
        End If
        If EventoTarea(i).NroConsecuentes > 0 Then
           For C = 1 To EventoTarea(i).NroConsecuentes
                Put #FileNum, , EventoTarea(i).Consecuente(C).RelacionActiva      'As Boolean
                Put #FileNum, , EventoTarea(i).Consecuente(C).RelacionInfinita    'As Boolean
                Put #FileNum, , EventoTarea(i).Consecuente(C).TareaPrescedente    'As Integer
                Put #FileNum, , EventoTarea(i).Consecuente(C).NumeroRPrescedente  'As Integer
                Put #FileNum, , EventoTarea(i).Consecuente(C).TareaConsecuente    'As Integer
                Put #FileNum, , EventoTarea(i).Consecuente(C).NumeroRConsecuente  'As Integer
                Put #FileNum, , EventoTarea(i).Consecuente(C).NoElementoRelacion  'As Integer
                Put #FileNum, , EventoTarea(i).Consecuente(C).RelacionTipo        'As Integer
                Put #FileNum, , EventoTarea(i).Consecuente(C).PosicionXYPersonal  'As Boolean
                Put #FileNum, , EventoTarea(i).Consecuente(C).VH                  'As String * 2
                Put #FileNum, , EventoTarea(i).Consecuente(C).Ruta                'As String * 1
                Put #FileNum, , EventoTarea(i).Consecuente(C).X1                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).Y1                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).X2                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).Y2                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).X3                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).Y3                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).X4                  'As Single
                Put #FileNum, , EventoTarea(i).Consecuente(C).Y4                  'As Single
           Next C
        End If
        Put #FileNum, , EventoPropiedades(i).Personalizada
        Put #FileNum, , EventoPropiedades(i).IdProceso
        Put #FileNum, , EventoPropiedades(i).IdTarea
        Put #FileNum, , EventoPropiedades(i).NoTarea
        Put #FileNum, , EventoPropiedades(i).Definicion
        Put #FileNum, , EventoPropiedades(i).TareaTipo
        Put #FileNum, , EventoPropiedades(i).Asunto
        Put #FileNum, , EventoPropiedades(i).Mensaje
        Put #FileNum, , EventoPropiedades(i).Para
        Put #FileNum, , EventoPropiedades(i).NoCondiciones
        Put #FileNum, , EventoPropiedades(i).DiasMaximo
        For C = 1 To EventoPropiedades(i).NoCondiciones
            Put #FileNum, , EventoPropiedades(i).Condicion(C).CondicionActiva
            Put #FileNum, , EventoPropiedades(i).Condicion(C).NoCondicion
            Put #FileNum, , EventoPropiedades(i).Condicion(C).Definicion
            Put #FileNum, , EventoPropiedades(i).Condicion(C).Tipo
        Next C
  Next i

  For i = 1 To MaxRelaObj
      Put #FileNum, , Relacion(i).ConsecuentesOrigen
      Put #FileNum, , Relacion(i).PrescedentesDestin
      Put #FileNum, , Relacion(i).RelacionActiva
      Put #FileNum, , Relacion(i).TareaDestin
      Put #FileNum, , Relacion(i).TareaOrigen
  Next
 
  da = UBound(DatosProceso) - 1
  Put #FileNum, , da
  If da > 0 Then
     For i = 1 To da
         Put #FileNum, , DatosProceso(i).Activo
         Put #FileNum, , DatosProceso(i).Campo
         Put #FileNum, , DatosProceso(i).Definicion
         Put #FileNum, , DatosProceso(i).Tipo
         Put #FileNum, , DatosProceso(i).Longitud
         Put #FileNum, , DatosProceso(i).VDefecto
         Put #FileNum, , DatosProceso(i).Clave
     Next
  End If

  Put #FileNum, , MaxNotaObj
  If MaxNotaObj > 0 Then
     For i = 1 To MaxNotaObj
        Put #FileNum, , EventoNota(i).NotaActiva
        Put #FileNum, , EventoNota(i).IdNota
        Put #FileNum, , EventoNota(i).IdProceso
        Put #FileNum, , EventoNota(i).Nota
        Put #FileNum, , EventoNota(i).posx
        Put #FileNum, , EventoNota(i).posy
        Put #FileNum, , EventoNota(i).Titulo
        Put #FileNum, , EventoNota(i).Definicion
        Put #FileNum, , EventoNota(i).Fecha
     Next
  End If

  Dim K As String * 20
  Dim t As String * 50
  da = Adjuntos.ListItems.Count
  Put #FileNum, , da
  If da > 0 Then
     For C = 1 To da
        K = Adjuntos.ListItems.Item(C).Key
        t = Adjuntos.ListItems.Item(C).Text
        Put #FileNum, , K
        Put #FileNum, , t
     Next
  End If

  Close FileNum

  RequiereGrabar = False
  RequiereReproc = False
  
  EscribeArchivo = False
  Exit Function

EscribeError:
  EscribeArchivo = True
  Exit Function

End Function


Function LeerAdjunto(ByVal Fname$)
    On Error GoTo leeError
    Dim FileNum%
    Dim Adj As Variant
    Dim Whole As Long
    Dim Part  As Long
    Dim Start As Long
    Dim Buffer1 As String
    Dim X As Long
    
    FileNum = FreeFile
    Me.MousePointer = 11
    
    'Abre el archvo de modo binario
    Open Fname For Binary As FileNum

    Whole = LOF(FileNum) \ 20000          'number of whole 10,000 byte chunks
    Part = LOF(FileNum) Mod 20000         'remaining bytes at end of file
    Buffer1 = String$(20000, 0)
    Start = 1
        
    For X = 1 To Whole                    'this for-next loop will get 10,000
        Get #FileNum, Start, Buffer1       'byte chunks at a time.
        Adj = Adj + Buffer1
        Start = Start + 20000
    Next
        
    Buffer1$ = String$(Part, 0)
    Get #FileNum, Start, Buffer1          'get the remaining bytes at the end
        
    Adj = Adj + Buffer1
        
    Close #FileNum
    LeerAdjunto = False
    Debug.Print Adj
    Me.MousePointer = 0
    
    Exit Function

leeError:
  Me.MousePointer = 0

  LeerAdjunto = True
  Exit Function

End Function

Function LeerAcceso(ByVal Fname$)
    On Error GoTo leeError
    Dim FileNum%
    Dim Adj As Variant
    Dim Whole As Long
    Dim Part  As Long
    Dim Start As Long
    Dim Buffer1 As String
    Dim X As Long
    
    FileNum = FreeFile
    Me.MousePointer = 11
    
    'Abre el archvo de modo binario
    Open Fname For Binary As FileNum

    Whole = LOF(FileNum) \ 20000          'number of whole 10,000 byte chunks
    Part = LOF(FileNum) Mod 20000         'remaining bytes at end of file
    Buffer1 = String$(20000, 0)
    Start = 1
        
    For X = 1 To Whole                    'this for-next loop will get 10,000
        Get #FileNum, Start, Buffer1       'byte chunks at a time.
        Adj = Adj + Buffer1
        Start = Start + 20000
    Next
        
    Buffer1$ = String$(Part, 0)
    Get #FileNum, Start, Buffer1          'get the remaining bytes at the end
        
    Adj = Adj + Buffer1
        
    Close #FileNum
    LeerAcceso = False
    Debug.Print Adj
    Me.MousePointer = 0
    
    Exit Function

leeError:
  Me.MousePointer = 0

  LeerAcceso = True
  Exit Function

End Function


Private Function LeeArchivo(ByVal Fname$)

  Dim FileNum%
  Dim i%, p%, C%, a As String * 10
  Dim da%

  If Not ArchivoValido(Fname$) Then
     MsgBox "No es un archivo de Procesos valido.", MB_OK + MB_ICONEXCLAMATION, "Abrir"
     GoTo leeError
  End If

  'Salta al manejador de errors cuando exista error
'  On Error GoTo leeError
  
  ContenedorDatos.Tab = 3
  PizarraDatos(0).ZOrder 0
  
  Call LimpiarPizarra(False)
  
  FileNum = FreeFile

  'Abre el archvo de modo binario
  Open Fname For Binary As FileNum

  Get #FileNum, , a
  Get #FileNum, , Plantilla.XMax      'As Single
  Get #FileNum, , Plantilla.YMax      'As Single
  Get #FileNum, , Plantilla.XObjMax   'As Single
  Get #FileNum, , Plantilla.YObjMax   'As Single
  Get #FileNum, , MaxProcObj
  Get #FileNum, , MaxRelaObj
  ReDim EventoTarea(MaxProcObj)
  ReDim EventoPropiedades(MaxProcObj)
  ReDim Relacion(MaxRelaObj)
  
  'lee registros y los carga a la matriz
  For i = 1 To MaxProcObj
        Get #FileNum, , EventoTarea(i).ProcesoActivo 'As Boolean
        Get #FileNum, , EventoTarea(i).IdProceso     'As String * 6
        Get #FileNum, , EventoTarea(i).IdTarea       'As String * 6
        Get #FileNum, , EventoTarea(i).NoTarea       'As Integer
        Get #FileNum, , EventoTarea(i).TareaTipo       'As Byte
        Get #FileNum, , EventoTarea(i).posx            'As Single
        Get #FileNum, , EventoTarea(i).posy            'As Single
        Get #FileNum, , EventoTarea(i).Definicion      'As String * 25
        Get #FileNum, , EventoTarea(i).NroPrescedentes 'As Byte
        Get #FileNum, , EventoTarea(i).NroConsecuentes 'As Byte
        Get #FileNum, , EventoTarea(i).Terminal        'As Boolean
        ReDim EventoTarea(i).Prescedente(EventoTarea(i).NroPrescedentes)
        ReDim EventoTarea(i).Consecuente(EventoTarea(i).NroConsecuentes)
        If EventoTarea(i).NroPrescedentes > 0 Then
           For p = 1 To EventoTarea(i).NroPrescedentes
                Get #FileNum, , EventoTarea(i).Prescedente(p).RelacionActiva       'As Boolean
                Get #FileNum, , EventoTarea(i).Prescedente(p).RelacionInfinita     'As Boolean
                Get #FileNum, , EventoTarea(i).Prescedente(p).TareaPrescedente     'As Integer
                Get #FileNum, , EventoTarea(i).Prescedente(p).NumeroRPrescedente   'As Integer
                Get #FileNum, , EventoTarea(i).Prescedente(p).TareaConsecuente     'As Integer
                Get #FileNum, , EventoTarea(i).Prescedente(p).NumeroRConsecuente   'As Integer
                Get #FileNum, , EventoTarea(i).Prescedente(p).NoElementoRelacion   'As Integer
                Get #FileNum, , EventoTarea(i).Prescedente(p).RelacionTipo         'As Integer
                Get #FileNum, , EventoTarea(i).Prescedente(p).PosicionXYPersonal   'As Boolean
                Get #FileNum, , EventoTarea(i).Prescedente(p).VH                   'As String * 2
                Get #FileNum, , EventoTarea(i).Prescedente(p).Ruta                 'As String * 1
                Get #FileNum, , EventoTarea(i).Prescedente(p).X1                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).Y1                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).X2                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).Y2                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).X3                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).Y3                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).X4                   'As Single
                Get #FileNum, , EventoTarea(i).Prescedente(p).Y4                   'As Single
           Next p
        End If
        If EventoTarea(i).NroConsecuentes > 0 Then
           For C = 1 To EventoTarea(i).NroConsecuentes
                Get #FileNum, , EventoTarea(i).Consecuente(C).RelacionActiva       'As Boolean
                Get #FileNum, , EventoTarea(i).Consecuente(C).RelacionInfinita     'As Boolean
                Get #FileNum, , EventoTarea(i).Consecuente(C).TareaPrescedente     'As Integer
                Get #FileNum, , EventoTarea(i).Consecuente(C).NumeroRPrescedente   'As Integer
                Get #FileNum, , EventoTarea(i).Consecuente(C).TareaConsecuente     'As Integer
                Get #FileNum, , EventoTarea(i).Consecuente(C).NumeroRConsecuente   'As Integer
                Get #FileNum, , EventoTarea(i).Consecuente(C).NoElementoRelacion   'As Integer
                Get #FileNum, , EventoTarea(i).Consecuente(C).RelacionTipo         'As Integer
                Get #FileNum, , EventoTarea(i).Consecuente(C).PosicionXYPersonal   'As Boolean
                Get #FileNum, , EventoTarea(i).Consecuente(C).VH                   'As String * 2
                Get #FileNum, , EventoTarea(i).Consecuente(C).Ruta                 'As String * 1
                Get #FileNum, , EventoTarea(i).Consecuente(C).X1                   'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).Y1                   'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).X2                   'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).Y2                   'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).X3                   'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).Y3                   'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).X4                    'As Single
                Get #FileNum, , EventoTarea(i).Consecuente(C).Y4                   'As Single
           Next C
        End If
        Get #FileNum, , EventoPropiedades(i).Personalizada
        Get #FileNum, , EventoPropiedades(i).IdProceso
        Get #FileNum, , EventoPropiedades(i).IdTarea
        Get #FileNum, , EventoPropiedades(i).NoTarea
        Get #FileNum, , EventoPropiedades(i).Definicion
        Get #FileNum, , EventoPropiedades(i).TareaTipo
        Get #FileNum, , EventoPropiedades(i).Asunto
        Get #FileNum, , EventoPropiedades(i).Mensaje
        Get #FileNum, , EventoPropiedades(i).Para
        Get #FileNum, , EventoPropiedades(i).NoCondiciones
        Get #FileNum, , EventoPropiedades(i).DiasMaximo
        ReDim EventoPropiedades(i).Condicion(EventoPropiedades(i).NoCondiciones)
        For C = 1 To EventoPropiedades(i).NoCondiciones
            Get #FileNum, , EventoPropiedades(i).Condicion(C).CondicionActiva
            Get #FileNum, , EventoPropiedades(i).Condicion(C).NoCondicion
            Get #FileNum, , EventoPropiedades(i).Condicion(C).Definicion
            Get #FileNum, , EventoPropiedades(i).Condicion(C).Tipo
        Next C
  Next i
  
  For i = 1 To MaxRelaObj
      Get #FileNum, , Relacion(i).ConsecuentesOrigen
      Get #FileNum, , Relacion(i).PrescedentesDestin
      Get #FileNum, , Relacion(i).RelacionActiva
      Get #FileNum, , Relacion(i).TareaDestin
      Get #FileNum, , Relacion(i).TareaOrigen
  Next
  
  Get #FileNum, , da
  ReDim Preserve DatosProceso(da + 1)
  
  If da > 0 Then
     'DatosAplicMtx.RemoveRow 1
     
     For i = 1 To da
         Get #FileNum, , DatosProceso(i).Activo
         Get #FileNum, , DatosProceso(i).Campo
         Get #FileNum, , DatosProceso(i).Definicion
         Get #FileNum, , DatosProceso(i).Tipo
         Get #FileNum, , DatosProceso(i).Longitud
         Get #FileNum, , DatosProceso(i).VDefecto
         Get #FileNum, , DatosProceso(i).Clave
         On Error Resume Next

'         If i > 1 Then
         DatosAplicMtx.AddRow , , (i = 1)
'         Else
'            DatosAplicMtx.RowVisible(1) = True
'         End If
         
         With DatosAplicMtx
            
            .RowVisible(i) = DatosProceso(i).Activo
            .CellText(i, 1) = DatosProceso(i).Campo
            .CellText(i, 2) = DatosProceso(i).Definicion
            .CellText(i, 3) = DatosProceso(i).Tipo
            .CellText(i, 4) = DatosProceso(i).Longitud
            .CellText(i, 5) = DatosProceso(i).VDefecto
            .CellText(i, 6) = DatosProceso(i).Clave
            .CellForeColor(i, 1) = vbWindowText
            .CellForeColor(i, 2) = vbWindowText
            .CellForeColor(i, 3) = vbWindowText
            .CellForeColor(i, 4) = vbWindowText
            .CellForeColor(i, 5) = vbWindowText
        End With
        On Error GoTo 0
     Next
  End If
  i = da + 1
  On Error Resume Next
  
  With DatosAplicMtx
      .AddRow , , True
      .CellText(i, 1) = "(ninguno)"
      .CellText(i, 2) = "(ninguno)"
      .CellText(i, 3) = "Caracter"
      .CellText(i, 4) = "10"
      .CellText(i, 5) = "Nulo"
      .CellForeColor(i, 1) = vbButtonFace
      .CellForeColor(i, 2) = vbButtonFace
      .CellForeColor(i, 3) = vbButtonFace
      .CellForeColor(i, 4) = vbButtonFace
      .CellForeColor(i, 5) = vbButtonFace
  End With

  On Error GoTo 0

  Get #FileNum, , MaxNotaObj
  ReDim EventoNota(MaxNotaObj)
  If MaxNotaObj > 0 Then
     For i = 1 To MaxNotaObj
        Get #FileNum, , EventoNota(i).NotaActiva
        Get #FileNum, , EventoNota(i).IdNota
        Get #FileNum, , EventoNota(i).IdProceso
        Get #FileNum, , EventoNota(i).Nota
        Get #FileNum, , EventoNota(i).posx
        Get #FileNum, , EventoNota(i).posy
        Get #FileNum, , EventoNota(i).Titulo
        Get #FileNum, , EventoNota(i).Definicion
        Get #FileNum, , EventoNota(i).Fecha
     Next
  End If

  Get #FileNum, , da
  Dim listX As ListItem
  
  Dim K As String * 20
  Dim t As String * 50
  Dim Nicon%
  If da > 0 Then
     For i = 1 To da
        Get #FileNum, , K
        Get #FileNum, , t
        Nicon = 1
        Select Case UCase(Right(Trim(t), 3))
               Case "TXT", "DOC"
                    Nicon = 3
               Case "EXE", "COM"
                    Nicon = 5
               Case "PRC"
                    Nicon = 4
               Case "MSG"
                    Nicon = 2
               Case Else
                    Nicon = 1
        End Select
        
        Set listX = Adjuntos.ListItems.Add(i, Trim(K), Trim(t), Nicon)
     Next
  End If


  Close FileNum

  If Adjuntos.ListItems.Count > 0 Then
     HerramientasAdjuntos.Buttons.Item(3).Enabled = True
     HerramientasAdjuntos.Buttons.Item(4).Enabled = True
     HerramientasAdjuntos.Buttons.Item(6).Enabled = True
  End If


  'Limpia la pantalla y fuerza a pintar una nueva
  Pizarra.Cls
  Pizarra.Width = Plantilla.XMax
  Pizarra.Height = Plantilla.YMax
  VentanaMovil.Refresh
  On Error Resume Next
  
  For i = 1 To MaxProcObj
      If EventoTarea(i).ProcesoActivo Then
          DibujaTarea i, EventoTarea(i).posx, EventoTarea(i).posy, EventoTarea(i).Definicion, EventoTarea(i).TareaTipo
          If EventoTarea(i).NroPrescedentes > 0 Then
             For p = 1 To EventoTarea(i).NroPrescedentes
                 If EventoTarea(i).Prescedente(p).RelacionActiva Then
                    Load ObjRelP1(EventoTarea(i).Prescedente(p).NoElementoRelacion): Load ObjRelP4(EventoTarea(i).Prescedente(p).NoElementoRelacion)
                    Load ObjRelP2(EventoTarea(i).Prescedente(p).NoElementoRelacion): Load ObjRelP3(EventoTarea(i).Prescedente(p).NoElementoRelacion)
                    
                    Load ObjRelS1(EventoTarea(i).Prescedente(p).NoElementoRelacion): Load ObjRelS2(EventoTarea(i).Prescedente(p).NoElementoRelacion): Load ObjRelS3(EventoTarea(i).Prescedente(p).NoElementoRelacion)
                    Load ObjRelT1(EventoTarea(i).Prescedente(p).NoElementoRelacion)
                    
                    r = EventoTarea(i).Prescedente(p)
                    Call DibujaRelacion( _
                    EventoTarea(i).Prescedente(p).NoElementoRelacion, _
                    EventoTarea(i).Prescedente(p).TareaPrescedente, _
                    EventoTarea(i).Prescedente(p).NumeroRPrescedente, _
                    EventoTarea(i).Prescedente(p).TareaConsecuente, _
                    EventoTarea(i).Prescedente(p).NumeroRConsecuente)
                 End If
             Next p
          End If
          If EventoTarea(i).NroConsecuentes > 0 Then
             For C = 1 To EventoTarea(i).NroConsecuentes
                 If EventoTarea(i).Consecuente(C).RelacionActiva Then
                    Load ObjRelP1(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    Load ObjRelP4(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    Load ObjRelP2(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    Load ObjRelP3(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    
                    Load ObjRelS1(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    Load ObjRelS2(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    Load ObjRelS3(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    Load ObjRelT1(EventoTarea(i).Consecuente(C).NoElementoRelacion)
                    
                    r = EventoTarea(i).Consecuente(C)
                    Call DibujaRelacion( _
                    EventoTarea(i).Consecuente(C).NoElementoRelacion, _
                    EventoTarea(i).Consecuente(C).TareaPrescedente, _
                    EventoTarea(i).Consecuente(C).NumeroRConsecuente, _
                    EventoTarea(i).Consecuente(C).TareaConsecuente, _
                    EventoTarea(i).Consecuente(C).NumeroRPrescedente)
                 End If
             Next C
          End If
      End If
  Next i
  Pizarra.Refresh

  If MaxNotaObj > 0 Then
     For i = 1 To MaxNotaObj
         If EventoNota(i).NotaActiva Then
            DibujaNota i, EventoNota(i).posx, EventoNota(i).posy, EventoNota(i).Definicion
         End If
     Next
  End If
    

  Dim NActivosT%, NActivosR%
  
  For NActivosT = 1 To MaxProcObj
      If NActivosT > 1 And EventoTarea(NActivosT).ProcesoActivo Then Exit For
  Next
  For NActivosR = 1 To MaxRelaObj
      If NActivosR > 1 And Relacion(NActivosR).RelacionActiva Then Exit For
  Next
  
  
  Herramientas.Buttons.Item(10).Enabled = IIf(NActivosT <= MaxProcObj, True, False)
  Herramientas.Buttons.Item(11).Enabled = IIf(NActivosT <= MaxProcObj, True, False)
  Herramientas.Buttons.Item(12).Enabled = IIf(NActivosR <= MaxRelaObj, True, False)

  Herramientas.Buttons.Item(1).Value = tbrPressed
  HerramientaSeleccionada = 0
  
  NombreDeArchivo = Fname
  Flujo(FORM_NUM).Caption = GetFileName(NombreDeArchivo)
  mnuFlujos(FORM_NUM - 1).Caption = Flujo(FORM_NUM).Caption
  
  RequiereGrabar = False
  RequiereReproc = True
  
'  PizarraDatos(0).Contents(SF_TEXT) = ""
'  PizarraDatos(1).Contents(SF_TEXT) = ""
  PizarraDatos(0).Text = ""
  PizarraDatos(1).Text = ""
  
  Call MuestraProceso
  LeeArchivo = False
  
  Exit Function

leeError:
  LeeArchivo = True
  Exit Function

End Function


Private Function ArchivoValido(ByVal Fname$)
  Dim FileSize&
  Dim FigSize%
  Dim FileNum%
  Dim FirstStr As String * 10

  On Error GoTo ValidoError
  FileNum = FreeFile

  Open Fname For Binary As FileNum
  Get #FileNum, , FirstStr
  Close FileNum

  If FirstStr <> "WSWORKFLOW" Then
    ArchivoValido = False
    Exit Function
  End If
  ArchivoValido = True
  Exit Function

ValidoError:
  ArchivoValido = False
  Exit Function
End Function


Private Function GetFileName$(ByVal Path$)
  Dim Length%
  Dim p%, q%

  Length = Len(Path)
  q = 1
  p = InStr(Path, "\")

  While (p > 0)
    q = p + 1
    p = InStr(q, Path, "\")
  Wend

  GetFileName = Right$(Path, Length - q + 1)

End Function


Private Sub AbrirArchivo()
  Dim ret%
  ret = GrabarCambios("Abrir")
  If ret = IDCANCEL Then
    Exit Sub
  End If
  On Error GoTo abrirError
  CMDialog1.CancelError = True
  CMDialog1.Filter = "Todos los Archivos (*.*)|*.*|Archivos de Proceso (*.PRC)|*.prc"
  CMDialog1.FilterIndex = 2
  CMDialog1.Action = 1
  If LeeArchivo(CMDialog1.FileName) Then
    GoTo abrirError
  End If

  NombreDeArchivo = CMDialog1.FileName
  Me.Caption = "Proceso - " + GetFileName(NombreDeArchivo)
  Flujo(FORM_NUM).Caption = GetFileName(NombreDeArchivo)
  mnuFlujos(FORM_NUM - 1).Caption = Flujo(FORM_NUM).Caption
  
  Exit Sub

abrirError:
  Exit Sub

End Sub

Private Sub AbrirAdjunto()
  On Error GoTo abrirError
  Dim NombreAdjunto$
  CMDialog1.CancelError = True
  CMDialog1.Filter = "Todos los Archivos (*.*)|*.*|Archivos de Documento (*.DOC)|*.doc|Archivos de Hojas de Calculo (*.XLS)|*.xls"
  CMDialog1.FilterIndex = 1
  CMDialog1.Action = 1
  If LeerAdjunto(CMDialog1.FileName) Then
    GoTo abrirError
  End If

  NombreAdjunto = CMDialog1.FileName
  
  Dim listX As ListItem
  Dim X%
    
  X = Adjuntos.ListItems.Count + 1
  Dim Nicon%
  Nicon = 1
  Select Case UCase(Right(GetFileName(NombreAdjunto), 3))
         Case "TXT", "DOC"
              Nicon = 3
         Case "EXE", "COM"
              Nicon = 5
         Case "PRC"
              Nicon = 4
         Case "MSG"
              Nicon = 2
         Case Else
              Nicon = 1
  End Select
  Set listX = Adjuntos.ListItems.Add(X, "Adjunto" + Str(X), GetFileName(NombreAdjunto), Nicon)
  Exit Sub

abrirError:
  Exit Sub

End Sub

Sub AbrirAcceso()
  On Error GoTo abrirError
  Dim NombreAcceso$
  CMDialog1.CancelError = True
  CMDialog1.Filter = "Todos los Aplicativos (*.EXE)(*.COM)|*.exe;*.com|Archivos de Inicio (*.BAT)|*.bat"
  CMDialog1.FilterIndex = 1
  CMDialog1.Action = 1
  If LeerAcceso(CMDialog1.FileName) Then
    GoTo abrirError
  End If

  NombreAcceso = CMDialog1.FileName
  
  Dim listX As ListItem
  Dim X%
    
  X = Accesos.ListItems.Count + 1
  Dim Nicon%
  Nicon = 1
  Select Case UCase(Right(GetFileName(NombreAcceso), 3))
         Case "BAT"
              Nicon = 6
         Case "EXE", "COM"
              Nicon = 7
         Case Else
              Nicon = 1
  End Select
  Set listX = Accesos.ListItems.Add(X, "Acceso" + Str(X), GetFileName(NombreAcceso), Nicon)
  Exit Sub

abrirError:
  Exit Sub

End Sub

Private Sub LimpiarPizarra(Optional Crear As Boolean = True)
  
  
  Dim i%, Nt%, NActivos%
  RequiereGrabar = False
  RequiereReproc = True
  
  
  If MaxProcObj > 0 Then
     For i = MaxProcObj To 1 Step -1
       If EventoTarea(i).ProcesoActivo Then
          Unload ObjEvento(i)
          Unload ShapeStat(i)
       End If
     Next i
     ReDim EventoTarea(0)
     ReDim EventoPropiedades(0)
  End If
  
  If MaxRelaObj > 0 Then
     For i = MaxRelaObj To 1 Step -1
       If Relacion(i).RelacionActiva Then
          Unload ObjRelP1(i): Unload ObjRelP2(i): Unload ObjRelP3(i): Unload ObjRelP4(i)
          Unload ObjRelS1(i): Unload ObjRelS2(i): Unload ObjRelS3(i): Unload ObjRelT1(i)
       End If
     Next i
     ReDim Relacion(0)
  End If
  
  If MaxNotaObj > 0 Then
     For i = MaxNotaObj To 1 Step -1
       If EventoNota(i).NotaActiva Then
          Unload Notas(i)
          Unload DefNotas(i)
          Unload Note(i)
       End If
     Next i
     ReDim EventoNota(0)
  End If
  
  Dim da%
  da = Adjuntos.ListItems.Count
  If da > 0 Then
     For i = da To 1 Step -1
         Adjuntos.ListItems.Remove (i)
     Next
  End If
  
  HerramientasAdjuntos.Buttons.Item(3).Enabled = False
  HerramientasAdjuntos.Buttons.Item(4).Enabled = False
  HerramientasAdjuntos.Buttons.Item(6).Enabled = False
  
  MaxProcObj = 0: MaxRelaObj = 0: MaxNotaObj = 0
  
  NombreDeArchivo = ""
  Me.Caption = "Proceso - (Intitulado)"
  Flujo(FORM_NUM).Caption = "Intitulado"
  mnuFlujos(FORM_NUM - 1).Caption = Flujo(FORM_NUM).Caption
  
  Pizarra.Cls
  Pizarra.Height = IIf(VentanaMovil.Height * 15 < Pizarra.Height, VentanaMovil.Height * 15, Pizarra.Height)
  Pizarra.Width = IIf(VentanaMovil.Width * 15 < Pizarra.Width, VentanaMovil.Width * 15, Pizarra.Width)
  
  VentanaMovil.Refresh
  
  Call VerificaHerramientas
  Herramientas.Height = Me.ScaleHeight
  
  If HerramientaSeleccionada = 0 Then ShapeDel.Visible = False
  
  
  RAnterior = 0
  PAnterior = 0
  For Nt = 0 To 8: NoTareasTipo(Nt) = 0: Next
  
  If Crear Then
     TareaTipo = 7
     Call CreaTarea(100, 40, 7, Nt)
  End If

    
  On Error Resume Next
  ReDim DatosProceso(0)
  If Crear Then
     If DatosAplicMtx.Visible Then
       For i = DatosAplicMtx.Rows To 1 Step -1
           DatosAplicMtx.RemoveRow i
       Next i
       With DatosAplicMtx
         .AddRow , , (1 = 1)
         .CellText(1, 1) = "(ninguno)"
         .CellText(1, 2) = "(ninguno)"
         .CellText(1, 3) = "Caracter"
         .CellText(1, 4) = "10"
         .CellText(1, 5) = "Nulo"
         .CellForeColor(1, 1) = vbButtonFace
       End With
     End If
  Else
     If DatosAplicMtx.Visible Then
        For i = DatosAplicMtx.Rows To 1 Step -1
            DatosAplicMtx.RemoveRow i
        Next i
     End If
  End If
  
  RequiereGrabar = False
  RequiereReproc = True

End Sub


Private Sub DibujaTarea(Index As Integer, ByVal X As Double, ByVal Y As Double, Detalle As String, ByVal Tipo As Single, Optional BordeVisible As Boolean = False, Optional BordeColor As ColorConstants = vbWhite)
    
    Load ObjEvento(Index): Load ShapeStat(Index)
       
    If (EventoTarea(Index).NroConsecuentes = 0 Or EventoTarea(Index).Terminal) And BordeColor = vbWhite And EventoTarea(Index).TareaTipo > 0 Then
       BordeVisible = True
       BordeColor = vbCyan
    End If
    
    With ObjEvento(Index)
        .ZOrder 0
        .Picture = TareaImagen(Tipo)
        ' Posiciona y muestra el Control (Tarea)
        .Move X - (.Width / 2), Y - (.Height / 2), .Width, .Height
        .Visible = True
    End With
    
    ShapeStat(Index).Move ObjEvento(Index).Left - 2, ObjEvento(Index).Top - 2, ObjEvento(Index).Width + 4, ObjEvento(Index).Height + 4
    ShapeStat(Index).BorderColor = BordeColor
    ShapeStat(Index).Visible = BordeVisible
    
    ObjEvento(Index).Print
    ObjEvento(Index).Font = "Mirror"
    ObjEvento(Index).FontSize = 7
    ObjEvento(Index).Print Space(10) + Detalle
    
    
End Sub
    
Private Sub DibujaRelacion(Elm As Integer, Torigen As Integer, ROrigen As Integer, TDestino As Integer, RDestino As Integer, Optional Estado As Boolean = True, Optional RelacionColor As ColorConstants = vbBlack)
        
    If EventoTarea(Torigen).Consecuente(ROrigen).RelacionInfinita Then
       ObjRelS1(Elm).BorderStyle = 3: ObjRelS2(Elm).BorderStyle = 3: ObjRelS3(Elm).BorderStyle = 3
       ObjRelS1(Elm).BorderWidth = 1: ObjRelS2(Elm).BorderWidth = 1: ObjRelS3(Elm).BorderWidth = 1
    Else
       ObjRelS1(Elm).BorderStyle = 1: ObjRelS2(Elm).BorderStyle = 1: ObjRelS3(Elm).BorderStyle = 1
       ObjRelS1(Elm).BorderWidth = 2: ObjRelS2(Elm).BorderWidth = 2: ObjRelS3(Elm).BorderWidth = 2
    End If
    
    ObjRelS1(0).Visible = False: ObjRelS2(0).Visible = False: ObjRelS3(0).Visible = False
        
    'Elemento indicador de incio de Relacion
    ObjRelP1(Elm).Top = IIf(Mid(r.VH, 1, 1) = "D", r.Y1, IIf(Mid(r.VH, 1, 1) = "U", r.Y1 - ObjRelP1(Elm).Height, r.Y1 - 5))
    ObjRelP1(Elm).Left = IIf(Mid(r.VH, 1, 1) = "N", IIf(Mid(r.VH, 2, 1) = "R", r.X1, IIf(Mid(r.VH, 2, 1) = "L", r.X1 - ObjRelP1(Elm).Width, r.X1 - 5)), r.X1 - 5)
    ObjRelP1(Elm).Visible = True: ObjRelP1(Elm).ZOrder 0
    
    ObjRelS1(Elm).BorderColor = IIf(Estado, RelacionColor, &H8000000B)
    ObjRelS1(Elm).X1 = IIf(Mid(r.VH, 1, 1) = "N", IIf(Mid(r.VH, 2, 1) = "R", r.X1 + ObjRelP1(Elm).Width, IIf(Mid(r.VH, 2, 1) = "L", r.X1 - ObjRelP1(Elm).Width, r.X1 - ObjRelP1(Elm).Width)), r.X1)
    ObjRelS1(Elm).X2 = r.X2
    ObjRelS1(Elm).Y1 = IIf(Mid(r.VH, 1, 1) = "D", r.Y1 + ObjRelP1(Elm).Height, IIf(Mid(r.VH, 1, 1) = "U", r.Y1 - ObjRelP1(Elm).Height, r.Y1))
    ObjRelS1(Elm).Y2 = r.Y2: ObjRelS1(Elm).Visible = True
    
    ObjRelP2(Elm).Left = r.X2 - 5: ObjRelP2(Elm).Top = r.Y2 - 5
    ObjRelP2(Elm).Visible = IIf(ArrastrandoLinea, True, False)
    
    ObjRelS2(Elm).BorderColor = IIf(Estado, RelacionColor, &H8000000B)
    ObjRelS2(Elm).X1 = r.X2: ObjRelS2(Elm).X2 = r.X3
    ObjRelS2(Elm).Y1 = r.Y2: ObjRelS2(Elm).Y2 = r.Y3: ObjRelS2(Elm).Visible = True
    
    ObjRelT1(Elm).Caption = Trim(EventoPropiedades(Torigen).Condicion(EventoTarea(Torigen).Consecuente(ROrigen).RelacionTipo).Definicion) + IIf(EventoTarea(Torigen).Consecuente(ROrigen).RelacionInfinita, Chr(10) + "(Posible Flujo Circular)", "")
    ObjRelT1(Elm).Top = r.Y2 + IIf(r.Y1 < r.Y4, (Abs(r.Y2 - r.Y3) / 2) - (ObjRelT1(Elm).Height / 2) - 7, ((Abs(r.Y2 - r.Y3) / 2) - (ObjRelT1(Elm).Height / 2)) * -1)
    ObjRelT1(Elm).Left = r.X2 + IIf(r.X1 < r.X4, (Abs(r.X2 - r.X3) / 2) - (ObjRelT1(Elm).Width / 2), ((Abs(r.X2 - r.X3) / 2) + (ObjRelT1(Elm).Width / 2)) * -1)
    ObjRelT1(Elm).Visible = True: ObjRelT1(Elm).ZOrder 0
    
    ObjRelP3(Elm).Left = r.X3 - 5: ObjRelP3(Elm).Top = r.Y3 - 5
    ObjRelP3(Elm).Visible = IIf(ArrastrandoLinea, True, False)
    
    ObjRelS3(Elm).BorderColor = IIf(Estado, RelacionColor, &H8000000B)
    ObjRelS3(Elm).X1 = r.X3
    ObjRelS3(Elm).X2 = IIf(Mid(r.VH, 1, 1) = "N", IIf(Mid(r.VH, 2, 1) = "L", r.X4 + ObjRelP1(Elm).Width, IIf(Mid(r.VH, 2, 1) = "R", r.X4 - ObjRelP1(Elm).Width, r.X4 - ObjRelP1(Elm).Width)), r.X4)
    ObjRelS3(Elm).Y1 = r.Y3
    ObjRelS3(Elm).Y2 = IIf(Mid(r.VH, 1, 1) = "U", r.Y4 + ObjRelP1(Elm).Height, IIf(Mid(r.VH, 1, 1) = "D", r.Y4 - ObjRelP1(Elm).Height, r.Y4))
    ObjRelS3(Elm).Visible = True
    
    'Elemento indicador de Fin de Relacion
    ObjRelP4(Elm).Top = IIf(Mid(r.VH, 1, 1) = "U", r.Y4, IIf(Mid(r.VH, 1, 1) = "D", r.Y4 - ObjRelP1(Elm).Height, r.Y4 - 5))
    ObjRelP4(Elm).Left = IIf(Mid(r.VH, 1, 1) = "N", IIf(Mid(r.VH, 2, 1) = "L", r.X4, IIf(Mid(r.VH, 2, 1) = "R", r.X4 - ObjRelP1(Elm).Width, r.X4 - 5)), r.X4 - 5)
    ObjRelP4(Elm).Visible = True:     ObjRelP4(Elm).ZOrder 0
    
    EventoTarea(Torigen).Terminal = IIf(EventoTarea(Torigen).NroConsecuentes > 0, False, True)
    If ShapeStat(Torigen).Visible And ShapeStat(Torigen).BorderColor = vbCyan And Not EventoTarea(Torigen).Terminal Then
       ShapeStat(Torigen).BorderColor = vbWhite
       ShapeStat(Torigen).Visible = False
    End If
    
    EventoTarea(TDestino).Terminal = IIf(EventoTarea(TDestino).NroConsecuentes > 0, False, True)
    If ShapeStat(TDestino).Visible And ShapeStat(TDestino).BorderColor = vbCyan And Not EventoTarea(TDestino).Terminal Then
       ShapeStat(TDestino).BorderColor = vbWhite
       ShapeStat(TDestino).Visible = False
    End If
End Sub


Private Sub CalculaCoordenadasRelacion(ProcI As Control, ProcF As Object, Optional Nor As Integer = 0, Optional Nds As Integer = 0, Optional CreadPrev As Boolean = False, Optional NroRelacion As Integer = 0, Optional ODEST As Boolean = False)
    Dim Pi As ElementoCordenadas
    Dim PF As ElementoCordenadas
    
    With Pi
        .XCordLeft = ProcI.Left
        .XCordMed = ProcI.Left + (ProcI.Width / 2)
        .XCordRigth = ProcI.Left + ProcI.Width
        .YCordTop = ProcI.Top
        .YCordMed = ProcI.Top + (ProcI.Height / 2)
        .YCordBottom = ProcI.Top + ProcI.Height
    End With
        
    With PF
        .XCordLeft = ProcF.Left
        .XCordMed = ProcF.Left + (ProcF.Width / 2)
        .XCordRigth = ProcF.Left + ProcF.Width
        .YCordTop = ProcF.Top
        .YCordMed = ProcF.Top + (ProcF.Height / 2)
        .YCordBottom = ProcF.Top + ProcF.Height
    End With
    If PF.XCordLeft = Pi.XCordLeft And PF.YCordTop = Pi.YCordTop Then
       EsPosibleRelacion = False
       Exit Sub
    Else
       EsPosibleRelacion = True
    End If
    With r
        .Ruta = "A"
'        .VH = IIf(Pi.YCordTop > PF.YCordBottom, "U", IIf(Pi.YCordBottom < PF.YCordTop, "D", "N")) + _
'              IIf(Pi.XCordLeft > PF.XCordRigth, "L", IIf(Pi.XCordRigth < PF.XCordLeft, "R", "N"))
        .VH = IIf(Pi.YCordTop > PF.YCordBottom, "U", IIf(Pi.YCordBottom < PF.YCordTop, "D", "N")) + _
              IIf(Pi.XCordMed > PF.XCordMed, "L", IIf(Pi.XCordMed < PF.XCordMed, "R", "N"))
        
        Select Case .VH
            Case "NL", "NR"
                 .Ruta = "B"
                 .Y1 = Pi.YCordMed: .Y4 = PF.YCordMed
                 If .VH = "NL" Then
                    .X1 = Pi.XCordLeft: .X4 = PF.XCordRigth
                 Else
                    .X1 = Pi.XCordRigth: .X4 = PF.XCordLeft
                 End If
            
            Case "UL", "UR"
                 Dim Co%
                 .X1 = Pi.XCordMed: .Y1 = Pi.YCordTop
                 .X4 = PF.XCordMed: .Y4 = PF.YCordBottom
            
            Case "DL", "DR"
                 .X1 = Pi.XCordMed: .Y1 = Pi.YCordBottom
                 .X4 = PF.XCordMed: .Y4 = PF.YCordTop
            
            Case "UN", "DN"
                 .X1 = Pi.XCordMed: .X4 = PF.XCordMed
                 If .VH = "UN" Then
                    .Y1 = Pi.YCordTop: .Y4 = PF.YCordBottom
                 Else
                    .Y1 = Pi.YCordBottom: .Y4 = PF.YCordTop
                 End If
        End Select
        
        Dim Rcalc%
        Dim t%
        Dim RN() As Integer
        Dim First As Boolean
        Dim Rf As Integer
        Dim RC As Integer
        
        Dim X1I%, Y1I%, X4I%, Y4I%
        
        If NroRelacion > 0 Then
            If ODEST Then
                RC = EventoTarea(Nds).Prescedente(NroRelacion).NumeroRPrescedente
            Else
                X1I = EventoTarea(Nor).Consecuente(NroRelacion).X1 - Pi.XCordLeft
                Y1I = EventoTarea(Nor).Consecuente(NroRelacion).Y1 - Pi.YCordTop
                RC = EventoTarea(Nor).Consecuente(NroRelacion).NumeroRConsecuente
            End If
        End If
        
        X1I = Pi.YCordTop
        
        
        Select Case .Ruta
            Case "B"
                .X2 = .X1 + (((.X1 - .X4) / 2) * -1)
                .Y2 = .Y1: .X3 = .X2: .Y3 = .Y4
            Case "A"
                .Y2 = .Y1 + (((.Y1 - .Y4) / 2) * -1)
                 
                 If Nds > 0 Then
                     If NroRelacion = 0 Then
                        If EventoTarea(Nor).NroConsecuentes > 0 Then
                           Rcalc = 1
                           For t = 1 To EventoTarea(Nor).NroConsecuentes
                               If EventoTarea(Nor).Consecuente(t).TareaConsecuente = Nds Then
                                  Rcalc = Rcalc + 1
                               End If
                           Next
                           If Rcalc > 1 Then
                              If Rcalc Mod 2 = 0 Then
                                 .X1 = EventoTarea(Nor).Consecuente(Rcalc - IIf(Rcalc > 2, 2, 1)).X1 + 10
                                 .Y2 = EventoTarea(Nor).Consecuente(Rcalc - IIf(Rcalc > 2, 2, 1)).Y2 - 10
                              Else
                                 .X1 = EventoTarea(Nor).Consecuente(Rcalc - 2).X1 - 10
                                 .Y2 = EventoTarea(Nor).Consecuente(Rcalc - 2).Y2 + 10
                              End If
                           End If
                        End If
                     ElseIf EventoTarea(Nor).NroConsecuentes > 0 Then
                            Rcalc = 0
                            First = True
                            For t = 1 To EventoTarea(Nor).NroConsecuentes
                                If EventoTarea(Nor).Consecuente(t).TareaConsecuente = Nds Then
                                   If EventoTarea(Nor).Consecuente(t).RelacionActiva Then
                                      Rcalc = Rcalc + 1
                                      If First Then Rf = t
                                      First = False
                                      ReDim Preserve RN(Rcalc)
                                      RN(Rcalc) = t
                                   End If
                                End If
                            Next
                            For Co = 1 To Rcalc
                                If EventoTarea(Nor).Consecuente(Co).RelacionActiva Then
                                   If RN(Co) <> Rf Then
                                      If ODEST Then
                                         RC = EventoTarea(Nds).Prescedente(NroRelacion).NumeroRPrescedente
                                      Else
                                         RC = NroRelacion
                                      End If
                                      If RN(Co) = RC Then
                                         If Co Mod 2 = 0 Then
                                           .X1 = EventoTarea(Nor).Consecuente(Co - IIf(Co > 2, 2, 1)).X1 + 10
                                           .Y2 = EventoTarea(Nor).Consecuente(Co - IIf(Co > 2, 2, 1)).Y2 - 10
                                         Else
                                           .X1 = EventoTarea(Nor).Consecuente(Co - IIf(Co = 1, 0, 2)).X1 - 10
                                           .Y2 = EventoTarea(Nor).Consecuente(Co - IIf(Co = 1, 0, 2)).Y2 + 10
                                         End If
                                      End If
                                   End If
                                End If
                            Next
                     End If
                     ReDim RN(0)
                     
                     If NroRelacion = 0 Then
                        If EventoTarea(Nds).NroPrescedentes > 0 Then
                           Rcalc = 1
                           For t = 1 To EventoTarea(Nds).NroPrescedentes
                               If EventoTarea(Nds).Prescedente(t).TareaPrescedente = Nor Then
                                  Rcalc = Rcalc + 1
                               End If
                           Next
                           If Rcalc > 1 Then
                              If Rcalc Mod 2 = 0 Then
                                 If Left(.VH, 1) = "D" Then
                                    If Right(.VH, 1) = "R" Then
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).X4 + 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).Y2 - 10
                                    Else
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).X4 + 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).Y2 + 10
                                    End If
                                 Else
                                    If Right(.VH, 1) = "R" Then
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).X4 + 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).Y2 + 10
                                    Else
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).X4 + 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - IIf(Rcalc > 2, 2, 1)).Y2 - 10
                                    End If
                                 End If
                              Else
                                 If Left(.VH, 1) = "D" Then
                                    If Right(.VH, 1) = "R" Then
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - 2).X4 - 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - 2).Y2 + 10
                                    Else
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - 2).X4 - 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - 2).Y2 - 10
                                    End If
                                 Else
                                    If Right(.VH, 1) = "R" Then
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - 2).X4 - 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - 2).Y2 - 10
                                    Else
                                       .X4 = EventoTarea(Nds).Prescedente(Rcalc - 2).X4 - 10
                                       .Y2 = EventoTarea(Nds).Prescedente(Rcalc - 2).Y2 + 10
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     ElseIf EventoTarea(Nds).NroPrescedentes > 0 Then
                            Rcalc = 0
                            First = True
                            For t = 1 To EventoTarea(Nds).NroPrescedentes
                                If EventoTarea(Nds).Prescedente(t).TareaPrescedente = Nor Then
                                   If EventoTarea(Nds).Prescedente(t).RelacionActiva Then
                                      Rcalc = Rcalc + 1
                                      If First Then Rf = t
                                      First = False
                                      ReDim Preserve RN(Rcalc)
                                      RN(Rcalc) = t
                                   End If
                                End If
                            Next
                            For Co = 1 To Rcalc
                                If EventoTarea(Nds).Prescedente(Co).RelacionActiva Then
                                   If RN(Co) <> Rf Then
                                      If ODEST Then
                                         RC = NroRelacion
                                      Else
                                         RC = EventoTarea(Nor).Consecuente(NroRelacion).NumeroRConsecuente
                                      End If
                                      
                                      If RN(Co) = RC Then
                                         If Co Mod 2 = 0 Then
                                            If Left(.VH, 1) = "D" Then
                                               If Right(.VH, 1) = "R" Then
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).X4 + 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).Y2 - 10
                                               Else
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).X4 + 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).Y2 + 10
                                               End If
                                            Else
                                               If Right(.VH, 1) = "R" Then
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).X4 + 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).Y2 + 10
                                               Else
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).X4 + 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co > 2, 2, 1)).Y2 - 10
                                               End If
                                            End If
                                         Else
                                            If Left(.VH, 1) = "D" Then
                                               If Right(.VH, 1) = "R" Then
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).X4 - 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).Y2 + 10
                                               Else
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).X4 - 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).Y2 - 10
                                               End If
                                            Else
                                               If Right(.VH, 1) = "R" Then
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).X4 - 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).Y2 - 10
                                               Else
                                                  .X4 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).X4 - 10
                                                  .Y2 = EventoTarea(Nds).Prescedente(Co - IIf(Co = 1, 0, 2)).Y2 + 10
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                            Next
                     End If
                 End If
                .X2 = .X1: .X3 = .X4: .Y3 = .Y2
        
        
        
        End Select
    End With

End Sub


Private Function DetectaRelacion(X, Y, ByRef CtrlNum)

    DetectaRelacion = -1
    If MaxRelaObj = 0 Then Exit Function
    
    Dim i%, XP1%, YP1%, XP2%, YP2%, XP3%, YP3%, XP4%, YP4%
    Dim RAxI%, RAxF%, RAyI%, RAyF%
    Dim RBxI%, RBxF%, RByI%, RByF%
    Dim RCxI%, RCxF%, RCyI%, RCyF%

    CtrlNum = 0

    For i = 1 To MaxRelaObj
      If Relacion(i).RelacionActiva Then
        XP1 = ObjRelP2(i).Left: YP1 = ObjRelP2(i).Top
        XP2 = ObjRelP3(i).Left: YP2 = ObjRelP3(i).Top
        XP3 = ObjRelP1(i).Left: YP3 = ObjRelP1(i).Top
        XP4 = ObjRelP4(i).Left: YP4 = ObjRelP4(i).Top
        RAyI = IIf(ObjRelS1(i).Y1 = ObjRelS1(i).Y2, ObjRelS1(i).Y1 - 2, IIf(ObjRelS1(i).Y1 > ObjRelS1(i).Y2, ObjRelS1(i).Y2 + 4, ObjRelS1(i).Y1 + 4))
        RAyF = IIf(ObjRelS1(i).Y1 = ObjRelS1(i).Y2, ObjRelS1(i).Y1 + 2, IIf(ObjRelS1(i).Y1 > ObjRelS1(i).Y2, ObjRelS1(i).Y1 + 4, ObjRelS1(i).Y2 + 4))
        RAxI = IIf(ObjRelS1(i).X1 = ObjRelS1(i).X2, ObjRelS1(i).X1 - 2, IIf(ObjRelS1(i).X1 > ObjRelS1(i).X2, ObjRelS1(i).X2 + 4, ObjRelS1(i).X1 + 4))
        RAxF = IIf(ObjRelS1(i).X1 = ObjRelS1(i).X2, ObjRelS1(i).X1 + 2, IIf(ObjRelS1(i).X1 > ObjRelS1(i).X2, ObjRelS1(i).X1 + 4, ObjRelS1(i).X2 + 4))
        RByI = IIf(ObjRelS2(i).Y1 = ObjRelS2(i).Y2, ObjRelS2(i).Y1 - 2, IIf(ObjRelS2(i).Y1 > ObjRelS2(i).Y2, ObjRelS2(i).Y2 + 4, ObjRelS2(i).Y1 + 4))
        RByF = IIf(ObjRelS2(i).Y1 = ObjRelS2(i).Y2, ObjRelS2(i).Y1 + 2, IIf(ObjRelS2(i).Y1 > ObjRelS2(i).Y2, ObjRelS2(i).Y1 + 4, ObjRelS2(i).Y2 + 4))
        RBxI = IIf(ObjRelS2(i).X1 = ObjRelS2(i).X2, ObjRelS2(i).X1 - 2, IIf(ObjRelS2(i).X1 > ObjRelS2(i).X2, ObjRelS2(i).X2 + 4, ObjRelS2(i).X1 + 4))
        RBxF = IIf(ObjRelS2(i).X1 = ObjRelS2(i).X2, ObjRelS2(i).X1 + 2, IIf(ObjRelS2(i).X1 > ObjRelS2(i).X2, ObjRelS2(i).X1 + 4, ObjRelS2(i).X2 + 4))
        RCyI = IIf(ObjRelS3(i).Y1 = ObjRelS3(i).Y2, ObjRelS3(i).Y1 - 2, IIf(ObjRelS3(i).Y1 > ObjRelS3(i).Y2, ObjRelS3(i).Y2 + 4, ObjRelS3(i).Y1 + 4))
        RCyF = IIf(ObjRelS3(i).Y1 = ObjRelS3(i).Y2, ObjRelS3(i).Y1 + 2, IIf(ObjRelS3(i).Y1 > ObjRelS3(i).Y2, ObjRelS3(i).Y1 + 4, ObjRelS3(i).Y2 + 4))
        RCxI = IIf(ObjRelS3(i).X1 = ObjRelS3(i).X2, ObjRelS3(i).X1 - 2, IIf(ObjRelS3(i).X1 > ObjRelS3(i).X2, ObjRelS3(i).X2 + 4, ObjRelS3(i).X1 + 4))
        RCxF = IIf(ObjRelS3(i).X1 = ObjRelS3(i).X2, ObjRelS3(i).X1 + 2, IIf(ObjRelS3(i).X1 > ObjRelS3(i).X2, ObjRelS3(i).X1 + 4, ObjRelS3(i).X2 + 4))

        If (Abs(XP1 - X) < 9 And Abs(YP1 - Y) < 9) Then
           If ObjRelP2(i).Visible Then
              DetectaRelacion = i
              CtrlNum = 1
              Exit Function
           End If
        End If
        If (Abs(XP2 - X) < 9 And Abs(YP2 - Y) < 9) Then
           If ObjRelP2(i).Visible Then
              DetectaRelacion = i
              CtrlNum = 2
              Exit Function
           End If
        End If
        If (Abs(XP3 - X) < 9 And Abs(YP3 - Y) < 9) Then
           If ObjRelP2(i).Visible Then
              DetectaRelacion = i
              CtrlNum = 3
              Exit Function
           End If
        End If
        If (Abs(XP4 - X) < 9 And Abs(YP4 - Y) < 9) Then
           If ObjRelP2(i).Visible Then
              DetectaRelacion = i
              CtrlNum = 4
              Exit Function
           End If
        End If
        If (X >= RAxI And X <= RAxF) And (Y >= RAyI And Y <= RAyF) Then
            DetectaRelacion = i
            CtrlNum = 5
            Exit Function
        End If
        If (X >= RBxI And X <= RBxF) And (Y >= RByI And Y <= RByF) Then
            DetectaRelacion = i
            CtrlNum = 6
            Exit Function
        End If
        If (X >= RCxI And X <= RCxF) And (Y >= RCyI And Y <= RCyF) Then
            DetectaRelacion = i
            CtrlNum = 7
            Exit Function
        End If
      End If

    Next i
    DetectaRelacion = -1
End Function



Sub EliminaTarea(NoProc As Integer)
    If NoProc = 1 Then Exit Sub
    Dim PO%, pd%, Rsp%, Opt%, NActivos%, i%, Rt%
    Rt = 0
    If EventoTarea(NoProc).NroConsecuentes = 1 And EventoTarea(NoProc).NroConsecuentes = EventoTarea(NoProc).NroPrescedentes Then
       PO% = EventoTarea(NoProc).Prescedente(1).TareaPrescedente
       pd% = EventoTarea(NoProc).Consecuente(1).TareaConsecuente
       Rsp = MsgBox("Existe solo una relacion continua, Desea conservarla?", vbYesNoCancel + vbQuestion, "Eliminar Tarea")
       Rt = EventoTarea(NoProc).Prescedente(1).RelacionTipo
       Opt = 1
    ElseIf EventoTarea(NoProc).NroConsecuentes > 1 Or EventoTarea(NoProc).NroPrescedentes > 1 Then
       Rsp = MsgBox("Existen varias relaciones ya establecidas, las cuales seran eliminadas." + Chr(10) + Chr(13) + "Esta operacion podria afectar a otras tareas del proceso" + Chr(10) + Chr(13) + "Esta seguro de querer eliminar esta tarea y sus relaciones?", vbYesNo + vbQuestion, "Eliminar Tarea")
       Opt = 2
    ElseIf EventoPropiedades(NoProc).Personalizada Then
       Rsp = MsgBox("Eliminara toda definicion y propiedades conferidas a esta tarea." + Chr(10) + Chr(13) + "Esta operacion podria 'NO' afectar a otras tareas del proceso" + Chr(10) + Chr(13) + "Esta seguro de querer eliminar esta tarea?", vbYesNo + vbQuestion, "Eliminar Tarea")
       Opt = 2
    End If
    If Rsp = 2 Or (Rsp = 7 And Opt <> 1) Then Exit Sub
    
    If MaxRelaObj > 0 Then
        Call EliminaRelaciones(NoProc)
    End If
    
    If MaxProcObj > 0 Then
        Unload ObjEvento(NoProc)
        Unload ShapeStat(NoProc)
        NoTareasTipo(EventoTarea(NoProc).TareaTipo) = NoTareasTipo(EventoTarea(NoProc).TareaTipo) - 1
        EventoTarea(NoProc).ProcesoActivo = False
    End If
    
'    For NActivos = 1 To MaxProcObj
'        If NActivos > 1 And EventoTarea(NActivos).ProcesoActivo Then Exit For
'    Next
'    Herramientas.Buttons.Item(9).Enabled = IIf(NActivos <= MaxProcObj, True, False)
'    Herramientas.Buttons.Item(10).Enabled = IIf(NActivos <= MaxProcObj, True, False)
'    Herramientas.Buttons.Item(1).Value = IIf(NActivos <= MaxProcObj, tbrUnpressed, tbrPressed)
    
'    HerramientaSeleccionada = IIf(NActivos <= MaxProcObj, 3, 0)
    Call VerificaHerramientas
    If HerramientaSeleccionada = 0 Then ShapeDel.Visible = False
    If Rsp = 6 And Opt = 1 Then Call CreaRelacion(PO, pd, Rt)
    RequiereGrabar = True
    RequiereReproc = True
    
    Call MuestraProceso

End Sub

Sub EliminaRelacion(Obj As Integer)
    Dim pd%, Rd%, MxPred%, MxCond%, I2%
    Dim PE%, RE%, PD2%, RD2%, Elm%, PO%, Ro%, NActivos%
            
    PO = Relacion(Obj).TareaOrigen
    Ro = Relacion(Obj).ConsecuentesOrigen
    MxCond = EventoTarea(PO).NroConsecuentes
    If Ro < MxCond Then
       For I2 = Ro + 1 To MxCond
           EventoTarea(PO).Consecuente(I2 - 1) = EventoTarea(PO).Consecuente(I2)
           PD2 = EventoTarea(PO).Consecuente(I2 - 1).TareaConsecuente
           RD2 = EventoTarea(PO).Consecuente(I2 - 1).NumeroRConsecuente
           Elm = EventoTarea(PO).Consecuente(I2 - 1).NoElementoRelacion
           If EventoTarea(PD2).NroPrescedentes > 0 Then
              EventoTarea(PD2).Prescedente(RD2).NumeroRPrescedente = I2 - 1
           End If
           Relacion(Elm).ConsecuentesOrigen = RD2
       Next
    End If
    EventoTarea(PO).NroConsecuentes = EventoTarea(PO).NroConsecuentes - 1
    ReDim Preserve EventoTarea(PO).Consecuente(EventoTarea(PO).NroConsecuentes)
    
    EventoTarea(PO).Terminal = IIf(EventoTarea(PO).NroConsecuentes > 0, False, True)
    If Not ShapeStat(PO).Visible And ShapeStat(PO).BorderColor = vbWhite And EventoTarea(PO).Terminal Then
       ShapeStat(PO).BorderColor = vbCyan
       ShapeStat(PO).Visible = True
    End If
            
    pd = Relacion(Obj).TareaDestin
    Rd = Relacion(Obj).PrescedentesDestin
    MxPred = EventoTarea(pd).NroPrescedentes
    If Rd < MxPred Then
       For I2 = Rd + 1 To MxPred
           EventoTarea(pd).Prescedente(I2 - 1) = EventoTarea(pd).Prescedente(I2)
           PD2 = EventoTarea(pd).Prescedente(I2 - 1).TareaPrescedente
           RD2 = EventoTarea(pd).Prescedente(I2 - 1).NumeroRPrescedente
           Elm = EventoTarea(pd).Prescedente(I2 - 1).NoElementoRelacion
           If EventoTarea(PD2).NroConsecuentes > 0 Then
              EventoTarea(PD2).Consecuente(RD2).NumeroRConsecuente = I2 - 1
           End If
           Relacion(Elm).PrescedentesDestin = RD2
       Next
    End If
    EventoTarea(pd).NroPrescedentes = EventoTarea(pd).NroPrescedentes - 1
    ReDim Preserve EventoTarea(PO).Prescedente(EventoTarea(PO).NroPrescedentes)
    
    EventoTarea(pd).Terminal = IIf(EventoTarea(pd).NroConsecuentes > 0, False, True)
    If Not ShapeStat(pd).Visible And ShapeStat(pd).BorderColor = vbWhite And EventoTarea(pd).Terminal Then
       ShapeStat(pd).BorderColor = vbCyan
       ShapeStat(pd).Visible = True
    End If
        
    Relacion(Obj).RelacionActiva = False
    Unload ObjRelP1(Obj): Unload ObjRelP4(Obj)
    Unload ObjRelP2(Obj): Unload ObjRelP3(Obj)
    
    Unload ObjRelS1(Obj): Unload ObjRelS2(Obj): Unload ObjRelS3(Obj)
    Unload ObjRelT1(Obj)

    Call VerificaHerramientas
End Sub

Sub EliminaRelaciones(Torigen As Integer)

        
    Dim I1%, I2%
    Dim MxPrescedentes As Integer, MxConsecuentes As Integer
    Dim pd%, Rd%, MxPred%, MxCond%, PD2%, RD2%, NE%
    Dim NRC%, NRP%, OBjM%, OBjO%, OBjD%, Elm%
    Dim NRCD%, NRPD%
    
    MxConsecuentes = EventoTarea(Torigen).NroConsecuentes
    MxPrescedentes = EventoTarea(Torigen).NroPrescedentes
        
    Do While MxConsecuentes > 0
       pd = EventoTarea(Torigen).Consecuente(1).TareaConsecuente
       Rd = EventoTarea(Torigen).Consecuente(1).NumeroRConsecuente
       NE = EventoTarea(Torigen).Consecuente(1).NoElementoRelacion
       Call EliminaRelacion(NE)
       MxConsecuentes = EventoTarea(Torigen).NroConsecuentes
    Loop
    
    Do While MxPrescedentes > 0
       pd = EventoTarea(Torigen).Prescedente(1).TareaPrescedente
       Rd = EventoTarea(Torigen).Prescedente(1).NumeroRPrescedente
       NE = EventoTarea(Torigen).Prescedente(1).NoElementoRelacion
       Call EliminaRelacion(NE)
       MxPrescedentes = EventoTarea(Torigen).NroPrescedentes
    Loop
    

End Sub


Sub RemarcaRelaciones(Obj As Integer, Status As Integer)
    Dim MxPdrsD As Integer, MxHijsO As Integer
    Dim NR%, OBjO%, OBjD%, NElm%
    If PAnterior = Obj Then Exit Sub
    MxHijsO = EventoTarea(Obj).NroConsecuentes
    MxPdrsD = EventoTarea(Obj).NroPrescedentes
    
    
    If MxHijsO > 0 Then
        For NR = 1 To MxHijsO
            NElm = EventoTarea(Obj).Consecuente(NR).NoElementoRelacion
            ObjRelS1(NElm).BorderColor = IIf(Status = 1, &H8000000B, &H80000008)
            ObjRelS2(NElm).BorderColor = IIf(Status = 1, &H8000000B, &H80000008)
            ObjRelS3(NElm).BorderColor = IIf(Status = 1, &H8000000B, &H80000008)
            ObjRelS1(NElm).ZOrder 0
            ObjRelS2(NElm).ZOrder 0
            ObjRelS3(NElm).ZOrder 0
        Next NR
    End If
    If MxPdrsD > 0 Then
        For NR = 1 To MxPdrsD
            NElm = EventoTarea(Obj).Prescedente(NR).NoElementoRelacion
            ObjRelS1(NElm).BorderColor = IIf(Status = 1, &H8000000B, &H80000008)
            ObjRelS2(NElm).BorderColor = IIf(Status = 1, &H8000000B, &H80000008)
            ObjRelS3(NElm).BorderColor = IIf(Status = 1, &H8000000B, &H80000008)
            ObjRelS1(NElm).ZOrder 0
            ObjRelS2(NElm).ZOrder 0
            ObjRelS3(NElm).ZOrder 0
        Next NR
    End If
End Sub

Sub OcultaElementos()
     ObjRelS1(0).Visible = False: ObjRelS2(0).Visible = False: ObjRelS3(0).Visible = False
     ObjEvento(ObjEnMovimiento).Drag 0: ObjEvento(ObjEnMovimiento).DragMode = 0: ObjEvento(ObjEnMovimiento).Visible = True
     ShapeDel.Visible = False
     ShapeMov.Visible = False
End Sub

Sub MuestraProceso(Optional BloqMnu As Boolean = True)
    Dim Temp%
    Temp = Me.MousePointer
    Me.MousePointer = 11
    If BloqMnu Then
       Prv1.Item(0).Enabled = False
       Prv1.Item(1).Enabled = False
       prn1.Item(1).Enabled = False
       prn1.Item(2).Enabled = False
       If RequiereReproc Then
          RequiereReproc1 = True
          RequiereReproc2 = True
       End If
    End If
    
    If (ContenedorDatos = 4 And Not PDCancel) Or Not ContenedorDatos.Visible Then
       ContenedorDatos.Tab = 3
       PizarraDatos(0).ZOrder 0
    End If
    PDCancel = False
'    If PizarraDatos(0).Contents(SF_TEXT) = "" Or RequiereReproc Then
    If PizarraDatos(0).Text = "" Or RequiereReproc1 Then
       Call MuestraErroresProceso
    End If
    If ContenedorDatos.Visible Then
'       If PizarraDatos(1).Contents(SF_TEXT) = "" Or RequiereReproc Then
       If PizarraDatos(1).Text = "" Or RequiereReproc2 Then
           Call MuestraSecuenciaProceso
       End If
    End If
    If RequiereReproc And Not RequiereReproc1 And Not RequiereReproc2 Then
       RequiereReproc = False
    End If
    Me.MousePointer = Temp
End Sub

Sub MuestraErroresProceso()
    If ContenedorDatos.Tab <> 3 Then Exit Sub
    Prv1.Item(0).Enabled = True
    prn1.Item(1).Enabled = True
    Dim I1%, I2%, I3%, txt$, Inn$, Tn$, Te$, TNE%
    Dim Wrn%, Rec%
    Rec = 0
    Wrn = 0
    txt = ""
    Tn = ""
    Te = ""
    For I1 = 1 To MaxProcObj
        Te = ""
        TNE = 0
        If EventoTarea(I1).ProcesoActivo Then
           Tn = WRtf("t", "      ", True)
           Tn = Tn + WRtf("t", "   •  ", False, True, , , 1) + WRtf("t", UCase(Trim(EventoTarea(I1).Definicion)), , True, , True, , , 18)
           Tn = Tn + WRtf("t", IIf(EventoTarea(I1).Terminal, "   -  (Tarea Terminal)", ""), True, True, True)

           If Mid(Trim(EventoTarea(I1).Definicion), 1, Len(DefTareaTipo(EventoTarea(I1).TareaTipo + 1))) = DefTareaTipo(EventoTarea(I1).TareaTipo + 1) Then
              Rec = Rec + 1
              TNE = TNE + 1
              Te = Te + WRtf("t", "      ")
              Te = Te + WRtf("t", "Recuerde" + vbTab, , True, False, False, 1)
              Te = Te + WRtf("t", ":  Especificar el titulo de la tarea en la ventana de Propiedades.", True, False, True, False, 1)
           End If
           If I1 > 1 And EventoTarea(I1).NroPrescedentes = 0 Then
              Rec = Rec + 1
              TNE = TNE + 1
              Te = Te + WRtf("t", "      ")
              Te = Te + WRtf("t", "Recuerde" + vbTab, , True, False, False, 1)
              Te = Te + WRtf("t", ":  Este paso no tiene relacionada ninguna tarea prescedente.", True, False, True, False, 1)
           End If
           
           If Len(Trim(EventoPropiedades(I1).Para)) = 0 Or Trim(EventoPropiedades(I1).Para) = "(ninguno)" Then
              Wrn = Wrn + 1
              TNE = TNE + 1
              Te = Te + WRtf("t", "      ")
              Te = Te + WRtf("t", "Advertencia" + vbTab, , True, False, False, 6)
              Te = Te + WRtf("t", ":  Debe indicar para que rol(es) sera definida esta tarea en la ventana de Propiedades.", True, False, True, False, 6)
'              Te = Te + WRtf("t", "   El no indicarlo causara que el proceso posea una tarea sin propietario.", True, False, True, False, 6)
           End If
           If Len(Trim(EventoPropiedades(I1).Asunto)) = 0 Or Trim(EventoPropiedades(I1).Asunto) = "(ninguno)" Then
              Rec = Rec + 1
              TNE = TNE + 1
              Te = Te + WRtf("t", "      ")
              Te = Te + WRtf("t", "Recuerde" + vbTab, , True, False, False, 1)
              Te = Te + WRtf("t", ":  Debe indicar el motivo o asunto de este tarea en la ventana de Propiedades.", True, False, True, False, 1)
           End If
           If Len(Trim(EventoPropiedades(I1).Mensaje)) = 0 Or Trim(EventoPropiedades(I1).Mensaje) = "(ninguno)" Then
              Rec = Rec + 1
              TNE = TNE + 1
              Te = Te + WRtf("t", "      ")
              Te = Te + WRtf("t", "Recuerde" + vbTab, , True, False, False, 1)
              Te = Te + WRtf("t", ":  Debe indicar el mensaje a proporcionar con esta tarea en la ventana de Propiedades.", True, False, True, False, 1)
           End If
           
           
           Select Case EventoTarea(I1).TareaTipo
                  Case 7, 5, 4, 3, 1
                       If EventoTarea(I1).NroConsecuentes = 0 Then
                          Wrn = Wrn + 1
                          TNE = TNE + 1
                          Te = Te + WRtf("t", "      ")
                          Te = Te + WRtf("t", "Advertencia" + vbTab, , True, False, False, 6)
                          Te = Te + WRtf("t", ":  Este paso no tiene relacionada ninguna tarea de consecuente.", True, False, True, False, 6)
                       End If
           End Select
           If EventoTarea(I1).NroConsecuentes > 0 And EventoPropiedades(I1).NoCondiciones > 1 Then
              Dim NOk As Boolean
              NOk = False
              For I2 = 2 To EventoPropiedades(I1).NoCondiciones
                  For I3 = 1 To EventoTarea(I1).NroConsecuentes
                      If EventoTarea(I1).Consecuente(I3).RelacionTipo = EventoPropiedades(I1).Condicion(I2).NoCondicion Then Exit For
                  Next
                  If I3 > EventoTarea(I1).NroConsecuentes Then
                     NOk = True: Exit For
                  End If
              Next
              If NOk Then
                 Wrn = Wrn + 1
                 TNE = TNE + 1
                 Te = Te + WRtf("t", "      ")
                 Te = Te + WRtf("t", "Advertencia" + vbTab, , True, False, False, 6)
                 Te = Te + WRtf("t", ":  Existen posibles respuestas sin relacion alguna.", True, False, True, False, 6)
              End If
           End If
           
        End If
        If TNE > 0 Then
           txt = txt + Tn + Te
        End If
    Next
    
    Te = WRtf("t", "", True)
    If MaxNotaObj > 0 Then
       Te = Te + WRtf("t", "ANOTACIONES", True, True, , True)
       For I1 = 1 To MaxNotaObj
           Te = Te + WRtf("t", "      ")
           Te = Te + WRtf("t", EventoNota(I1).Titulo + " -" + vbTab + EventoNota(I1).Fecha + vbTab, , True, False, False)
           Te = Te + WRtf("t", ":" + Trim(EventoNota(I1).Definicion), True, False)
       Next
    End If
    txt = txt + Te
    
    Inn = WRtf("i", "")
    Inn = Inn + WRtf("t", " Errores -   ", , True, , , , , 18)
    Inn = Inn + WRtf("t", Str(Wrn) + " Advertencia" + IIf(Wrn = 1, "", "s"), , True, , , IIf(Wrn > 0, 6, 1), , 18)
    Inn = Inn + WRtf("t", " , ", , True, , , , , 18)
    Inn = Inn + WRtf("t", Str(Rec) + " Recordatorio" + IIf(Rec = 1, "", "s"), , True, , , IIf(Rec > 0, 6, 1), , 18)
    Inn = Inn + WRtf("t", "    -   En " + Str(MaxProcObj) + IIf(MaxProcObj = 1, " Tarea Definida", " Tareas Definidas"), True, True, , , , , 18)
    
    txt = Inn + txt
    
    txt = txt + WRtf("f", "")
    
    HerramientasMenu.Buttons.Item(12).Enabled = IIf(Wrn > 0, False, True)
    HerramientasMenu.Buttons.Item(13).Enabled = IIf(Wrn > 0, False, HerramientasMenu.Buttons.Item(15).Enabled)
    
'    PizarraDatos(0).Contents(SF_RTF) = txt
    PizarraDatos(0).TextRTF = txt
    PizarraDatos(0).ZOrder 0
    RequiereReproc1 = False
    
End Sub

Sub MuestraSecuenciaProceso()
    If ContenedorDatos.Tab <> 4 Then Exit Sub
    Prv1.Item(1).Enabled = True
    prn1.Item(2).Enabled = True
    
    Dim C%
    SecObj = WRtf("i", "")
    SecObj = SecObj & WRtf("t", "NOMBRE DEL PROCESO", , True, , True)
    If Len(GetFileName(NombreDeArchivo)) > 0 Then
       SecObj = SecObj & WRtf("t", ":    " + UCase(Mid(GetFileName(NombreDeArchivo), 1, InStr(1, GetFileName(NombreDeArchivo), ".") - 1)), True, True, True)
    Else
       SecObj = SecObj & WRtf("t", ":    INTITULADO", True, True, True)
    End If
    SecObj = SecObj & WRtf("t", "PROPIEDADES:", True, True, , True)
    SecObj = SecObj & WRtf("t", "    Fecha de creación:", True)
    SecObj = SecObj & WRtf("t", "    Fecha de publicación:", True)
    SecObj = SecObj & WRtf("t", "    Autor del proceso:", True)
    SecObj = SecObj & WRtf("t", "", True)
    SecObj = SecObj & WRtf("t", "DATOS DEL PROCESO:", True, True, , True)
    If UBound(DatosProceso) > 1 Then
       For C = 1 To UBound(DatosProceso)
           If DatosProceso(C).Activo Then
              SecObj = SecObj & _
              WRtf("t", "  Campo: " + DatosProceso(C).Definicion + vbTab + "Tipo: " + DatosProceso(C).Tipo + vbTab + "Longitud: " + DatosProceso(C).Longitud + vbTab + " Valor Defecto: " + DatosProceso(C).VDefecto, True, , , , , , , 2)
           End If
       Next
    Else
       SecObj = SecObj & WRtf("t", "    (Ninguno Definido)", True)
    End If
    
    SecObj = SecObj & WRtf("t", "", True)
    SecObj = SecObj & WRtf("t", "", True)
    SecObj = SecObj & WRtf("t", "DEFINICIONES Y FLUJO DEL PROCESO", True, True, , True)
    
    NivelProc = 0
    Procesos = 0
    DiasProc = 0
    DiasMax = 0
    DiasMin = 10000
    ProcMax = 0
    ProcMin = 10000
    
    
    Call SecuenciaObjeto(1, 0, Complejo())
    SecObj = SecObj & WRtf("t", "", True)
    
    NivelProc = 1
    For C = 1 To MaxProcObj
        If EventoTarea(C).NroPrescedentes = 0 Then
           If EventoTarea(C).ProcesoActivo Then
              If (C = 1 And MaxProcObj = 1) Or (C > 1) Then
                    SecObj = SecObj + _
                    WRtf("t", Space(5 * NivelProc - 1) + "Nivel 0" + " -- ", , True, , , 6) + _
                    WRtf("t", Space(8) + "  Acciones de ", , True) + _
                    WRtf("t", UCase(Trim(EventoTarea(C).Definicion) + "        "), , True, True, , 9) + _
                    WRtf("t", "'Precaucion Esta Es Una Tarea " + IIf(EventoTarea(C).NroConsecuentes > 0, "Quebrada'", "Aislada'"), True, True, , True, 6)
                    If EventoTarea(C).NroConsecuentes > 0 Then
                       Call SecuenciaObjeto(C, 0, Complejo())
                    End If
              End If
           End If
        End If
    Next
    SecObj = SecObj & WRtf("t", "", True)
    SecObj = SecObj & WRtf("t", "ESCENARIO DE TIEMPOS DEL PROCESO", True, True, , True)
    SecObj = SecObj & WRtf("t", "    OPTIMISTA - " + Str(DiasMin) + "  DIAS.", , True, , , 2)
    SecObj = SecObj & WRtf("t", "          (Realizando el menor número de tareas) " + Str(ProcMin), True)
    SecObj = SecObj & WRtf("t", "    PESIMISTA - " + Str(DiasMax) + "  DIAS.", , True, , , 6)
    SecObj = SecObj & WRtf("t", "          (Realizando el mayor número de tareas) " + Str(ProcMax), True)
    
    
    SecObj = SecObj & WRtf("f", "")
'    PizarraDatos(1).Contents(SF_RTF) = SecObj
    PizarraDatos(1).TextRTF = SecObj
    PizarraDatos(1).ZOrder 0
    RequiereReproc2 = False

End Sub


Sub SecuenciaObjeto(Nobj As Integer, Optional Parar As Integer = 0, Optional Complx As Boolean = False)
    If Not EventoTarea(Nobj).ProcesoActivo Then Exit Sub
    Dim Tn%, Pa%, Nc%, C1%, Rt%, Rd$, RC%, DTt$, DTd$, C%
    NivelProc = NivelProc + 1
    Procesos = Procesos + 1
    DiasProc = DiasProc + EventoPropiedades(Nobj).DiasMaximo
    
    SecObj = SecObj + _
        WRtf("t", Space(8 * NivelProc - 1) + "Nivel " + Str(NivelProc) + " -- ", , True) + _
        WRtf("t", Space(8) + UCase(Trim(EventoTarea(Nobj).Definicion)), , True, True, , 9) + _
        WRtf("t", IIf(EventoTarea(Nobj).Terminal, "   -  (Tarea Terminal)", ""), True, True, True)

    
    Nc = EventoTarea(Nobj).NroConsecuentes
    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "PROPIEDADES:", True, True, , True)
    
    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "Tipo de acción a efectuarse con esta tarea :       ") + Space(4) + WRtf("t", DefTareaTipo(EventoTarea(Nobj).TareaTipo + 1), True, True)
    
    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "Tiempo necesario para completar esta tarea :  ") + Space(4) + WRtf("t", Str(EventoPropiedades(Nobj).DiasMaximo), True, True)
    
    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "Tarea definida para los siguientes Roles :          ") + Space(4)
    
    If Len(Trim(EventoPropiedades(Nobj).Para)) = 0 Or Trim(EventoPropiedades(Nobj).Para) = "(ninguno)" Then
        SecObj = SecObj + _
        WRtf("t", "(ninguno)", True, True)
    Else
        SecObj = SecObj + _
        WRtf("t", EventoPropiedades(Nobj).Para, True, True)
    End If

    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "Decisiones posibles a tomar en esta tarea :       ") + Space(4)
    If EventoPropiedades(Nobj).NoCondiciones > 1 Then
       For C = 2 To EventoPropiedades(Nobj).NoCondiciones
           SecObj = SecObj + _
           WRtf("t", Trim(EventoPropiedades(Nobj).Condicion(C).Definicion) + IIf(C < EventoPropiedades(Nobj).NoCondiciones, "; ", ""), False, True)
       Next
       SecObj = SecObj + _
       WRtf("t", "", True, True)
    Else
       SecObj = SecObj + _
       WRtf("t", "(ninguna)", True, True)
    End If

    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "ANALISIS DE ETAPA DEL PROCESO:", True, True, , True)

    SecObj = SecObj + _
    WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
    WRtf("t", "Tareas efectuadas hasta este momento :          ") + Space(4) + _
    WRtf("t", Str(Procesos), True, True)
    
    If Nc > 0 Then
       SecObj = SecObj + _
       WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
       WRtf("t", "Tiempo transcurrido al concluir con esta tarea :") + Space(4) + _
       WRtf("t", Str(DiasProc), True, True)
       
       
       SecObj = SecObj + _
       WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
       WRtf("t", "ACCIONES CONSECUENTES:", True, True, , True)
       For C = 1 To Nc
           If EventoTarea(Nobj).Consecuente(C).RelacionActiva Then
              RC = EventoTarea(Nobj).Consecuente(C).TareaConsecuente
              Rt = EventoTarea(Nobj).Consecuente(C).RelacionTipo
              Rd = EventoPropiedades(Nobj).Condicion(Rt).Definicion
              DTt = DefTareaTipo(EventoTarea(RC).TareaTipo + 1)
              DTd = EventoTarea(RC).Definicion
              SecObj = SecObj + _
                WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
                WRtf("t", "El proceso continuará con la Tarea de:  ") + _
                WRtf("t", UCase(DTt), , True, , , 12) + _
                WRtf("t", " : a   ") + _
                WRtf("t", Trim(UCase(DTd)), , True, True, True, 10) + _
                WRtf("t", "  Cuando la tarea de " + UCase(Trim(EventoTarea(Nobj).Definicion)) + " concluya con su función") + _
                IIf(Rt > 1, " en '" + WRtf("t", Trim(Rd), , True, True, True) + WRtf("t", "'.", True), WRtf("t", ".", True))
                If EventoTarea(Nobj).Consecuente(C).RelacionInfinita Then
                
                   SecObj = SecObj + _
                   WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
                   WRtf("t", "NOTA:", , True, , True) + WRtf("t", "  " + Trim(EventoTarea(Nobj).Definicion) + " al ejecutar la tarea de nivel a " + Trim(UCase(DTd)) + " podria producir un flujo circular o infinito en el Proceso.", True, True, True, , 6)
                End If
           End If
       Next
       If (Parar > 0 And Not Complx) Or Parar = 0 Then
          For C = 1 To Nc
              If EventoTarea(Nobj).Consecuente(C).RelacionActiva Then
                 RC = EventoTarea(Nobj).Consecuente(C).TareaConsecuente
                 If Nobj <> Parar Or Not EventoTarea(Nobj).Consecuente(C).RelacionInfinita Then
                    Call SecuenciaObjeto(RC, IIf(EventoTarea(Nobj).Consecuente(C).RelacionInfinita, Nobj, IIf(Parar > 0, Parar, 0)), Complx)
                 Else
                    If Parar = Nobj Then
                       SecObj = SecObj + _
                          WRtf("t", Space(8 * (NivelProc + 1)) + "Nivel " + Str(NivelProc + 1) + " -- ", , True) + _
                          WRtf("t", Space(8) + "" + UCase(Trim(EventoTarea(RC).Definicion)) + " : ", , True, True, , 9) + _
                          WRtf("t", "No puede continuar descendiendo de nivel ya que podria producirse un flujo circular o infinito en el Reporte", True, True, True, , 6)
                       DiasMax = IIf(DiasMax < DiasProc, DiasProc, DiasMax)
                       DiasMin = IIf(DiasMin > DiasProc, DiasProc, DiasMin)
                       ProcMax = IIf(ProcMax < Procesos, Procesos, ProcMax)
                       ProcMin = IIf(ProcMin > Procesos, Procesos, ProcMin)
                          
                       SecObj = SecObj + WRtf("t", "", True)
                    End If
                 End If
              End If
          Next
       End If
    Else
       DiasMax = IIf(DiasMax < DiasProc, DiasProc, DiasMax)
       DiasMin = IIf(DiasMin > DiasProc, DiasProc, DiasMin)
       ProcMax = IIf(ProcMax < Procesos, Procesos, ProcMax)
       ProcMin = IIf(ProcMin > Procesos, Procesos, ProcMin)
       
       
       SecObj = SecObj + _
       WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
       WRtf("t", "Tiempo transcurrido al concluir con esta tarea :    ") + _
       WRtf("t", Str(DiasProc), True, True)
       
       SecObj = SecObj + _
       WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
       WRtf("t", "ACCIONES CONSECUENTES:", True, True, , True)
       
       SecObj = SecObj + _
       WRtf("t", Space(8 * NivelProc - 1) + Space(36)) + _
       WRtf("t", "Completar su función y concluir sin tareas consecuentes.", True)
       
    End If
    NivelProc = NivelProc - 1
    Procesos = Procesos - 1
    DiasProc = DiasProc - EventoPropiedades(Nobj).DiasMaximo
End Sub


Sub SecuenciaReversaObjeto(Nobj As Integer, NobjC As Integer)
    If Not EventoTarea(Nobj).ProcesoActivo Then Exit Sub
    If Nobj = NobjC Then
       Circular = True
       Exit Sub
    End If
    Dim Tn%, Pa%, Np%, C1%, Rt%, Rd$, Rp%, DTt$, DTd$, C%
    Np = EventoTarea(Nobj).NroPrescedentes
    If Np > 0 Then
       For C = 1 To Np
           If EventoTarea(Nobj).Prescedente(C).RelacionActiva Then
              Rp = EventoTarea(Nobj).Prescedente(C).TareaPrescedente
              If Not EventoTarea(Nobj).Prescedente(C).RelacionInfinita Then
                 Call SecuenciaReversaObjeto(Rp, NobjC)
              End If
           End If
       Next
    Else
    End If
End Sub

Function Complejo() As Boolean
    Dim o%, C%, Nc%, Ni%
    Complejo = False
    For o = 1 To MaxProcObj
        If EventoTarea(o).ProcesoActivo Then
           Nc = EventoTarea(o).NroConsecuentes
           If Nc > 0 Then
              For C = 1 To Nc
                  If EventoTarea(o).Consecuente(C).RelacionActiva Then
                     Ni = Ni + IIf(EventoTarea(o).Consecuente(C).RelacionInfinita, 1, 0)
                     If Ni > 1 Then
                        Complejo = True
                        Exit Function
                     End If
                  End If
              Next
           End If
        End If
    Next
End Function

Sub InicializaMtxDatosAplic()
    Dim i As Long
    With DatosAplicMtx
      .Editable = True
      .AddColumn "campo", "Campo", , , 140
      .AddColumn "descripcion", "Descripcion", , , 140
      .AddColumn "tipo", "Tipo", , , 100
      .AddColumn "longitud", "Longitud", , , 80
      .AddColumn "inicial", "Valor Inicial", , , 140
      .AddColumn "key", , , , , False
      .GridLines = True
      .HeaderButtons = False
      .MultiSelect = False
      .RowMode = False
   
   
      For i = 1 To 1
         .AddRow , , (i = 1)
         .CellText(i, 1) = "(ninguno)"
         .CellText(i, 2) = "(ninguno)"
         .CellText(i, 3) = "Caracter"
         .CellText(i, 4) = "10"
         .CellText(i, 5) = "Nulo"
         .CellForeColor(i, 1) = vbButtonFace
         .CellForeColor(i, 2) = vbButtonFace
         .CellForeColor(i, 3) = vbButtonFace
         .CellForeColor(i, 4) = vbButtonFace
         .CellForeColor(i, 5) = vbButtonFace
      Next i
    End With
    cboTipo.AddItem "Caracter"
    cboTipo.AddItem "Numerico"
    cboTipo.AddItem "Moneda"
    cboTipo.AddItem "Fecha - Hora"
    cboTipo.AddItem "Logico"
    

End Sub

Sub CreaNotas(X As Single, Y As Single, Tipo As Integer, ByRef Nt As Integer)
    ' Cargar Nuevo Control
    Dim Mx%, Nota%
    
    If MaxNotaObj > 0 Then
       For Mx = 1 To MaxNotaObj
           If Not EventoNota(Mx).NotaActiva Then Exit For
       Next Mx
    Else
       Mx = 1
    End If
    If Mx > MaxNotaObj Then
       MaxNotaObj = MaxNotaObj + 1
       Nota = MaxNotaObj
       ReDim Preserve EventoNota(MaxNotaObj)
    Else
       Nota = Mx
    End If
    
        
    
    Herramientas.Buttons.Item(10).Enabled = IIf(Nota > 0, True, False)
    
    ' Guarda coordenadas del control creado
    EventoNota(Nota).IdProceso = "PXXXXX"
    EventoNota(Nota).IdNota = "N" + Format(MaxNotaObj, "00000")
    EventoNota(Nota).Nota = Nota
    EventoNota(Nota).Definicion = "Nota #" + Str(MaxNotaObj)
    EventoNota(Nota).Titulo = "Nota #" + Str(MaxNotaObj)
    EventoNota(Nota).Fecha = Format$(Now(), "dddd d, mmmm yyyy")
    
    EventoNota(Nota).posx = X
    EventoNota(Nota).posy = Y
    EventoNota(Nota).NotaActiva = True
    
    Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - (Notas(0).Width / 2)) + "," + Str(Y - (Notas(0).Height / 2)) + " "

    DibujaNota Nota, X, Y, EventoNota(Nota).Definicion
    If Nota >= 1 Then
       RequiereGrabar = True
    End If
       
    Nt = Nota

End Sub

Sub DibujaNota(Index As Integer, ByVal X As Double, ByVal Y As Double, Detalle As String)
    Load Notas(Index): Load Note(Index): Load DefNotas(Index)
    With Note(Index)
            .NoteLeft = (30 * Screen.TwipsPerPixelX)
            .NoteTop = (30 * Screen.TwipsPerPixelY)
    End With
    
    With Notas(Index)
        .ZOrder 0
        ' Posiciona y muestra el Control (Tarea)
        .Move X - (.Width / 2), Y - (.Height / 2), .Width, .Height
        .Visible = True
    End With
    With DefNotas(Index)
        .ZOrder 0
        ' Posiciona y muestra el Control (Tarea)
        .Move X - (.Width / 2), Notas(Index).Top + Notas(Index).Height + 2, .Width, .Height
        .Alignment = 2
        .Visible = True
    End With
    
    Note(Index).NoteCaption = EventoNota(Index).Titulo
    Note(Index).NoteText = Detalle
    Note(Index).NoteInfo = EventoNota(Index).Fecha
    DefNotas(Index).Caption = IIf(Len(Trim(Detalle)) > 24, Mid(Detalle, 1, 20) + "...", Trim(Detalle))

End Sub

Private Sub Note_UnloadNote(Index As Integer)
    If Index > 0 And Herramientas.Visible Then
        If EventoNota(Index).Definicion <> Note(Index).NoteText Then
           EventoNota(Index).Fecha = Format$(Now(), "dddd d, mmmm yyyy")
        End If
        EventoNota(Index).Definicion = Note(Index).NoteText
        DefNotas(Index).Caption = IIf(Len(Trim(Note(Index).NoteText)) > 24, Mid(Note(Index).NoteText, 1, 20) + "...", Trim(Note(Index).NoteText))
        RequiereGrabar = True
        RequiereReproc = True
        Call MuestraProceso
        
    End If
End Sub



Sub EliminaNota(NoNota As Integer)
    Dim PO%, pd%, Rsp%, Opt%, NActivos%, i%, Rt%
    Rt = 0
    Rsp = MsgBox("Eliminara toda anotación y comentarios anotados aqui." + Chr(10) + Chr(13) + "Esta operacion podria afectar a referencias futuras del proceso" + Chr(10) + Chr(13) + "Esta seguro de querer eliminar esta Nota?", vbYesNo + vbQuestion, "Eliminar Nota")
    Opt = 2
    If Rsp = 2 Or (Rsp = 7 And Opt <> 1) Then Exit Sub
    If MaxNotaObj > 0 Then
        Unload Notas(NoNota)
        Unload DefNotas(NoNota)
        Unload Note(NoNota)
        EventoNota(NoNota).NotaActiva = False
    End If
    Call VerificaHerramientas
    
    RequiereGrabar = True
    RequiereReproc = True
    Call MuestraProceso

End Sub



Sub MueveRelaciones(X As Single, Y As Single)
    Dim NRC%
    ObjEvento(ObjEnMovimiento).Move X - DespX, Y - DespY
    ShapeStat(ObjEnMovimiento).Move ObjEvento(ObjEnMovimiento).Left - 2, ObjEvento(ObjEnMovimiento).Top - 2, ObjEvento(ObjEnMovimiento).Width + 4, ObjEvento(ObjEnMovimiento).Height + 4
    
    RequiereGrabar = True
    RequiereReproc = True
    EventoTarea(ObjEnMovimiento).posx = (X - DespX) + (ObjEvento(0).Width / 2)
    EventoTarea(ObjEnMovimiento).posy = (Y - DespY) + (ObjEvento(0).Height / 2)
    Plantilla.XObjMax = 0
    Plantilla.YObjMax = 0
    
    Dim MxPrescedentes As Integer, MxConsecuentes As Integer
    Dim NRP%, OBjM%, OBjO%, OBjD%, NElm%
    Dim NRCD%, NRPD%
    Dim T1%, T2%
    
    For NRC = 1 To MaxProcObj
        Plantilla.XObjMax = IIf(EventoTarea(NRC).posx + ObjEvento(0).Width > Plantilla.XObjMax, EventoTarea(NRC).posx + ObjEvento(0).Width, Plantilla.XObjMax)
        Plantilla.YObjMax = IIf(EventoTarea(NRC).posy + ObjEvento(0).Height > Plantilla.YObjMax, EventoTarea(NRC).posy + ObjEvento(0).Height, Plantilla.YObjMax)
    Next NRC
    
    MxConsecuentes = EventoTarea(ObjEnMovimiento).NroConsecuentes
    MxPrescedentes = EventoTarea(ObjEnMovimiento).NroPrescedentes
    OBjM = ObjEnMovimiento
    Dim X1%, X2%, X3%, X4%
    Dim Y1%, Y2%, Y3%, Y4%
    Dim V$
    
    If MxConsecuentes > 0 Then
        For NRC = 1 To MxConsecuentes
            OBjD = EventoTarea(ObjEnMovimiento).Consecuente(NRC).TareaConsecuente
            NRCD = EventoTarea(ObjEnMovimiento).Consecuente(NRC).NumeroRConsecuente
            NElm = EventoTarea(ObjEnMovimiento).Consecuente(NRC).NoElementoRelacion
            T1 = EventoTarea(OBjD).Prescedente(NRCD).TareaPrescedente
            T2 = EventoTarea(OBjD).Prescedente(NRCD).NumeroRPrescedente
            
            r = EventoTarea(OBjM).Consecuente(NRC)
            If EventoTarea(OBjM).Consecuente(NRC).PosicionXYPersonal Then
                X2 = r.X2: Y2 = r.Y2: X3 = r.X3: Y3 = r.Y3: X4 = r.X4: Y4 = r.Y4
                V = Left(r.VH, 1)
            End If
            Call CalculaCoordenadasRelacion(ObjEvento(OBjM), ObjEvento(OBjD), OBjM, OBjD, False, NRC)
            If EventoTarea(OBjM).Consecuente(NRC).PosicionXYPersonal And V = Left(r.VH, 1) Then
                Select Case r.VH
                       Case "DR", "DL", "DN"
                           r.X4 = X4: r.Y4 = Y4
                           If r.Y1 >= Y2 Then
                              r.X3 = X3
                           Else
                              r.Y2 = Y2: r.X3 = X3: r.Y3 = Y3
                           End If
                       Case "UR", "UL", "UN"
                           r.X4 = X4: r.Y4 = Y4
                           If r.Y1 <= Y2 Then
                              r.X3 = X3
                           Else
                              r.Y2 = Y2: r.X3 = X3: r.Y3 = Y3
                           End If
                End Select
            Else
                EventoTarea(OBjM).Consecuente(NRC).PosicionXYPersonal = False
                r.PosicionXYPersonal = False
            End If
            Call DibujaRelacion(NElm, OBjM, NRC, OBjD, NRCD)
            EventoTarea(OBjM).Consecuente(NRC) = r
            EventoTarea(OBjD).Prescedente(NRCD) = r
            EventoTarea(OBjD).Prescedente(NRCD).TareaPrescedente = T1
            EventoTarea(OBjD).Prescedente(NRCD).NumeroRPrescedente = T2
        Next NRC
        RequiereGrabar = True
        RequiereReproc = True
    End If
    If MxPrescedentes > 0 Then
        For NRP = 1 To MxPrescedentes
             OBjO = EventoTarea(ObjEnMovimiento).Prescedente(NRP).TareaPrescedente
             NRPD = EventoTarea(ObjEnMovimiento).Prescedente(NRP).NumeroRPrescedente
             NElm = EventoTarea(ObjEnMovimiento).Prescedente(NRP).NoElementoRelacion
             T1 = EventoTarea(OBjO).Consecuente(NRPD).TareaConsecuente
             T2 = EventoTarea(OBjO).Consecuente(NRPD).NumeroRConsecuente
             
             r = EventoTarea(OBjM).Prescedente(NRP)
             If EventoTarea(OBjM).Prescedente(NRP).PosicionXYPersonal Then
                X1 = r.X1: Y1 = r.Y1: X2 = r.X2: Y2 = r.Y2: X3 = r.X3: Y3 = r.Y3
                V = Left(r.VH, 1)
             End If
             
             Call CalculaCoordenadasRelacion(ObjEvento(OBjO), ObjEvento(OBjM), OBjO, OBjM, False, NRP, True)
             If EventoTarea(OBjM).Prescedente(NRP).PosicionXYPersonal And V = Left(r.VH, 1) Then
                Select Case r.VH
                       Case "DR", "DL", "DN"
                           r.X1 = X1: r.Y1 = Y1
                           If r.Y4 <= Y2 Then
                              r.X2 = X2
                           Else
                              r.X2 = X2: r.Y2 = Y2:  r.Y3 = Y3
                           End If
                       Case "UR", "UL", "UN"
                           r.X1 = X1: r.Y1 = Y1
                           If r.Y4 >= Y2 Then
                              r.X2 = X2
                           Else
                              r.X2 = X2: r.Y2 = Y2: r.Y3 = Y3
                           End If
                End Select
             Else
                EventoTarea(OBjM).Prescedente(NRP).PosicionXYPersonal = False
                r.PosicionXYPersonal = False
             End If
             
             Call DibujaRelacion(NElm, OBjO, NRPD, OBjM, NRP)
             EventoTarea(OBjM).Prescedente(NRP) = r
             EventoTarea(OBjO).Consecuente(NRPD) = r
             EventoTarea(OBjO).Consecuente(NRPD).TareaConsecuente = T1
             EventoTarea(OBjO).Consecuente(NRPD).NumeroRConsecuente = T2
        Next NRP
        RequiereGrabar = True
        RequiereReproc = True
    End If

End Sub

Sub VerificaDimiensionPizarra(X As Single, Y As Single, State As Integer)
    Dim HXMax As Integer, HXMin As Integer
    Dim HYMax As Integer, HYMin As Integer
    Dim PXMin As Integer, PXMax As Integer
    Dim PYMin As Integer, PYMax As Integer
    Dim oB%
    
    
    HXMin = X - IIf(HerramientaSeleccionada <> 2, DespX, 0)
    HXMax = HXMin + ObjEvento(ObjEnMovimiento).Width
    HYMin = Y - IIf(HerramientaSeleccionada <> 2, DespY, 0)
    HYMax = HYMin + ObjEvento(ObjEnMovimiento).Height
    
    PXMin = VentanaMovil.HValue / Screen.TwipsPerPixelX
    PXMax = PXMin + VentanaMovil.Width
        
    PYMin = VentanaMovil.VValue / Screen.TwipsPerPixelY
    PYMax = PYMin + VentanaMovil.Height

    Ventana.StatusBar1.Panels.Item(4).Text = "" + Str(X - DespX) + IIf(HXMax >= PXMax, "*", "") + "," + Str(Y - DespY) + IIf(HYMax >= PYMax, "* ", " ")

        
    If HXMin <= PXMin And PXMin + Screen.TwipsPerPixelX >= Screen.TwipsPerPixelX Then
       If VentanaMovil.HValue - Screen.TwipsPerPixelX > 0 Then
          VentanaMovil.HValue = VentanaMovil.HValue - IIf(State = 0, Screen.TwipsPerPixelX, Screen.TwipsPerPixelX * 3)
          VentanaMovil.Refresh
          Pizarra.Refresh
       End If
    End If
    If HXMax >= PXMax And PXMin + Screen.TwipsPerPixelX <= PXMax Then
       If VentanaMovil.HValue + Screen.TwipsPerPixelX < VentanaMovil.HMaxValue Then
          VentanaMovil.HValue = VentanaMovil.HValue + IIf(State = 0, Screen.TwipsPerPixelX, Screen.TwipsPerPixelX * 3)
          VentanaMovil.Refresh
          Pizarra.Refresh
       End If
    End If
    
    If HYMin <= PYMin + Abs(HYMax - HYMin) And PYMin + Screen.TwipsPerPixelY >= Screen.TwipsPerPixelY Then
       If VentanaMovil.VValue - Screen.TwipsPerPixelY > 0 Then
          VentanaMovil.VValue = VentanaMovil.VValue - IIf(State = 0, Screen.TwipsPerPixelY, Screen.TwipsPerPixelY * 3)
          VentanaMovil.Refresh
          Pizarra.Refresh
       End If
    End If
    If HYMax >= PYMax - Abs(HYMax - HYMin) And PYMin + Screen.TwipsPerPixelY <= PYMax Then
       If VentanaMovil.VValue + Screen.TwipsPerPixelY < VentanaMovil.VMaxValue Then
          VentanaMovil.VValue = VentanaMovil.VValue + IIf(State = 0, Screen.TwipsPerPixelY, Screen.TwipsPerPixelY * 3)
          VentanaMovil.Refresh
          Pizarra.Refresh
       End If
    End If
End Sub


' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
    Dim pt As POINTAPI
    Dim pos As Integer
    Dim start_pos As Integer
    Dim end_pos As Integer
    Dim ch As String
    Dim txt As String
    Dim txtlen As Integer
    
        ' Convert the position to pixels.
        pt.X = X \ Screen.TwipsPerPixelX
        pt.Y = Y \ Screen.TwipsPerPixelY
    
        ' Get the character number
        pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
        If pos <= 0 Then Exit Function
    
        ' Find the start of the word.
        txt = rch.Text
        For start_pos = pos To 1 Step -1
            ch = Mid$(rch.Text, start_pos, 1)
            ' Allow digits, letters, and underscores.
            If Not ( _
                (ch >= "0" And ch <= "9") Or _
                (ch >= "a" And ch <= "z") Or _
                (ch >= "A" And ch <= "Z") Or _
                ch = "_" _
            ) Then Exit For
        Next start_pos
        start_pos = start_pos + 1
    
        ' Find the end of the word.
        txtlen = Len(txt)
        For end_pos = pos To txtlen
            ch = Mid$(txt, end_pos, 1)
            ' Allow digits, letters, and underscores.
            If Not ( _
                (ch >= "0" And ch <= "9") Or _
                (ch >= "a" And ch <= "z") Or _
                (ch >= "A" And ch <= "Z") Or _
                ch = "_" _
            ) Then Exit For
        Next end_pos
        end_pos = end_pos - 1
    
        If start_pos <= end_pos Then RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
End Function


Sub SeñalaTarea(NoTarea As Integer, Optional Status As Boolean = False, Optional Color As ColorConstants = vbWhite)
    If Status Then
       EventoTarea(NoTarea).BordeVisible = ShapeStat(NoTarea).Visible
       EventoTarea(NoTarea).BordeColor = ShapeStat(NoTarea).BorderColor
       EventoTarea(NoTarea).Señalado = True
       ShapeStat(NoTarea).Visible = True
       ShapeStat(NoTarea).BorderColor = Color
    ElseIf EventoTarea(NoTarea).Señalado Then
       ShapeStat(NoTarea).Visible = EventoTarea(NoTarea).BordeVisible
       ShapeStat(NoTarea).BorderColor = EventoTarea(NoTarea).BordeColor
       EventoTarea(NoTarea).Señalado = False
    End If
End Sub

Sub VerificaHerramientas()
    Dim Nts As Boolean
    Dim Rls As Boolean
    Dim Prs As Boolean
    Dim C%
    
    If MaxNotaObj > 0 Then
       For C = 1 To MaxNotaObj
           If EventoNota(C).NotaActiva Then Exit For
       Next
    Else
       C = 1
    End If
    Nts = IIf(C <= MaxNotaObj, True, False)
    
    
    If MaxRelaObj > 0 Then
       For C = 1 To MaxRelaObj
           If Relacion(C).RelacionActiva Then Exit For
       Next
    Else
       C = 1
    End If
    Rls = IIf(C <= MaxRelaObj, True, False)
    
    
    If MaxProcObj > 0 Then
       For C = 1 To MaxProcObj
           If C > 1 And EventoTarea(C).ProcesoActivo Then Exit For
       Next
    Else
       C = 1
    End If
    Prs = IIf(C <= MaxProcObj, True, False)
    
    Herramientas.Buttons.Item(10).Enabled = IIf(Not Prs, False, True)
    Herramientas.Buttons.Item(11).Enabled = IIf(Not Nts And Not Prs And Not Rls, False, True)
    Herramientas.Buttons.Item(12).Enabled = IIf(Not Rls, False, True)
    
    If (Not Herramientas.Buttons.Item(10).Enabled And Herramientas.Buttons.Item(10).Value = tbrPressed) Or _
       (Not Herramientas.Buttons.Item(11).Enabled And Herramientas.Buttons.Item(11).Value = tbrPressed) Or _
       (Not Herramientas.Buttons.Item(12).Enabled And Herramientas.Buttons.Item(12).Value = tbrPressed) Then
       Herramientas.Buttons.Item(1).Value = tbrPressed
       HerramientaSeleccionada = 0
    End If

End Sub
