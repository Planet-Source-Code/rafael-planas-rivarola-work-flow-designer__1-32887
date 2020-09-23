Attribute VB_Name = "Definiciones"

Type Frms
     Caption As String
     Handle As Long
     Secuencia As Integer
End Type

Type ElementoCordenadas
     XCordLeft   As Single
     XCordMed    As Single
     XCordRigth  As Single
     YCordTop    As Single
     YCordMed    As Single
     YCordBottom As Single
End Type

Type WorkArea
     XMax    As Single
     YMax    As Single
     XObjMax As Single
     YObjMax As Single
End Type

Type RelacionCordenadas
     RelacionActiva      As Boolean
     RelacionInfinita    As Boolean
     TareaPrescedente    As Integer
     NumeroRPrescedente  As Integer
     TareaConsecuente    As Integer
     NumeroRConsecuente  As Integer
     NoElementoRelacion  As Integer
     RelacionTipo        As Integer
     PosicionXYPersonal  As Boolean
     VH                  As String * 2
     Ruta                As String * 1
     X1                  As Single
     Y1                  As Single
     X2                  As Single
     Y2                  As Single
     X3                  As Single
     Y3                  As Single
     X4                  As Single
     Y4                  As Single
End Type

Type Condiciones
     CondicionActiva As Boolean
     NoCondicion     As Byte
     Definicion      As String * 25
     Tipo            As Byte
End Type

Type PropiedadesTarea
     Personalizada   As Boolean
     IdProceso       As String * 6
     IdTarea         As String * 6
     NoTarea         As Integer
     TareaTipo       As Byte
     Definicion      As String * 25
     NoCondiciones   As Byte
     Condicion()     As Condiciones
     Para            As String * 100
     Asunto          As String * 60
     Mensaje         As String * 500
     DiasMinimo      As Single
     DiasMaximo      As Single
End Type

Type ElementoTarea
     ProcesoActivo   As Boolean
     IdProceso       As String * 6
     IdTarea         As String * 6
     NoTarea         As Integer
     TareaTipo       As Byte
     posx            As Single
     posy            As Single
     Definicion      As String * 25
     NroPrescedentes As Byte
     NroConsecuentes As Byte
     Prescedente()   As RelacionCordenadas
     Consecuente()   As RelacionCordenadas
     Terminal        As Boolean
     BordeVisible    As Boolean
     BordeColor      As ColorConstants
     Señalado        As Boolean
End Type

Type ElementoRelacion
     RelacionActiva      As Boolean
     TareaOrigen         As Integer
     ConsecuentesOrigen  As Integer
     TareaDestin         As Integer
     PrescedentesDestin  As Integer
End Type
     
Type ElementoNota
     NotaActiva      As Boolean
     IdProceso       As String * 6
     IdNota          As String * 6
     Nota            As Integer
     posx            As Single
     posy            As Single
     Titulo          As String * 20
     Definicion      As String * 255
     Fecha           As String * 30
End Type
     
Type DatosProc
     Activo              As Boolean
     Campo               As String * 20
     Definicion          As String * 20
     Tipo                As String * 10
     Longitud            As String * 5
     VDefecto            As String * 20
     Clave               As String * 20
End Type
     
Type Grilla
     Fila    As Integer
     Columna As Integer
End Type
     
     
Global Flujo() As Frms
     
Global GEventoTarea       As ElementoTarea
Global GEventoPropiedades As PropiedadesTarea

Global ArchivoACargar As String

'Show parameter for form.
Global Const MODAL = 1

'Settings for MsgBox.
Global Const MB_OK = 0
Global Const MB_YESNOCANCEL = 3
Global Const MB_YESNO = 4
Global Const MB_ICONEXCLAMATION = 48
'Specifies default button if more than one.
Global Const MB_DEFBUTTON1 = 0
Global Const MB_DEFBUTTON2 = 256
Global Const MB_DEFBUTTON3 = 512

'Return values from YES/NO/CANCEL Message Box.
Global Const IDNOSAVE = -1
Global Const IDOK = 1
Global Const IDCANCEL = 2
Global Const IDYES = 6
Global Const IDNO = 7

'CMDIALOG.VBX error (user clicked Cancel).
Global Const CDERR_CANCEL = &H7FF3


