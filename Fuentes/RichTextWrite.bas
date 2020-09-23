Attribute VB_Name = "RichTextWrite"

Function WRtf(Caso As String, Cadena As String, Optional CRLF As Boolean = False, Optional Negrita As Boolean = False, Optional Italica As Boolean = False, Optional SubRayado As Boolean = False, Optional Color As Integer = 0, Optional Justificado As String = "l", Optional Tamaño As Integer = 18, Optional Letra As Integer = 1, Optional Estilo As Integer = 0) As String
    Dim txt As String
    txt = "\plain"
    Select Case UCase(Caso)
           Case "T"
                txt = txt & IIf(Negrita, "\b", "")
                txt = txt & IIf(Italica, "\i", "")
                txt = txt & IIf(SubRayado, "\ul", "")
                Select Case UCase(Justificado)
                       Case "L"
                            txt = txt & "\ql"
                       Case "C"
                            txt = txt & "\qc"
                       Case "R"
                            txt = txt & "\qr"
                       Case Else
                            txt = txt & "\ql"
                End Select
                txt = txt & IIf(Color >= 0 And Color < 16, "\cf" & LTrim(Str(Color)), "\cf0")
                txt = txt & IIf(Tamaño > 0 And Tamaño < 50, "\fs" & LTrim(Str(Tamaño)), "\fs20")
                txt = txt & IIf(Letra >= 0 And Letra <= 15, "\f" & LTrim(Str(Letra)), "\f1")
                txt = txt & IIf(Estilo >= 0 And Estilo <= 19, "\cs" & LTrim(Str(Estilo)), "")
                txt = txt & " " & Cadena
                
                txt = txt & IIf(CRLF, vbCrLf & "\par }{", "")
           Case "I"
                txt = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033"
                txt = txt & "{\fonttbl"
                txt = txt & "{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}"
                txt = txt & "{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0604020202020204}Arial;}"
                txt = txt & "{\f2\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Courier New;}"
                txt = txt & "{\f3\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Mirror;}"
                txt = txt & "{\f15\fswiss\fcharset0\fprq2{\*\panose 020b0604030504040204}Verdana;}}"
                txt = txt & "{\colortbl;"
                txt = txt & "\red0\green0\blue0;"        'color 0,1  NEGRO
                txt = txt & "\red0\green0\blue255;"      'color 2    AZUL
                txt = txt & "\red0\green255\blue255;"    'color 3
                txt = txt & "\red0\green255\blue0;"      'color 4    VERDE
                txt = txt & "\red255\green0\blue255;"    'color 5
                txt = txt & "\red255\green0\blue0;"      'color 6    ROJO
                txt = txt & "\red255\green255\blue0;"    'color 7
                txt = txt & "\red255\green255\blue255;"  'color 8
                txt = txt & "\red0\green0\blue128;"      'color 9
                txt = txt & "\red0\green128\blue128;"    'color 10
                txt = txt & "\red0\green128\blue0;"      'color 11
                txt = txt & "\red128\green0\blue128;"    'color 12
                txt = txt & "\red128\green0\blue0;"      'color 13
                txt = txt & "\red128\green128\blue0;"    'color 14
                txt = txt & "\red128\green128\blue128;"  'color 15
                txt = txt & "\red192\green192\blue192;}" 'color 16
                txt = txt & "{\stylesheet"
                txt = txt & "{\widctlpar\adjustright \fs20\lang2057\cgrid \snext0 Normal;}"
                txt = txt & "{\s1\sb240\sa60\keepn\widctlpar\adjustright \b\f15\fs28\lang2057\kerning28\cgrid \sbasedon0 \snext0 heading 1;}"
                txt = txt & "{\s3\sb240\sa60\keepn\widctlpar\adjustright \b\f15\lang2057\cgrid \sbasedon0 \snext0 heading 3;}{\*\cs10 \additive Default Paragraph Font;}"
                txt = txt & "{\s15\qc\widctlpar\adjustright \b\f15\fs16\lang2057\cgrid \sbasedon0 \snext0 caption;}"
                txt = txt & "{\s16\li720\widctlpar\adjustright \f2\fs16\lang2057\cgrid \sbasedon0 \snext16 Code;}{\*\cs17 \additive \ul\cf12 \sbasedon10 FollowedHyperlink;}{\*\cs18 \additive \ul\cf2 \sbasedon10 Hyperlink;}"
                txt = txt & "{\s19\widctlpar\adjustright \f15\fs20\lang2057\cgrid \sbasedon0 \snext19 Paragraph;}}"
                txt = txt & "{\info"
                txt = txt & "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard\plain \widctlpar\adjustright \fs20\lang2057\cgrid {"
           Case "F"
                txt = txt & "\par }}"
    End Select
    WRtf = txt
End Function
