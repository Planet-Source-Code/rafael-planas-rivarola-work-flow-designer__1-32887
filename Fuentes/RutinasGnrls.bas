Attribute VB_Name = "RutinasGnrls"
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long


Sub InicializaAny(pvInitilizeAs As Variant, ParamArray psPossibleValArray() As Variant)
' Inputs:pvInitilizeAs = Initialize variable(s) to this - can be anything
'
' Returns:None
'
' Assumes:None
'
' Side Effects:It improves the use in Dll's, Ole-Server's in VB 4.
'     0 mostly. The result wil be faster Dll's and less memory usage.



    ' ==============================================================================
    '*\ PURPOSE:This will initialize passed in variables
    ' ==============================================================================
    Dim liArrayCnt, i As Integer
    liArrayCnt = UBound(psPossibleValArray)
    For i = 0 To liArrayCnt
        If VarType(pvInitilizeAs) = VarType(psPossibleValArray(i)) Then
            psPossibleValArray(i) = pvInitilizeAs
        Else
            'Debug.Print "Error in passed types"
            psPossibleValArray(i) = pvInitilizeAs
        End If
    Next
End Sub


Public Sub SetLoaded()


    'put this in your main forms' Load procedure
    'this will set the count
    Dim lTemp As Long, sPath As String
    lTemp& = GetLoaded&
    If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & App.EXEName & ".tmp" Else sPath$ = App.Path & App.EXEName & ".tmp"
    Open sPath$ For Output As #1
    Print #1, lTemp& + 1
    Close #1
End Sub


Public Function GetLoaded() As Long


    'call this to get how many times program has been loaded
    On Error Resume Next
    Dim sPath As String, sTemp As String
    If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & App.EXEName & ".tmp" Else sPath$ = App.Path & App.EXEName & ".tmp"
    Open sPath$ For Input As #1
    sTemp$ = Input(LOF(1), #1)
    Close #1
    If sTemp$ = "" Then GetLoaded& = 0 Else GetLoaded& = CLng(sTemp$)
End Function

