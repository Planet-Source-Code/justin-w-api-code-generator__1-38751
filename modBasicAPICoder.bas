Attribute VB_Name = "modBasicAPICoder"
'*******************************************
'Program Filename:  Basic API Coder
'Author          :  Greenmonkey
'Date            :  6-5-01
'Description     :  Creates API code
'*******************************************

Option Explicit
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3

'API Declarations
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Type RGB
    Red As Long
    Green As Long
    Blue As Long
End Type
Global lngHandle        As Long
Global strCodeOption    As String

Global blnLoadedNames   As Boolean
Global strOptionNames() As String
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Dim strVariableNames()  As String    'stores variables

'**********************************************
'Procedure: CreateCode
'Purpose:   Create code to find a window
'Called by: frmWindowProperties.cmdCreateCode
'**********************************************
Public Function CreateCode(hwnd As Long) As String
    On Error GoTo ErrHandler
    
    Dim Txt             As String
    Dim strClassNames() As String
    Dim strTemp         As String
    Dim lngHandles()    As Long
    Dim lngHandle       As Long
    Dim lngChild        As Long
    Dim lngRetVal       As Long
    Dim intCount        As Integer
     
    'initialize
    lngHandle = hwnd
    
    'add first window
    strTemp = String(256, " ")
    lngRetVal = GetClassName(lngHandle, strTemp, 255)
    strTemp = Left(strTemp, InStr(strTemp, vbNullChar) - 1)
    ReDim Preserve strClassNames(intCount)
    ReDim Preserve lngHandles(intCount)
    strClassNames(intCount) = strTemp
    lngHandles(intCount) = lngHandle
    
    'add parents
    Do
        lngHandle = GetParent(lngHandle)
        
        If lngHandle <> 0 Then
            intCount = intCount + 1
            strTemp = String(256, " ")
            lngRetVal = GetClassName(lngHandle, strTemp, 255)
            strTemp = Left(strTemp, InStr(strTemp, vbNullChar) - 1)
            ReDim Preserve strClassNames(intCount)
            ReDim Preserve lngHandles(intCount)
            strClassNames(intCount) = strTemp
            lngHandles(intCount) = lngHandle
        End If
    
    Loop Until lngHandle = 0
    
    'create variables
    For intCount = 0 To UBound(strClassNames)
        ReDim Preserve strVariableNames(intCount)
        strVariableNames(intCount) = VariableName(strClassNames(intCount))
    Next
     
    'create code
    For intCount = 0 To UBound(strClassNames)
        If intCount <> UBound(strClassNames) Then
            lngHandle = GetParent(lngHandles(intCount))
            lngChild = FindWindowEx(lngHandle, 0, strClassNames(intCount), vbNullString)
            Do While lngChild <> 0 And lngChild <> lngHandles(intCount)
                Txt = strVariableNames(intCount) & " = FindWindowEx(" & strVariableNames(intCount + 1) & ", " & strVariableNames(intCount) & ", " & Chr(34) & strClassNames(intCount) & Chr(34) & ", vbNullString)" & vbCrLf & Txt
                lngChild = FindWindowEx(lngHandle, lngChild, strClassNames(intCount), vbNullString)
            Loop
            Txt = strVariableNames(intCount) & " = FindWindowEx(" & strVariableNames(intCount + 1) & ", 0, " & Chr(34) & strClassNames(intCount) & Chr(34) & ", vbNullString)" & vbCrLf & Txt
        Else
            Txt = strVariableNames(intCount) & " = FindWindow(" & Chr(34) & strClassNames(intCount) & Chr(34) & ", vbNullString)" & vbCrLf & Txt
        End If
    Next
    
    'add the code option
    Txt = Txt & GetCodeOption(strCodeOption)
    
    'declare variables
        Txt = VariableDeclarations(CInt(1)) & vbCrLf & vbCrLf & Txt
    
    CreateCode = Txt
    
ErrHandler:
End Function

'**********************************************
'Procedure: GetCodeOption
'Purpose:   Get code from a BAC code file
'Called by: CreateCode
'**********************************************
Public Function GetCodeOption(OptionName As String) As String
    On Error GoTo GetCodeErrHandler
    
    Dim strVariables()  As String
    Dim strOption       As String
    Dim strCode         As String
    Dim lngPos1         As Long
    Dim lngPos2         As Long
    Dim intNum          As Integer
    Dim intCount        As Integer
    
    Open App.Path & "\code\" & OptionName For Input As #1
    strOption = Input(LOF(1), 1)
    Close #1
    
    'get variables
    lngPos1 = InStr(strOption, "Dim") + 3
    lngPos2 = InStr(lngPos1, strOption, Chr(13))
    
    strVariables = Split(Mid(strOption, lngPos1, lngPos2 - lngPos1), ",")
    intCount = UBound(strVariableNames)
    'add variables to variablename array
    For intNum = 0 To UBound(strVariables)
        intCount = intCount + 1
        ReDim Preserve strVariableNames(intCount)
        strVariableNames(intCount) = strVariables(intNum)
        strVariableNames(intCount) = Trim(strVariableNames(intCount))
    Next
    
    'get the code from the file
    lngPos1 = InStrRev(strOption, "-")
    strCode = Right(strOption, Len(strOption) - lngPos1)
    'add the window's handle to the code option
    strCode = Replace(strCode, "*", strVariableNames(0))
    
    GetCodeOption = strCode
    
GetCodeErrHandler:
End Function

'**********************************************
'Procedure: VariableDeclarations
'Purpose:   Returns variable declarations
'Called by: CreateCode
'**********************************************
Public Function VariableDeclarations(PerLine As Integer) As String
    Dim intCount    As Integer
    Dim intNum      As Integer
    Dim strTemp     As String
    Dim strLine     As String
    Dim strType     As String
    
   
    For intNum = 0 To UBound(strVariableNames)
        
        If Left(strVariableNames(intNum), 3) <> "str" Then
            strType = "Long"
        Else
            strType = "String"
        End If
        
        intCount = intCount + 1
        
        If strLine = "" Then
            strLine = "Dim " & strVariableNames(intNum) & " As " & strType
        Else
            strLine = strLine & ", " & strVariableNames(intNum) & " As " & strType
        End If
        
        If intCount = PerLine Or intNum = UBound(strVariableNames) Then
            If strTemp = "" Then
                strTemp = strLine
            Else
                strTemp = strTemp & vbCrLf & strLine
            End If
            
            intCount = 0
            strLine = ""
        End If
    
    Next
        
    VariableDeclarations = strTemp
    
End Function

'**********************************************
'Procedure: VariableName
'Purpose:   Change a classname to a variable
'Called by: CreateCode
'**********************************************
Public Function VariableName(ClassName As String) As String
    Dim intNum      As Integer
    Dim strLetter   As String
    Dim strTemp     As String
    
    For intNum = 1 To Len(ClassName)
        strLetter = Mid(ClassName, intNum, 1)
        
        If Asc(strLetter) >= Asc("a") And Asc(strLetter) <= Asc("z") Then
            strTemp = strTemp & strLetter
        ElseIf Asc(strLetter) >= Asc("A") And Asc(strLetter) <= Asc("Z") Then
            strTemp = strTemp & strLetter
        End If
    
    Next
    
    If strTemp = "" Then
        strTemp = "X"
    ElseIf LCase(strTemp) = "static" Then
        strTemp = "StaticX"
    End If
    
    VariableName = strTemp
    
End Function
Public Function GetRGB(ByVal CVal As Long) As RGB
Dim TempColor As RGB
    
    TempColor.Blue = Int(CVal / 65536)
    TempColor.Green = Int((CVal - (65536 * TempColor.Blue)) / 256)
    TempColor.Red = CVal - (65536 * TempColor.Blue + 256 * TempColor.Green)

GetRGB = TempColor
  
End Function

Function GetWindowInformation(WindowHandle2 As Label, WindowClassName2 As Label, WindowText2 As Label, ParentList As ListBox)
    Dim CursorPos As POINTAPI
    Dim BufferAll&, WindowHandle&, TextLength&, PrevHandle&
    Dim WindowClassName$, WindowText$
    Call GetCursorPos(CursorPos)
    WindowHandle& = WindowFromPoint(CursorPos.X, CursorPos.Y)
    WindowClassName$ = String(100, Chr(0))
    BufferAll& = GetClassName(WindowHandle&, WindowClassName$, 100)
    WindowClassName$ = Left(WindowClassName$, BufferAll&)
    WindowText$ = String(100, Chr(0))
    BufferAll& = GetWindowTextLength(WindowHandle&)
    BufferAll& = GetWindowText(WindowHandle&, WindowText$, BufferAll& + 1)
    WindowText$ = Left(WindowText$, BufferAll&)
    WindowHandle2 = WindowHandle&
    WindowClassName2 = WindowClassName$
    WindowText2 = WindowText$
    PrevHandle& = GetParent(WindowHandle&)
    ParentList.Clear
    Do While PrevHandle& <> 0
        PrevHandle& = GetParent(WindowHandle&)
        WindowHandle& = PrevHandle&
        WindowClassName$ = String(100, Chr(0))
        BufferAll& = GetClassName(WindowHandle&, WindowClassName$, 100)
        WindowClassName$ = Left(WindowClassName$, BufferAll&)
        ParentList.AddItem (WindowClassName$)
    Loop
End Function

