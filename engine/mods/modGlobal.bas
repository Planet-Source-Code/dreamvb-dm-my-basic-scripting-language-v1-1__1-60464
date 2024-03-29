Attribute VB_Name = "modGlobal"
Public LineHolder() As String 'Used to hold the code Lines
Public LineCount As Long ' Current line count of total lines
Public CurrentLine As Long ' The current line number we are executeing
Public ProcessLine As String ' The current line number we are executeing
Public inSideQuotes As Boolean 'Check if a string is within quotes

'Reserved words in our Basic script
Public Const ReservedWords As String = "AS,REM,CLS,CLG,DIM,BEEP,COLOR" _
& ",GOTO,INPUT,LET,LOCATE,END,STRING,INTEGER,LONG,VARIANT,DOUBLE,BOOLEAN" _
& ",CONST,TRUE,FALSE,RND,DATE,TIME"

Public Const COMPILE_RESULT As Long = &H512 'Used as an error flag for the IDE

Public Sub Reset()
    ' Here we must reset any global vars been used
    MaxVars = -1 'Reset variable counter
    ConstMax = -1 'Reset const counter
    Erase mVarStack ' Erase the variable stack
    Erase mConstStack 'Erease consts stack
    ProcessLine = ""
    LineCount = -1
    CurrentLine = -1
    Erase LineHolder
End Sub

Function ReturnData(lpExpr As Variant) As Variant
Dim StrTmp As Variant, var_Idx As Integer
    
    'Edited to support consts
    
    StrTmp = Trim(lpExpr) 'Trim away any spaces in lpStr
    
    If IsNumeric(StrTmp) Then
        'If the expression is Numeric send it back
        ReturnData = Val(lpExpr)
        StrTmp = ""
        Exit Function
    End If
    
    'Look for a variable index
    var_Idx = VariableIndex(CVar(StrTmp))
        
    If var_Idx <> -1 Then
        If inSideQuotes Then ReturnData = StrTmp: Exit Function
        'The line above is used to stop the pressing of Variables inside stings
        ' eg print width whould print the value of width
        ' print "width" whould only print the word
        
        'Variable index is found so we need to get the data
        ' from that variable and return it back
        ReturnData = GetVar(CStr(StrTmp))
        StrTmp = ""
        Exit Function
    ElseIf ConstIndex(Trim(CStr(StrTmp))) <> -1 Then
        'Looks like we have found a const
        If inSideQuotes Then ReturnData = StrTmp: Exit Function
        ReturnData = GetConst(CStr(StrTmp)) 'Get the const data from the stack and pass it back
        StrTmp = ""
        Exit Function
    Else
        ' Ok we asume for now this is a string
        StrTmp = lpExpr
        ReturnData = StrTmp
        StrTmp = ""
        Exit Function
    End If
    
End Function

Public Sub GetCodeLines(lpScript As String)
Dim vStr As Variant, x As Long
    ' This function is used to fill LineHolder with all the current script lines
    
    vStr = Split(lpScript, vbCrLf, , vbBinaryCompare) 'Split the script by vbcrlf
    For x = 0 To UBound(vStr) - 1 ' loop though all the lines in vStr
        ReDim Preserve LineHolder(x) ' Resize LineHolder based on the number of lines
        LineHolder(x) = vStr(x) ' Add the current line to LineHolder
    Next
    
    If Not IsEmpty(vStr) Then
        'if vStr did contain lines we can area the array
        Erase vStr
    End If
    
    x = 0 ' Reset counter
    
End Sub

Public Sub Abort(code As Integer, LineIdx As Long, Optional Extra As String, Optional Other As String)
Dim StrMess As String, Ide_Hwnd As Long, atIdx As Integer


    ' We used this abort sub to repond to different error we may have incountered
    ' code is used to allow us to know what error code was found
    ' LineIdx is the current line that the error was found
    ' Extra and Other are optional and are used to add extra error information
    
    cFree ' close down the console first
    StrMess = ""
    StrMess = "Error: " & code & vbCrLf
    
    Select Case code
        Case 0
            'No script code was found
            StrMess = StrMess & "No program script found"
            StrMess = StrMess & vbCrLf & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 1
            ' An unkown command has been found
            StrMess = StrMess & "Syntex error unkown command: '" & Extra & "'"
            StrMess = StrMess & vbCrLf & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 2
            ' used for custom defined error message
            StrMess = StrMess & Extra & vbCrLf
            StrMess = StrMess & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 3
            StrMess = StrMess & "Syntex error unkown DataType: '" & Extra & "'"
            StrMess = StrMess & vbCrLf & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 4
            StrMess = StrMess & "Variable identifier required" & vbCrLf
            StrMess = StrMess & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 5
            StrMess = StrMess & "Duplication variable '" & Extra & "' in current scope" & vbCrLf
            StrMess = StrMess & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 6
            StrMess = StrMess & "Undefined variable '" & Extra & "' in current scope" & vbCrLf
            StrMess = StrMess & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 7
            StrMess = StrMess & "unkown variable Datatype '" & Extra & "'" & vbCrLf
            StrMess = StrMess & "Line: " & LineIdx + 1
            GoTo Terminate:
        Case 8
            StrMess = StrMess & "Syntex error: '" & Extra & "'" & vbCrLf
            StrMess = StrMess & "Use: " & Extra & " " & Other & vbCrLf
            StrMess = StrMess & "Line: " & LineIdx + 1
            GoTo Terminate:
    End Select
    
Terminate:
    'Ok this small if statement check to see if the DM MyBasic Script IDE is Open
    ' if the window is found then we send a message informing of an error.
    ' Why do we do this. well the reason is so we can phase our error messages into
    ' the error listbox and also move to the correct error line in the editor.
    ' If no IDE window was not found. we just use a standared message box
    
    Ide_Hwnd = GetIde
    
    If Ide_Hwnd <> 0 Then
        atIdx = GlobalAddAtom(StrMess) 'Add the error message to the Atom
        'Below we send the error flag and the atIdx
        SendMessage Ide_Hwnd, COMPILE_RESULT, ByVal atIdx, ByVal 0
    Else
        MsgBox StrMess, vbInformation, "Execute Error"
    End If
    
    Call Reset ' this calls the reset to reset Global vars
    End ' end the program
End Sub

Public Sub AddSystemVars()
    'This is were we place any system consts eg TRUE / FALSE etc
    AddConst "true", -1, True
    AddConst "false", 0, True
    AddConst "rnd", Rnd, True
    AddConst "time", Time, True
    AddConst "date", Date, True
End Sub
