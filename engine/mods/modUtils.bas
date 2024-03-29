Attribute VB_Name = "modUtils"
'Any tools we use for the scripting engine will be placed in here.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer

Public Function FixStr(s As String) As String
Dim I As Integer, ch As String * 1, sBuffer As String

    For I = 1 To Len(s)
        ch = Mid(s, I, 1)
        If Not ch = vbNullChar Then
            sBuffer = sBuffer & ch
        End If
    Next
    
    ch = ""
    I = 0
    
    FixStr = sBuffer
    sBuffer = ""
    
End Function

Public Function isVaildVarName(lpVarName As String) As Boolean
Dim I As Long, lChr As Byte, lpVarKeyLst As Variant

    isVaildVarName = True
    lpVarKeyLst = Split(ReservedWords, ",")
    'Check that the variable name is not a reserved word
    'Level 1 check
    For I = 0 To UBound(lpVarKeyLst)
        If lpVarKeyLst(I) = UCase(lpVarName) Then
        'invaild variable found
        isVaildVarName = False
        Erase lpVarKeyLst
        Exit Function
        End If
    Next
    
    'Level 2 check
    
    'Now to allow our variable to be added we need to check that it
    ' has a vaild name eg we do not want numbers as variable or extanded chars
    ' eg Dim 5Age as integer <- is invaild because of 5 at the first start
    ' Dim Age10 as string <- this is a vaild name
    ' Dim _Age as integer <- this is not becuase underscore is the first start
    ' Dim A*ge <- nor is this becuase of *
    ' Dim How_Old_25 <- this of cource is fine
    
    ' This is first part that checks for Alpha and numeric
    For I = 1 To Len(lpVarName)
        lChr = Asc(Mid(lpVarName, I, 1))
            If Not (lChr = Asc("_")) Then
                If IsCharAlphaNumeric(lChr) <> 1 Then
                    isVaildVarName = False
                    Exit Function
                End If
            End If
        Next
    I = 0
    
    'Level 3 check
    ' the final part checks the first char to see if it is a digit or an underscore
    lChr = Asc(Mid(lpVarName, 1, 1))
    
    If lChr = Asc("_") Or IsNumeric(Chr(lChr)) Then
        isVaildVarName = False
    End If
    
    lChr = 0
    
End Function

Public Sub GetLineCount(lpScript As String)
    ' All the sub does is store the number of lines the script has into LineCount
    LineCount = UBound(Split(lpScript, vbCrLf)) - 1
End Sub

Function GetIde() As Long
    'We use this to get the hangle of the ide window
    GetIde = FindWindow(vbNullString, "DM MyBasic-Script")
End Function

Function CharPos(lpStr As String, nChr As String) As Integer
Dim x As Integer, idx As Integer
    idx = 0
    'Function used to find the position of nChr in lpStr
    'Ex CharPos("hello world","r") returns 9
    
    For x = 1 To Len(lpStr) 'Loop tho lpStr
        If Mid(lpStr, x, 1) = nChr Then 'check if we have a match
            idx = x ' yes we have so store it's index
            Exit For ' get out of this loop
        End If
    Next
    
    x = 0
    CharPos = idx ' Return the index
    
End Function

Function isEmptyLine(expLine As String) As Boolean
    'Checks if the current executeing line is a nullchar
    isEmptyLine = (expLine = vbNullChar)
End Function

Function FixPath(lzPath As String) As String
    'Appends a \ to a given path
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    'Used to checking if a file exsits
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Public Function OpenFile(Filename As String) As String
Dim iFile As Long
Dim mByte() As Byte 'Byte array to hold the contents of the file

    'Opens a given file
    iFile = FreeFile 'Pointer to a free file
    Open Filename For Binary As #iFile 'Open file in binary mode
        'Resize the array to hold the data based on the length of the file
        If LOF(iFile) = 0 Then
            ReDim Preserve mByte(0 To LOF(iFile))
        Else
            ReDim Preserve mByte(0 To LOF(iFile) - 1)
        End If
        Get #iFile, , mByte 'Stote the data into the byte array
    Close #iFile
    
    OpenFile = StrConv(mByte, vbUnicode) 'Convert the array to a VB string and return
    
    Erase mByte 'Erase the array conents
    
End Function

Public Function SerachList(StartIdx As Integer, FindStr As String) As Integer
'This serach's LineHolder array looking for a match for FindStr
' StartIdx is the start position index that we seraching from
Dim x As Integer, idx As Integer
    
    idx = -1 'error flag index
    For x = StartIdx To UBound(LineHolder)
        If LCase(LineHolder(x)) = LCase(FindStr) Then
            idx = x
            Exit For
        End If
    Next
    
    SerachList = idx
    x = 0
    
End Function
