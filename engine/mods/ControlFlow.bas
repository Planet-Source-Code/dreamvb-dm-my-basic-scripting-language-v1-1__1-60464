Attribute VB_Name = "ControlFlow"
Sub doGoto()
Dim sTmp_Label As String
Dim e_pos As Integer, bFoundIdx As Integer
    'this function is used for the goto Statement
    ' at the moment it does seem to works quite well and does move between the lines.
    ' tho it may need some tweaks made here and there in some up comming versions.
    ' but as it now stands it does a good job
    
    'Below line check if the executeing ProcessLine is empty
    If isEmptyLine(ProcessLine) Then Abort 8, CurrentLine, "GOTO", "<Label>"
    
    'Below is used to locate the goto label in the script
    bFoundIdx = SerachList(0, ProcessLine)
    
    If bFoundIdx = -1 Then
        'If the label was not found we show the error
        Abort 2, CurrentLine, "Label not defined"
    Else
        'Move to the current line
        CurrentLine = bFoundIdx
    End If
    
    
    
End Sub
