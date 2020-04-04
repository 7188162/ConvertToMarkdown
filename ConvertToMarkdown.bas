Attribute VB_Name = "ConvertToMarkdown"
Sub ConvertTo_MDTable()
    Attribute ConvertTo_MDTable.VB_ProcData.VB_Invoke_Func = "M\n14"
    Dim md As String
    md = ""

    For j = Selection(1).Row To (Selection(Selection.Count).Row)
        Dim theline As String
        theline = "|"
        
        For i = Selection(1).Column To (Selection(Selection.Count).Column)
                    If Cells(j, i).Value = "" Then
                    cellvalue = " "
                    ElseIf Cells(j, i).Value = "-" Then
                    cellvalue = "-"
                    Else
                        cellvalue = Cells(j, i).Text
                    End If
                    
                    If Cells(j, i).Font.Bold Then
                    cellvalue = "**" & cellvalue & "**"
                    End If
                    
                    theline = theline & cellvalue & "|"
        Next i
        
        theline = theline & vbCrLf

        If j = (Selection(1).Row + 1) Then
        Dim hr As String
        
            hr = "|"
            For k = Selection(1).Column To (Selection(Selection.Count).Column)
                Select Case Cells(Selection(1).Row, k).DisplayFormat.HorizontalAlignment
                    Case xlLeft:    strAlign = ":---"
                    Case xlCenter:  strAlign = ":---:"
                    Case xlRight:   strAlign = "---:"
                    Case Else:      strAlign = "---"
                End Select
                hr = hr & strAlign & "|"
            Next k
            md = md & hr & vbCrLf
        End If
        md = md & theline
    Next j

    Call copyToClipboard(md)

    MsgBox "copied Markdown ", vbInformation + vbOKOnly

End Sub

Sub copyToClipboard(str)
    Dim CB As Object
    Set CB = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    CB.SetText str
    CB.PutInClipboard
End Sub
