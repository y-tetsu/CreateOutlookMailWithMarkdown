Attribute VB_Name = "LineArrange"
'==============================================================================================================
' Line Arrange
'==============================================================================================================
Option Explicit


'---------------------------------------------------------------------------------
' Add Br Tag Each Line
'---------------------------------------------------------------------------------
Function AddBrTag(lines As Variant) As String
    Dim i As Long
    Dim ret As String
    
    ret = ""
    
    For i = 0 To UBound(lines)
        ret = ret & lines(i) & "<br>"
    Next i
    
    AddBrTag = ret
End Function


'---------------------------------------------------------------------------------
' Replace Space Character To "&nbsp;"
'---------------------------------------------------------------------------------
Function ReplaceSpace(lines As Variant) As Variant
    Dim i As Long
    Dim ret As String
    
    For i = 0 To UBound(lines)
        lines(i) = Replace(lines(i), " ", "&nbsp;")
    Next i
    
    ReplaceSpace = lines
End Function


'==============================================================================================================
' END
'==============================================================================================================
