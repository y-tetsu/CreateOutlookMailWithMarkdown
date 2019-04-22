Attribute VB_Name = "Markdown"
'==============================================================================================================
' Markdown
'==============================================================================================================
Option Explicit

Const SPLIT_CHAR = "`split`"


'---------------------------------------------------------------------------------
' Parse
'---------------------------------------------------------------------------------
Function Parse(lines As Variant) As String
    Dim state As String
    Dim code_lines As Variant
    Dim buffer As String
    Dim python_code As Boolean
    Dim simple_code As Boolean
    Dim skip_flag As Boolean
    Dim tmp As String
    Dim i As Long
    Dim re
    
    Parse = ""
    state = "normal_line"
    buffer = ""
    python_code = False
    
    Set re = CreateObject("VBScript.RegExp")
    re.ignoreCase = False
    re.Global = True
    
    For i = 0 To UBound(lines)
        skip_flag = False
        
        If state = "normal_line" Then
            ' Python Syntax
            re.Pattern = "^```python$"
            
            If re.Test(lines(i)) Then
                python_code = True
                state = "syntax_line"
            End If
            
            ' Python Syntax with Text
            re.Pattern = "^```python (.*)$"
            
            If re.Test(lines(i)) Then
                python_code = True
                state = "syntax_line"
                
                tmp = re.Replace(lines(i), "$1")
                
                lines(i) = ""
                lines(i) = lines(i) & "<span style=color:#FFFFFF>"
                lines(i) = lines(i) & "<div style=""background-color: #364549;"">"
                lines(i) = lines(i) & "<span style=background-color:#777777>"
                lines(i) = lines(i) & tmp
                lines(i) = lines(i) & "</span>"
                lines(i) = lines(i) & "</div>"
                lines(i) = lines(i) & "</span>"
                
                Parse = Parse & lines(i)
            End If
            
            ' Simple Syntax
            re.Pattern = "^```$"
            
            If re.Test(lines(i)) Then
                simple_code = True
                state = "syntax_line"
            End If
            
            ' Simple Syntax with Text
            re.Pattern = "^``` (.*)$"
            
            If re.Test(lines(i)) Then
                simple_code = True
                state = "syntax_line"
                
                tmp = re.Replace(lines(i), "$1")
                
                lines(i) = ""
                lines(i) = lines(i) & "<span style=color:#FFFFFF>"
                lines(i) = lines(i) & "<div style=""background-color: #364549;"">"
                lines(i) = lines(i) & "<span style=background-color:#777777>"
                lines(i) = lines(i) & tmp
                lines(i) = lines(i) & "</span>"
                lines(i) = lines(i) & "</div>"
                lines(i) = lines(i) & "</span>"
                
                Parse = Parse & lines(i)
            End If
            
            ' Output Normal Line
            If state = "normal_line" Then
                ' ---
                re.Pattern = "^---$"
                
                If re.Test(lines(i)) Then
                    Parse = Parse & "<p><hr></p>"
                    
                    skip_flag = True
                End If
                
                ' # <string>
                re.Pattern = "^# (.*)$"
                
                If re.Test(lines(i)) And Not skip_flag Then
                    tmp = re.Replace(lines(i), "$1")
                    
                    lines(i) = "<h1>" & tmp & "</h1>"
                    
                    Parse = Parse & lines(i)
                    
                    skip_flag = True
                End If
                
                ' ## <string>
                re.Pattern = "^## (.*)$"
                
                If re.Test(lines(i)) And Not skip_flag Then
                    tmp = re.Replace(lines(i), "$1")
                    
                    lines(i) = "<h2>" & tmp & "</h2>"
                    
                    Parse = Parse & lines(i)
                    
                    skip_flag = True
                End If
                
                ' ### <string>
                re.Pattern = "^### (.*)$"
                
                If re.Test(lines(i)) And Not skip_flag Then
                    tmp = re.Replace(lines(i), "$1")
                    
                    lines(i) = "<h3>" & tmp & "</h3>"
                    
                    Parse = Parse & lines(i)
                    
                    skip_flag = True
                End If
                
                ' #### <string>
                re.Pattern = "^#### (.*)$"
                
                If re.Test(lines(i)) And Not skip_flag Then
                    tmp = re.Replace(lines(i), "$1")
                    
                    lines(i) = "<h4>" & tmp & "</h4>"
                    
                    Parse = Parse & lines(i)
                    
                    skip_flag = True
                End If
                
                ' ##### <string>
                re.Pattern = "^##### (.*)$"
                
                If re.Test(lines(i)) And Not skip_flag Then
                    tmp = re.Replace(lines(i), "$1")
                    
                    lines(i) = "<h5>" & tmp & "</h5>"
                    
                    Parse = Parse & lines(i)
                    
                    skip_flag = True
                End If
                
                ' ###### <string>
                re.Pattern = "^###### (.*)$"
                
                If re.Test(lines(i)) And Not skip_flag Then
                    tmp = re.Replace(lines(i), "$1")
                    
                    lines(i) = "<h6>" & tmp & "</h6>"
                    
                    Parse = Parse & lines(i)
                    
                    skip_flag = True
                End If
                
                If Not skip_flag Then
                    Dim tmp_lines(0) As Variant
                    tmp_lines(0) = lines(i)
                    Parse = Parse & LineArrange.AddBrTag(LineArrange.ReplaceSpace(tmp_lines))
                End If
            End If
        
        ' Output Syntax Line
        ElseIf state = "syntax_line" Then
            ' Check End of Syntax Highlight
            re.Pattern = "^```$"
            
            If re.Test(lines(i)) Then
                state = "normal_line"
                code_lines = Split(buffer, SPLIT_CHAR)
                buffer = ""
                
                Parse = Parse & "<span style=color:#FFFFFF>"
                Parse = Parse & "<div style=""background-color: #364549"">"
                
                If python_code Then
                    Parse = Parse & LineArrange.AddBrTag(PythonSyntax.HighLight(ReplaceSpace(code_lines)))
                    python_code = False
                ElseIf simple_code Then
                    Parse = Parse & LineArrange.AddBrTag(LineArrange.ReplaceSpace(code_lines))
                    simple_code = False
                End If
                
                Parse = Parse & "</div>"
                Parse = Parse & "</span>"
            Else
                If buffer = "" Then
                    buffer = lines(i)
                Else
                    buffer = buffer & SPLIT_CHAR & lines(i)
                End If
            End If
        End If
    Next i
End Function


'==============================================================================================================
' END
'==============================================================================================================
