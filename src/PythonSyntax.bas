Attribute VB_Name = "PythonSyntax"
'==============================================================================================================
' Python Syntax
'==============================================================================================================
Option Explicit

Const PYTHON_OPERATOR1 = "=|<|>"
Const PYTHON_OPERATOR2 = "==|\!=|\+=|-=|\*\*=|\*=|//=|/=|%=|&=|\|=|\^=|<<=|>>=|<=|>=|<>|->|\+|-|\*\*|\*|//|/|%|&|\||\^|<<|>>"
Const PYTHON_OPERATOR3 = "\."
Const PYTHON_OPERATOR4 = "is\s+not|is|not\s+in|in"
Const PYTHON_SYNTAX1 = "print|and|as|assert|break|class|continue|def|del|elif|else|except|finally|for|global|if|is|lambda|not|or|pass|raise|return|try|while|with|yield"
Const PYTHON_SYNTAX2 = "from|import"


'---------------------------------------------------------------------------------
' HighLight
'---------------------------------------------------------------------------------
Function HighLight(lines As Variant) As Variant
    Dim line As Variant
    Dim i As Long
    Dim tmp1 As String
    Dim tmp2 As String
    Dim code As String
    Dim comment As String
    Dim tdquote, tsquote As Boolean
    Dim re
    Dim index As Long
    
    Set re = CreateObject("VBScript.RegExp")
    re.ignoreCase = False
    re.Global = True
    
    tdquote = False
    tsquote = False
    
    For i = 0 To UBound(lines)
        tmp1 = ""
        tmp2 = ""
        code = ""
        comment = ""
        
        ' Pre-Process
        re.Pattern = "<(.*?)>"
        lines(i) = re.Replace(lines(i), "&lt;$1&gt;")
        
        ' """ Comment
        re.Pattern = "("""""".*)$"
        If tdquote = True Then
            comment = "<span style=color:#41B7D7>" & lines(i) & "</span>"
            
            If re.Test(lines(i)) Then
                tdquote = False
            End If
        Else
            If re.Test(lines(i)) Then
                tdquote = True
                comment = "<span style=color:#41B7D7>" & lines(i) & "</span>"
            Else
                tmp1 = lines(i)
            End If
        End If
        
        ' ''' Comment
        re.Pattern = "('''.*)$"
        If tsquote = True Then
            comment = "<span style=color:#41B7D7>" & tmp1 & "</span>"
            
            If re.Test(tmp1) Then
                tsquote = False
            End If
        Else
            If re.Test(tmp1) Then
                tsquote = True
                comment = "<span style=color:#41B7D7"">" & tmp1 & "</span>"
            Else
                tmp2 = tmp1
            End If
        End If
        
        ' # Comment
        re.Pattern = "^(.*?)(#.*)$"

        If re.Test(tmp2) Then
            code = re.Replace(tmp2, "$1")
            comment = re.Replace(tmp2, "<span style=color:#82ABA4>$2</span>")
        Else
            code = tmp2
        End If
        
        ' Save "string"
        re.Pattern = "([urfb]*"".*?"")"
        
        Dim matches_dc, match
        Set matches_dc = re.Execute(code)
        index = 0
        For Each match In matches_dc
            code = Replace(code, match.Value, "{#" & CStr(index) & "}", 1, 1)
            index = index + 1
        Next match
        
        ' Save 'string'
        re.Pattern = "([urfb]*'.*?')"
        
        Dim matches_sc
        Set matches_sc = re.Execute(code)
        index = 0
        For Each match In matches_sc
            code = Replace(code, match.Value, "{#" & CStr(index) & "}", 1, 1)
            index = index + 1
        Next match
        
        ' OPERATOR1
        re.Pattern = "([^=\!\+\*/%&\|\^<>])(" & PYTHON_OPERATOR1 & ")([^=<>])"
        code = re.Replace(code, "$1<span style=color:#FF8095>$2</span>$3")
        
        ' OPERATOR2
        re.Pattern = "([^<>a-zA-Z])(" & PYTHON_OPERATOR2 & ")([^/(span)<>]|$)"
        code = re.Replace(code, "$1<span style=color:#FF8095>$2</span>$3")
        
        ' OPERATOR3
        re.Pattern = "(import|from|@)"
        If Not re.Test(code) Then
            re.Pattern = "([^\d])(" & PYTHON_OPERATOR3 & ")"
            code = re.Replace(code, "$1<span style=color:#FF8095>$2</span>")
        End If
        
        ' OPERATOR4
        re.Pattern = "([W+])(" & PYTHON_OPERATOR4 & ")([W+])"
        code = re.Replace(code, "$1<span style=color:#FF8095>$2</span>$3")
        
        ' Number
        re.Pattern = "([^#a-zA-Z\d]+)(\d+\.*\d*)"
        code = re.Replace(code, "$1<span style=color:#A980F5>$2</span>")
        
        ' Decorator
        re.Pattern = "^(@)(\S+?)(\(|$)"
        code = re.Replace(code, "<span style=color:#80D34B>$1$2</span>$3")
        
        ' FunctionÅEClass
        re.Pattern = "(\W+|^)(def|class)(\W+)(.*)(\()"
        code = re.Replace(code, "$1$2$3<span style=color:#80D34B>$4</span>$5")
        
        ' Syntax1
        re.Pattern = "(\W+|^)(" & PYTHON_SYNTAX1 & ")(\W+|$)"
        code = re.Replace(code, "$1<span style=color:#EBD247>$2</span>$3")
        
        ' Syntax
        re.Pattern = "(\W+|^)(" & PYTHON_SYNTAX2 & ")(\W+)"
        code = re.Replace(code, "$1<span style=color:#FF8095>$2</span>$3")
        
        ' "string"
        index = 0
        For Each match In matches_dc
            code = Replace(code, "{#" & CStr(index) & "}", "<span style=color:#41B7D7>" & match.Value & "</span>", 1, 1)
            index = index + 1
        Next match
        
        ' 'string'
        index = 0
        For Each match In matches_sc
            code = Replace(code, "{#" & CStr(index) & "}", "<span style=color:#41B7D7>" & match.Value & "</span>", 1, 1)
            index = index + 1
        Next match
        
        ' Join code and comments
        lines(i) = code & comment
    Next i
    
    HighLight = lines
End Function


'==============================================================================================================
' END
'==============================================================================================================
