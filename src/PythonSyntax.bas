Attribute VB_Name = "PythonSyntax"
'==============================================================================================================
' Python Syntax
'==============================================================================================================
Option Explicit

Const PYTHON_OPERATOR1 = "=|<|>|\.|\+|-|\*|/|%|&|\||\^"
Const PYTHON_OPERATOR2 = "and|or|is|in|from|import"
Const PYTHON_OPERATOR3 = "not"
Const PYTHON_SYNTAX1 = "print|as|assert|break|class|continue|def|del|elif|else|except|finally|for|global|if|lambda|pass|raise|return|try|while|with|yield"

Const SPAN_STYLE1 = "<span style=color:#41B7D7>"  ' """Comment or '''Comment of String
Const SPAN_STYLE2 = "<span style=color:#82ABA4>"  ' #Comment
Const SPAN_STYLE3 = "<span style=color:#FF8095>"  ' Operator
Const SPAN_STYLE4 = "<span style=color:#A980F5>"  ' Number
Const SPAN_STYLE5 = "<span style=color:#80D34B>"  ' Decorator-name Function-name Class-name
Const SPAN_STYLE6 = "<span style=color:#EBD247>"  ' Syntax

'---------------------------------------------------------------------------------
' HighLight
'---------------------------------------------------------------------------------
Function HighLight(lines As Variant) As Variant
    Dim line As Variant
    Dim i As Long
    Dim tmp1 As String
    Dim tmp2 As String
    Dim tmp3 As String
    Dim tmp4 As String
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
        re.Pattern = "&nbsp;"
        lines(i) = re.Replace(lines(i), "{space}")
        
        ' """ Comment """
        re.Pattern = "(.*)("""""".*"""""")$"
        
        If re.Test(lines(i)) Then
            tmp1 = re.Replace(lines(i), "$1")
            comment = "{span1}" & re.Replace(lines(i), "$2") & "{endspan}"
        Else
            tmp1 = lines(i)
        End If
        
        ' """ Comment
        re.Pattern = "(.*)("""""".*)$"
        If tdquote = True Then
            comment = "{span1}" & tmp1 & "{endspan}"
            
            If re.Test(tmp1) Then
                tdquote = False
            End If
        Else
            If re.Test(tmp1) Then
                tdquote = True
                tmp2 = re.Replace(tmp1, "$1")
                comment = "{span1}" & re.Replace(tmp1, "$2") & "{endspan}"
            Else
                tmp2 = tmp1
            End If
        End If
        
        ' ''' Comment '''
        re.Pattern = "(.*)('''.*''')$"
        
        If re.Test(tmp2) Then
            tmp3 = re.Replace(tmp2, "$1")
            comment = "{span1}" & re.Replace(tmp2, "$2") & "{endspan}"
        Else
            tmp3 = tmp2
        End If
        
        ' ''' Comment
        re.Pattern = "(.*)('''.*)$"
        If tsquote = True Then
            comment = "{span1}" & tmp3 & "{endspan}"
            
            If re.Test(tmp3) Then
                tsquote = False
            End If
        Else
            If re.Test(tmp3) Then
                tsquote = True
                tmp4 = re.Replace(tmp3, "$1")
                comment = "{span1}" & re.Replace(tmp3, "$2") & "{endspan}"
            Else
                tmp4 = tmp3
            End If
        End If
        
        ' Save 'string'
        re.Pattern = "([urfb]*'.*?')"
        
        Dim match
        Dim matches_sc
        Set matches_sc = re.Execute(tmp4)
        index = 0
        For Each match In matches_sc
            tmp4 = Replace(tmp4, match.Value, "{`" & CStr(index) & "}", 1, 1)
            index = index + 1
        Next match
        
        ' Save "string"
        re.Pattern = "([urfb]*"".*?"")"
        
        Dim matches_dc
        Set matches_dc = re.Execute(tmp4)
        index = 0
        For Each match In matches_dc
            tmp4 = Replace(tmp4, match.Value, "{`" & CStr(index) & "}", 1, 1)
            index = index + 1
        Next match
        
        ' # Comment
        re.Pattern = "^(.*?)(#.*)$"
        
        If re.Test(tmp4) Then
            code = re.Replace(tmp4, "$1")
            comment = re.Replace(tmp4, "{span2}$2{endspan}")
        Else
            code = tmp4
        End If
        
        ' OPERATOR1
        re.Pattern = "(" & PYTHON_OPERATOR1 & ")"
        code = re.Replace(code, "{span3}$1{endspan}")
        
        ' OPERATOR2
        re.Pattern = "(\W+|^)(" & PYTHON_OPERATOR2 & ")(\W+)"
        code = re.Replace(code, "$1{span3}$2{endspan}$3")
        
        ' OPERATOR3
        re.Pattern = "(\W+|^)(" & PYTHON_OPERATOR3 & ")(\W+)"
        code = re.Replace(code, "$1{span3}$2{endspan}$3")
        
        ' Number(dec, float)
        re.Pattern = "([^`\da-zA-Z]+)(\d+\.*\d*)(\W+|$)"
        code = re.Replace(code, "$1{span4}$2{endspan}$3")
        
        ' Number(hex)
        re.Pattern = "(\W+)(0x)([a-fA-F0-9]+)(\W+|$)"
        code = re.Replace(code, "$1{span4}$2$3{endspan}$4")
        
        ' Number(oct)
        re.Pattern = "(\W+)(0o)([0-7]+)(\W+|$)"
        code = re.Replace(code, "$1{span4}$2$3</span>$4")
        
        ' Number(bin)
        re.Pattern = "(\W+)(0b)([0-1]+)(\W+|$)"
        code = re.Replace(code, "$1{span4}$2$3{endspan}$4")
        
        ' Decorator
        re.Pattern = "^(@)(\S+?)(\(|$)"
        code = re.Replace(code, "{span5}$1$2{endspan}$3")
        
        ' FunctionÅEClass
        re.Pattern = "(\W+|^)(def|class)([{space}]+})(\S+)(\()"
        code = re.Replace(code, "$1$2$3{span5}$4{endspan}$5")
        
        ' Syntax1
        re.Pattern = "(\W+|^)(" & PYTHON_SYNTAX1 & ")(\W+|$)"
        code = re.Replace(code, "$1{span6}$2{endspan}$3")
        
        ' "string"
        index = 0
        For Each match In matches_dc
            code = Replace(code, "{`" & CStr(index) & "}", "{span1}" & match.Value & "{endspan}", 1, 1)
            index = index + 1
        Next match
        
        index = 0
        For Each match In matches_dc
            comment = Replace(comment, "{`" & CStr(index) & "}", match.Value, 1, 1)
            index = index + 1
        Next match
        
        ' 'string'
        index = 0
        For Each match In matches_sc
            code = Replace(code, "{`" & CStr(index) & "}", "{span1}" & match.Value & "{endspan}", 1, 1)
            index = index + 1
        Next match
        
        index = 0
        For Each match In matches_sc
            comment = Replace(comment, "{`" & CStr(index) & "}", match.Value, 1, 1)
            index = index + 1
        Next match
        
        ' Join code and comments
        lines(i) = code & comment
        
        ' Post-Process
        re.Pattern = "{space}"
        lines(i) = re.Replace(lines(i), "&nbsp;")
        
        re.Pattern = "{endspan}"
        lines(i) = re.Replace(lines(i), "</span>")
        
        re.Pattern = "{span1}"
        lines(i) = re.Replace(lines(i), SPAN_STYLE1)
        
        re.Pattern = "{span2}"
        lines(i) = re.Replace(lines(i), SPAN_STYLE2)
        
        re.Pattern = "{span3}"
        lines(i) = re.Replace(lines(i), SPAN_STYLE3)
        
        re.Pattern = "{span4}"
        lines(i) = re.Replace(lines(i), SPAN_STYLE4)
        
        re.Pattern = "{span5}"
        lines(i) = re.Replace(lines(i), SPAN_STYLE5)
        
        re.Pattern = "{span6}"
        lines(i) = re.Replace(lines(i), SPAN_STYLE6)
    Next i
    
    HighLight = lines
End Function


'==============================================================================================================
' END
'==============================================================================================================
