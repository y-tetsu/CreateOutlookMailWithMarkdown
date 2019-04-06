Attribute VB_Name = "main"
'==============================================================================================================
' Create Mail with Markdown y-tetsu
'==============================================================================================================
Option Explicit

Const VERSION = "1.0.0"
Const AUTHOR = "y-tetsu"

Const SHEET_NAME = "Template"

Const FONTNAME_CELL = "B1"
Const FONTSIZE_CELL = "B2"
Const TO_CELL = "B3"
Const CC_CELL = "B4"
Const BCC_CELL = "B5"
Const SUBJECT_CELL = "B6"
Const GREETING_CELL = "B7"
Const BODY_CELL = "B8"
Const SIGNATURE_CELL = "B9"


' -------------------------------------------------------------------------------------------------------------
' Main Function
' -------------------------------------------------------------------------------------------------------------
Sub Main()
    Call CreateOutlookMailWithMarkdown.Display(SHEET_NAME, FONTNAME_CELL, FONTSIZE_CELL, TO_CELL, CC_CELL, BCC_CELL, SUBJECT_CELL, GREETING_CELL, BODY_CELL, SIGNATURE_CELL)
End Sub


'==============================================================================================================
' END
'==============================================================================================================
