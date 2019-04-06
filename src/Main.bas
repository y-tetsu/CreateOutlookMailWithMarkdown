Attribute VB_Name = "main"
'==============================================================================================================
' Create Mail with Markdown y-tetsu
'==============================================================================================================
Option Explicit

Const VERSION = "1.0.0"
Const AUTHOR = "y-tetsu"

Const FONT_NAME = "ÇlÇr ÉSÉVÉbÉN"
Const FONT_SIZE = "2"
Const SHEET_NAME = "Template"
Const TO_CELL = "B1"
Const CC_CELL = "B2"
Const BCC_CELL = "B3"
Const SUBJECT_CELL = "B4"
Const GREETING_CELL = "B5"
Const BODY_CELL = "B6"
Const SIGNATURE_CELL = "B7"


' -------------------------------------------------------------------------------------------------------------
' Main Function
' -------------------------------------------------------------------------------------------------------------
Sub Main()
    Call CreateOutlookMailWithMarkdown.Display(FONT_NAME, FONT_SIZE, SHEET_NAME, TO_CELL, CC_CELL, BCC_CELL, SUBJECT_CELL, GREETING_CELL, BODY_CELL, SIGNATURE_CELL)
End Sub


'==============================================================================================================
' END
'==============================================================================================================
