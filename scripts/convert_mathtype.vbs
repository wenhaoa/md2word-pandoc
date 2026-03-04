' convert_mathtype.vbs - Convert OMML equations to MathType objects
' Uses MathType's Word add-in ConvertEquations with correct parameters
'
' Usage:
'   cscript //nologo convert_mathtype.vbs <docx_path>

Option Explicit

Dim objWord, objDoc
Dim docPath, fso

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript //nologo convert_mathtype.vbs <docx_path>"
    WScript.Quit 1
End If

Set fso = CreateObject("Scripting.FileSystemObject")
docPath = fso.GetAbsolutePathName(WScript.Arguments(0))

If Not fso.FileExists(docPath) Then
    WScript.Echo "Error: File not found: " & docPath
    WScript.Quit 1
End If

On Error Resume Next

' Create Word instance
Set objWord = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot create Word instance: " & Err.Description
    WScript.Quit 1
End If

objWord.Visible = False
objWord.DisplayAlerts = 0

' Ensure MathType add-in is loaded
' WHY: MathType installs to STARTUP, should auto-load, but force it just in case
Dim mtPaths(2)
mtPaths(0) = "C:\Program Files (x86)\MathType\Office Support\64\MathType Commands 2016.dotm"
mtPaths(1) = "C:\Program Files (x86)\MathType\Office Support\64\MathType Commands 2013.dotm"
mtPaths(2) = "C:\Program Files (x86)\MathType\Office Support\32\MathType Commands.dotm"

Dim i
For i = 0 To 2
    If fso.FileExists(mtPaths(i)) Then
        Err.Clear
        objWord.AddIns.Add mtPaths(i), True
    End If
Next
Err.Clear

' Open document
Set objDoc = objWord.Documents.Open(docPath)
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot open document: " & Err.Description
    objWord.Quit False
    WScript.Quit 1
End If

' Count OMML equations
Dim omathCount
omathCount = objDoc.OMaths.Count
WScript.Echo "  OMML equations: " & omathCount

If omathCount = 0 Then
    WScript.Echo "  No equations to convert"
    objDoc.Close False
    objWord.Quit False
    WScript.Quit 0
End If

' Strategy: Convert each OMML equation by selecting it and using
' MathType's "Insert Equation" which replaces selection with MathType object
' WHY: Direct macro ConvertEquations requires SDK info structure,
' but selecting OMML and running Insert will use MathType to open/convert it

Dim converted
converted = False

' Method 1: Try ConvertEquations with variant parameter (some versions accept Empty)
Dim templateNames(1)
templateNames(0) = "MathType Commands 2016"
templateNames(1) = "MathType Commands 2013"

Dim j
For j = 0 To 1
    Err.Clear
    objWord.Run templateNames(j) & ".dotm!ConvertEquations", 1
    If Err.Number = 0 Then
        WScript.Echo "  OK: ConvertEquations with param=1"
        converted = True
        Exit For
    End If
    
    Err.Clear
    objWord.Run templateNames(j) & ".dotm!ConvertEquations", Empty
    If Err.Number = 0 Then
        WScript.Echo "  OK: ConvertEquations with Empty param"
        converted = True
        Exit For
    End If
Next

' Method 2: Try to find and execute the ribbon command
If Not converted Then
    WScript.Echo "  Trying ribbon command method..."
    Err.Clear
    
    ' Try using CommandBars to find MathType Convert button
    Dim bar, ctrl
    For Each bar In objWord.CommandBars
        For Each ctrl In bar.Controls
            If InStr(ctrl.Caption, "Convert") > 0 And InStr(ctrl.Caption, "Equation") > 0 Then
                WScript.Echo "  Found button: " & ctrl.Caption
                Err.Clear
                ctrl.Execute
                If Err.Number = 0 Then
                    WScript.Echo "  OK: Executed Convert Equations button"
                    converted = True
                    Exit For
                End If
            End If
        Next
        If converted Then Exit For
    Next
End If

' Method 3: Select all and use Toggle (converts OMML <-> MathType)
If Not converted Then
    WScript.Echo "  Trying Toggle method on each equation..."
    Err.Clear
    
    ' Select entire document
    objDoc.Content.Select
    
    For j = 0 To 1
        Err.Clear
        objWord.Run templateNames(j) & ".dotm!ToggleMathType"
        If Err.Number = 0 Then
            WScript.Echo "  OK: ToggleMathType"
            converted = True
            Exit For
        End If
    Next
End If

' Method 4: Enable VBA trust and try accessing macros list
If Not converted Then
    WScript.Echo "  All automatic methods failed."
    WScript.Echo "  Please convert manually in Word:"
    WScript.Echo "    MathType tab > Convert Equations"
End If

' Save and close
Err.Clear
If converted Then
    objDoc.Save
End If
objDoc.Close False
objWord.Quit False

Set objDoc = Nothing
Set objWord = Nothing
Set fso = Nothing

If converted Then
    WScript.Echo "  MathType conversion complete"
    WScript.Quit 0
Else
    WScript.Quit 1
End If
