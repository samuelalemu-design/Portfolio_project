' ==============================================================================
' SOLIDWORKS 2026 SP2.1 VBA Macro - Sheet Metal Flat Pattern DXF Exporter
' Compatibility: SOLIDWORKS Design 2026 SP2.1 (.swp compatible)
' Language: Visual Basic for Applications (VBA 7.1 / 64-bit & 32-bit Windows)
' Author: Senior SOLIDWORKS API Developer
' Description: Master-level SOLIDWORKS API macro to programmatically export Sheet 
'              Metal Flat Patterns directly to DXF with 3mm bend mark ticks.
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' SOLIDWORKS 2026 API CONSTANTS & ENUMERATIONS
' ------------------------------------------------------------------------------
Private Const swDocPART As Long = 1
Private Const swExportToDWG_ExportSheetMetal As Long = 1

' Options Bitmask Flags for Sheet Metal Flat Pattern DXF:
' Bit 0 (1) : Flat Pattern Geometry = YES
' Bit 2 (4) : Bend Lines = YES (4)
' Bit 4 (16): Merge Coplanar Faces = YES (16)
' Bitmask Total = 1 + 4 + 16 = 21
Private Const OPTION_FLAT_PATTERN_GEOMETRY As Long = 1
Private Const OPTION_BEND_LINES As Long = 4
Private Const OPTION_MERGE_COPLANAR_FACES As Long = 16

' ==============================================================================
' MAIN MACRO ENTRY POINT
' ==============================================================================
Public Sub main()
    On Error GoTo ErrorHandler

10: Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        MsgBox "Failed to connect to SOLIDWORKS session.", vbCritical, "API Error"
        Exit Sub
    End If

20: Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    ' Requirement 2: Display clear error and stop if no active document
    If swModel Is Nothing Then
        MsgBox "No document is currently open. Please open a Sheet Metal Part file (.SLDPRT).", vbCritical, "No Active Document"
        Exit Sub
    End If

    ' Requirement 1 & 2: Ensure document is a Part (.SLDPRT)
30: If swModel.GetType() <> swDocPART Then
        MsgBox "The active document is not a Part document (.SLDPRT)." & vbCrLf & _
               "This macro only works on Sheet Metal Parts.", vbCritical, "Invalid Document Type"
        Exit Sub
    End If

    ' Requirement 11: Display specific message if part has never been saved
40: Dim sModelPath As String
    sModelPath = swModel.GetPathName()

    If Trim$(sModelPath) = "" Then
        MsgBox "Please save the part before exporting.", vbExclamation, "Part Not Saved"
        Exit Sub
    End If

    ' --------------------------------------------------------------------------
    ' SECTION 2: Validate Sheet Metal Document & Find Flat Pattern Feature
    ' --------------------------------------------------------------------------
50: Dim swPart As SldWorks.PartDoc
    Set swPart = swModel

60: Dim swFeat As SldWorks.Feature
    Dim swFlatPatternFeat As SldWorks.Feature
    Dim bIsSheetMetal As Boolean
    bIsSheetMetal = False

    Set swFeat = swModel.FirstFeature()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2() = "FlatPattern" Then
            bIsSheetMetal = True
            Set swFlatPatternFeat = swFeat
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature()
    Loop

70: If Not bIsSheetMetal Then
        MsgBox "The active part document does not contain a Sheet Metal Flat Pattern feature." & vbCrLf & _
               "This macro only works on Sheet Metal Parts.", vbCritical, "Not a Sheet Metal Part"
        Exit Sub
    End If

    ' Requirement 9: Rebuild model prior to export
80: swModel.EditRebuild3

    ' --------------------------------------------------------------------------
    ' SECTION 3: Convert Bend Lines to 3mm End Marks in Sketch
    ' --------------------------------------------------------------------------
90: Dim bModifiedBendLines As Boolean
    bModifiedBendLines = ConvertBendLinesTo3mmMarks(swModel, swFlatPatternFeat)

    ' --------------------------------------------------------------------------
    ' SECTION 4: File Path Preparation & Windows "Save As" Dialog
    ' --------------------------------------------------------------------------
100: Dim sInitialFolder As String
    Dim sPartFileName As String
    Dim sSuggestedDxfName As String

    sInitialFolder = Left$(sModelPath, InStrRev(sModelPath, "\"))
    sPartFileName = Mid$(sModelPath, InStrRev(sModelPath, "\") + 1)
    
    If InStrRev(sPartFileName, ".") > 0 Then
        sPartFileName = Left$(sPartFileName, InStrRev(sPartFileName, ".") - 1)
    End If

    sSuggestedDxfName = sPartFileName & ".dxf"

110: Dim sSelectedDxfPath As String
    sSelectedDxfPath = ShowSaveAsDialog(sInitialFolder, sSuggestedDxfName)
    sSelectedDxfPath = Replace(sSelectedDxfPath, vbCr, "")
    sSelectedDxfPath = Replace(sSelectedDxfPath, vbLf, "")
    sSelectedDxfPath = Trim$(sSelectedDxfPath)

    If sSelectedDxfPath = "" Then
        If bModifiedBendLines Then swModel.EditUndo2 1
        Exit Sub
    End If

    ' Ensure path has .dxf extension
    If LCase$(Right$(sSelectedDxfPath, 4)) <> ".dxf" Then
        sSelectedDxfPath = sSelectedDxfPath & ".dxf"
    End If

    ' Guarantee destination directory exists on disk before saving
    Dim sTargetFolder As String
    sTargetFolder = Left$(sSelectedDxfPath, InStrRev(sSelectedDxfPath, "\"))
    EnsureFolderExists sTargetFolder

    ' Delete existing destination file if present to guarantee clean write
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sSelectedDxfPath) Then
        On Error Resume Next
        fso.DeleteFile sSelectedDxfPath, True
        On Error GoTo ErrorHandler
    End If

    ' --------------------------------------------------------------------------
    ' SECTION 5: Export Flat Pattern to DXF
    ' --------------------------------------------------------------------------
120: Dim lSheetMetalOptions As Long
    lSheetMetalOptions = OPTION_FLAT_PATTERN_GEOMETRY + OPTION_BEND_LINES + OPTION_MERGE_COPLANAR_FACES

    Dim dataAlignment(0 To 11) As Double
    dataAlignment(0) = 0#
    dataAlignment(1) = 0#
    dataAlignment(2) = 0#
    dataAlignment(3) = 1#
    dataAlignment(4) = 0#
    dataAlignment(5) = 0#
    dataAlignment(6) = 0#
    dataAlignment(7) = 1#
    dataAlignment(8) = 0#
    dataAlignment(9) = 0#
    dataAlignment(10) = 0#
    dataAlignment(11) = 1#

    Dim varAlignment As Variant
    varAlignment = dataAlignment
    Dim varTarget As Variant
    varTarget = dataAlignment

    Dim dataManifest(0 To 0) As String
    dataManifest(0) = ""
    Dim varManifest As Variant
    varManifest = dataManifest

    Dim bExportSuccess As Boolean
    bExportSuccess = False

    On Error Resume Next
    bExportSuccess = swPart.ExportToDWG2( _
        sSelectedDxfPath, _
        sModelPath, _
        swExportToDWG_ExportSheetMetal, _
        True, _
        varAlignment, _
        False, _
        varTarget, _
        lSheetMetalOptions, _
        varManifest)
    On Error GoTo ErrorHandler

    If fso.FileExists(sSelectedDxfPath) Then
        bExportSuccess = True
    End If

    If Not bExportSuccess Then
        Dim varEmptyAlignment As Variant
        varEmptyAlignment = Empty
        On Error Resume Next
        swPart.ExportToDWG2 _
            sSelectedDxfPath, _
            sModelPath, _
            swExportToDWG_ExportSheetMetal, _
            True, _
            varEmptyAlignment, _
            False, _
            varEmptyAlignment, _
            lSheetMetalOptions, _
            varManifest
        On Error GoTo ErrorHandler

        If fso.FileExists(sSelectedDxfPath) Then
            bExportSuccess = True
        End If
    End If

    ' Restore original Bend-Lines sketch in SolidWorks
    If bModifiedBendLines Then
        On Error Resume Next
        swModel.EditUndo2 1
        On Error GoTo ErrorHandler
    End If

    ' --------------------------------------------------------------------------
    ' SECTION 6: Physical File Verification & User Notification
    ' --------------------------------------------------------------------------
140: If fso.FileExists(sSelectedDxfPath) Then
        MsgBox "DXF exported successfully with 3mm bend marks:" & vbCrLf & sSelectedDxfPath, vbInformation, "Export Complete"
    Else
        MsgBox "DXF export failed. Could not write DXF file to:" & vbCrLf & sSelectedDxfPath, vbCritical, "Export Failed"
    End If

    Exit Sub

ErrorHandler:
    If bModifiedBendLines Then
        On Error Resume Next
        swModel.EditUndo2 1
    End If

    MsgBox "An unexpected error occurred during execution:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & " (Line " & Erl & ")", vbCritical, "Macro Execution Error"
End Sub

' ==============================================================================
' Helper Function: ConvertBendLinesTo3mmMarks
' Purpose: Replaces full bend lines in the FlatPattern Bend-Lines sketch with
'          3mm tick mark lines at both endpoints.
' ==============================================================================
Private Function ConvertBendLinesTo3mmMarks(ByVal swModel As SldWorks.ModelDoc2, ByVal swFlatPatternFeat As SldWorks.Feature) As Boolean
    On Error GoTo ErrorHandler

    Dim swSubFeat As SldWorks.Feature
    Set swSubFeat = swFlatPatternFeat.GetFirstSubFeature()

    Dim swBendFeat As SldWorks.Feature
    Dim swBendSketch As SldWorks.Sketch
    Dim vSegments As Variant

    Do While Not swSubFeat Is Nothing
        Dim sName As String
        sName = LCase$(swSubFeat.GetNameForSelection())
        If InStr(sName, "bend") > 0 Or swSubFeat.GetTypeName2() = "ProfileFeature" Or swSubFeat.GetTypeName2() = "SubSketch" Then
            Set swBendSketch = swSubFeat.GetSpecificFeature2()
            If Not swBendSketch Is Nothing Then
                vSegments = swBendSketch.GetSketchSegments()
                If Not IsEmpty(vSegments) Then
                    Set swBendFeat = swSubFeat
                    Exit Do
                End If
            End If
        End If
        Set swSubFeat = swSubFeat.GetNextSubFeature()
    Loop

    If IsEmpty(vSegments) Or swBendFeat Is Nothing Then
        ConvertBendLinesTo3mmMarks = False
        Exit Function
    End If

    Dim countLines As Long
    countLines = UBound(vSegments) + 1

    Dim ptsX1() As Double
    Dim ptsY1() As Double
    Dim ptsX2() As Double
    Dim ptsY2() As Double

    ReDim ptsX1(0 To countLines - 1)
    ReDim ptsY1(0 To countLines - 1)
    ReDim ptsX2(0 To countLines - 1)
    ReDim ptsY2(0 To countLines - 1)

    Dim i As Long
    Dim swSketchSeg As SldWorks.SketchSegment
    Dim swLine As SldWorks.SketchLine
    Dim swStartPt As SldWorks.SketchPoint
    Dim swEndPt As SldWorks.SketchPoint

    For i = 0 To countLines - 1
        Set swSketchSeg = vSegments(i)
        If swSketchSeg.GetType() = 0 Then ' swSketchLINE = 0
            Set swLine = swSketchSeg
            Set swStartPt = swLine.GetStartPoint2()
            Set swEndPt = swLine.GetEndPoint2()

            ptsX1(i) = swStartPt.X
            ptsY1(i) = swStartPt.Y
            ptsX2(i) = swEndPt.X
            ptsY2(i) = swEndPt.Y
        End If
    Next i

    ' Edit Bend-Lines sketch
    swModel.ClearSelection2 True
    swBendFeat.Select2 False, 0
    swModel.EditSketch

    ' Delete full bend lines
    For i = 0 To countLines - 1
        Set swSketchSeg = vSegments(i)
        If Not swSketchSeg Is Nothing Then
            swModel.ClearSelection2 True
            swSketchSeg.Select2 False, 0
            swModel.EditDelete
        End If
    Next i

    ' Create 3mm end mark lines in local sketch coordinates
    Dim markLen As Double
    markLen = 0.003 ' 3mm in meters

    For i = 0 To countLines - 1
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        x1 = ptsX1(i): y1 = ptsY1(i)
        x2 = ptsX2(i): y2 = ptsY2(i)

        Dim dx As Double, dy As Double, L As Double
        dx = x2 - x1
        dy = y2 - y1
        L = Sqr(dx * dx + dy * dy)

        If L > 0.0001 Then
            Dim ux As Double, uy As Double
            ux = dx / L
            uy = dy / L

            Dim actualLen As Double
            If L >= (markLen * 2#) Then
                actualLen = markLen
            Else
                actualLen = L / 2#
            End If

            ' Draw 3mm line at start point
            swModel.CreateLine2D x1, y1, x1 + (ux * actualLen), y1 + (uy * actualLen)

            ' Draw 3mm line at end point
            swModel.CreateLine2D x2 - (ux * actualLen), y2 - (uy * actualLen), x2, y2
        End If
    Next i

    ' Close sketch edit
    swModel.InsertSketch2 True
    ConvertBendLinesTo3mmMarks = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    swModel.InsertSketch2 True
    ConvertBendLinesTo3mmMarks = False
End Function

' ==============================================================================
' Helper Function: EnsureFolderExists
' Purpose: Recursively creates parent directory structure if it does not exist
' ==============================================================================
Private Sub EnsureFolderExists(ByVal sFolderPath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Right$(sFolderPath, 1) = "\" Then
        sFolderPath = Left$(sFolderPath, Len(sFolderPath) - 1)
    End If

    If sFolderPath <> "" And Not fso.FolderExists(sFolderPath) Then
        Dim sParent As String
        sParent = fso.GetParentFolderName(sFolderPath)
        If sParent <> "" And Not fso.FolderExists(sParent) Then
            EnsureFolderExists sParent
        End If
        fso.CreateFolder sFolderPath
    End If
End Sub

' ==============================================================================
' Helper Function: ShowSaveAsDialog
' Purpose: Displays native Windows "Save As" dialog using PowerShell SaveFileDialog
' Returns: Clean destination file path (stripped of control chars) or empty string if cancelled
' ==============================================================================
Private Function ShowSaveAsDialog(ByVal sInitialDir As String, ByVal sDefaultFile As String) As String
    On Error GoTo DialogFallback

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim sCleanFolder As String
    Dim sCleanFile As String
    sCleanFolder = Replace(sInitialDir, "'", "''")
    sCleanFile = Replace(sDefaultFile, "'", "''")

    Dim sCmd As String
    sCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command """ & _
           "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; " & _
           "$d = New-Object System.Windows.Forms.SaveFileDialog; " & _
           "$d.Filter = 'DXF Files (*.dxf)|*.dxf|All Files (*.*)|*.*'; " & _
           "$d.InitialDirectory = '" & sCleanFolder & "'; " & _
           "$d.FileName = '" & sCleanFile & "'; " & _
           "$d.Title = 'Save Flat Pattern DXF As'; " & _
           "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Host $d.FileName }"""

    Dim execObj As Object
    Set execObj = sh.Exec(sCmd)

    Dim sResult As String
    sResult = execObj.StdOut.ReadAll()
    sResult = Replace(sResult, vbCr, "")
    sResult = Replace(sResult, vbLf, "")
    sResult = Trim$(sResult)

    ShowSaveAsDialog = sResult
    Exit Function

DialogFallback:
    ShowSaveAsDialog = sInitialDir & sDefaultFile
End Function
