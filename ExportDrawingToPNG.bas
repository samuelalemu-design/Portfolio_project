' ==============================================================================
' SOLIDWORKS VBA Macro: Export Active Drawing Sheet to PNG Image
' Compatibility: SOLIDWORKS 2016 through SOLIDWORKS 2026+
' Language: Visual Basic for Applications (VBA 7.1 / 64-bit & 32-bit Windows)
' Description:
'   1. Verifies an active Drawing document (.slddrw) is open.
'   2. Detects active sheet paper size (A0-A4, ANSI A-E, Custom) and dimensions.
'   3. Opens native Windows Save As dialog prompting user for save location & filename.
'   4. Exports active drawing sheet as high-resolution PNG image.
'   5. Automatically opens and views saved PNG image using default Windows image viewer.
' ==============================================================================

Option Explicit

' SOLIDWORKS API Constant Definitions
Private Const swDocDRAWING As Long = 3
Private Const swSaveAsCurrentVersion As Long = 0
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swImageExportDPI As Long = 203 ' Preference ID for Image Export Resolution (DPI)
Private Const swTiffPrintAllSheets As Long = 55 ' Toggle ID for Export All Sheets vs Current Sheet Only

' ==============================================================================
' MAIN ENTRY POINT
' ==============================================================================
Public Sub main()
    On Error GoTo ErrorHandler

    ' 1. Connect to active SOLIDWORKS instance
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        MsgBox "Failed to connect to active SOLIDWORKS session.", vbCritical, "SolidWorks API Error"
        Exit Sub
    End If

    ' 2. Retrieve active document
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No document is currently open." & vbCrLf & _
               "Please open a SOLIDWORKS Drawing document (.slddrw) first.", vbExclamation, "No Active Document"
        Exit Sub
    End If

    ' 3. Verify document type is Drawing (.slddrw)
    If swModel.GetType() <> swDocDRAWING Then
        MsgBox "The active document is not a Drawing document (.slddrw)." & vbCrLf & _
               "This macro requires an active Drawing file.", vbExclamation, "Invalid Document Type"
        Exit Sub
    End If

    ' 4. Cast model to DrawingDoc and retrieve active sheet
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet()

    If swSheet Is Nothing Then
        MsgBox "Unable to access the active drawing sheet.", vbCritical, "Sheet Error"
        Exit Sub
    End If

    Dim sSheetName As String
    sSheetName = swSheet.GetName()

    ' Retrieve physical paper dimensions (returned in meters)
    Dim dPaperWidthMeters As Double
    Dim dPaperHeightMeters As Double
    swSheet.GetSize dPaperWidthMeters, dPaperHeightMeters

    ' Convert dimensions to millimeters and inches for clear reporting
    Dim dWidthMm As Double
    Dim dHeightMm As Double
    dWidthMm = Round(dPaperWidthMeters * 1000#, 1)
    dHeightMm = Round(dPaperHeightMeters * 1000#, 1)

    Dim dWidthIn As Double
    Dim dHeightIn As Double
    dWidthIn = Round(dPaperWidthMeters * 39.3700787, 2)
    dHeightIn = Round(dPaperHeightMeters * 39.3700787, 2)

    ' Get paper size standard enumeration
    Dim vSheetProps As Variant
    vSheetProps = swSheet.GetProperties2()

    Dim lPaperSizeEnum As Long
    lPaperSizeEnum = -1
    If Not IsEmpty(vSheetProps) Then
        lPaperSizeEnum = CLng(vSheetProps(0))
    End If

    ' Format paper size description string
    Dim sPaperSizeDesc As String
    sPaperSizeDesc = DeterminePaperSizeName(lPaperSizeEnum, dWidthMm, dHeightMm)

    Dim sPaperSummary As String
    sPaperSummary = "Sheet Name: " & sSheetName & vbCrLf & _
                    "Paper Standard: " & sPaperSizeDesc & vbCrLf & _
                    "Dimensions: " & dWidthMm & " mm x " & dHeightMm & " mm (" & dWidthIn & """ x " & dHeightIn & """)"

    ' 5. Determine initial directory & suggested filename
    Dim sDocPath As String
    sDocPath = swModel.GetPathName()

    Dim sInitialFolder As String
    Dim sDefaultFileName As String

    If sDocPath <> "" Then
        sInitialFolder = Left$(sDocPath, InStrRev(sDocPath, "\"))
        Dim sBaseName As String
        sBaseName = Mid$(sDocPath, InStrRev(sDocPath, "\") + 1)
        If InStrRev(sBaseName, ".") > 0 Then
            sBaseName = Left$(sBaseName, InStrRev(sBaseName, ".") - 1)
        End If
        sDefaultFileName = sBaseName & "_" & sSheetName & ".png"
    Else
        ' Fallback for unsaved unsaved drawings
        sInitialFolder = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
        sDefaultFileName = SanitiseFileName(swModel.GetTitle()) & "_" & sSheetName & ".png"
    End If

    sDefaultFileName = SanitiseFileName(sDefaultFileName)

    ' 6. Prompt user for save destination via native Save File Dialog
    Dim sSelectedSavePath As String
    sSelectedSavePath = ShowSaveAsDialog(sInitialFolder, sDefaultFileName)
    sSelectedSavePath = Trim$(sSelectedSavePath)

    ' If user cancels dialog, stop execution gracefully
    If sSelectedSavePath = "" Then
        Exit Sub
    End If

    ' Ensure .png extension
    If LCase$(Right$(sSelectedSavePath, 4)) <> ".png" Then
        sSelectedSavePath = sSelectedSavePath & ".png"
    End If

    ' Ensure target directory exists on disk
    Dim sTargetFolder As String
    sTargetFolder = Left$(sSelectedSavePath, InStrRev(sSelectedSavePath, "\"))
    EnsureFolderStructure sTargetFolder

    ' 7. Configure SolidWorks PNG export preferences (Set DPI to 300 and export active sheet only)
    On Error Resume Next
    swApp.SetUserPreferenceIntegerValue swImageExportDPI, 300
    swApp.SetUserPreferenceToggle swTiffPrintAllSheets, False
    On Error GoTo ErrorHandler

    ' 8. Export active drawing to PNG
    Dim swModelExt As SldWorks.ModelDocExtension
    Set swModelExt = swModel.Extension

    Dim lErrors As Long
    Dim lWarnings As Long
    Dim bSuccess As Boolean

    bSuccess = swModelExt.SaveAs3( _
        sSelectedSavePath, _
        swSaveAsCurrentVersion, _
        swSaveAsOptions_Silent, _
        Nothing, _
        Nothing, _
        lErrors, _
        lWarnings)

    ' 9. Verify physical file creation & notify/view image
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If bSuccess And fso.FileExists(sSelectedSavePath) Then
        ' View image immediately using Windows default image viewer
        OpenImageInDefaultViewer sSelectedSavePath

        MsgBox "Drawing successfully exported to PNG!" & vbCrLf & vbCrLf & _
               sPaperSummary & vbCrLf & vbCrLf & _
               "File Location:" & vbCrLf & sSelectedSavePath, _
               vbInformation, "Export Complete"
    Else
        MsgBox "Failed to export drawing sheet to PNG image." & vbCrLf & _
               "Error Code: " & lErrors & vbCrLf & _
               "Target Path: " & sSelectedSavePath, _
               vbCritical, "Export Failed"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred during macro execution:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & " (Line " & Erl & ")", _
           vbCritical, "Macro Execution Error"
End Sub

' ==============================================================================
' Helper Function: DeterminePaperSizeName
' Purpose: Resolves paper size standard name from SW enum & physical dimensions
' ==============================================================================
Private Function DeterminePaperSizeName(ByVal lPaperEnum As Long, ByVal dWidthMm As Double, ByVal dHeightMm As Double) As String
    Select Case lPaperEnum
        Case 0: DeterminePaperSizeName = "ANSI A (Horizontal - 11"" x 8.5"")"
        Case 1: DeterminePaperSizeName = "ANSI A (Vertical - 8.5"" x 11"")"
        Case 2: DeterminePaperSizeName = "ANSI B (17"" x 11"")"
        Case 3: DeterminePaperSizeName = "ANSI C (22"" x 17"")"
        Case 4: DeterminePaperSizeName = "ANSI D (34"" x 22"")"
        Case 5: DeterminePaperSizeName = "ANSI E (44"" x 34"")"
        Case 6: DeterminePaperSizeName = "ISO A4 (Landscape - 297mm x 210mm)"
        Case 7: DeterminePaperSizeName = "ISO A4 (Portrait - 210mm x 297mm)"
        Case 8: DeterminePaperSizeName = "ISO A3 (420mm x 297mm)"
        Case 9: DeterminePaperSizeName = "ISO A2 (594mm x 420mm)"
        Case 10: DeterminePaperSizeName = "ISO A1 (841mm x 594mm)"
        Case 11: DeterminePaperSizeName = "ISO A0 (1189mm x 841mm)"
        Case 12: DeterminePaperSizeName = "Custom Sheet Size"
        Case Else
            ' Dimensional fallback comparison (+/- 5mm tolerance)
            If (Abs(dWidthMm - 297) < 5 And Abs(dHeightMm - 210) < 5) Or (Abs(dWidthMm - 210) < 5 And Abs(dHeightMm - 297) < 5) Then
                DeterminePaperSizeName = "ISO A4"
            ElseIf (Abs(dWidthMm - 420) < 5 And Abs(dHeightMm - 297) < 5) Or (Abs(dWidthMm - 297) < 5 And Abs(dHeightMm - 420) < 5) Then
                DeterminePaperSizeName = "ISO A3"
            ElseIf (Abs(dWidthMm - 594) < 5 And Abs(dHeightMm - 420) < 5) Or (Abs(dWidthMm - 420) < 5 And Abs(dHeightMm - 594) < 5) Then
                DeterminePaperSizeName = "ISO A2"
            ElseIf (Abs(dWidthMm - 841) < 5 And Abs(dHeightMm - 594) < 5) Or (Abs(dWidthMm - 594) < 5 And Abs(dHeightMm - 841) < 5) Then
                DeterminePaperSizeName = "ISO A1"
            ElseIf (Abs(dWidthMm - 1189) < 5 And Abs(dHeightMm - 841) < 5) Or (Abs(dWidthMm - 841) < 5 And Abs(dHeightMm - 1189) < 5) Then
                DeterminePaperSizeName = "ISO A0"
            Else
                DeterminePaperSizeName = "Custom Paper Size"
            End If
    End Select
End Function

' ==============================================================================
' Helper Sub: OpenImageInDefaultViewer
' Purpose: Launches the default Windows image viewer application for the PNG
' ==============================================================================
Private Sub OpenImageInDefaultViewer(ByVal sFilePath As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ' Shell execute to open image with system default application
    wsh.Run """" & sFilePath & """", 1, False
End Sub

' ==============================================================================
' Helper Function: SanitiseFileName
' Purpose: Removes illegal file path characters (\ / : * ? " < > |)
' ==============================================================================
Private Function SanitiseFileName(ByVal sFileName As String) As String
    Dim sBadChars As String
    sBadChars = "\/:*?""<>|"
    Dim i As Long
    For i = 1 To Len(sBadChars)
        sFileName = Replace(sFileName, Mid$(sBadChars, i, 1), "_")
    Next i
    SanitiseFileName = sFileName
End Function

' ==============================================================================
' Helper Sub: EnsureFolderStructure
' Purpose: Recursively creates path directories if missing
' ==============================================================================
Private Sub EnsureFolderStructure(ByVal sFolderPath As String)
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
            EnsureFolderStructure sParent
        End If
        fso.CreateFolder sFolderPath
    End If
End Sub

' ==============================================================================
' Helper Function: ShowSaveAsDialog
' Purpose: Opens native Windows Save As dialog using PowerShell SaveFileDialog
' Returns: Full chosen file path string or empty string if cancelled
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
           "$d.Filter = 'PNG Image (*.png)|*.png|All Files (*.*)|*.*'; " & _
           "$d.InitialDirectory = '" & sCleanFolder & "'; " & _
           "$d.FileName = '" & sCleanFile & "'; " & _
           "$d.Title = 'Export Active Drawing to PNG - Select Destination'; " & _
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
