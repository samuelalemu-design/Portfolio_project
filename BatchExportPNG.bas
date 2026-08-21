' ==============================================================================
' SOLIDWORKS VBA Macro: Batch Export Drawings / Models to High-Res PNG Images
' Compatibility: SOLIDWORKS 2016 through SOLIDWORKS 2026+
' Language: Visual Basic for Applications (VBA 7.1 / 64-bit & 32-bit Windows)
' Description:
'   1. Interactively prompts user to batch export either:
'      - All CAD documents in a chosen folder (Drawings, Parts, Assemblies)
'      - All currently open documents in the active SOLIDWORKS session
'   2. Automatically iterates through all drawing sheets for multi-sheet drawings.
'   3. Sets export resolution to 300 DPI for high-quality image rendering.
'   4. Saves output PNGs to a dedicated "PNG_Exports" subfolder or user target folder.
'   5. Displays execution summary and offers to open the export directory in Explorer.
' ==============================================================================

Option Explicit

' SOLIDWORKS API Constant Definitions
Private Const swDocNONE As Long = 0
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3

Private Const swSaveAsCurrentVersion As Long = 0
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swOpenDocOptions_Silent As Long = 1
Private Const swOpenDocOptions_ReadOnly As Long = 2
Private Const swImageExportDPI As Long = 203 ' Preference ID for Image Export Resolution (DPI)
Private Const swTiffPrintAllSheets As Long = 55 ' Toggle ID for Export All Sheets vs Current Sheet Only

' Global counters for reporting
Private g_lSuccessCount As Long
Private g_lFailCount As Long
Private g_lSkippedCount As Long

' ==============================================================================
' MAIN ENTRY POINT
' ==============================================================================
Public Sub main()
    On Error GoTo ErrorHandler

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        MsgBox "Failed to connect to active SOLIDWORKS session.", vbCritical, "SolidWorks API Error"
        Exit Sub
    End If

    ' Reset counters
    g_lSuccessCount = 0
    g_lFailCount = 0
    g_lSkippedCount = 0

    ' Configure SolidWorks PNG export preferences (300 DPI for high resolution, active sheet export only)
    On Error Resume Next
    swApp.SetUserPreferenceIntegerValue swImageExportDPI, 300
    swApp.SetUserPreferenceToggle swTiffPrintAllSheets, False
    On Error GoTo ErrorHandler

    ' Present user selection dialog (Folder Batch vs Currently Open Files Batch)
    Dim sPrompt As String
    sPrompt = "Select Batch Export Mode:" & vbCrLf & vbCrLf & _
              "[ Yes ]   - Export all files in a Folder (Drawings .slddrw, Parts .sldprt, Assemblies .sldasm)" & vbCrLf & _
              "[ No ]    - Export all Currently Open Documents in SOLIDWORKS" & vbCrLf & _
              "[ Cancel ] - Exit Macro"

    Dim nChoice As VbMsgBoxResult
    nChoice = MsgBox(sPrompt, vbYesNoCancel + vbQuestion + vbDefaultButton1, "SOLIDWORKS Batch PNG Exporter")

    If nChoice = vbCancel Then
        Exit Sub
    ElseIf nChoice = vbYes Then
        ' Mode 1: Folder Batch
        ProcessFolderBatch swApp
    ElseIf nChoice = vbNo Then
        ' Mode 2: Currently Open Documents Batch
        ProcessOpenDocsBatch swApp
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred during macro execution:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & " (Line " & Erl & ")", _
           vbCritical, "Macro Execution Error"
End Sub

' ==============================================================================
' Sub Routine: ProcessFolderBatch
' Purpose: Prompts user for source folder, scans for SW files, and exports PNGs
' ==============================================================================
Private Sub ProcessFolderBatch(ByRef swApp As SldWorks.SldWorks)
    On Error GoTo ErrorHandler

    ' Prompt user to select source directory
    Dim sSourceFolder As String
    sSourceFolder = BrowseForFolder("Select Folder Containing SOLIDWORKS Files to Export")

    If sSourceFolder = "" Then Exit Sub
    If Right$(sSourceFolder, 1) <> "\" Then sSourceFolder = sSourceFolder & "\"

    ' Ask if user wants to process subfolders recursively
    Dim nRecurseChoice As VbMsgBoxResult
    nRecurseChoice = MsgBox("Include subfolders recursively?", vbYesNo + vbQuestion, "Subfolder Option")
    Dim bIncludeSubfolders As Boolean
    bIncludeSubfolders = (nRecurseChoice = vbYes)

    ' Output directory path
    Dim sOutputFolder As String
    sOutputFolder = sSourceFolder & "PNG_Exports\"
    EnsureFolderStructure sOutputFolder

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim sourceFolderObj As Object
    Set sourceFolderObj = fso.GetFolder(sSourceFolder)

    ' Gather list of matching files
    Dim fileList As Collection
    Set fileList = New Collection

    CollectCADFiles fso, sourceFolderObj, bIncludeSubfolders, fileList

    If fileList.Count = 0 Then
        MsgBox "No SOLIDWORKS drawing (.slddrw), part (.sldprt), or assembly (.sldasm) files found in:" & vbCrLf & sSourceFolder, _
               vbInformation, "No Files Found"
        Exit Sub
    End If

    ' Process each file
    Dim i As Long
    Dim sFilePath As String
    Dim swModel As SldWorks.ModelDoc2
    Dim lErrors As Long
    Dim lWarnings As Long
    Dim nDocType As Long
    Dim sExt As String

    For i = 1 To fileList.Count
        sFilePath = CStr(fileList.Item(i))
        sExt = LCase$(fso.GetExtensionName(sFilePath))

        ' Update status bar
        swApp.SendMsgToUser2 "Processing " & i & " of " & fileList.Count & ": " & fso.GetFileName(sFilePath), swMbInformation, swMbOk

        ' Determine document type
        Select Case sExt
            Case "slddrw": nDocType = swDocDRAWING
            Case "sldprt": nDocType = swDocPART
            Case "sldasm": nDocType = swDocASSEMBLY
            Case Else: nDocType = swDocNONE
        End Select

        If nDocType <> swDocNONE Then
            ' Open document silently
            Set swModel = swApp.OpenDoc6(sFilePath, nDocType, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", lErrors, lWarnings)

            If Not swModel Is Nothing Then
                ExportDocumentPNG swApp, swModel, sOutputFolder
                ' Close document after exporting to conserve memory
                swApp.CloseDoc swModel.GetTitle()
            Else
                g_lFailCount = g_lFailCount + 1
            End If
        End If
    Next i

    ' Display summary and prompt to open export folder
    DisplaySummaryAndOpenFolder sOutputFolder, fileList.Count

    Exit Sub

ErrorHandler:
    MsgBox "Error during folder batch processing: " & Err.Description, vbCritical, "Batch Error"
End Sub

' ==============================================================================
' Sub Routine: CollectCADFiles
' Purpose: Recursively or flatly gathers .slddrw, .sldprt, .sldasm file paths
' ==============================================================================
Private Sub CollectCADFiles(ByRef fso As Object, ByRef folderObj As Object, ByVal bRecurse As Boolean, ByRef fileList As Collection)
    Dim fileObj As Object
    Dim sExt As String

    For Each fileObj In folderObj.Files
        sExt = LCase$(fso.GetExtensionName(fileObj.Path))
        If sExt = "slddrw" Or sExt = "sldprt" Or sExt = "sldasm" Then
            ' Skip temporary ~$ files created by SOLIDWORKS
            If Left$(fileObj.Name, 2) <> "~$" Then
                fileList.Add fileObj.Path
            End If
        End If
    Next fileObj

    If bRecurse Then
        Dim subFolderObj As Object
        For Each subFolderObj In folderObj.SubFolders
            ' Skip the PNG_Exports directory to avoid endless recursion
            If LCase$(subFolderObj.Name) <> "png_exports" Then
                CollectCADFiles fso, subFolderObj, True, fileList
            End If
        Next subFolderObj
    End If
End Sub

' ==============================================================================
' Sub Routine: ProcessOpenDocsBatch
' Purpose: Iterates through all currently open documents in SOLIDWORKS
' ==============================================================================
Private Sub ProcessOpenDocsBatch(ByRef swApp As SldWorks.SldWorks)
    On Error GoTo ErrorHandler

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.GetFirstDocument()

    If swModel Is Nothing Then
        MsgBox "No documents are currently open in SOLIDWORKS.", vbExclamation, "No Open Documents"
        Exit Sub
    End If

    ' Ask for destination folder
    Dim sOutputFolder As String
    sOutputFolder = BrowseForFolder("Select Destination Folder for Exported PNG Images")

    If sOutputFolder = "" Then Exit Sub
    If Right$(sOutputFolder, 1) <> "\" Then sOutputFolder = sOutputFolder & "\"
    sOutputFolder = sOutputFolder & "PNG_Exports\"
    EnsureFolderStructure sOutputFolder

    Dim lTotalOpen As Long
    lTotalOpen = 0

    Do While Not swModel Is Nothing
        lTotalOpen = lTotalOpen + 1
        ExportDocumentPNG swApp, swModel, sOutputFolder
        Set swModel = swModel.GetNext()
    Loop

    ' Display summary and prompt to open export folder
    DisplaySummaryAndOpenFolder sOutputFolder, lTotalOpen

    Exit Sub

ErrorHandler:
    MsgBox "Error during open documents batch processing: " & Err.Description, vbCritical, "Batch Error"
End Sub

' ==============================================================================
' Sub Routine: ExportDocumentPNG
' Purpose: Handles exporting individual model doc (Drawing sheet loop or Part/Assem)
' ==============================================================================
Private Sub ExportDocumentPNG(ByRef swApp As SldWorks.SldWorks, ByRef swModel As SldWorks.ModelDoc2, ByVal sOutputFolder As String)
    On Error GoTo ErrorHandler

    Dim sDocPath As String
    sDocPath = swModel.GetPathName()

    Dim sBaseName As String
    If sDocPath <> "" Then
        sBaseName = Mid$(sDocPath, InStrRev(sDocPath, "\") + 1)
        If InStrRev(sBaseName, ".") > 0 Then
            sBaseName = Left$(sBaseName, InStrRev(sBaseName, ".") - 1)
        End If
    Else
        sBaseName = swModel.GetTitle()
    End If

    sBaseName = SanitiseFileName(sBaseName)

    Dim swModelExt As SldWorks.ModelDocExtension
    Set swModelExt = swModel.Extension

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim lErrors As Long
    Dim lWarnings As Long
    Dim bSuccess As Boolean

    If swModel.GetType() = swDocDRAWING Then
        ' Force SolidWorks user preference to export CURRENT SHEET ONLY (prevents duplicate exports)
        On Error Resume Next
        Dim bOrigTiffAllSheets As Boolean
        bOrigTiffAllSheets = swApp.GetUserPreferenceToggle(swTiffPrintAllSheets)
        swApp.SetUserPreferenceToggle swTiffPrintAllSheets, False
        On Error GoTo ErrorHandler

        ' Export drawing sheets
        Dim swDraw As SldWorks.DrawingDoc
        Set swDraw = swModel

        Dim vSheetNames As Variant
        vSheetNames = swDraw.GetSheetNames()

        If Not IsEmpty(vSheetNames) Then
            Dim i As Long
            Dim sSheetName As String
            Dim sPNGPath As String
            Dim lSheetNum As Long

            ' Active original sheet name to restore at end
            Dim swOrigSheet As SldWorks.Sheet
            Set swOrigSheet = swDraw.GetCurrentSheet()
            Dim sOrigSheetName As String
            If Not swOrigSheet Is Nothing Then sOrigSheetName = swOrigSheet.GetName()

            For i = 0 To UBound(vSheetNames)
                sSheetName = CStr(vSheetNames(i))
                lSheetNum = i + 1

                ' Activate target sheet
                swDraw.ActivateSheet sSheetName

                ' Construct target path: BaseName_1.png, BaseName_2.png, etc.
                sPNGPath = sOutputFolder & sBaseName & "_" & lSheetNum & ".png"

                ' Delete existing output file if present
                If fso.FileExists(sPNGPath) Then
                    On Error Resume Next
                    fso.DeleteFile sPNGPath, True
                    On Error GoTo ErrorHandler
                End If

                ' Save current sheet as PNG
                bSuccess = swModelExt.SaveAs3(sPNGPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, lErrors, lWarnings)

                If bSuccess And fso.FileExists(sPNGPath) Then
                    g_lSuccessCount = g_lSuccessCount + 1
                Else
                    g_lFailCount = g_lFailCount + 1
                End If
            Next i

            ' Restore original sheet focus
            If sOrigSheetName <> "" Then swDraw.ActivateSheet sOrigSheetName
        Else
            g_lFailCount = g_lFailCount + 1
        End If

        ' Restore original SolidWorks system setting
        On Error Resume Next
        swApp.SetUserPreferenceToggle swTiffPrintAllSheets, bOrigTiffAllSheets
        On Error GoTo ErrorHandler
    Else
        ' Part or Assembly model export
        Dim sModelPNGPath As String
        sModelPNGPath = sOutputFolder & sBaseName & ".png"

        If fso.FileExists(sModelPNGPath) Then
            On Error Resume Next
            fso.DeleteFile sModelPNGPath, True
            On Error GoTo ErrorHandler
        End If

        bSuccess = swModelExt.SaveAs3(sModelPNGPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, lErrors, lWarnings)

        If bSuccess And fso.FileExists(sModelPNGPath) Then
            g_lSuccessCount = g_lSuccessCount + 1
        Else
            g_lFailCount = g_lFailCount + 1
        End If
    End If

    Exit Sub

ErrorHandler:
    g_lFailCount = g_lFailCount + 1
End Sub

' ==============================================================================
' Helper Sub: DisplaySummaryAndOpenFolder
' Purpose: Displays completion stats and offers to open PNG output folder
' ==============================================================================
Private Sub DisplaySummaryAndOpenFolder(ByVal sOutputFolder As String, ByVal lProcessedFiles As Long)
    Dim sSummary As String
    sSummary = "Batch PNG Export Completed!" & vbCrLf & vbCrLf & _
               "Processed Input Files: " & lProcessedFiles & vbCrLf & _
               "Exported Images Created: " & g_lSuccessCount & vbCrLf & _
               "Failed / Errors: " & g_lFailCount & vbCrLf & vbCrLf & _
               "Destination Folder:" & vbCrLf & sOutputFolder & vbCrLf & vbCrLf & _
               "Would you like to open the export folder in Windows File Explorer?"

    Dim nResult As VbMsgBoxResult
    nResult = MsgBox(sSummary, vbYesNo + vbInformation, "Batch Export Summary")

    If nResult = vbYes Then
        On Error Resume Next
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run "explorer.exe """ & sOutputFolder & """", 1, False
    End If
End Sub

' ==============================================================================
' Helper Function: BrowseForFolder
' Purpose: Native Windows folder picker using PowerShell FolderBrowserDialog
' ==============================================================================
Private Function BrowseForFolder(ByVal sTitle As String) As String
    On Error GoTo Fallback

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim sCleanTitle As String
    sCleanTitle = Replace(sTitle, "'", "''")

    Dim sCmd As String
    sCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command """ & _
           "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; " & _
           "$f = New-Object System.Windows.Forms.FolderBrowserDialog; " & _
           "$f.Description = '" & sCleanTitle & "'; " & _
           "$f.ShowNewFolderButton = $true; " & _
           "if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Host $f.SelectedPath }"""

    Dim execObj As Object
    Set execObj = sh.Exec(sCmd)

    Dim sResult As String
    sResult = execObj.StdOut.ReadAll()
    sResult = Replace(sResult, vbCr, "")
    sResult = Replace(sResult, vbLf, "")
    sResult = Trim$(sResult)

    BrowseForFolder = sResult
    Exit Function

Fallback:
    ' Shell.Application fallback picker
    Dim objShell As Object
    Dim objFolder As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, sTitle, 0, 0)
    If Not objFolder Is Nothing Then
        BrowseForFolder = objFolder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

' ==============================================================================
' Helper Function: SanitiseFileName
' Purpose: Strips illegal file characters (\ / : * ? " < > |)
' ==============================================================================
Private Function SanitiseFileName(ByVal sFileName As String) As String
    Dim sBadChars As String
    sBadChars = "\/:*?""<>|"
    Dim i As Long
    For i = 1 To Len(sBadChars)
        sFileName = Replace(sFileName, Mid$(sBadChars, i, 1), "_")
    Next i
    SanitiseFileName = Trim$(sFileName)
End Function

' ==============================================================================
' Helper Sub: EnsureFolderStructure
' Purpose: Recursively builds subdirectories if missing
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
