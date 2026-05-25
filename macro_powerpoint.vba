Attribute VB_Name = "Module2"
' =============================================================================
' MODULE : Phase1_ChartsAndTables
' PURPOSE: Phase 1 of the report automation.
'          Reads layout instructions from the Control2 sheet and exports
'          charts, named ranges, and shapes from Excel into the active
'          PowerPoint presentation, positioning each object with exact
'          coordinates. Applies the Aspect-Ratio Cloning strategy for the
'          Aspect S & A tables to guarantee a pixel-perfect fit
'          inside the destination slide. Replaces all [Name]
'          placeholders across every slide using the value in Control!H5.
' =============================================================================

Sub Phase1_ChartsAndTables2()

    ' =========================================================================
    ' VARIABLE DECLARATIONS
    ' =========================================================================

    ' PowerPoint objects
    Dim pptApp   As Object
    Dim pptPres  As Object
    Dim pptSlide As Object
    Dim shpRange As Object

    ' Worksheets
    Dim wsControl2 As Worksheet
    Dim ws         As Worksheet

    ' Loop counters
    Dim i        As Long
    Dim lastRow  As Long
    Dim slideNum As Long

    ' Object identity and position parameters (read from Control2)
    Dim objName  As String
    Dim objType  As String
    Dim txtTop   As String
    Dim txtLeft  As String
    Dim txtH     As String
    Dim txtW     As String

    ' Converted position/size values in points (1 inch = 72 pts)
    Dim numTop  As Double
    Dim numLeft As Double
    Dim numH    As Double
    Dim numW    As Double

    ' Control flags
    Dim found              As Boolean
    Dim attempts           As Integer
    Dim isAspectS_A        As Boolean   ' TRUE for tables using Aspect-Ratio Cloning
    Dim masterRulerActive  As Boolean   ' TRUE if column widths were modified and must be restored

    ' Dynamic range variables
    Dim rngOriginal  As Range
    Dim rngTrimmed   As Range
    Dim totalRows    As Long
    Dim totalCols    As Long
    Dim rowScan      As Long
    Dim colScan      As Long
    Dim cellVal      As String

    ' ---- Master Ruler / Aspect-Ratio Cloning variables (Aspect S & A only) ----
    Dim colAWidth_pts    As Double   ' Actual width of Column A in the range (points)
    Dim remainingW_pts   As Double   ' Width available for variable columns (points)
    Dim varCols          As Long     ' Number of variable columns (all except Col A)
    Dim widthPerCol_pts  As Double   ' Target width per variable column (points)
    Dim cwRef            As Double   ' Current ColumnWidth of reference column (chars)
    Dim wRef             As Double   ' Current Width of reference column (points)
    Dim convFactor       As Double   ' Conversion factor: points per ColumnWidth unit
    Dim newColCW         As Double   ' New ColumnWidth in character units
    Dim originalCW()     As Double   ' Array storing original ColumnWidth values for restore
    Dim colIdx           As Long     ' Column iterator
    Dim absCol           As Long     ' Absolute column index on the sheet
    Dim aspectRatio      As Double   ' Target H/W ratio from PowerPoint destination box
    Dim excelRangeH_pts  As Double   ' Actual height of the trimmed range in Excel (fixed)
    Dim neededExcelW_pts As Double   ' Excel range width needed to match the target aspect ratio

    ' ---- Name replacement variables ----
    Dim Name          As String      ' Value read from Control!H5
    Dim oSlide        As Object      ' Slide iterator
    Dim oShape        As Object      ' Shape iterator within each slide
    Dim oTF           As Object      ' TextFrame of each shape
    Dim oPara         As Object      ' Paragraph within the TextFrame
    Dim oRun          As Object      ' Text run (fragment) within the paragraph

    ' ---- File rename variables ----
    Dim ID            As String      ' Active Group ID read from Control!G2
    Dim newFileName   As String      ' Final file name: "ID - Name"
    Dim newFullPath   As String      ' Full path for the renamed PPT file
    Dim wsControl     As Worksheet   ' Reference to the Control sheet
    Dim illegalChars  As Variant     ' Array of characters illegal in Windows file names
    Dim illChar       As Variant     ' Single illegal character iterator
    Dim currentFolder As String      ' Folder path of the current PPT file

    ' =========================================================================
    ' INITIALIZATION
    ' =========================================================================
    Set wsControl2 = ThisWorkbook.Worksheets("Control2")

    ' Verify PowerPoint is open before proceeding
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        MsgBox "Please open PowerPoint before running this macro.", vbCritical, "PowerPoint Not Found"
        Exit Sub
    End If
    On Error GoTo 0

    Set pptPres = pptApp.ActivePresentation
    lastRow = wsControl2.Cells(wsControl2.Rows.Count, 2).End(xlUp).Row

    ' Read Name once from Control sheet, cell H5
    Name = Trim(ThisWorkbook.Worksheets("Control").Cells(5, 8).Value)

    ' =========================================================================
    ' MAIN LOOP — process each row in Control2
    ' =========================================================================
    For i = 2 To lastRow

        ' Read layout parameters from Control2
        slideNum = wsControl2.Cells(i, 1).Value
        objName = Trim(wsControl2.Cells(i, 2).Value)
        txtTop = wsControl2.Cells(i, 3).Value
        txtLeft = wsControl2.Cells(i, 4).Value
        txtH = wsControl2.Cells(i, 5).Value
        txtW = wsControl2.Cells(i, 6).Value
        objType = UCase(Trim(wsControl2.Cells(i, 7).Value))

        ' Convert inches to points (1 inch = 72 points)
        numTop = Val(Replace(txtTop, """", "")) * 72
        numLeft = Val(Replace(txtLeft, """", "")) * 72
        numH = Val(Replace(txtH, """", "")) * 72
        numW = Val(Replace(txtW, """", "")) * 72

        ' Reset per-iteration flags
        found = False
        masterRulerActive = False
        isAspectS_A = False

        ' =====================================================================
        ' BLOCK A — CHARTS
        ' =====================================================================
        If objType = "CHART" Then

            For Each ws In ThisWorkbook.Worksheets
                On Error Resume Next
                ws.ChartObjects(objName).Chart.ChartArea.Copy
                If Err.Number = 0 Then
                    found = True
                    On Error GoTo 0
                    Exit For
                End If
                On Error GoTo 0
            Next ws

        ' =====================================================================
        ' BLOCK B — IMAGES / NAMED RANGES / SHAPES
        ' =====================================================================
        ElseIf objType = "IMAGE" Or objType = "PICTURE" Then

            Set rngOriginal = Nothing
            On Error Resume Next
            Set rngOriginal = Range(objName)
            On Error GoTo 0

            ' -----------------------------------------------------------------
            ' PATH 1: Object is a Named Range
            ' -----------------------------------------------------------------
            If Not rngOriginal Is Nothing Then

                ' -------------------------------------------------------------
                ' CASE 1: OVERLAP TABLES — vertical expansion only
                '         Scans down Col A until the first empty cell.
                ' -------------------------------------------------------------
                If objName = "ANCHOR_1_Overlap_Table" Or _
                   objName = "ANCHOR_2_Overlap_Table" Or _
                   objName = "ANCHOR_3_Overlap_Table" Then

                    totalRows = 1
                    Do While Len(Trim(rngOriginal.Cells(1, 1).Offset(totalRows, 0).Value)) > 0
                        totalRows = totalRows + 1
                        If totalRows > 1000 Then Exit Do   ' Safety cap
                    Loop
                    Set rngTrimmed = rngOriginal.Resize(totalRows, rngOriginal.Columns.Count)

                ' -------------------------------------------------------------
                ' CASE 2: Aspect S & A TABLES
                '
                '   STEP 1 — Precise range delimitation
                '     Scans active columns (header <> "" and <> "0") and active
                '     rows using a dual stop condition:
                '       a) Empty or "0" cell in Col A  -> natural end of data
                '       b) "Sum of Available" text      -> top of the lower pivot
                '     This prevents capturing the pivot table below, which was
                '     inflating the image height and causing it to overflow the slide.
                '
                '   STEP 2 — Aspect-Ratio Cloning
                '     Calculates the H/W ratio of the PowerPoint destination box,
                '     then resizes Excel columns so that the captured range has
                '     the exact same proportions. When pasted with LockAspectRatio
                '     = msoTrue and only .Width imposed, PowerPoint derives the
                '     height automatically and the table fits the slide perfectly.
                '
                '     Formula:
                '       aspectRatio      = numH / numW
                '       excelRangeH_pts  = rngTrimmed.Height   (fixed, never touched)
                '       neededExcelW_pts = excelRangeH_pts / aspectRatio
                '       colAWidth_pts    = width of Col A (fixed)
                '       remainingW_pts   = neededExcelW_pts - colAWidth_pts
                '       widthPerCol_pts  = remainingW_pts / varCols
                ' -------------------------------------------------------------
                ElseIf objName = "ANCHOR_S_Table" Or _
                       objName = "ANCHOR_A_Table" Then

                    isAspectS_A = True
                    totalRows = 1
                    totalCols = 1

                    ' STEP 1A: Count active columns (header must be non-empty and non-zero)
                    For colScan = 2 To 20
                        cellVal = Trim(CStr(rngOriginal.Cells(1, colScan).Value))
                        If cellVal = "0" Or cellVal = "" Then Exit For
                        totalCols = colScan
                    Next colScan

                    ' STEP 1B: Count active rows — dual stop condition
                    For rowScan = 2 To 500
                        cellVal = Trim(CStr(rngOriginal.Cells(rowScan, 1).Value))
                        If cellVal = "" Or cellVal = "0" Then Exit For
                        If InStr(1, cellVal, "Sum of Available", vbTextCompare) > 0 Then Exit For
                        totalRows = rowScan
                    Next rowScan

                    ' Trim range to exact data boundaries
                    Set rngTrimmed = rngOriginal.Resize(totalRows, totalCols)

                    ' STEP 2: Aspect-Ratio Cloning
                    If numW > 0 And numH > 0 And totalCols > 1 Then

                        aspectRatio = numH / numW
                        excelRangeH_pts = rngTrimmed.Height           ' Fixed — never modified
                        neededExcelW_pts = excelRangeH_pts / aspectRatio

                        varCols = totalCols - 1
                        colAWidth_pts = rngTrimmed.Columns(1).Width
                        remainingW_pts = neededExcelW_pts - colAWidth_pts

                        If remainingW_pts > 0 Then

                            widthPerCol_pts = remainingW_pts / varCols

                            ' Dynamic conversion factor: points -> ColumnWidth units
                            ' Derived from the actual reference column to avoid hardcoding DPI
                            absCol = rngTrimmed.Columns(2).Column
                            cwRef = rngTrimmed.Worksheet.Columns(absCol).ColumnWidth
                            wRef = rngTrimmed.Worksheet.Columns(absCol).Width

                            If cwRef > 0 Then
                                convFactor = wRef / cwRef
                            Else
                                convFactor = 7.5   ' Standard fallback
                            End If

                            newColCW = widthPerCol_pts / convFactor

                            ' Save original column widths so they can be restored after capture
                            ReDim originalCW(1 To varCols)
                            For colIdx = 1 To varCols
                                absCol = rngTrimmed.Columns(colIdx + 1).Column
                                originalCW(colIdx) = rngTrimmed.Worksheet.Columns(absCol).ColumnWidth
                            Next colIdx

                            ' Apply the new uniform width to all variable columns
                            For colIdx = 1 To varCols
                                absCol = rngTrimmed.Columns(colIdx + 1).Column
                                rngTrimmed.Worksheet.Columns(absCol).ColumnWidth = newColCW
                            Next colIdx

                            masterRulerActive = True

                        End If
                    End If
                    ' --- END ASPECT-RATIO CLONING ---

                ' -------------------------------------------------------------
                ' CASE 3: Any other named range — no modifications
                ' -------------------------------------------------------------
                Else
                    Set rngTrimmed = rngOriginal
                End If

                ' -------------------------------------------------------------
                ' COPY ENGINE — Anti-Error 1004 with up to 3 retries
                ' -------------------------------------------------------------
                attempts = 0
                found = False
                rngTrimmed.Worksheet.Activate   ' Activate sheet to force rendering

                Do While attempts < 3 And Not found
                    On Error Resume Next
                    rngTrimmed.CopyPicture Appearance:=1, Format:=-4147   ' xlScreen, xlPicture
                    If Err.Number = 0 Then
                        found = True
                    Else
                        attempts = attempts + 1
                        Err.Clear
                        DoEvents
                        Application.Wait (Now + TimeValue("0:00:01"))
                    End If
                    On Error GoTo 0
                Loop

                ' -------------------------------------------------------------
                ' RESTORE: Return all modified column widths to their original
                ' values immediately after the range has been captured.
                ' Only executes if the Master Ruler modified columns this iteration.
                ' -------------------------------------------------------------
                If masterRulerActive Then
                    For colIdx = 1 To varCols
                        absCol = rngTrimmed.Columns(colIdx + 1).Column
                        rngTrimmed.Worksheet.Columns(absCol).ColumnWidth = originalCW(colIdx)
                    Next colIdx
                End If

            ' -----------------------------------------------------------------
            ' PATH 2: Not a named range — search for a Shape across all sheets
            ' -----------------------------------------------------------------
            Else
                For Each ws In ThisWorkbook.Worksheets
                    On Error Resume Next
                    ws.Shapes(objName).Copy
                    If Err.Number = 0 Then
                        found = True
                        On Error GoTo 0
                        Exit For
                    End If
                    On Error GoTo 0
                Next ws
            End If

        End If

        ' =====================================================================
        ' PASTE AND POSITION IN POWERPOINT
        ' =====================================================================
        If found Then
            If slideNum <= pptPres.Slides.Count Then

                Set pptSlide = pptPres.Slides(slideNum)

                ' Paste as PNG (DataType:=18) to keep file size manageable
                On Error Resume Next
                Set shpRange = pptSlide.Shapes.PasteSpecial(DataType:=18)
                If Err.Number <> 0 Then
                    Err.Clear
                    Set shpRange = pptSlide.Shapes.Paste   ' Fallback to standard paste
                End If
                On Error GoTo 0

                With shpRange
                    .Top = numTop
                    .Left = numLeft
                    If isAspectS_A Then
                        ' Aspect S & A: lock aspect ratio and impose width only.
                        ' PowerPoint derives the height automatically from the cloned ratio,
                        ' guaranteeing the table fits the slide without overflow.
                        .LockAspectRatio = msoTrue
                        If numW > 0 Then .Width = numW
                    Else
                        ' All other objects: impose exact width and height from Control2.
                        .LockAspectRatio = msoFalse
                        If numH > 0 Then .Height = numH
                        If numW > 0 Then .Width = numW
                    End If
                End With

            End If
            Application.CutCopyMode = False
        End If

    Next i

    ' =========================================================================
    ' NAME REPLACEMENT
    ' Iterates every shape on every slide, replacing [Name] inside
    ' each text Run. Operating at Run level preserves font, size, and color.
    ' =========================================================================
    wsControl2.Activate

    If Len(Name) > 0 Then
        For Each oSlide In pptPres.Slides
            For Each oShape In oSlide.Shapes
                If oShape.HasTextFrame Then
                    Set oTF = oShape.TextFrame
                    For Each oPara In oTF.TextRange.Paragraphs
                        For Each oRun In oPara.Runs
                            If InStr(1, oRun.Text, "[Name]", vbTextCompare) > 0 Then
                                oRun.Text = Replace(oRun.Text, "[Name]", _
                                            Name, 1, -1, vbTextCompare)
                            End If
                        Next oRun
                    Next oPara
                End If
            Next oShape
        Next oSlide
    End If
    ' =========================================================================
    ' END NAME REPLACEMENT
    ' =========================================================================

    ' =========================================================================
    ' FILE RENAME
    ' Reads the active Group ID from Control!G2 (linked to the Group ID slicer),
    ' then renames the active PowerPoint file to: "ID - Name"
    ' Example output: "2 - Client Name.pptx"
    ' =========================================================================
    Set wsControl = ThisWorkbook.Worksheets("Control")

    ' Read the active Group ID directly from Control!G2
    ID = Trim(CStr(wsControl.Cells(2, 7).Value))

    ' Build new file name only if both values are available
    If Len(ID) > 0 And Len(Name) > 0 Then

        ' Sanitize: remove characters that are illegal in Windows file names
        illegalChars = Array("", "/", ":", "*", "?", """", "<", ">", "|")
        newFileName = ID & " - " & Name
        For Each illChar In illegalChars
            newFileName = Replace(newFileName, illChar, "")
        Next illChar

        ' Build full destination path: same folder as the current PPT, new name
        currentFolder = Left(pptPres.FullName, _
                             Len(pptPres.FullName) - Len(pptPres.Name))
        newFullPath = currentFolder & newFileName & ".pptx"

        ' Save the file under the new name (SaveAs preserves format)
        On Error Resume Next
        pptPres.SaveAs newFullPath, 24   ' 24 = ppSaveAsOpenXMLPresentation (.pptx)
        If Err.Number <> 0 Then
            MsgBox "File rename failed." & vbCrLf & _
                   "The report was NOT renamed. Please rename it manually to:" & vbCrLf & _
                   newFileName & ".pptx", vbExclamation, "Rename Warning"
            Err.Clear
        End If
        On Error GoTo 0

    End If
    ' =========================================================================
    ' END FILE RENAME
    ' =========================================================================

    MsgBox "Phase 1 of report complete." & vbCrLf & _
           "Charts, tables, and name have been successfully exported to PowerPoint.", _
           vbInformation, "Phase 1 Complete"

End Sub
