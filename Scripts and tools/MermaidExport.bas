Attribute VB_Name = "MermaidExport"
Option Explicit

' =====================================================================
' Generates Mermaid flowchart files from the "publication" sheet.
' - Full export (all thematiques)
' - One file per Thematique
' - One file per Verticale (wrapped in its parent Thematique)
' Output folder: "{WorkbookBaseName} - Export mermaid" next to the file.
' Output encoding: UTF-8 without BOM via ADODB.Stream.
' All code and comments in English. No accented characters in code.
' =====================================================================

' ----- Module-level dataset arrays (filled once from the sheet) -----
Private gID() As String
Private gThem() As String
Private gIntitule() As String
Private gPalier() As Double
Private gVert() As String
Private gType() As String
Private gAssoc() As String
Private gRowCount As Long


Public Sub GenerateMermaidFiles()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim baseName As String
    Dim folderPath As String
    Dim fso As Object
    Dim included() As Boolean
    Dim outText As String
    Dim safeName As String
    Dim themList As Object
    Dim vertList As Object
    Dim keysArr As Variant
    Dim idx As Long

    Set wb = ActiveWorkbook

    ' Locate the publication sheet
    On Error Resume Next
    Set ws = wb.Worksheets("publication")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'publication' not found.", vbExclamation
        Exit Sub
    End If

    ' Always start by activating the publication sheet
    ws.Activate

    ' Workbook must have been saved
    If Len(wb.Path) = 0 Then
        MsgBox "Please save the workbook before running this macro.", vbExclamation
        Exit Sub
    End If

    ' Read data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in sheet 'publication'.", vbExclamation
        Exit Sub
    End If

    gRowCount = lastRow - 1
    ReDim gID(1 To gRowCount)
    ReDim gThem(1 To gRowCount)
    ReDim gIntitule(1 To gRowCount)
    ReDim gPalier(1 To gRowCount)
    ReDim gVert(1 To gRowCount)
    ReDim gType(1 To gRowCount)
    ReDim gAssoc(1 To gRowCount)

    For r = 2 To lastRow
        i = r - 1
        gID(i) = Trim$(CStr(ws.Cells(r, 1).Value))
        gThem(i) = Trim$(CStr(ws.Cells(r, 2).Value))
        gIntitule(i) = CStr(ws.Cells(r, 3).Value)
        gPalier(i) = ToDouble(ws.Cells(r, 4).Value)
        gVert(i) = Trim$(CStr(ws.Cells(r, 5).Value))
        gType(i) = Trim$(CStr(ws.Cells(r, 12).Value))
        gAssoc(i) = Trim$(CStr(ws.Cells(r, 13).Value))
    Next r

    ' Output folder
    baseName = GetBaseName(wb.Name)
    folderPath = wb.Path & "\" & baseName & " - Export mermaid"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    ' 1) Full export
    ReDim included(1 To gRowCount)
    For i = 1 To gRowCount
        included(i) = True
    Next i
    outText = BuildDoc(included, True, False)
    WriteUTF8NoBOM folderPath & "\" & baseName & ".mermaid", outText

    ' 1b) Full export, plain text (no HTML tags) for Excalidraw import
    outText = BuildDoc(included, True, True)
    WriteUTF8NoBOM folderPath & "\" & baseName & "_excalidraw.mermaid", outText

    ' 2) One file per Thematique
    Set themList = OrderedUniqueArr(gThem)
    keysArr = themList.keys
    For idx = 0 To themList.Count - 1
        ReDim included(1 To gRowCount)
        For i = 1 To gRowCount
            included(i) = (gThem(i) = CStr(keysArr(idx)))
        Next i
        outText = BuildDoc(included, False, False)
        safeName = SafeID(CStr(keysArr(idx)))
        WriteUTF8NoBOM folderPath & "\" & baseName & "_" & safeName & ".mermaid", outText
    Next idx

    ' 3) One file per Verticale
    Set vertList = OrderedUniqueArr(gVert)
    keysArr = vertList.keys
    For idx = 0 To vertList.Count - 1
        ReDim included(1 To gRowCount)
        For i = 1 To gRowCount
            included(i) = (gVert(i) = CStr(keysArr(idx)))
        Next i
        outText = BuildDoc(included, False, False)
        safeName = SafeID(CStr(keysArr(idx)))
        WriteUTF8NoBOM folderPath & "\" & baseName & "_" & safeName & ".mermaid", outText
    Next idx

    MsgBox "Mermaid files generated in:" & vbCrLf & folderPath, vbInformation
End Sub


' =====================================================================
' Builds one complete Mermaid document for the rows flagged in included().
' isFullExport = True only for the global file (controls inter-thematique
' links, which are omitted in per-thematique and per-verticale files).
' =====================================================================
Private Function BuildDoc(included() As Boolean, isFullExport As Boolean, plainText As Boolean) As String
    Dim sb As String
    Dim NL As String
    Dim q As String
    Dim i As Long, j As Long, t As Long, vIdx As Long, pIdx As Long
    Dim a As Long, b As Long
    Dim c As Long, prevPos As Long, colNodes As Long
    Dim hasArrow As Boolean
    Dim nodeLabel As String
    Dim arrowNodes As Object
    Dim idToRow As Object
    Dim parts As Variant
    Dim src As String
    Dim themDict As Object
    Dim vertDict As Object
    Dim tmpKeys As Variant
    Dim themNames() As String, themN As Long
    Dim vertNames() As String, vertN As Long
    Dim themName As String, vertName As String
    Dim safeThemID As String, safeVertID As String
    Dim themTitle As String, vertTitle As String, palTitle As String
    Dim palVals() As Double, palEmpty() As Boolean, palN As Long
    Dim palIDarr() As String
    Dim rowIdx() As Long, rowCnt As Long
    Dim keyIdx As Long, keyPrio As Long
    Dim allVertIDs As Collection
    Dim allPalIDs As Collection
    Dim nodeClassLines As Collection
    Dim vv As Variant

    NL = vbCrLf
    q = Chr$(34)
    Set allVertIDs = New Collection
    Set allPalIDs = New Collection
    Set nodeClassLines = New Collection

    ' ----- Map of included node IDs -> row, and global arrow-node set -----
    Set idToRow = CreateObject("Scripting.Dictionary")
    For i = 1 To gRowCount
        If included(i) Then
            If Not idToRow.Exists(gID(i)) Then idToRow.Add gID(i), i
        End If
    Next i

    Set arrowNodes = CreateObject("Scripting.Dictionary")
    For i = 1 To gRowCount
        If included(i) And Len(gAssoc(i)) > 0 Then
            parts = Split(gAssoc(i), ",")
            For j = 0 To UBound(parts)
                src = Trim$(CStr(parts(j)))
                If Len(src) > 0 Then
                    If idToRow.Exists(src) Then
                        If Not arrowNodes.Exists(src) Then arrowNodes.Add src, True
                        If Not arrowNodes.Exists(gID(i)) Then arrowNodes.Add gID(i), True
                    End If
                End If
            Next j
        End If
    Next i

    ' ----- 1. YAML config block -----
    sb = "---" & NL
    sb = sb & "config:" & NL
    sb = sb & "  flowchart:" & NL
    sb = sb & "    defaultRenderer: " & q & "elk" & q & NL
    sb = sb & "  look: classic" & NL
    sb = sb & "  theme: 'base'" & NL
    sb = sb & "  themeVariables:" & NL
    sb = sb & "    lineColor: '#284370'" & NL
    sb = sb & "    nodeBorder: '#284370'" & NL
    sb = sb & "    primaryTextColor: '#284370'" & NL
    sb = sb & "    primaryColor: '#E0E3E6'" & NL
    sb = sb & "    clusterBkg: '#D6C8DB'" & NL
    sb = sb & "    clusterBorder: '#6E1C73'" & NL
    sb = sb & "    titleColor: '#6E1C73'" & NL
    sb = sb & "---" & NL

    ' ----- 2. Global direction -----
    sb = sb & "flowchart LR" & NL

    ' ----- 3. classDef declarations -----
    sb = sb & "classDef Thematique font-size:14px,fill:transparent,stroke:#000000,color:#000000" & NL
    sb = sb & "classDef Verticale font-size:26px,fill:#D6C8DB,stroke:#6E1C73,color:#6E1C73" & NL
    sb = sb & "classDef Pallier font-size:20px,fill:#D6C8DB,stroke:#6E1C73,color:#284370" & NL
    sb = sb & "classDef Mesure font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370" & NL
    sb = sb & "classDef MesurePlus font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370" & NL
    sb = sb & "classDef MesureMoins font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370" & NL
    sb = sb & "classDef Recommendation font-size:16px,fill:#E0E3E6,stroke:#284370,color:#284370" & NL
    sb = sb & NL

    ' ----- 4. Node declarations (all included nodes) + collect node classes -----
    For i = 1 To gRowCount
        If included(i) Then
            If plainText Then
                nodeLabel = gID(i) & " - [" & gType(i) & "] " & EscapePlain(gIntitule(i))
            Else
                nodeLabel = "<i>" & gID(i) & "</i><br><b>[" & gType(i) & "] " _
                            & WrapTitle(gIntitule(i), 25) & "</b>"
            End If
            sb = sb & "  " & gID(i) & "[" & q & nodeLabel & q & "]" & NL
            nodeClassLines.Add "class " & gID(i) & " " & TypeToClass(gType(i))
        End If
    Next i
    sb = sb & NL

    ' ----- 5. Subgraph structure : Thematique > Verticale > Palier -----
    Set themDict = OrderedUniqueArrFiltered(gThem, included)
    themN = themDict.Count
    If themN > 0 Then
        ReDim themNames(1 To themN)
        tmpKeys = themDict.keys
        For t = 0 To themN - 1
            themNames(t + 1) = CStr(tmpKeys(t))
        Next t
    End If

    ReDim rowIdx(1 To gRowCount)

    For t = 1 To themN
        themName = themNames(t)
        safeThemID = SafeID(themName)
        If plainText Then
            themTitle = FrenchMonth(Month(Date)) & " " & CStr(Year(Date)) _
                        & " - Dalton - Thematique " & EscapePlain(themName)
        Else
            themTitle = "<i>" & FrenchMonth(Month(Date)) & " " & CStr(Year(Date)) _
                        & "</i><br><b>Dalton - Thematique " & EscapeHTML(themName) & "</b>"
        End If
        sb = sb & "subgraph " & safeThemID & "[" & q & themTitle & q & "]" & NL
        sb = sb & "  direction LR" & NL

        ' Verticales of this thematique, in order of first appearance
        Set vertDict = CreateObject("Scripting.Dictionary")
        For i = 1 To gRowCount
            If included(i) And gThem(i) = themName Then
                If Not vertDict.Exists(gVert(i)) Then vertDict.Add gVert(i), i
            End If
        Next i
        vertN = vertDict.Count
        ReDim vertNames(1 To IIf(vertN = 0, 1, vertN))
        If vertN > 0 Then
            tmpKeys = vertDict.keys
            For vIdx = 0 To vertN - 1
                vertNames(vIdx + 1) = CStr(tmpKeys(vIdx))
            Next vIdx
        End If

        For vIdx = 1 To vertN
            vertName = vertNames(vIdx)
            safeVertID = safeThemID & "__" & SafeID(vertName)
            allVertIDs.Add safeVertID

            If plainText Then
                vertTitle = EscapePlain(vertName)
            Else
                vertTitle = "<center>" & WrapTitle(vertName, 30) & "</center>"
            End If
            sb = sb & "  subgraph " & safeVertID & "[" & q & vertTitle & q & "]" & NL
            sb = sb & "    direction BT" & NL

            ' Paliers (sorted ascending, with integer gap-filling)
            BuildPaliers included, themName, vertName, palVals, palEmpty, palN
            ReDim palIDarr(1 To IIf(palN = 0, 1, palN))

            For pIdx = 1 To palN
                palIDarr(pIdx) = safeVertID & "_Palier_" & FmtPalier(palVals(pIdx))
                allPalIDs.Add palIDarr(pIdx)

                If palEmpty(pIdx) Then
                    palTitle = "Palier " & FmtPalier(palVals(pIdx)) & " (vide)"
                Else
                    palTitle = "Palier " & FmtPalier(palVals(pIdx))
                End If

                sb = sb & "    subgraph " & palIDarr(pIdx) & "[" & q & palTitle & q & "]" & NL

                If palEmpty(pIdx) Then
                    sb = sb & "      direction LR" & NL
                Else
                    ' Collect rows of this palier
                    rowCnt = 0
                    For i = 1 To gRowCount
                        If included(i) Then
                            If gThem(i) = themName And gVert(i) = vertName _
                               And SamePalier(gPalier(i), palVals(pIdx)) Then
                                rowCnt = rowCnt + 1
                                rowIdx(rowCnt) = i
                            End If
                        End If
                    Next i

                    ' Stable sort rows by type priority (M+, M-, M, R+, R-, R--, R)
                    For a = 2 To rowCnt
                        keyIdx = rowIdx(a)
                        keyPrio = TypePriority(gType(keyIdx))
                        b = a - 1
                        Do While b >= 1
                            If TypePriority(gType(rowIdx(b))) > keyPrio Then
                                rowIdx(b + 1) = rowIdx(b)
                                b = b - 1
                            Else
                                Exit Do
                            End If
                        Loop
                        rowIdx(b + 1) = keyIdx
                    Next a

                    ' Does any measure of this palier take part in an arrow?
                    hasArrow = False
                    For a = 1 To rowCnt
                        If arrowNodes.Exists(gID(rowIdx(a))) Then
                            hasArrow = True
                            Exit For
                        End If
                    Next a

                    ' Paliers driven by arrows keep LR (the arrows give the layout).
                    ' Pure-measure paliers use TB so the two invisible columns
                    ' sit side by side instead of stacking in a single column.
                    If hasArrow Then
                        sb = sb & "      direction LR" & NL
                    Else
                        sb = sb & "      direction TB" & NL
                    End If

                    If hasArrow Then
                        ' Arrows : sourceID = column M value, targetID = current row
                        For a = 1 To rowCnt
                            i = rowIdx(a)
                            If Len(gAssoc(i)) > 0 Then
                                parts = Split(gAssoc(i), ",")
                                For j = 0 To UBound(parts)
                                    src = Trim$(CStr(parts(j)))
                                    If Len(src) > 0 Then
                                        If idToRow.Exists(src) Then
                                            sb = sb & "      " & src & " --> " & gID(i) & NL
                                        End If
                                    End If
                                Next j
                            End If
                        Next a

                        ' Plain node lines for nodes not used in any arrow
                        For a = 1 To rowCnt
                            i = rowIdx(a)
                            If Not arrowNodes.Exists(gID(i)) Then
                                sb = sb & "      " & gID(i) & NL
                            End If
                        Next a
                    Else
                        ' Two-column layout. Each column is chained top-down with
                        ' invisible links (~~~). Row-major reading order:
                        ' left column = positions 1,3,5...  right = 2,4,6...
                        For c = 0 To 1
                            colNodes = 0
                            prevPos = 0
                            For a = 1 + c To rowCnt Step 2
                                If prevPos > 0 Then
                                    sb = sb & "      " & gID(rowIdx(prevPos)) _
                                         & " ~~~ " & gID(rowIdx(a)) & NL
                                End If
                                colNodes = colNodes + 1
                                prevPos = a
                            Next a
                            ' A single-node column has no link, so emit it plainly
                            If colNodes = 1 Then
                                sb = sb & "      " & gID(rowIdx(1 + c)) & NL
                            End If
                        Next c
                    End If
                End If

                sb = sb & "    end" & NL
            Next pIdx

            sb = sb & "  end" & NL

            ' 6. Inter-palier links (just after the verticale end)
            If palN >= 2 Then
                sb = sb & "  %% Links between paliers for " & vertName & NL
                For pIdx = 1 To palN - 1
                    sb = sb & "  " & palIDarr(pIdx) & " --- " & palIDarr(pIdx + 1) & NL
                Next pIdx
            End If
        Next vIdx

        sb = sb & "end" & NL

        ' 7. Inter-verticale links (just after the thematique end)
        If vertN >= 2 Then
            sb = sb & "%% Links between verticales for " & themName & NL
            For vIdx = 1 To vertN - 1
                sb = sb & "  " & safeThemID & "__" & SafeID(vertNames(vIdx)) _
                     & " --- " & safeThemID & "__" & SafeID(vertNames(vIdx + 1)) & NL
            Next vIdx
        End If
    Next t

    ' 8. Inter-thematique links (full export only)
    If isFullExport And themN >= 2 Then
        sb = sb & "%% Space thematiques" & NL
        For t = 1 To themN - 1
            sb = sb & "  " & SafeID(themNames(t)) & " --- " & SafeID(themNames(t + 1)) & NL
        Next t
    End If

    ' 9. Class assignments
    sb = sb & "%% Apply style" & NL
    For t = 1 To themN
        sb = sb & "class " & SafeID(themNames(t)) & " Thematique" & NL
    Next t
    For Each vv In allVertIDs
        sb = sb & "class " & CStr(vv) & " Verticale" & NL
    Next vv
    For Each vv In allPalIDs
        sb = sb & "class " & CStr(vv) & " Pallier" & NL
    Next vv
    For Each vv In nodeClassLines
        sb = sb & CStr(vv) & NL
    Next vv

    BuildDoc = sb
End Function


' =====================================================================
' Builds the ordered, gap-filled palier list for a given verticale.
' palVals / palEmpty are returned as 1-based parallel arrays of size palN.
' =====================================================================
Private Sub BuildPaliers(included() As Boolean, themName As String, vertName As String, _
                         ByRef palVals() As Double, ByRef palEmpty() As Boolean, ByRef palN As Long)
    Dim i As Long, k As Long, a As Long, b As Long, n As Long
    Dim valDict As Object
    Dim present As Object
    Dim rawVals() As Double
    Dim rawN As Long
    Dim minInt As Long, maxInt As Long
    Dim entriesVal() As Double
    Dim entriesEmpty() As Boolean
    Dim cnt As Long
    Dim keyStr As String
    Dim tmp As Double
    Dim tv As Double, te As Boolean

    Set valDict = CreateObject("Scripting.Dictionary")
    Set present = CreateObject("Scripting.Dictionary")

    rawN = 0
    ReDim rawVals(1 To gRowCount)
    For i = 1 To gRowCount
        If included(i) Then
            If gThem(i) = themName And gVert(i) = vertName Then
                keyStr = Trim$(Str$(gPalier(i)))
                If Not valDict.Exists(keyStr) Then
                    valDict.Add keyStr, gPalier(i)
                    rawN = rawN + 1
                    rawVals(rawN) = gPalier(i)
                End If
            End If
        End If
    Next i

    If rawN = 0 Then
        palN = 0
        ReDim palVals(1 To 1)
        ReDim palEmpty(1 To 1)
        Exit Sub
    End If

    ' Sort raw values ascending
    For a = 2 To rawN
        tmp = rawVals(a)
        b = a - 1
        Do While b >= 1
            If rawVals(b) > tmp Then
                rawVals(b + 1) = rawVals(b)
                b = b - 1
            Else
                Exit Do
            End If
        Loop
        rawVals(b + 1) = tmp
    Next a

    ' Record present integer paliers
    For k = 1 To rawN
        If rawVals(k) = Int(rawVals(k)) Then
            present(CStr(CLng(rawVals(k)))) = True
        End If
    Next k

    minInt = Int(rawVals(1))
    maxInt = Int(rawVals(rawN))

    ' Combine real values with gap-fill integers
    ReDim entriesVal(1 To rawN + (maxInt - minInt + 1) + 1)
    ReDim entriesEmpty(1 To rawN + (maxInt - minInt + 1) + 1)
    cnt = 0
    For k = 1 To rawN
        cnt = cnt + 1
        entriesVal(cnt) = rawVals(k)
        entriesEmpty(cnt) = False
    Next k
    For n = minInt To maxInt
        If Not present.Exists(CStr(n)) Then
            cnt = cnt + 1
            entriesVal(cnt) = CDbl(n)
            entriesEmpty(cnt) = True
        End If
    Next n

    ' Sort combined entries by value ascending
    For a = 2 To cnt
        tv = entriesVal(a)
        te = entriesEmpty(a)
        b = a - 1
        Do While b >= 1
            If entriesVal(b) > tv Then
                entriesVal(b + 1) = entriesVal(b)
                entriesEmpty(b + 1) = entriesEmpty(b)
                b = b - 1
            Else
                Exit Do
            End If
        Loop
        entriesVal(b + 1) = tv
        entriesEmpty(b + 1) = te
    Next a

    palN = cnt
    ReDim palVals(1 To cnt)
    ReDim palEmpty(1 To cnt)
    For k = 1 To cnt
        palVals(k) = entriesVal(k)
        palEmpty(k) = entriesEmpty(k)
    Next k
End Sub


' =====================================================================
' Helper functions
' =====================================================================

Private Function ToDouble(v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then
        ToDouble = CDbl(v)
    Else
        ToDouble = 0
    End If
    On Error GoTo 0
End Function

Private Function GetBaseName(fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then
        GetBaseName = Left$(fileName, p - 1)
    Else
        GetBaseName = fileName
    End If
End Function

' Ordered unique values over all rows (insertion order preserved)
Private Function OrderedUniqueArr(arr() As String) As Object
    Dim d As Object
    Dim k As Long
    Set d = CreateObject("Scripting.Dictionary")
    For k = 1 To gRowCount
        If Not d.Exists(arr(k)) Then d.Add arr(k), k
    Next k
    Set OrderedUniqueArr = d
End Function

' Ordered unique values over included rows only
Private Function OrderedUniqueArrFiltered(arr() As String, included() As Boolean) As Object
    Dim d As Object
    Dim k As Long
    Set d = CreateObject("Scripting.Dictionary")
    For k = 1 To gRowCount
        If included(k) Then
            If Not d.Exists(arr(k)) Then d.Add arr(k), k
        End If
    Next k
    Set OrderedUniqueArrFiltered = d
End Function

' Format a palier value without trailing zeros (uses "." decimal separator)
Private Function FmtPalier(v As Double) As String
    FmtPalier = Trim$(Str$(v))
End Function

Private Function SamePalier(a As Double, b As Double) As Boolean
    SamePalier = (Abs(a - b) < 0.0000001)
End Function

Private Function TypeToClass(t As String) As String
    Select Case t
        Case "M": TypeToClass = "Mesure"
        Case "M+": TypeToClass = "MesurePlus"
        Case "M-": TypeToClass = "MesureMoins"
        Case "R", "R+", "R-", "R--": TypeToClass = "Recommendation"
        Case Else: TypeToClass = "Mesure"
    End Select
End Function

Private Function TypePriority(t As String) As Long
    Select Case t
        Case "M+": TypePriority = 1
        Case "M-": TypePriority = 2
        Case "M": TypePriority = 3
        Case "R+": TypePriority = 4
        Case "R-": TypePriority = 5
        Case "R--": TypePriority = 6
        Case "R": TypePriority = 7
        Case Else: TypePriority = 8
    End Select
End Function

Private Function FrenchMonth(m As Long) As String
    Select Case m
        Case 1: FrenchMonth = "Janvier"
        Case 2: FrenchMonth = "F" & ChrW$(233) & "vrier"
        Case 3: FrenchMonth = "Mars"
        Case 4: FrenchMonth = "Avril"
        Case 5: FrenchMonth = "Mai"
        Case 6: FrenchMonth = "Juin"
        Case 7: FrenchMonth = "Juillet"
        Case 8: FrenchMonth = "Ao" & ChrW$(251) & "t"
        Case 9: FrenchMonth = "Septembre"
        Case 10: FrenchMonth = "Octobre"
        Case 11: FrenchMonth = "Novembre"
        Case 12: FrenchMonth = "D" & ChrW$(233) & "cembre"
        Case Else: FrenchMonth = ""
    End Select
End Function

' HTML-escape displayed text. & must be replaced first.
Private Function EscapeHTML(s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "&", "&amp;")
    t = Replace(t, Chr$(34), "&quot;")        ' "
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, "'", "&#39;")              ' straight apostrophe U+0027
    t = Replace(t, ChrW$(96), "&#96;")        ' backtick
    t = Replace(t, ChrW$(8216), "&#39;")      ' left curly quote U+2018
    t = Replace(t, ChrW$(8217), "&#39;")      ' right curly quote U+2019
    t = Replace(t, ChrW$(171), "&quot;")      ' left guillemet U+00AB
    t = Replace(t, ChrW$(187), "&quot;")      ' right guillemet U+00BB
    EscapeHTML = t
End Function

' Plain-text label cleaner for the Excalidraw export: no HTML tags and no
' characters that would break a quoted Mermaid label. Double quotes become
' single quotes; angle brackets become parentheses. Accents are kept.
Private Function EscapePlain(s As String) As String
    Dim t As String
    t = s
    t = Replace(t, Chr$(34), "'")     ' double quote would close the label
    t = Replace(t, "<", "(")
    t = Replace(t, ">", ")")
    EscapePlain = t
End Function

' Wrap a title at word boundaries to lines <= maxLen, then HTML-escape
' each line and join with <br>. Only wraps if the title exceeds maxLen.
Private Function WrapTitle(s As String, maxLen As Long) As String
    Dim words() As String
    Dim k As Long
    Dim line As String
    Dim result As String
    Dim w As String

    If Len(s) <= maxLen Then
        WrapTitle = EscapeHTML(s)
        Exit Function
    End If

    words = Split(s, " ")
    line = ""
    result = ""
    For k = LBound(words) To UBound(words)
        w = words(k)
        If line = "" Then
            line = w
        ElseIf Len(line) + 1 + Len(w) <= maxLen Then
            line = line & " " & w
        Else
            If result = "" Then
                result = EscapeHTML(line)
            Else
                result = result & "<br>" & EscapeHTML(line)
            End If
            line = w
        End If
    Next k
    If Len(line) > 0 Then
        If result = "" Then
            result = EscapeHTML(line)
        Else
            result = result & "<br>" & EscapeHTML(line)
        End If
    End If
    WrapTitle = result
End Function

' Transliterate accents to ASCII, then replace every non-alphanumeric
' character (spaces, hyphens, anything else) with underscore.
Private Function SafeID(s As String) As String
    Dim t As String
    Dim out As String
    Dim k As Long
    Dim ch As String
    Dim code As Long

    t = s
    ' lowercase accented
    t = Replace(t, ChrW$(224), "a")   ' a grave
    t = Replace(t, ChrW$(226), "a")   ' a circ
    t = Replace(t, ChrW$(228), "a")   ' a diaeresis
    t = Replace(t, ChrW$(231), "c")   ' c cedilla
    t = Replace(t, ChrW$(233), "e")   ' e acute
    t = Replace(t, ChrW$(232), "e")   ' e grave
    t = Replace(t, ChrW$(234), "e")   ' e circ
    t = Replace(t, ChrW$(235), "e")   ' e diaeresis
    t = Replace(t, ChrW$(238), "i")   ' i circ
    t = Replace(t, ChrW$(239), "i")   ' i diaeresis
    t = Replace(t, ChrW$(244), "o")   ' o circ
    t = Replace(t, ChrW$(246), "o")   ' o diaeresis
    t = Replace(t, ChrW$(249), "u")   ' u grave
    t = Replace(t, ChrW$(251), "u")   ' u circ
    t = Replace(t, ChrW$(252), "u")   ' u diaeresis
    t = Replace(t, ChrW$(339), "oe")  ' oe ligature
    t = Replace(t, ChrW$(230), "ae")  ' ae ligature
    ' uppercase accented
    t = Replace(t, ChrW$(192), "A")
    t = Replace(t, ChrW$(194), "A")
    t = Replace(t, ChrW$(196), "A")
    t = Replace(t, ChrW$(199), "C")
    t = Replace(t, ChrW$(201), "E")
    t = Replace(t, ChrW$(200), "E")
    t = Replace(t, ChrW$(202), "E")
    t = Replace(t, ChrW$(203), "E")
    t = Replace(t, ChrW$(206), "I")
    t = Replace(t, ChrW$(207), "I")
    t = Replace(t, ChrW$(212), "O")
    t = Replace(t, ChrW$(214), "O")
    t = Replace(t, ChrW$(217), "U")
    t = Replace(t, ChrW$(219), "U")
    t = Replace(t, ChrW$(220), "U")
    t = Replace(t, ChrW$(338), "OE")
    t = Replace(t, ChrW$(198), "AE")

    out = ""
    For k = 1 To Len(t)
        ch = Mid$(t, k, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) _
           Or (code >= 65 And code <= 90) _
           Or (code >= 97 And code <= 122) Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next k
    ' Never return an empty identifier (e.g. blank Thematique/Verticale cell)
    If Len(out) = 0 Then out = "EMPTY"
    SafeID = out
End Function

' Write text as UTF-8 without a BOM (overwrite if exists)
Private Sub WriteUTF8NoBOM(filePath As String, content As String)
    Dim st As Object
    Dim bin As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "UTF-8"
    st.Open
    st.WriteText content

    ' Re-read as binary and skip the 3-byte BOM
    st.Position = 0
    st.Type = 1                 ' adTypeBinary
    st.Position = 3

    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    st.CopyTo bin
    st.Close

    bin.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    bin.Close
End Sub
