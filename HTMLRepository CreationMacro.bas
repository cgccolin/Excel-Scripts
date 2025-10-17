Sub ExtractLinksFromHTMLwithPageType()
    Dim fso As Object, folder As Object, file As Object
    Dim wsAll As Worksheet, wsGroups As Worksheet
    Dim htmlDoc As Object
    Dim rowAll As Long, rowGroup As Long
    Dim linkDict As Object, groupDict As Object
    Dim filePath As String
    Dim processedFiles As String, skippedFiles As String
    
    ' Initialize objects
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set linkDict = CreateObject("Scripting.Dictionary")
    Set groupDict = CreateObject("Scripting.Dictionary")
    Set htmlDoc = CreateObject("htmlfile")
    processedFiles = ""
    skippedFiles = ""
    
    ' Select folder with HTML files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing HTML Files"
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    
    ' Create sheets
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("All Links").Delete
    ThisWorkbook.Sheets("Grouped Links").Delete
    Application.DisplayAlerts = True
    Set wsAll = ThisWorkbook.Sheets.Add
    wsAll.Name = "All Links"
    Set wsGroups = ThisWorkbook.Sheets.Add
    wsGroups.Name = "Grouped Links"
    On Error GoTo 0
    
    ' Process each HTML file
    Set folder = fso.GetFolder(filePath)
    For Each file In folder.Files
        If LCase(Right(file.Name, 5)) = ".html" Or LCase(Right(file.Name, 4)) = ".htm" Then
            On Error Resume Next
            Dim htmlContent As String
            htmlContent = ReadFile(file.Path)
            If Err.Number <> 0 Then
                skippedFiles = skippedFiles & vbCrLf & file.Name & " (Read error: " & Err.Description & ")"
                Err.Clear
                GoTo NextFile
            End If
            
            htmlDoc.body.innerHTML = htmlContent
            If Err.Number <> 0 Then
                skippedFiles = skippedFiles & vbCrLf & file.Name & " (Parse error: " & Err.Description & ")"
                Err.Clear
                GoTo NextFile
            End If
            On Error GoTo 0
            
            processedFiles = processedFiles & vbCrLf & file.Name
            
            ' Get group title
            Dim groupTitle As String
            On Error Resume Next
            groupTitle = htmlDoc.getElementsByClassName("exlheader")(0).getElementsByTagName("b")(0).innerText
            If Err.Number <> 0 Then
                Err.Clear
                groupTitle = htmlDoc.getElementsByClassName("sectiontext")(0).innerText
            End If
            If Err.Number <> 0 Or Len(groupTitle) = 0 Then
                Err.Clear
                groupTitle = Left(file.Name, InStrRev(file.Name, ".") - 1)
            End If
            On Error GoTo 0
            
            ' Determine page type
            Dim pageType As String
            pageType = DeterminePageType(htmlDoc)
            
            ' Get all links
            Dim links As Object, link As Object
            Set links = htmlDoc.getElementsByTagName("a")
            Dim groupLinks As Object
            Set groupLinks = New Collection
            
            For Each link In links
                Dim linkURL As String, linkName As String
                linkURL = link.href
                If Len(linkURL) > 0 And (Left(LCase(linkURL), 4) = "http" Or Left(LCase(linkURL), 6) = "mailto") Then
                    linkName = Trim(link.innerText)
                    linkName = CleanString(linkName)
                    
                    If Len(linkName) = 0 Then
                        On Error Resume Next
                        linkName = Trim(link.ParentNode.innerText)
                        linkName = CleanString(linkName)
                        If Err.Number <> 0 Or Len(linkName) = 0 Then
                            Err.Clear
                            linkName = "Unnamed Link"
                        End If
                        On Error GoTo 0
                    End If
                    
                    Dim pairKey As String
                    pairKey = linkName & vbTab & linkURL & vbTab & pageType
                    If Not linkDict.Exists(pairKey) Then
                        linkDict.Add pairKey, Array(linkName, linkURL, pageType)
                    End If
                    
                    Dim linkPair As Variant
                    linkPair = Array(linkName, linkURL, pageType)
                    groupLinks.Add linkPair
                End If
            Next link
            
            If groupLinks.Count > 0 Then
                If Not groupDict.Exists(groupTitle) Then
                    groupDict.Add groupTitle, groupLinks
                Else
                    Dim existingLinks As Collection
                    Set existingLinks = groupDict(groupTitle)
                    For Each linkPair In groupLinks
                        existingLinks.Add linkPair
                    Next linkPair
                End If
            End If
NextFile:
            On Error GoTo 0
        End If
    Next file
    
    ' Populate All Links sheet
    With wsAll
        .Cells(1, 1).Value = "Link Name"
        .Cells(1, 2).Value = "URL"
        .Cells(1, 3).Value = "Page Type"
        .Range("A1:C1").Font.Bold = True
        
        rowAll = 2
        Dim keyVar As Variant
        For Each keyVar In linkDict.Keys
            Dim pair As Variant
            pair = linkDict(keyVar)
            .Cells(rowAll, 1).Value = pair(0)
            .Cells(rowAll, 2).Value = pair(1)
            .Cells(rowAll, 3).Value = pair(2)
            rowAll = rowAll + 1
        Next keyVar
        
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 60
        .Columns("C").ColumnWidth = 20
        .Rows.RowHeight = 15
        .Cells.WrapText = True
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End With
    
    ' Populate Grouped Links sheet (updated to bold Page Type value)
    With wsGroups
        rowGroup = 1
        Dim groupVar As Variant
        For Each groupVar In groupDict.Keys
            ' Page Name header and value
            .Cells(rowGroup, 1).Value = "Page Name"
            .Cells(rowGroup, 1).Font.Bold = True
            .Cells(rowGroup, 2).Value = groupVar
            .Cells(rowGroup, 2).Font.Bold = True
            rowGroup = rowGroup + 1
            
            ' Page Type header and value (value bolded)
            Dim firstLink As Variant
            Set groupLinks = groupDict(groupVar)
            If groupLinks.Count > 0 Then
                firstLink = groupLinks(1) ' Get page type from first link
                .Cells(rowGroup, 1).Value = "Page Type"
                .Cells(rowGroup, 1).Font.Bold = True
                .Cells(rowGroup, 2).Value = firstLink(2)
                .Cells(rowGroup, 2).Font.Bold = True ' Bold the page type value
                rowGroup = rowGroup + 1
            End If
            
            ' Link Name header
            .Cells(rowGroup, 1).Value = "Link Name"
            .Cells(rowGroup, 1).Font.Bold = True
            rowGroup = rowGroup + 1
            
            ' Links (Link Name in Column A, URL in Column B)
            Dim groupLinksColl As Object
            Set groupLinksColl = groupDict(groupVar)
            Dim i As Long
            For i = 1 To groupLinksColl.Count
                Dim linkPairVar As Variant
                linkPairVar = groupLinksColl(i)
                .Cells(rowGroup, 1).Value = linkPairVar(0) ' Link Name
                .Cells(rowGroup, 2).Value = linkPairVar(1) ' URL
                rowGroup = rowGroup + 1
            Next i
            
            rowGroup = rowGroup + 1 ' Extra row for spacing between groups
        Next groupVar
        
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 60
        .Columns("C").ColumnWidth = 20 ' Still set for consistency, but unused
        .Rows.RowHeight = 15
        .Cells.WrapText = True
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End With
    
    ' Report results
    Dim msg As String
    msg = "Link extraction complete!" & vbCrLf & vbCrLf & _
          "Processed files:" & processedFiles
    If Len(skippedFiles) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Skipped files:" & skippedFiles
    End If
    MsgBox msg, vbInformation
    
    ' Clean up
    Set fso = Nothing
    Set htmlDoc = Nothing
    Set linkDict = Nothing
    Set groupDict = Nothing
End Sub

' Helper function to read file content
Private Function ReadFile(filePath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile
    On Error Resume Next
    Open filePath For Input As #fileNum
    If Err.Number = 0 Then
        ReadFile = Input$(LOF(fileNum), fileNum)
    End If
    Close #fileNum
    On Error GoTo 0
End Function

' Helper function to clean string of non-printable characters
Private Function CleanString(str As String) As String
    Dim i As Long
    Dim clean As String
    clean = ""
    For i = 1 To Len(str)
        If Asc(Mid(str, i, 1)) >= 32 And Asc(Mid(str, i, 1)) <= 126 Then
            clean = clean & Mid(str, i, 1)
        End If
    Next i
    CleanString = clean
End Function

' Updated function to determine page type
Private Function DeterminePageType(htmlDoc As Object) As String
    Dim hasTiles As Boolean
    Dim hasList As Boolean
    Dim hasText As Boolean
    
    ' Check for tiles
    On Error Resume Next
    hasTiles = (htmlDoc.getElementsByClassName("tile").Length > 0 Or _
                htmlDoc.getElementsByClassName("w3-card").Length > 2) ' More than 2 for reliability
    Err.Clear
    On Error GoTo 0
    
    ' Check for lists (specifically tables with exllink)
    On Error Resume Next
    Dim tables As Object, table As Object
    Set tables = htmlDoc.getElementsByTagName("table")
    hasList = False
    For Each table In tables
        If InStr(table.className, "w3-table") > 0 Or table.getElementsByClassName("exllink").Length > 0 Then
            hasList = True
            Exit For
        End If
    Next table
    Err.Clear
    On Error GoTo 0
    
    ' Check for significant text content (excluding tiles and lists)
    On Error Resume Next
    Dim bodyText As String
    bodyText = htmlDoc.body.innerText
    hasText = (Len(Trim(bodyText)) > 100) ' Significant text length
    If hasTiles Then
        ' Filter out tile text to check for standalone text
        Dim tileText As String
        Dim tiles As Object, tile As Object
        Set tiles = htmlDoc.getElementsByClassName("tiletext")
        For Each tile In tiles
            tileText = tileText & tile.innerText
        Next tile
        hasText = hasText And (Len(Trim(bodyText)) - Len(Trim(tileText)) > 100)
    End If
    ' If no tiles, check for li tags as an additional text indicator
    If Not hasTiles And Not hasList Then
        Dim liCount As Long
        liCount = htmlDoc.getElementsByTagName("li").Length
        hasText = hasText Or (liCount > 0) ' Presence of li tags can indicate text content
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Determine page type based on content
    If hasTiles And Not hasList And Not hasText Then
        DeterminePageType = "Tiles"
    ElseIf hasTiles And Not hasList And hasText Then
        DeterminePageType = "Tiles + Text"
    ElseIf Not hasTiles And Not hasList And hasText Then
        DeterminePageType = "Text"
    ElseIf Not hasTiles And hasList And Not hasText Then
        DeterminePageType = "List"
    ElseIf Not hasTiles And hasList And hasText Then
        DeterminePageType = "List + Text"
    Else
        DeterminePageType = "Unknown"
    End If
End Function

