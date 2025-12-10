Sub ConvertAndLinkReferences()
    ' This macro combines ConvertURLsAndDOIsToHyperlinks and LinkCitationsToURLs
    ' It first converts URLs and DOIs to hyperlinks, then links citations to them
    
    Dim doc As Document
    Dim rng As Range
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim linkText As String
    Dim linkAddress As String
    Dim counter As Long
    Dim totalCounter As Long
    Dim response As VbMsgBoxResult
    Dim keepGoing As Boolean
   
    ' Initialize for URL/DOI conversion
    Set doc = ActiveDocument
    totalCounter = 0
   
    ' Ask user if they want to convert only selected text or entire document
    response = MsgBox("Convert URLs and DOIs in:" & vbCrLf & vbCrLf & _
                      "YES = Selected text only" & vbCrLf & _
                      "NO = Entire document" & vbCrLf & vbCrLf & _
                      "(Select your bibliography first if you choose YES)", _
                      vbYesNoCancel + vbQuestion, "URL & DOI Converter")
   
    If response = vbCancel Then Exit Sub
   
    ' Set range based on user choice
    If response = vbYes Then
        If Selection.Type = wdSelectionIP Then
            MsgBox "Please select the text containing URLs/DOIs first.", vbExclamation
            Exit Sub
        End If
        Set rng = Selection.Range
    Else
        Set rng = doc.Content
    End If
   
    ' Create regex object
    Set regEx = CreateObject("VBScript.RegExp")
   
    With regEx
        .Global = True
        .IgnoreCase = True
    End With
   
    ' Keep looping until no new links are created
    Do
        counter = 0
       
        ' Store original range end to avoid issues
        Dim originalEnd As Long
        originalEnd = rng.End
       
        ' STEP 1: Convert regular URLs (http/https)
        regEx.Pattern = "https?://[^\s\]\)>,]+"
        Set matches = regEx.Execute(rng.Text)
       
        Dim searchRng As Range
        For Each match In matches
            linkText = match.Value
            linkAddress = linkText
           
            Set searchRng = rng.Duplicate
            With searchRng.Find
                .Text = linkText
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
               
                Do While .Execute
                    If searchRng.End <= originalEnd And searchRng.Hyperlinks.Count = 0 Then
                        doc.Hyperlinks.Add Anchor:=searchRng, Address:=linkAddress
                        counter = counter + 1
                        searchRng.Collapse wdCollapseEnd
                        searchRng.End = originalEnd
                    Else
                        Exit Do
                    End If
                Loop
            End With
        Next match
       
        ' Update range after adding hyperlinks
        originalEnd = rng.End
       
        ' STEP 2: Convert DOIs (format: doi: 10.xxxx/xxxxx)
        ' Match everything after "doi: 10." until we hit whitespace or sentence-ending punctuation
        ' Allow periods within the DOI but not at the very end
        regEx.Pattern = "doi:\s*10\.[0-9]+/[^\s\)>,]+[^\s\)>,\.]"
        Set matches = regEx.Execute(rng.Text)
       
        For Each match In matches
            linkText = match.Value
            ' Extract just the DOI part (remove "doi: " prefix)
            linkAddress = "https://doi.org/" & Trim(Mid(linkText, 5))
           
            Set searchRng = rng.Duplicate
            With searchRng.Find
                .Text = linkText
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
               
                Do While .Execute
                    If searchRng.End <= originalEnd And searchRng.Hyperlinks.Count = 0 Then
                        doc.Hyperlinks.Add Anchor:=searchRng, Address:=linkAddress
                        counter = counter + 1
                        searchRng.Collapse wdCollapseEnd
                        searchRng.End = originalEnd
                    Else
                        Exit Do
                    End If
                Loop
            End With
        Next match
       
        ' Add to total counter
        totalCounter = totalCounter + counter
       
        ' If no new links were created this round, we're done
        keepGoing = (counter > 0)
       
    Loop While keepGoing
   
    ' Now proceed to link citations
    Dim bibRng As Range
    Dim urlRegEx As Object
    Dim doiRegEx As Object
    Dim urlMatches As Object
    Dim doiMatches As Object
    Dim citationNum As String
    Dim linkCounter As Long
    Dim targetURL As String
    Dim paragraphText As String
    Dim bibStartPos As Long
    Dim paraText As String
    Dim limitPos As Long
  
    linkCounter = 0
  
    ' Find the start of the Bibliography section
    Set bibRng = doc.Content
    bibStartPos = 0
    With bibRng.Find
        .Text = "Bibliography"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            bibRng.Expand wdParagraph
            paraText = Trim(Replace(bibRng.Text, vbCr, ""))
            If paraText = "Bibliography" Then
                bibStartPos = bibRng.Start
                Exit Do
            End If
            bibRng.Collapse wdCollapseEnd
        Loop
    End With
  
    If bibStartPos > 0 Then
        limitPos = bibStartPos
    Else
        limitPos = doc.Content.End
    End If
  
    ' Set range for regex to find citations before limit
    Set rng = doc.Range(Start:=doc.Content.Start, End:=limitPos)
  
    ' Create regex to find citation numbers like [1], [18], [22], etc.
    With regEx
        .Global = True
        .Pattern = "\[(\d+)\]"
    End With
  
    ' Create regex to find URLs
    Set urlRegEx = CreateObject("VBScript.RegExp")
    With urlRegEx
        .Global = True
        .Pattern = "https?://[^\s\]\)>,]+"
    End With
  
    ' Create regex to find DOIs
    Set doiRegEx = CreateObject("VBScript.RegExp")
    With doiRegEx
        .Global = True
        .Pattern = "doi:\s*10\.[0-9]+/[^\s\)>,]+[^\s\)>,\.]"
    End With
  
    ' Search through the defined range (up to Bibliography)
    Set matches = regEx.Execute(rng.Text)
  
    ' Process each citation found
    Dim i As Long
    For i = 0 To matches.Count - 1
        citationNum = matches(i).SubMatches(0) ' Get just the number without brackets
      
        ' Find this citation in the text (e.g., [18])
        Set rng = doc.Range(Start:=doc.Content.Start, End:=doc.Content.Start) ' Start at beginning as insertion point
        With rng.Find
            .Text = "[" & citationNum & "]"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
          
            Do While .Execute
                If rng.Start >= limitPos Then Exit Do
              
                ' Check if this is already a hyperlink
                If rng.Hyperlinks.Count = 0 Then
                    ' Now find the corresponding bibliography entry (search whole doc for bib)
                    Set bibRng = doc.Content
                  
                    ' Search for the bibliography entry (starts with [num] and tab)
                    With bibRng.Find
                        .Text = "[" & citationNum & "]" & vbTab
                        .Forward = True
                        .Wrap = wdFindStop
                        .Format = False
                      
                        If .Execute Then
                            ' Found the bibliography entry, now get the whole paragraph
                            bibRng.Expand wdParagraph
                            paragraphText = bibRng.Text
                          
                            ' Try to find a URL in this bibliography entry
                            Set urlMatches = urlRegEx.Execute(paragraphText)
                          
                            If urlMatches.Count > 0 Then
                                ' Found a URL, use it
                                targetURL = urlMatches(0).Value
                            Else
                                ' No URL found, check for DOI
                                Set doiMatches = doiRegEx.Execute(paragraphText)
                              
                                If doiMatches.Count > 0 Then
                                    ' Found a DOI, convert it to URL
                                    Dim doiText As String
                                    doiText = doiMatches(0).Value
                                    targetURL = "https://doi.org/" & Trim(Mid(doiText, 5))
                                Else
                                    ' No URL or DOI found, skip this citation
                                    targetURL = ""
                                End If
                            End If
                          
                            ' If we found a target URL, create the hyperlink
                            If targetURL <> "" Then
                                doc.Hyperlinks.Add Anchor:=rng, _
                                    Address:=targetURL, _
                                    ScreenTip:="Open reference [" & citationNum & "]"
                              
                                linkCounter = linkCounter + 1
                            End If
                        End If
                    End With
                End If
              
                ' Move to next occurrence
                rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
   
    ' Show completion message
    MsgBox "Conversion and linking complete!" & vbCrLf & _
           totalCounter & " URL(s) and DOI(s) converted to hyperlinks." & vbCrLf & _
           linkCounter & " citation(s) linked to their URLs/DOIs.", _
           vbInformation, "Reference Converter and Linker"
End Sub
