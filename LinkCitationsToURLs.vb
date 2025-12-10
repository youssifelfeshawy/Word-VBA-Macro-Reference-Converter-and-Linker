Sub LinkCitationsToURLs()
    Dim doc As Document
    Dim rng As Range
    Dim bibRng As Range
    Dim regEx As Object
    Dim urlRegEx As Object
    Dim doiRegEx As Object
    Dim matches As Object
    Dim urlMatches As Object
    Dim doiMatches As Object
    Dim match As Object
    Dim citationNum As String
    Dim counter As Long
    Dim targetURL As String
    Dim bibText As String
    Dim paragraphText As String
    
    Set doc = ActiveDocument
    counter = 0
    
    ' Create regex to find citation numbers like [1], [18], [22], etc.
    Set regEx = CreateObject("VBScript.RegExp")
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
    
    ' Search through entire document
    Set rng = doc.Content
    Set matches = regEx.Execute(rng.Text)
    
    ' Process each citation found
    Dim i As Long
    For i = 0 To matches.Count - 1
        citationNum = matches(i).SubMatches(0) ' Get just the number without brackets
        
        ' Find this citation in the text (e.g., [18])
        Set rng = doc.Content
        With rng.Find
            .Text = "[" & citationNum & "]"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            
            Do While .Execute
                ' Check if this is already a hyperlink
                If rng.Hyperlinks.Count = 0 Then
                    ' Now find the corresponding bibliography entry
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
                                
                                counter = counter + 1
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
    MsgBox "Citation linking complete!" & vbCrLf & _
           counter & " citation(s) linked to their URLs/DOIs.", _
           vbInformation, "Citation Linker"
End Sub
