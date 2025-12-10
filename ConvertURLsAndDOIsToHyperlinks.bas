Sub ConvertURLsAndDOIsToHyperlinks()
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
    
    ' Initialize
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
    
    ' Show completion message
    MsgBox "Conversion complete!" & vbCrLf & _
           totalCounter & " URL(s) and DOI(s) converted to hyperlinks.", _
           vbInformation, "URL & DOI Converter"
End Sub
