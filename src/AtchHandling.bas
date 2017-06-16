Attribute VB_Name = "AtchHandling"
Option Explicit

Sub EditForAtchAnnot()
    Dim insp As Inspector, wdEd As Word.Document
    Dim itm As Object
    
    If ActiveInspector Is Nothing Then Exit Sub
    
    Set insp = ActiveInspector
    Set wdEd = insp.WordEditor
    Set itm = insp.CurrentItem
    
    With itm
        ' Close and reopen item to provide consistent state
        Call .Close(olSave)
        .Display
        Call .GetInspector.CommandBars.ExecuteMso("EditMessage")  ' Activate edit mode (presumes valid)
        Call .GetInspector.CommandBars.ExecuteMso("MessageFormatHtml")  ' Set to HTML mode
    End With
    
    wdEd.StoryRanges(wdMainTextStory).InsertBefore "[#" & vbCrLf & vbCrLf
    With wdEd.StoryRanges(wdMainTextStory).Paragraphs(1).Range
        .Font.TextColor.RGB = RGB(120, 113, 68)
        .Font.Italic = True
        .Characters(2).Select
        .Characters(2).delete wdCharacter, 1
    End With

End Sub

Public Sub ReattachAttachments()
    ' Will be better done by collecting all (re)attach-able files and presenting in a checkbox-enabled
    '  UserForm, permitting user to select which ones to (re)attach and whether or not to retain the
    '  out-of-message copies of any that were originally macro-detached
    '
    ' Macro in present form will reattach all linked, locally-accessible files, including any that were
    '  linked in the original text of the message.  This may not be the desired behavior.
    
    Dim wd As Word.Document, wdRg As Word.Range
    Dim fs As FileSystemObject
    Dim hl As Hyperlink
    Dim insp As Inspector
    Dim mi As MailItem
    Dim lf As LinkedFile, lfColl As New Collection, lfIter As LinkedFile, itm As Object
    Dim alreadyLinked As Boolean
    
    Dim rx As New RegExp
    
    Dim fNameFull As String, fName As String, fPath As String
    Dim fullAddress As String
    Dim iter As Long, atchsExist As Boolean
    
    ' If no active inspector, just silently drop
    If ActiveInspector Is Nothing Then Exit Sub
    
    Set insp = ActiveInspector
    Set wd = insp.WordEditor
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Set up RegEx
    With rx
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "[a-z]:\\"
    End With
    
    ' Check each hyperlink for whether it's local or not; if any found, indicate as such
    atchsExist = False
    For Each hl In wd.Hyperlinks
        If Left(hl.Address, 2) = "\\" Or rx.Test(hl.Address) Then
            atchsExist = True
            Exit For
        End If
                    '(rx.Test(hl.Address) And Not Left(hl.Address, 4) = "http") Then atchsExist = True
    Next hl
    
    ' If found attachments to reattach, perform reattachment
    If atchsExist Then
        Load UFAtchReatch
        Set mi = insp.CurrentItem  ' Comment these lines
        Call mi.Close(olSave)
        mi.Display
        Set insp = mi.GetInspector
        Set wd = insp.WordEditor
        If Not mi.Parent.EntryID = Session.GetDefaultFolder(olFolderDrafts).EntryID And Not _
                mi.Parent.EntryID = Session.GetDefaultFolder(olFolderOutbox).EntryID Then
            Call insp.CommandBars.ExecuteMso("EditMessage")  ' To here
            ' Breaks if drafting a message using 'Resend Message' w/o having saved a draft first
        End If
        
        iter = 1  ' Must initialize
        Do While iter <= wd.Hyperlinks.Count
            Set hl = wd.Hyperlinks(iter)
            If Left(hl.Address, 2) = "\\" Or rx.Test(hl.Address) Then
                        '(rx.Test(hl.Address) And Not Left(hl.Address, 4) = "http") Then
                ' Local file; attempt reattach
                ' Parse into folder and filename
                ' Have to tag on the anchor name, if present -- "#" is an assumption, but is the
                '   primary character encountered thus far in my uses
                fullAddress = hl.Address
                If Len(hl.SubAddress) > 0 Then fullAddress = fullAddress & "#" & hl.SubAddress
                
                fNameFull = fs.GetAbsolutePathName(fullAddress)
                fPath = fs.GetParentFolderName(fNameFull)
                fName = fs.GetFileName(fNameFull)
                
                If fs.FolderExists(fPath) Then      ' Folder exists (add notification if not exist)
                    If fs.FileExists(fNameFull) Then    ' File exists (add notification if not exist)
                        ' Check if hl already linked
                        alreadyLinked = False
                        For Each itm In lfColl
                            Set lfIter = itm
                            If lfIter.matchesHyperlink(hl) Then
                                alreadyLinked = True
                                Exit For
                            End If
                        Next itm
                        
                        ' Check for already linked
                        If Not alreadyLinked Then
                            ' Set up the linked file object if not already linked
                            Set lf = New LinkedFile
                            'lf.DeleteHashed = False
                            'lf.ID = lfColl.Count + 1
                            lf.setHyperlink hl, fName
                            'lf.ProcessFile = False
                        
                            ' Flag for whether it's hashed or not
                            Set wdRg = hl.Range.Paragraphs(1).Range
                            If Left(wdRg.Text, 3) = "###" And Right(wdRg.Text, 4) = "###" & Chr(13) Then
                                lf.isHashed = True
                            Else
                                lf.isHashed = False
                            End If
                            
                            ' Add the linked file to the collection
                            lfColl.Add lf
                        End If  ' alreadyLinked
                    End If  ' FileExists
                End If  ' FolderExists
            End If  ' local link
            
            ' Increase the iterator
            iter = iter + 1
        Loop    ' Until all hl's checked
        
        ' Populate the reattachment form
        UFAtchReatch.popFormStuff lfColl, mi
        
        ' Show the form
        UFAtchReatch.Show
        
    End If  ' Local links found
    
    ' Dereference variables
    Set fs = Nothing
    Set wd = Nothing
    Set insp = Nothing
    Set rx = Nothing
    Set hl = Nothing
    Set mi = Nothing
    
End Sub

Sub DetachAttachment()
    Dim atch As Attachment, atchSel As AttachmentSelection
    Dim insp As Inspector
    Dim itm As Object
    Dim sh As Shell32.Shell
    Dim fs As FileSystemObject
    Dim fld As Folder2
    Dim fName As String, extn As String, baseName As String, fullSavePath As String
    Dim atchName As String
    Dim okfName As Boolean
    Dim iter As Long, timeRef As Long
    Dim bodyHTML As String
    
    ' Bind script object, inspector,  and attachment selection
    Set sh = CreateObject("Shell.Application")
    Set insp = Application.ActiveInspector
    Set itm = insp.CurrentItem
    Set atchSel = insp.AttachmentSelection
    
    ' Create filesystem object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Initialize folder to Nothing; should be redundant
    Set fld = Nothing
    
    ' Close-with-save the Item and re-open for consistent, stable state DOESN'T WORK
    'Call itm.Close(olSave)
    'itm.Display
    'Set insp = itm.GetInspector
    
    ' Loop selected attachments, asking to detach or not
    For Each atch In atchSel
        ' Query for fld if none yet selected
        If fld Is Nothing Then Set fld = sh.BrowseForFolder(0, "Detach file(s) to:", 1, 17)
        
        ' Check whether folder was selected
        If Not fld Is Nothing Then
            ' Request filename from user, checking/confirming overwrite
            okfName = False ' Initialize no-good filename
            Do
                atchName = atch.fileName  ' Store attachment filename
                fName = atch.fileName ' Initialize working filename
                ' Split working filename into base and extension
                iter = Len(fName)
                Do Until Mid(fName, iter, 1) = "." Or iter = 1: iter = iter - 1: Loop
                If iter > 1 Then ' period found
                    baseName = Left(fName, iter - 1)
                    extn = Right(fName, Len(fName) - iter)
                Else ' period not found
                    baseName = fName
                    extn = ""
                End If
                
                ' Query filename; stop exec if zero-length return (user cancel); reconstruct with extension
                fName = InputBox("Save to filename:", "Enter File Name", baseName)
                If Len(fName) < 1 Then Exit For ' okfName already False; fragile if another surrounding For..Next added
                If Len(extn) > 0 Then fName = fName & "." & extn
                ' Check whether filename ok based on file existence
                okfName = Not fs.FileExists( _
                            IIf(Right(fld.Self.Path, 1) = "\", _
                                    fld.Self.Path, _
                                    fld.Self.Path & "\" _
                                ) & cleanFilename(fName))
                ' If exists (name is not okay), ask if overwrite ok
                If Not okfName Then
                    ' Need to deal with not-ok filename
                    Select Case MsgBox("File exists" & Chr(10) & Chr(10) & "Overwrite?", _
                                vbYesNoCancel + vbExclamation, "Confirm Overwrite")
                    Case vbYes
                        ' Go ahead and overwrite
                        okfName = True
                    Case vbNo
                        ' Just pass through the not-ok-filename flag
                    Case Else
                        ' Presumably just vbCancel is possible; exit sub
                        Exit For ' This is weak; addition of another wrapping For..Next will break code
                    End Select
                End If
            Loop Until okfName
                            
            ' If anything survives filename cleaning...
            If Not cleanFilename(fName) = "" Then
                ' Save atch to path\filename and delete
                fullSavePath = IIf(Right(fld.Self.Path, 1) = "\", _
                                        fld.Self.Path, _
                                        fld.Self.Path & "\" _
                                    ) & cleanFilename(fName)
                If Len(fName) = Len(cleanFilename(fName)) Then
                    ' To robustify, add wait loop that checks to be sure saved-out file exists
                    '  and has nonzero size (filesize match is not useful; attached file does not
                    '  have identical reported bytesize as on-disk file)
                    Call atch.SaveAsFile(fullSavePath)  ' FRAGILE if Save op unexpectedly fails!
                    Call atch.delete
                    'timeRef = Timer: Do While Timer <= timeRef + 0.5: DoEvents: Loop
                    Call tagTextIntoEmail(itm, atchName, fullSavePath)
                    Call itm.Save
                Else
                    ' Some invalid characters stripped; notify and save
                    Select Case MsgBox("Invalid characters have been stripped from the indicated " & _
                                "filename. File saved as:" & Chr(10) & Chr(10) & _
                                cleanFilename(fName), vbOKCancel + vbExclamation, _
                                "Invalid Characters Removed")
                    Case vbOK
                        Call atch.SaveAsFile(fullSavePath)  ' FRAGILE if Save op unexpectedly fails!
                        Call atch.delete
                        'timeRef = Timer: Do While Timer <= timeRef + 0.5: DoEvents: Loop
                        Call tagTextIntoEmail(itm, atchName, fullSavePath)
                        Call itm.Save
                    Case Else
                        ' Do nothing
                    End Select
                End If
            End If
        Else
            ' No folder is set; presume that user cancelled & wants to exit routine
            Exit For
        End If
    Next atch
    
    ' Dereference objects
    Set sh = Nothing
    Set fld = Nothing
    Set itm = Nothing
    Set insp = Nothing
    Set atchSel = Nothing
    Set atch = Nothing
    
End Sub

Public Sub tagTextIntoEmail(itm As Object, atchName As String, saveName As String)
    ' Might still be able to use Word.Document
    Dim wd As Word.Document, newRg As Word.Range, editRg As Word.Range, mi As MailItem
    Dim idx As Long
    Const verbStr As String = "' detached to "
    
    With itm
        ' Close and reopen item to provide consistent state
        Call .Close(olSave)
        .Display
        Call .GetInspector.CommandBars.ExecuteMso("EditMessage")  ' Activate edit mode (presumes valid)
        Call .GetInspector.CommandBars.ExecuteMso("MessageFormatHtml")  ' Set to HTML mode
    End With
    
    ' Attach Word Document for editing
    Set wd = itm.GetInspector.WordEditor
    
    ' Insert attachment detachment notification text
    Call wd.Content.InsertBefore("###Attachment '" & atchName & verbStr & saveName & Chr(13) & Chr(13))
    
    ' Bind the newly added paragraph's Range
    Set newRg = wd.Content.Paragraphs(1).Range
    With newRg
        ' Change color to red
        .Font.ColorIndex = wdRed
        ' Identify where link location starts
        idx = InStr(.Text, verbStr) + Len(verbStr)
        ' Set editing Range to first character
        Set editRg = .Characters(idx)
        ' Extend editing Range to end of paragraph
        Call editRg.MoveEnd(Word.WdUnits.wdParagraph, 1)
        ' Deselect hard return
        Call editRg.MoveEnd(Word.WdUnits.wdCharacter, -1)
        ' Append hashes
        Call editRg.InsertAfter("###")
        ' Deselect hashes
        Call editRg.MoveEnd(Word.WdUnits.wdCharacter, -3)
        ' Apply hyperlink -- EXCISED for security
        Call wd.Hyperlinks.Add(editRg, editRg.Text)

    End With
    
End Sub

Public Function cleanFilename(ByVal fn As String) As String
    Dim badchrs() As Variant, val As Long, st As String, ch As String
    badchrs = Array("\", "/", ":", """", "*", "?", "<", ">", "|")
    
    ' Set string to shorthand variable
    st = fn
    
    ' Search for and remove bad characters
    For val = LBound(badchrs) To UBound(badchrs)
        ch = badchrs(val)
        Do While InStr(st, ch) > 0
            st = Left(st, InStr(st, ch) - 1) & Right(st, Len(st) - InStr(st, ch))
        Loop
    Next val
    
    ' Set the cleaned filename to the output variable
    cleanFilename = st
    
End Function
