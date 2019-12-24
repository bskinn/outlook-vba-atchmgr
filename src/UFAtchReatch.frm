VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFAtchReatch 
   Caption         =   "Reattach Files"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   OleObjectBlob   =   "UFAtchReatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFAtchReatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




' # ------------------------------------------------------------------------------
' # Name:        UFAtchReatch.frm
' # Purpose:     Custom form, handling selection of files to re-attach
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     18 Jun 2017
' # Copyright:   (c) Brian Skinn 2017
' # License:     The MIT License; see "LICENSE.txt" for full license terms
' #                   and contributor agreement.
' #
' #       http://www.github.com/bskinn/outlook-vba-atchmgr
' #
' # ------------------------------------------------------------------------------

Option Explicit

Private filesColl As Collection, mi As MailItem

Private Sub BtnCancel_Click()
    Unload UFAtchReatch
End Sub

Private Sub BtnDoReattach_Click()
    Dim iter As Long, lf As LinkedFile, fs As FileSystemObject, wd As Word.Document
    Dim wdRg As Word.Range
    Dim hl As Hyperlink, iter2 As Long, foundHL As Boolean
    
    ' Link the file system
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Iterate through the listbox
    For iter = 0 To LBxFileList.ListCount - 1
        ' If checked, process the entry
        If LBxFileList.Selected(iter) Then
            ' Link the lf object
            Set lf = filesColl.Item(iter + 1)
        
            ' (Re-)attach the file; this returns the item to non-edit mode for a message
            '  that is not a draft-in-progress.  Check to ensure file still exists before
            '  attaching -- if the same file is also linked via a hashed block, it could have
            '  disappeared
            If fs.FileExists(lf.LinkAddress) Then
                Call mi.Attachments.Add(lf.LinkAddress)
            Else
                MsgBox "File """ & lf.dispName & """ no longer exists; cannot attach.", _
                            vbOKOnly + vbExclamation, "Cannot attach file"
            End If
            
            ' Save here in the event that something crashy happens
            mi.Save
            
            ' If was a detached link, cull the paragraph
            If lf.isHashed Then
                ' If not a draft-in-progress and running Office 2013, must restore edit mode
                If Not mi.Parent.EntryID = Session.GetDefaultFolder(olFolderDrafts).EntryID And Not _
                                mi.Parent.EntryID = Session.GetDefaultFolder(olFolderOutbox).EntryID And Not _
                                Left(Application.Version, 2) = "15" Then
                    ' Message is not a draft-in-progress and must be reset to edit mode
                    ' Sometimes the move into edit mode can fail, though, so enclose
                    ' with error trap that discards any errors
                    On Error Resume Next
                        Call mi.GetInspector.CommandBars.ExecuteMso("EditMessage")
                    Err.Clear: On Error GoTo 0
                End If
                
                ' Reattach the editor
                Set wd = mi.GetInspector.WordEditor
                
                ' Appear to always need to reassign Hyperlink after attaching the file
                ' Must re-search for the hyperlink in the document
                foundHL = False
                For iter2 = 1 To wd.Hyperlinks.Count
                    Set hl = wd.Hyperlinks(iter2)
                    If (InStr(lf.LinkAddress, hl.Address) > 0) And _
                                hl.TextToDisplay = lf.LinkText Then
                        foundHL = True
                        Exit For
                    End If
                Next iter2
                
                ' If the hyperlink was not retrieved, complain and do not process
                If Not foundHL Then
                    ' Block not found
                    MsgBox "Detached file annotation block for """ & lf.dispName & _
                            """ not found. Skipping removal of annotation block.", _
                            vbOKOnly + vbInformation, "Annotation block not found"
                Else
                    ' Block found; strip
                    ' Bind the whole paragraph
                    Set wdRg = hl.Range.Paragraphs(1).Range
                    
                    ' Simplest way to delete paragraph content
                    wdRg.Text = ""
                    Call wdRg.delete(wdCharacter, 1)
                    
                    ' Save here
                    mi.Save
                End If
                
                ' Check whether to delete the stored file
                If LBxFileList.List(iter, 1) = "Yes" Then
                    ' Do delete
                    fs.DeleteFile lf.LinkAddress, True
                End If
            End If  ' isHashed
        End If  ' .Selected
    Next iter
    
    ' Save the message one last time, just in case.
    mi.Save
    
    ' Close the form
    Unload UFAtchReatch
    
End Sub

Private Sub LBxFileList_Change()
    Dim iter As Long, somethingSelected As Boolean
    
    ' Initialize to nothing found
    somethingSelected = False
    
    ' See if anything selected
    For iter = 0 To LBxFileList.ListCount - 1
        somethingSelected = somethingSelected Or LBxFileList.Selected(iter)
    Next iter
    
    ' If nothing selected, disable the 'do reattach' button
    BtnDoReattach.Enabled = somethingSelected
    
End Sub

Private Sub LBxFileList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MsgBox "1: " & LBxFileList.List(LBxFileList.ListIndex, 0) & Chr(10) & Chr(10) & _
                "2: " & LBxFileList.Value
End Sub

Private Sub LBxFileList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim clickIndex As Long, chgIndex As Long
    
    ' Initialize the clickIndex to a flag value
    clickIndex = -1
    
    If Button = 2 Then  ' Right mouse button
        ' Check which row of the listbox was right-clicked
        If Y >= 0.75 And Y <= 11.25 Then
            clickIndex = 0
        ElseIf Y >= 14.25 And Y <= 24.75 Then
            clickIndex = 1
        ElseIf Y >= 26.25 And Y <= 37.55 Then
            clickIndex = 2
        ElseIf Y >= 39 And Y <= 51 Then
            clickIndex = 3
        ElseIf Y >= 52.55 And Y <= 63.05 Then
            clickIndex = 4
        ElseIf Y >= 64.5 And Y <= 75.8 Then
            clickIndex = 5
        ElseIf Y >= 78.05 And Y <= 87.8 Then
            clickIndex = 6
        ElseIf Y >= 90.05 And Y <= 100.55 Then
            clickIndex = 7
        End If
        
        ' If the click fell through, just exit the sub
        If clickIndex = -1 Then Exit Sub
        
        ' Toggle the 'delete' Yes/No value if the click was on an existing
        '  element of the list
        With LBxFileList
            chgIndex = clickIndex + .TopIndex
            If chgIndex < .ListCount Then
                .List(chgIndex, 1) = swapYesNo(.List(chgIndex, 1))
            End If
        End With
        
    End If
    
End Sub

Private Sub UserForm_Activate()
    Dim iter As Long, lf As LinkedFile
    
    ' Initialize all deletes to FALSE
    For iter = 0 To LBxFileList.ListCount - 1
        ' Typed object for convenience
        Set lf = filesColl.Item(iter + 1)
        
        ' Only set 'deleteable possible' if it's hashed
        If lf.isHashed Then
            LBxFileList.List(iter, 1) = "No"
        Else
            LBxFileList.List(iter, 1) = "N/A"
        End If
    Next iter
    
End Sub

Public Sub popFormStuff(listColl As Collection, mItem As MailItem)
    Dim itm As Object, lf As LinkedFile
    
    Set filesColl = listColl
    Set mi = mItem
    
    ' Iterate over all the items in the collection
    For Each itm In listColl
        ' Set to typed object for convenience
        Set lf = itm

        ' Add the display name to the listbox
        LBxFileList.AddItem lf.dispName
    Next itm

End Sub

Private Function swapYesNo(str As String) As String
    If str = "No" Then
        swapYesNo = "Yes"
    ElseIf str = "Yes" Then
        swapYesNo = "No"
    Else
        ' Do nothing; leave it the same
        swapYesNo = str
    End If
End Function
