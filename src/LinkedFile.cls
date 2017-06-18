VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkedFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' # ------------------------------------------------------------------------------
' # Name:        LinkedFile.cls
' # Purpose:     Class containing information on linked reattachable files
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

Private hashTagged As Boolean

' Should be enough to uniquely specify a Hyperlink, though in the rare case
'  of exactly the same file being linked with exactly the same text, confusion
'  might result.  Could proofread entire Hyperlinks at some point?
Private URLAddress As String  ' Matches to .Address
Private URLText As String    ' Matches to .TextToDisplay

Private hlObj As Hyperlink   ' Will need this to cull 'detached file' info blocks?
Private fileName As String ' Filename to display in listbox


Private Sub Class_Initialize()
    hashTagged = False
End Sub

Public Sub setHyperlink(hl As Hyperlink, name As String)
    ' Extracts the URL .Address and .TextToDisplay strings for storage
    '  Fragile to try to retrieve Hyperlink location within the WordEditor,
    '  so not going to mess with it
    If hl.SubAddress <> "" Then
        URLAddress = hl.Address & "#" & hl.SubAddress
    Else
        URLAddress = hl.Address
    End If
    
    URLText = hl.TextToDisplay
    Set hlObj = hl
    fileName = name
End Sub

Public Function getHyperlink() As Hyperlink
    Set getHyperlink = hlObj
End Function

Public Function matchesHyperlink(hl As Hyperlink)
    ' If target and text are identical, indicate the Hyperlink matches
    '  Potentially problematic if multiple Hyperlinks exist in the document
    '  with identical .Address and .TextToDisplay properties, but for
    '  reattachment purposes, one wouldn't want to reattach the same file twice
    '  anyways.
    '    Hm. Something to check for when parsing the Hyperlinks in the WordEditor:
    '    Crosscheck the .Address of the Hyperlink under examination, and if it
    '    points to an .Address already linked then exclude it from the list of files to reattach
    '    and from the Collection of LinkedFile's?
    If hlObj.Address = hl.Address And hlObj.TextToDisplay = hl.TextToDisplay Then
        matchesHyperlink = True
    Else
        matchesHyperlink = False
    End If
End Function

Property Get LinkAddress() As String
    LinkAddress = URLAddress
End Property

Property Get LinkText() As String
    LinkText = URLText
End Property

Property Get isHashed() As Boolean
    isHashed = hashTagged
End Property

Property Let isHashed(hashed As Boolean)
    hashTagged = hashed
End Property

Property Get dispName() As String
    dispName = fileName
End Property

