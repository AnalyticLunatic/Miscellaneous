' Purpose:      Word Macro to highlight in .docx files where Bookmark designations are overlapping same characters.
'
' Description:  A third party framework I used to work with heavily utilized Bookmarks in MS Word files to dynamically
'               create all manner of correspondence. Using their third party editor, it was not uncommon for Bookmark
'               designations to wind up overlapping same characters within a document, causing errors from the framework 
'               at generation. This macro code helped to immediately identify these problem areas in large and complicated templates.
'
' Author:       James Scurlock
'-----------------------------------------------------------------------------------------------------------------------
Sub Show_Overlapping_Bookmark_Designations()
  Dim objBookmark As bookmark
  Dim objDoc As Document
  Dim countBookmarksOnText As Integer
  
  Application.ScreenUpdating = False
 
  Set objDoc = ActiveDocument
 
  With objDoc
    For Each objBookmark In .Bookmarks
        ActiveDocument.Bookmarks(objBookmark).Select
        countBookmarksOnText = Selection.Bookmarks.Count
        If (countBookmarksOnText > 1) Then objBookmark.Range.HighlightColorIndex = wdBrightGreen
    Next objBookmark
  End With
  Application.ScreenUpdating = True

End Sub
