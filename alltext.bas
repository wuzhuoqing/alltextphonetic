Sub testInsertPhoneticGuide()
Call insertPhoneticGuide(Selection.Range)
End Sub
Sub insertPhoneticGuide(r As Word.Range)
'we then show field
'Therefore, here is my proposed workaround.
'Add the Phonetic Guide ruby text (i.e. 'furigana') in the usual way, using default settings for font, positioning, and font size.  Alternatively, you may have received a document created by somebody else with ruby text already added.
'Select all text (easy shortcut is CTRL A).
'Make the field codes visible with SHIFT F9.
'Unselect the text.  [optional]
'Do a find-and-replace operation (easy shortcut is CTRL H).  For example, search for all occurences of  Font:Arial  and replace them all automatically with (say)   Font:Times   .
'Select all text (easy shortcut is CTRL A).
'Make the field results visible either using SHIFT F9 (toggle field code visibility) or simply F9 (update field results).  Going to Print Preview might also work, depending upon your configuration.
'EQ \* jc2 \* "Font:DengXian" \* to EQ \* jc3 \* "Font:DengXian" \*

Dim d As Word.Dialog
Dim lng As Long
Dim lngChars As Long
Dim lng_end As Long
Dim r1 As Word.Range
Dim r2 As Word.Range
On Error Resume Next
Set d = Word.Dialogs(wdDialogPhoneticGuide)
Set r1 = r.Duplicate
r1.TextRetrievalMode.IncludeFieldCodes = False
For lng = Len(r1.Text) To 0 Step -20
  lng_end = lng - 20
  If lng_end < 0 Then
    lng_end = 0
  End If
  Set r2 = r.Duplicate
  r2.SetRange r.Start + lng_end, r.Start + lng
  ' Do not insert pinyin for any range that
  ' contains a field (this will prevent the code from re-inserting
  ' pinyin, but you can change the way this works if you like)
  If r2.Fields.Count = 0 Then
    r2.Select
    d.Show 1
    ' Error 6031 says there's no text to pinyin
    If Err.Number = 6031 Then
      ' Err.Clear
    Else
      ' On Error GoTo 0
    End If
  Else
    r2.Select
    Call insertPhoneticGuideOneByOne(r2)
  End If
Next
Set r2 = Nothing
Set r1 = Nothing
Set d = Nothing
End Sub

Sub insertPhoneticGuideOneByOne(r As Word.Range)
Dim d As Word.Dialog
Dim lng As Long
Dim lngChars As Long
Dim r1 As Word.Range
Dim r2 As Word.Range
On Error Resume Next
Set d = Word.Dialogs(wdDialogPhoneticGuide)
Set r1 = r.Duplicate
r1.TextRetrievalMode.IncludeFieldCodes = False
For lng = Len(r1.Text) To 1 Step -1
  Set r2 = r1.Characters(lng)
  ' Do not insert pinyin for any range that
  ' contains a field (this will prevent the code from re-inserting
  ' pinyin, but you can change the way this works if you like)
  If r2.Fields.Count = 0 Then
    r2.Select
    d.Show 1
    ' Error 6031 says there's no text to pinyin
    If Err.Number = 6031 Then
      ' Err.Clear
    Else
      ' On Error GoTo 0
    End If
  End If
Next
Set r2 = Nothing
Set r1 = Nothing
Set d = Nothing
End Sub
