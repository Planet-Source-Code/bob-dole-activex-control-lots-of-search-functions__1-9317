VERSION 5.00
Begin VB.UserControl Search 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ScaleHeight     =   735
   ScaleWidth      =   1530
   Begin VB.Label Label2 
      Caption         =   "By: Andrew Juergens"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Search Functions 1.1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Add item Search
Public Function AddItemSearch(recordset As ADODB.recordset, searchcolumn As String, searchfor As String, KeyColumn As String, AddItemObject As Object)
  recordset.MoveFirst
 While (recordset.EOF = False)
  If (recordset.Fields(searchcolumn).Value = searchfor) Then
   If (KeyColumn <> "") Then
    listobject.AddItem (recordset.Fields(KeyColumn) & " | " & recordset.Fields(searchcolumn))
   Else
    listobject.AddItem (recordset.Fields(searchcolumn))
   End If
   recordset.MoveNext
  End If
 Wend
 End Function
 
 'Advanced Add Item Search
 Public Function AdvAddItemSearch(recordset As ADODB.recordset, searchcolumn As String, searchfor As String, KeyColumn As String, Acuracy As Long, AddItemObject As Object)
 recordset.MoveFirst
 While (recordset.EOF = False)
  If (InStr(1, recordset.Fields(searchcolumn).Value, Left(searchfor, Acuracy), vbTextCompare) <> 0) Then
   If (KeyColumn <> "") Then
    listobject.AddItem (recordset.Fields(KeyColumn) & " | " & recordset.Fields(searchcolumn))
   Else
    listobject.AddItem (recordset.Fields(searchcolumn))
   End If
   recordset.MoveNext
  End If
 Wend
End Function

'Advanced Text Search
Public Function AdvTextSearch(TextObject As Object, Lookfor As String, message As String, Acuracy As Long)
 Dim done As Integer
 Dim start As String
 done = 1
 While (Len(TextObject.Text) < done)
  If (InStr(start, TextObject.Text, Left(Lookfor, Acuracy), vbTextCompare) <> 0) Then
   TextObject.SelStart = start
   TextObject.SelLength = Len(Lookfor)
   If (MsgBox(message & " Continue Looking?", vbYesNo, "Found") = vbNo) Then
    Exit Function
   End If
  End If
 Wend
End Function

'Letter Add Item Search
Public Function LetterAddItemSearch(recordset As ADODB.recordset, searchcolumn As String, searchfor As String, KeyColumn As String, Acuracy As Long, AddItemObject As Object)
 Dim word As Variant
 Dim x As Integer
 Dim start As Integer
 Dim done As Integer
 done = 1
 start = 1
 x = 1
 word = Array(x)
 While x < Len(searchfor)
  word(x) = Mid(searchfor, start, Acuracy)
  start = start + 1
  x = x + 1
 Wend
 recordset.MoveFirst
 While (recordset.EOF = False)
  While (done < x)
    If (InStr(1, recordset.Fields(searchcolumn), word(done), vbTextCompare <> 0)) Then
     If (KeyColumn <> "") Then
      listobject.AddItem (recordset.Fields(KeyColumn) & " | " & recordset.Fields(searchcolumn))
    Else
     listobject.AddItem (recordset.Fields(searchcolumn))
    End If
   recordset.MoveNext
   End If
  Wend
 Wend
End Function

'Letter Text Search
Public Function LetterTextSearch(TextObject As Object, Lookfor As String, message As String, Acuracy As Long)
 Dim word As Variant
 Dim x As Integer
 Dim start As Integer
 Dim done As Integer
 done = 1
 start = 1
 x = 1
 word = Array(x)
 While x < Len(searchfor)
  word(x) = Mid(searchfor, start, Acuracy)
  start = start + 1
  x = x + 1
 Wend
 While (Len(textobjects.Text) < done)
  While (done < x)
   If (InStr(start, TextObject.Text, word(done), vbTextCompare) <> 0) Then
    TextObject.SelStart = start
    TextObject.SelLength = Len(Lookfor)
    If (MsgBox(message & " Continue Looking?", vbYesNo, "Found") = vbNo) Then
     Exit Function
    End If
   End If
  Wend
 Wend
End Function

'Text Search
Public Function TextSearch(TextObject As Object, Lookfor As String, message As String)
 Dim done As Integer
 Dim start As String
 done = 1
 While (Len(TextObject.Text) < done)
  If (InStr(start, TextObject.Text, Lookfor, vbTextCompare) <> 0) Then
   TextObject.SelStart = start
   TextObject.SelLength = Len(Lookfor)
   If (MsgBox(message & " Continue Looking?", vbYesNo, "Found") = vbNo) Then
    Exit Function
   End If
  End If
 Wend
End Function

'Find and replace Text
'Text Search
Public Function FRTextSearch(TextObject As Object, Lookfor As String, Replace As String, message As String)
 Dim done As Integer
 Dim start As String
 done = 1
 While (Len(TextObject.Text) < done)
  If (InStr(start, TextObject.Text, Lookfor, vbTextCompare) <> 0) Then
   TextObject.SelStart = start
   TextObject.SelLength = Len(Lookfor)
   If (MsgBox(message & " Continue Looking?", vbYesNo, "Found") = vbNo) Then
    Exit Function
   End If
   TextObject.SelText = Replace
  End If
 Wend
End Function

'Letter Text Search
Public Function FRLetterTextSearch(TextObject As Object, Lookfor As String, Replace As String, message As String, Acuracy As Long)
 Dim word As Variant
 Dim x As Integer
 Dim start As Integer
 Dim done As Integer
 done = 1
 start = 1
 x = 1
 word = Array(x)
 While x < Len(searchfor)
  word(x) = Mid(searchfor, start, Acuracy)
  start = start + 1
  x = x + 1
 Wend
 While (Len(textobjects.Text) < done)
  While (done < x)
   If (InStr(start, TextObject.Text, word(done), vbTextCompare) <> 0) Then
    TextObject.SelStart = start
    TextObject.SelLength = Len(Lookfor)
    If (MsgBox(message & " Continue Looking?", vbYesNo, "Found") = vbNo) Then
     Exit Function
    End If
    TextObject.SetText = Replace
   End If
  Wend
 Wend
End Function

'Advanced Text Search
Public Function FRAdvTextSearch(TextObject As Object, Lookfor As String, Replace As String, message As String, Acuracy As Long)
 Dim done As Integer
 Dim start As String
 done = 1
 While (Len(TextObject.Text) < done)
  If (InStr(start, TextObject.Text, Left(Lookfor, Acuracy), vbTextCompare) <> 0) Then
   TextObject.SelStart = start
   TextObject.SelLength = Len(Lookfor)
   If (MsgBox(message & " Continue Looking?", vbYesNo, "Found") = vbNo) Then
    Exit Function
   End If
   TextObject.SelText = Replace
  End If
 Wend
End Function

'Find and Replace Table
'Letter Add Item Search
Public Function FRLetterAddItemSearch(recordset As ADODB.recordset, searchcolumn As String, searchfor As String, Replace As String, ReplaceColumn As String, Acuracy As Long, AddItemObject As Object)
 Dim word As Variant
 Dim x As Integer
 Dim start As Integer
 Dim done As Integer
 done = 1
 start = 1
 x = 1
 word = Array(x)
 While x < Len(searchfor)
  word(x) = Mid(searchfor, start, Acuracy)
  start = start + 1
  x = x + 1
 Wend
 recordset.MoveFirst
 While (recordset.EOF = False)
  While (done < x)
    If (InStr(1, recordset.Fields(searchcolumn), word(done), vbTextCompare <> 0)) Then
     recordset.Fields(ReplaceColumn) = Replace
    End If
  Wend
 Wend
End Function

'Add item Search
Public Function FRAddItemSearch(recordset As ADODB.recordset, searchcolumn As String, searchfor As String, Replace As String, ReplaceColumn As String, AddItemObject As Object)
  recordset.MoveFirst
 While (recordset.EOF = False)
  If (recordset.Fields(searchcolumn).Value = searchfor) Then
    recordset.Fields(ReplaceColumn) = Replace
   End If
   recordset.MoveNext
 Wend
 End Function
 
 'Advanced Add Item Search
 Public Function FRAdvAddItemSearch(recordset As ADODB.recordset, searchcolumn As String, searchfor As String, Replace As String, ReplaceColumn As String, Acuracy As Long, AddItemObject As Object)
 recordset.MoveFirst
 While (recordset.EOF = False)
  If (InStr(1, recordset.Fields(Field).Value, Left(searchfor, Acuracy), vbTextCompare) <> 0) Then
   recordset.Fields(ReplaceColumn) = Replace
  End If
 Wend
End Function



Private Sub Label1_Click()

End Sub
