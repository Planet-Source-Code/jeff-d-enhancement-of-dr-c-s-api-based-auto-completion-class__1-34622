Attribute VB_Name = "modPrimary"
Global db As Database
Global rs As Recordset

Sub Main()
Dim SQL As String

   Set db = OpenDatabase(App.Path & "\SampleDB.mdb")
 '  SQL = "Select * from tblCombo"
 '  Set rs = db.OpenRecordset(SQL, dbOpenDynaset, dbSeeChanges)
   Form1.Show
End Sub


Public Sub LoadComboBox(ctlList As Control)
   Dim oRS As Recordset
   Dim strSQL As String
   
   'Clear the Combo Box
   ctlList.Clear
   
   'Build SQL Statement and create recordset
   strSQL = "Select * from tblCombo order by ComboValue"
   
   Set oRS = db.OpenRecordset(strSQL, dbOpenSnapshot)
  ' Set oRS = db.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
   
   Do Until oRS.EOF
      With ctlList
         .AddItem oRS!ComboValue & ""
      End With
   '   MsgBox OpenReportRS!ReportPathFilename
      oRS.MoveNext
   Loop
End Sub
Function ListFindItem(lstCtrl As Control, lngSearch As Long) As Integer

'* Description  : This routine is used to find a long integer in the
'* ItemData() property of a list or combo box.  Returns the index
'* of where the item is located.
'*
'* Example: lstNames.ListIndex = ListFindItem(lstNames, 20)

   'just returns the position, does not set it
   'used to see if item is in list
   Dim intLen As Integer
   Dim intLoop As Integer
   Dim intPos As Integer

   intLen = lstCtrl.ListCount - 1
   intPos = -1
   For intLoop = 0 To intLen
      If lstCtrl.ItemData(intLoop) = lngSearch Then
         intPos = intLoop
         Exit For
      End If
   Next intLoop
   ListFindItem = intPos
End Function

