VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "AutoDropdown"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "www.dtdn.com/dev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1200
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsAC As New clsAutoComplete            'define a class instance

Private Sub Check1_Click()                  'turn autodropdown on/off
    If Check1.Value Then                    'if checked
        clsAC.AutoDropdown = True           'turn on
    Else                                    'if not checked
        clsAC.AutoDropdown = False          'turn off
    End If
End Sub

Private Sub cmdClose_Click()
    Set clsAC = Nothing                     'destroy the class
    Unload Me                               'unload the form
End Sub

Private Sub Form_Load()

    Call LoadComboBox(Combo1)
    
    Set clsAC.LinkedComboBox = Combo1       'assign the combo to the class
    
    'Combo1.AddItem "aabbcc"                 'add sample data to the combo
    'Combo1.AddItem "aaccbb"
    'Combo1.AddItem "bbaacc"
    'Combo1.AddItem "bbccaa"
    'Combo1.AddItem "ccaabb"
    'Combo1.AddItem "ccbbaa"
End Sub

Private Sub Combo1_GotFocus()
Dim strHoldComboText As String
    
    SendKeys "{BACKSPACE}"
    strHoldComboText = Combo1.Text
    Combo1.Tag = ""
      
    Call LoadComboBox(Combo1)
    Set clsAC.LinkedComboBox = Combo1
    Combo1.Text = strHoldComboText
End Sub

Private Sub Combo1_LostFocus()
Dim intAnswer As Integer
   If Not Combo1.Text = "" Then
      Set rs = db.OpenRecordset("Select * from tblCombo where ComboValue = '" & Combo1.Text & "'", dbOpenDynaset, dbSeeChanges)
      If rs.RecordCount = 0 Then
         intAnswer = MsgBox("DO YOU WANT TO ADD - " & Combo1.Text & " ?", vbYesNo + vbQuestion, "ADD VALUE TO DROPDOWN LIST?")
         If intAnswer = vbYes Then
            rs.AddNew
            rs!ComboValue = Combo1.Text
            rs.Update
            Set clsAC.LinkedComboBox = Combo1       'assign the combo to the class
            Combo1.Refresh
            rs.Close
         Else
            Combo1.Tag = ""
         End If
      End If
   End If
End Sub

Private Sub Label10_Click()
  Dim q As Variant
  q = "http://www.dtdn.com/dev"
  q = ShellExecute(0&, vbNullString, q, vbNullString, vbNullString, vbNormalFocus)
End Sub

