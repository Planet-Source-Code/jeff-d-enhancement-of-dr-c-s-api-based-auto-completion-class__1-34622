VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************
'* Copyright © 2002 Corné van Dyk *
'* All Rights Reserved            *
'**********************************
'***************************DISCLAIMER****************************
'*****************************************************************
'* No waranty of any kind, either explicit or implied is made as *
'* to the usability or the functionality of this code.           *
'* By using this code you indemnify the author of any and all    *
'* liability related to the use, or inability to use this code.  *
'*****************************************************************

'FEEDBACK
'========
'c@dtdn.com

'DESCRIPTION
'===========
'EASY TO IMPLEMENT API Based AutoCompletion Class featuring AutoDropdown option

'SYNOPSIS
'========
'This class is used to EASILY add AutoCompletion such as in Internet
'Explorer or MS Access to VB Combo Boxes.  It makes use of API calls
'which makes it EXTREMELY efficient.  It also gives you the ability
'to control the AutoDropdown property of the combo in code.

'USAGE
'=====
'The demo project should be self explanatory, but basically all you
'have to do is add the class to your project and upon the Load event
'of your form, simply initialize a class instance for each of your
'combo box controls.
'The AutoDropdown property can easily be set in code.  In the example
'project I tied it to a check box.  Play around with it to see how it
'works.

Private WithEvents memCBO As ComboBox       'member var to hold combo
Attribute memCBO.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any _
        ) As Long                           'internal windows messaging API

Private Const CB_ERR = (-1)                 'combo error code
Private Const CB_FINDSTRING = &H14C         'hex value to invoke find
Private Const CB_FINDSTRINGEXACT = &H158    'hex value to invoke find with exact matching
Private Const CB_SHOWDROPDOWN = &H14F       'hex value to drop down combo

Private membAutoDropdown As Boolean         'hold AutoDropdown property
Private intSelStart As Integer              'var to hold combo.SelStart val
Private intSelLength As Integer             'var to hold combo.SelLength val
Private intMatchLoc As Integer              'var to hold mathing index val
Private strText As String                   'var to hold search string

Public Property Let AutoDropdown(bAutoDropdown As Boolean)
    membAutoDropdown = bAutoDropdown        'set autodropdown property
End Property

Public Property Get AutoDropdown() As Boolean
    AutoDropdown = membAutoDropdown         'return autodropdown property
End Property

Public Property Set LinkedComboBox(cboTargetComboBox As ComboBox)
    Set memCBO = cboTargetComboBox          'set the linked combo
End Property

Public Property Get LinkedComboBox() As ComboBox
    Set LinkedComboBox = memCBO             'return the linked combo
End Property

Private Sub memCBO_KeyUp(KeyCode As Integer, Shift As Integer)
    'do nothing if certain special keys is pressed
    If KeyCode = 16 Then Exit Sub           'shift
    If KeyCode = 17 Then Exit Sub           'ctrl
    If KeyCode = 18 Then Exit Sub           'alt
    If KeyCode = 27 Then Exit Sub           'escape
    If KeyCode = 33 Then Exit Sub           'page up
    If KeyCode = 34 Then Exit Sub           'page down
    If KeyCode = 35 Then Exit Sub           'end
    If KeyCode = 36 Then Exit Sub           'home
    If KeyCode = 37 Then Exit Sub           'left arrow
    If KeyCode = 39 Then Exit Sub           'right arrow
    If KeyCode = 45 Then Exit Sub           'insert
    If KeyCode = 91 Then Exit Sub           'windows key
    If KeyCode = 112 Then Exit Sub          'F1
    If KeyCode = 113 Then Exit Sub          'F2
    If KeyCode = 114 Then Exit Sub          'F3
    If KeyCode = 115 Then Exit Sub          'F4
    If KeyCode = 116 Then Exit Sub          'F5
    If KeyCode = 117 Then Exit Sub          'F6
    If KeyCode = 118 Then Exit Sub          'F7
    If KeyCode = 119 Then Exit Sub          'F8
    If KeyCode = 120 Then Exit Sub          'F9
    If KeyCode = 121 Then Exit Sub          'F10
    If KeyCode = 122 Then Exit Sub          'F11
    If KeyCode = 123 Then Exit Sub          'F12

    With memCBO
        intSelStart = .SelStart             'save combo.SelStart
        intSelLength = .SelLength           'save combo.SelLength
        strText = .Text                     'save combo.Text
        intMatchLoc = FindInList(strText, , False)  'search for match
        If intMatchLoc <> CB_ERR Then       'if match found
            ShowDropDown True               'drop down combo
            .ListIndex = intMatchLoc        'position cursor in dropped list
            .Text = .List(intMatchLoc)      'set combo text
            .SelStart = intSelStart         'set selected text
            .SelLength = Len(.Text) - intSelStart
        Else                                'no match found
            ShowDropDown False              'hide combo drop down
            .Text = strText                 'set combo text to original text
            .SelStart = intSelStart         'set selected text
            .SelLength = intSelLength
        End If
    End With
End Sub

Private Sub memCBO_LostFocus()
    memCBO.Text = memCBO.List(intMatchLoc)  'set combo to index value if TAB pressed
End Sub

Private Function FindInList(strFindThis As String, _
        Optional lngIndexStart As Long = -1, _
        Optional bExactMatchRequired As Boolean = False) _
        As Long                             'function to search for match
                                            'if found, it returns the index
                                            'else it returns CB_ERR
    Dim wMsg As Long                        'temp var to hold type of search
    'check if exact match is required & set search type accordingly
    wMsg = IIf(bExactMatchRequired, CB_FINDSTRINGEXACT, CB_FINDSTRING)
    strFindThis = strFindThis & Chr(0)
    FindInList = SendMessage(memCBO.hwnd, wMsg, lngIndexStart, _
            ByVal strFindThis)              'invoke API search of combo
End Function

Private Sub ShowDropDown(bShowDropdown As Boolean)
    If membAutoDropdown Then                'if autodropdown property is set
        'make API call to drop down combo
        SendMessage memCBO.hwnd, CB_SHOWDROPDOWN, bShowDropdown, 0
    End If
End Sub

Private Sub memCBO_Keypress(KeyAscii As Integer)
   Dim strValidInput As String
   Dim intMsg As Integer

   If KeyAscii = 13 Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
End Sub
