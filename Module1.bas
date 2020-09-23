Attribute VB_Name = "Module1"
Option Compare Text
Option Explicit

#Const DebugMode = 1 ' whether or not to show debug information
'@===========================================================================
' SaveFormState:
'  Saves the state of controls to a file
'
'  Currently Supports: TextBox, CheckBox, OptionButton, Listbox, ComboBox
'=============================================================================
Sub SaveFormState(ByVal SourceForm As Form)
 Dim A As Long ' general purpose
 Dim B As Long
 Dim C As Long
 Dim FileName As String ' where to save to
 Dim FHandle As Long ' FileHandle
 ' error handling code
 On Error GoTo fError
 ' we create a filename based on the formname
 FileName = App.Path + "\" + SourceForm.Name + ".set"
 ' Get a filehandle
 FHandle = FreeFile()
 ' open the file
 #If DebugMode = 1 Then
  Debug.Print "--------------------------------------------------------->"
  Debug.Print "Saving Form State:" + SourceForm.Name
  Debug.Print "FileName=" + FileName
 #End If
 Open FileName For Output As FHandle
 ' loop through all controls
 ' first we save the type then the name
 For A = 0 To SourceForm.Controls.Count - 1
  #If DebugMode = 1 Then
   Debug.Print "Saving control:" + SourceForm.Controls(A).Name
  #End If
  ' if its textbox we save the .text property
  If TypeOf SourceForm.Controls(A) Is TextBox Then
   Print #FHandle, "TextBox"
   Print #FHandle, SourceForm.Controls(A).Name
   Print #FHandle, "StartText"
   Print #FHandle, SourceForm.Controls(A).Text
   Print #FHandle, "EndText"
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a checkbox we save the .value property
  If TypeOf SourceForm.Controls(A) Is CheckBox Then
   Print #FHandle, "CheckBox"
   Print #FHandle, SourceForm.Controls(A).Name
   Print #FHandle, Str(SourceForm.Controls(A).Value)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a option button we save its value
  If TypeOf SourceForm.Controls(A) Is OptionButton Then
   Print #FHandle, "OptionButton"
   Print #FHandle, SourceForm.Controls(A).Name
   Print #FHandle, Str(SourceForm.Controls(A).Value)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a listbox we save the .text and list contents
  If TypeOf SourceForm.Controls(A) Is ListBox Then
   Print #FHandle, "ListBox"
   Print #FHandle, SourceForm.Controls(A).Name
   Print #FHandle, SourceForm.Controls(A).Text
   Print #FHandle, "StartList"
   For B = 0 To SourceForm.Controls(A).ListCount - 1
    Print #FHandle, SourceForm.Controls(A).List(B)
   Next B
   Print #FHandle, "EndList"
   ' save listindex
   Print #FHandle, CStr(SourceForm.Controls(A).ListIndex)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a combobox, save .text and list items
  If TypeOf SourceForm.Controls(A) Is ComboBox Then
   Print #FHandle, "ComboBox"
   Print #FHandle, SourceForm.Controls(A).Name
   Print #FHandle, SourceForm.Controls(A).Text
   Print #FHandle, "StartList"
   For B = 0 To SourceForm.Controls(A).ListCount - 1
    Print #FHandle, SourceForm.Controls(A).List(B)
   Next B
   Print #FHandle, "EndList"
    ' print a separator
   Print #FHandle, "|<->|"
  End If
 Next A
' close file
 #If DebugMode = 1 Then
  Debug.Print "Closing File."
  Debug.Print "<----------------------------------------------------------"
 #End If
 Close #FHandle
 ' stop error handler
 On Error GoTo 0
 Exit Sub
fError: ' Simple error handler
 C = MsgBox("Error in SaveFormState. " + Err.Description + ", Number=" + CStr(Err.Number), vbAbortRetryIgnore)
 If C = vbIgnore Then Resume Next
 If C = vbRetry Then Resume
 ' else abort
End Sub
'@===========================================================================
' LoadFormState:
'  Loads the state of controls from file
'
'  Currently Supports: TextBox, CheckBox, OptionButton, Listbox, ComboBox
'=============================================================================
Sub LoadFormState(ByVal SourceForm As Form)
 Dim A As Long ' general purpose
 Dim B As Long
 Dim C As Long
 
 Dim TXT As String ' general purpose
 Dim fData As String ' used to hold File Data
' these are variables used for controls data
 Dim cType As String ' Type of control
 Dim cName As String ' Name of control
 Dim cNum As Integer ' number of control
' vars for the file
 Dim FileName As String ' where to save to
 Dim FHandle As Long ' FileHandle
 ' error handling code
 'On Error GoTo fError
 ' we create a filename based on the formname
 FileName = App.Path + "\" + SourceForm.Name + ".set"
 ' abort if file does not exist
 If Dir(FileName) = "" Then
  #If DebugMode = 1 Then
   Debug.Print "File Not found:" + FileName
  #End If
  Exit Sub
 End If
 ' Get a filehandle
 FHandle = FreeFile()
 ' open the file
 #If DebugMode = 1 Then
  Debug.Print "------------------------------------------------------>"
  Debug.Print "Loading FormState:" + SourceForm.Name
  Debug.Print "FileName:" + FileName
 #End If
 Open FileName For Input As FHandle
' go through file
 While Not EOF(FHandle)
  Line Input #FHandle, cType
  Line Input #FHandle, cName
  ' Get control number
  cNum = -1
  For A = 0 To SourceForm.Controls.Count - 1
   If SourceForm.Controls(A).Name = cName Then cNum = A
  Next A
  ' add some debug info if in debugmode
  #If DebugMode = 1 Then
   Debug.Print "Control Type=" + cType
   Debug.Print "Control Name=" + cName
   Debug.Print "Control Number=" + CStr(cNum)
  #End If
  ' if we find control
  If Not cNum = -1 Then
   ' Depending on type of control, what data we get
   Select Case cType
   Case "TextBox"
    Line Input #FHandle, fData
    fData = "": TXT = ""
    While Not fData = "EndText"
     If Not TXT = "" Then TXT = TXT + vbCrLf
     TXT = TXT + fData
     Line Input #FHandle, fData
    Wend
    ' update control
    SourceForm.Controls(cNum).Text = TXT
   Case "CheckBox"
    ' we get the value
    Line Input #FHandle, fData
    ' update control
    SourceForm.Controls(cNum).Value = fData
   Case "OptionButton"
    ' we get the value
    Line Input #FHandle, fData
    ' update control
    SourceForm.Controls(cNum).Value = fData
   Case "ListBox"
    ' clear listbox
    SourceForm.Controls(cNum).Clear
    ' get .text property
    Line Input #FHandle, fData
    SourceForm.Controls(cNum).Text = fData
    ' read past /startlist
    Line Input #FHandle, fData
    fData = "": TXT = ""
    ' Get List
    While Not fData = "EndList"
     If Not fData = "" Then SourceForm.Controls(cNum).AddItem fData
     Line Input #FHandle, fData
    Wend
    ' get listindex
     Line Input #FHandle, fData
     SourceForm.Controls(cNum).ListIndex = Val(fData)
   Case "ComboBox"
    ' Clear combobox
    SourceForm.Controls(cNum).Clear
    ' Get Text
    Line Input #FHandle, fData
    SourceForm.Controls(cNum).Text = fData
    ' readpast /startlist
    Line Input #FHandle, fData
    fData = "": TXT = ""
    ' get list
    While Not fData = "EndList"
     If Not fData = "" Then SourceForm.Controls(cNum).AddItem fData
     Line Input #FHandle, fData
    Wend
   End Select ' what type of control
  End If ' if we found control
  ' read till seperator
  fData = ""
  While Not fData = "|<->|"
   Line Input #FHandle, fData
  Wend
 Wend ' not end of File (EOF)
' close file
 #If DebugMode = 1 Then
  Debug.Print "Closing file.."
  Debug.Print "<------------------------------------------------------"
 #End If
 Close #FHandle
 Exit Sub
fError: ' Simple error handler
 C = MsgBox("Error in LoadFormState. " + Err.Description + ", Number=" + CStr(Err.Number), vbAbortRetryIgnore)
 If C = vbIgnore Then Resume Next
 If C = vbRetry Then Resume
 ' else abort
End Sub

