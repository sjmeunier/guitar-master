Attribute VB_Name = "mdlEdit"
Option Explicit

Public Sub CopyIt()
   If TypeOf Screen.ActiveControl Is TextBox Then
      Clipboard.SetText Screen.ActiveControl.SelText
  ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
      Clipboard.SetText Screen.ActiveControl.Text
   ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
      Clipboard.SetData Screen.ActiveControl.Picture
   ElseIf TypeOf Screen.ActiveControl Is ListBox Then
      Clipboard.SetText Screen.ActiveControl.Text
   Else
   End If
End Sub

Public Sub CutIt()
   CopyIt
   If TypeOf Screen.ActiveControl Is TextBox Then
      Screen.ActiveControl.SelText = ""
   ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
      Screen.ActiveControl.Text = ""
   ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
      Screen.ActiveControl.Picture = LoadPicture()
   ElseIf TypeOf Screen.ActiveControl Is ListBox Then
      If Screen.ActiveControl.ListIndex >= 0 Then
         Screen.ActiveControl.RemoveItem Screen.ActiveControl.ListIndex
      End If
   Else
   End If
End Sub

Public Sub PasteIt()
   If TypeOf Screen.ActiveControl Is TextBox Then
      Screen.ActiveControl.SelText = Clipboard.GetText()
   ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
      Screen.ActiveControl.Text = Clipboard.GetText()
   ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
      Screen.ActiveControl.Picture = Clipboard.GetData()
   ElseIf TypeOf Screen.ActiveControl Is ListBox Then
      Screen.ActiveControl.AddItem Clipboard.GetText()
   Else
   End If
End Sub

Public Sub DeleteIt()
   If TypeOf Screen.ActiveControl Is TextBox Then
      Screen.ActiveControl.SelText = ""
  ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
      Screen.ActiveControl.Text = ""
   ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
      Screen.ActiveControl.Picture = LoadPicture()
   ElseIf TypeOf Screen.ActiveControl Is ListBox Then
      If Screen.ActiveControl.ListIndex >= 0 Then
         Screen.ActiveControl.RemoveItem Screen.ActiveControl.ListIndex
      End If
   Else
   End If
End Sub
