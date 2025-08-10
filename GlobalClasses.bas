Attribute VB_Name = "GlobalClasses"
Option Compare Database
Option Explicit
Private g_FormCommunicator As clsMultiSelectEventSystem
Public Function GetFormCommunicator() As clsMultiSelectEventSystem
    If g_FormCommunicator Is Nothing Then
        Set g_FormCommunicator = New clsMultiSelectEventSystem
    End If
    Set GetFormCommunicator = g_FormCommunicator
End Function
Public Sub CleanupFormCommunicator()
    Set g_FormCommunicator = Nothing
End Sub
Public Sub RegisterForm(ByVal FormName As String)
    GetFormCommunicator.RegisterForm FormName
End Sub
Public Sub UnregisterForm(ByVal FormName As String)
    GetFormCommunicator.UnregisterForm FormName
End Sub
Public Sub SendMessage(ByVal SourceForm As String, ByVal MessageType As String, Optional ByVal MessageData As Variant = Null)
    GetFormCommunicator.SendCustomMessage SourceForm, MessageType, MessageData
End Sub
Public Sub SendMultiSelectUpdated(ByVal ParentFormForTextBox As String, ByVal TextBox As String, ByVal TextBoxData As Variant)
    GetFormCommunicator.SendMultiSelectUpdated ParentFormForTextBox, TextBox, TextBoxData
End Sub
