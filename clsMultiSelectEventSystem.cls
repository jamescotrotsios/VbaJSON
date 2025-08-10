VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultiSelectEventSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Event MultiSelectUpdated(ByVal ParentFormForTextBox As String, ByVal TextBox As String, ByVal TextBoxData As Variant)
Public Event CustomMessage(ByVal SourceForm As String, ByVal MessageType As String, ByVal MessageData As Variant)
Private m_RegisteredForms As Collection
Private m_IsActive As Boolean
Public Sub SendCustomMessage(ByVal SourceForm As String, ByVal MessageType As String, Optional ByVal MessageData As Variant = Null)
    If m_IsActive Then
        RaiseEvent CustomMessage(SourceForm, MessageType, MessageData)
    End If
End Sub
Public Sub SendMultiSelectUpdated(ByVal ParentFormForTextBox As String, ByVal TextBox As String, ByVal TextBoxData As Variant)
    If m_IsActive Then
        RaiseEvent MultiSelectUpdated(ParentFormForTextBox, TextBox, TextBoxData)
    End If
End Sub
Private Sub Class_Initialize()
    Set m_RegisteredForms = New Collection
    m_IsActive = True
End Sub

' Clean up when destroyed
Private Sub Class_Terminate()
    Set m_RegisteredForms = Nothing
End Sub

' Register a form with the communicator
Public Sub RegisterForm(ByVal FormName As String)
    On Error Resume Next
    ' Try to add the form (will error if already exists, which is fine)
    m_RegisteredForms.Add FormName, FormName
    On Error GoTo 0
End Sub

' Unregister a form from the communicator
Public Sub UnregisterForm(ByVal FormName As String)
    On Error Resume Next
    m_RegisteredForms.Remove FormName
    On Error GoTo 0
End Sub

' Check if a form is registered
Public Function IsFormRegistered(ByVal FormName As String) As Boolean
    On Error Resume Next
    Dim temp As String
    temp = m_RegisteredForms(FormName)
    IsFormRegistered = (Err.Number = 0)
    On Error GoTo 0
End Function

' Get list of registered forms
Public Function GetRegisteredForms() As String
    Dim result As String
    Dim i As Integer
    
    For i = 1 To m_RegisteredForms.Count
        If result <> "" Then result = result & ", "
        result = result & m_RegisteredForms(i)
    Next i
    
    GetRegisteredForms = result
End Function

' Enable/disable the communicator
Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

Public Property Let IsActive(ByVal value As Boolean)
    m_IsActive = value
End Property
Public Sub BroadcastMessage(ByVal SourceForm As String, ByVal MessageType As String, Optional ByVal MessageData As Variant = Null)
    Dim broadcastData As String
    broadcastData = "BROADCAST|" & GetRegisteredForms()
    SendCustomMessage SourceForm, MessageType, Array(MessageData, broadcastData)
End Sub

' Method to check if a specific form is open
Public Function IsFormOpen(ByVal FormName As String) As Boolean
    On Error Resume Next
    IsFormOpen = (Forms(FormName).Name = FormName)
    On Error GoTo 0
End Function

' Method to safely close a form
Public Sub SafeCloseForm(ByVal FormName As String)
    If IsFormOpen(FormName) Then
        DoCmd.Close acForm, FormName
        UnregisterForm FormName
    End If
End Sub
