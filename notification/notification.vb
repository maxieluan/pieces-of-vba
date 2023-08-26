Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Sub ShowNotification()
    Dim result As Long
    result = MessageBox(0, "This is your custom notification message.", "Notification", vbInformation)
    ' You can customize the message and options as needed
End Sub