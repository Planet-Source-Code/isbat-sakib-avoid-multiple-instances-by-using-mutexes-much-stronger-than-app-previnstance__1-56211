Attribute VB_Name = "modMain"
Option Explicit

'******************************************************************************
'******This code is downloaded from this page:
'****** http://www.Planet-Source-Code.com/vb/default.asp?lngCId=56211&lngWId=1
'****** Author: Isbat Sakib
'****** Email: sakib039@hotmail.com

'Just start your app from the Sub Main. And at the end of your app, call MutexCleanUp Sub.

'Now what is a mutex? A mutex object is a synchronization object whose state is set to
'signaled when it is not owned by any thread, and nonsignaled when it is owned. Only one
'thread at a time can own a mutex object, whose name comes from the fact that it is
'useful in coordinating mutually exclusive access to a shared resource. Now, how this
'method of avoiding multiple instances works is explained in the comments below. Enjoy.

'The API functions and constantfor mutex manipulation
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const WAIT_OBJECT_0         As Long = &H0

'These APIs are only used for showing the previous instance of the app.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const SW_RESTORE = 9


'This string should be as unique as possible but GENERALLY not more than 254 characters
'actually not more than the value of MAX_PATH constant as documented in MSDN.
Private Const UniqueString          As String = "StringForAvoidingMultipleInstance"

'This variable will have the handle of the mutex
Private GMutex      As Long


Sub Main()

    Dim OldHWnd      As Long
    
    If CheckAndCreateMutex Then         'No previous instance, so load the main form.
        frmDemo.Show
    Else
                                        'A previous instance exists. Find that window as
                                        'the caption is known.
        OldHWnd = FindWindow(vbNullString, "Demo Test Form")
        
        If OldHWnd <> 0 Then            'If the window is found,then show it and set focus
                                        'to it whether it is minimized or not.
            Call ShowWindow(OldHWnd, SW_RESTORE)
            Call SetForegroundWindow(OldHWnd)
        End If
    End If
    
End Sub

Public Function CheckAndCreateMutex() As Boolean
   
    GMutex = CreateMutex(0&, 0&, UniqueString)         'First, lets create the mutex
    
    If GMutex = 0 Then                                'Error occurred for some reason.
        MsgBox "The mechanism to ensure only one instance of this app has failed for unknown reasons.", vbCritical, "Error"
        CheckAndCreateMutex = True
    Else
    
        'Now this requires some explanation. The mutex has been created, but does not
        'belong to this specific application thread. This could be done by setting the
        'second parameter of CreateMutex function to 1, but I don't know why it doesn't
        'work in VB, though the same thing works perfectly in C++. So, another work-around
        'is here. The next function will only check if the mutex is signaled or not as the
        'second parameter is given zero. The mutex will be non-signaled if a thread owns
        'it already. Now calling once this function makes the calling thread the owner of
        'the mutex if it doesn't have an owner already.
        
        If WaitForSingleObject(GMutex, 0&) = WAIT_OBJECT_0 Then     'The mutex is signaled and
                                                                    'no other thread owns it.
                                                                    'But from now on, this thread
                                                                    'will own the mutex.
            CheckAndCreateMutex = True
        
        Else                                    'Several other things might have happened.
                                                'The mutex is may be non-signal, or it already
                                                'has an owner, or the time-out for the
                                                'function return is finished.
        
            MsgBox "This application is already running", vbInformation, "App Running"
            
            Call CloseHandle(GMutex)            'The already owned mutex has been opened
                                                'by this thread. Well, now close the handle
                                                'of it.
            CheckAndCreateMutex = False
            
        End If
        
    End If
    
End Function

Public Sub MutexCleanUp()

    Call ReleaseMutex(GMutex)            'A thread should release the mutex when no longer
                                         'needed.

End Sub

