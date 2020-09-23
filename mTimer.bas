Attribute VB_Name = "mTimer"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Private Type SubscriberData
    Obj As Long
    Tick As SYSTEMTIME
    Tag As Long
    Inc As Long
End Type
Const SDLen = 28

Private Data() As SubscriberData
Private miTimer As Long
Private mtSysTime As SYSTEMTIME

Public Function SetATimer(ByVal ForMe As iTimer, ByVal Interval As Long, Optional ByVal Tag As Long) As Boolean
    If Interval < 1 Then Err.Raise 5
    'Debug.Print "Set"
    Dim i As Long
    
    On Error Resume Next
    
    i = UBound(Data)
    
    If Err.Number = 0 Then
        'if array is already dimensioned
        i = i + 1
        ReDim Preserve Data(0 To i)
    Else
        'if array is not already dimensioned
        ReDim Data(0 To 0)
    End If
    
    With Data(i)
        'store the info about the subscriber
        .Inc = Interval
        
        GetSystemTime mtSysTime
        LSet .Tick = mtSysTime
        SysTimeAdd .Tick, Interval
        'The resulting time in the .tick member will be the time that we
        'want to fire the next timer event to this subscriber
        
        .Tag = Tag
        'illegal object reference so that the class does not have to have
        'a dispose-type method that it's clients are required to call
        .Obj = ObjPtr(ForMe)
    End With
    
    KillIt
    SetNextTimer
    
    SetATimer = miTimer <> 0
End Function

Public Function KillTimers(ByVal ForMe As iTimer, Optional piTag) As Long
    Dim i As Long, j As Long, b As Boolean
    On Error Resume Next
    'if pitag is supplied, only the timers with a matching tag are
    'removed, otherwise all timers for this subscriber are removed.
    b = Not IsMissing(piTag)
    j = piTag
    On Error GoTo ending
    Do While i <= UBound(Data)
        If ObjectFromPtr(Data(i).Obj) Is ForMe Then
            If b Then
                If Not Data(i).Tag = j Then GoTo iterateloop
            End If
            RemoveSubscriber i
            KillTimers = KillTimers + 1
        Else
iterateloop:
            i = i + 1
        End If
    Loop
ending:
    If Err.Number = 9 Then KillIt
End Function

Private Sub RemoveSubscriber(ByVal Index As Long)
    'Debug.Print "Remove " & Index
    Dim liTemp As Long, liLen As Long
    Dim liUbound As Long
    Dim lyTemp() As Byte
    
    On Error Resume Next
    liUbound = UBound(Data)
    liTemp = liUbound - Index
    
    If Err.Number = 0 Then
        On Error GoTo 0

        If liTemp > 0 Then
            'If we're deleting a subscriber in the middle of the array,
            'we must first move all of the subscribers that are above it
            'down one notch, then redim the array
            'You could loop through each element using LSet to copy it
            'to the previous one, but copymemory is much faster, especially
            'if there are numerous subscribers.
            'Plus copymemory is already declared to get illegal object references.
            liLen = liTemp * SDLen
            ReDim lyTemp(0 To liLen - 1)
            CopyMemory lyTemp(0), Data(Index + 1), liLen
            CopyMemory Data(Index), lyTemp(0), liLen
            ZeroMemory Data(liUbound), SDLen
            ReDim Preserve Data(0 To liUbound - 1)
        Else
            'This is if we're deleting the very last subscriber.
            'very easy, no shuffling around necessary.
            If UBound(Data) = 0 And Index = 0 Then Erase Data Else ReDim Preserve Data(0 To liTemp - 1)
        End If

    End If
End Sub

Private Function NextSubscriber() As Long
    On Error GoTo ending
    Dim i As Long
    'This function returns the index of the array whose next timer
    'tick is the least in chronological order.
    For i = 0 To UBound(Data)
        If i <> NextSubscriber Then
            If SysTimeComp(Data(NextSubscriber).Tick, Data(i).Tick) < 0 Then NextSubscriber = i
        End If
    Next

ending:
    If Err.Number <> 0 Then NextSubscriber = -1
    'Debug.Print "Next: " & NextSubscriber
End Function

Private Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    On Error Resume Next
    
    KillIt
    
    Dim liTemp As Long
    Dim liTag As Long
    Dim loFire As iTimer
    
TryNextOne:
    'Who's first in line?
    liTemp = NextSubscriber
    
    If liTemp >= 0 Then
        GetSystemTime mtSysTime
        
        With Data(liTemp)
            Just making sure that it's the right time to fire
            liTemp = SysTimeDiff(.Tick, mtSysTime)
            'Debug.Print vbTab & "Diff: " & liTemp
            If liTemp > -20 Then
                'Debug.Assert liTemp < 20 'Why so late?
                Set loFire = ObjectFromPtr(.Obj)
                liTag = .Tag
                'store the next time to fire for this subscriber
                SysTimeAdd .Tick, .Inc
            Else
                'Debug.Assert False 'Why so early?
            End If
        End With
    End If
    
    'Important..  do not call the fire event from inside the With block.
    'Oftentimes the subscriber will want to kill the timer during the fire
    'event, but since the array is locked for a with block, this will fail
    'and produce unpredictable(Worst-case: GPF) results.
    If Not loFire Is Nothing Then
        loFire.Fire liTag
        Set loFire = Nothing
        GoTo TryNextOne
    End If
    
    SetNextTimer
    
End Sub

Private Sub SetNextTimer()
    'Important! GetSystemTime mtSysTime must be called before calling this sub
    
    Dim liTemp As Long
    liTemp = NextSubscriber
    
    If liTemp >= 0 Then
        'Figure out what time the next subscriber want to be
        'notified, and how long we need to wait to accomplish that.
        'Then set the timer for the right time.
        liTemp = SysTimeDiff(mtSysTime, Data(liTemp).Tick)
        If liTemp < 1 Then liTemp = 1
        miTimer = SetTimer(0, 0, liTemp, AddressOf TimerProc)
        'Debug.Print "Set " & miTimer
    End If

End Sub

Private Function KillIt()
    If miTimer <> 0 Then
        'Debug.Print "Kill " & miTimer
        KillTimer 0, miTimer
        miTimer = 0
    End If
End Function

Private Function SysTimeComp(SysTime1 As SYSTEMTIME, SysTime2 As SYSTEMTIME) As Long
    'This could be done by testing the SysTimeDiff compared to 0, but this is faster
    'when only a relational comparison needs to be done
    With SysTime2
        Select Case True
            
            Case .wYear > SysTime1.wYear
                SysTimeComp = 1
            Case .wYear < SysTime1.wYear
                SysTimeComp = -1
            
            Case .wMonth > SysTime1.wMonth
                SysTimeComp = 1
            Case .wMonth > SysTime1.wMonth
                SysTimeComp = -1
            
            Case .wDay > SysTime1.wDay
                SysTimeComp = 1
            Case .wDay < SysTime1.wDay
                SysTimeComp = -1
                
            Case .wHour > SysTime1.wHour
                SysTimeComp = 1
            Case .wHour < SysTime1.wHour
                SysTimeComp = -1
                
            Case .wMinute > SysTime1.wMinute
                SysTimeComp = 1
            Case .wMinute < SysTime1.wMinute
                SysTimeComp = -1
                
            Case .wSecond > SysTime1.wSecond
                SysTimeComp = 1
            Case .wSecond < SysTime1.wSecond
                SysTimeComp = -1
                
            Case .wMilliseconds > SysTime1.wMilliseconds
                SysTimeComp = 1
            Case .wMilliseconds < SysTime1.wMilliseconds
                SysTimeComp = -1
                
        End Select
    End With
End Function

Private Function SysTimeDiff(SysTime1 As SYSTEMTIME, SysTime2 As SYSTEMTIME) As Long
    'How many milliseconds between these two times?
    Dim Date1 As Date, Date2 As Date
    With SysTime1
        Date1 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
    With SysTime2
        Date2 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        SysTimeDiff = DateDiff("s", Date1, Date2) * 1000 + .wMilliseconds - SysTime1.wMilliseconds
    End With
End Function

Private Sub SysTimeAdd(SysTime As SYSTEMTIME, ByVal Millisecs As Long)
    'Why does SYSTEMTIME have to be such a pain in the A$$?
    Dim Date1 As Date
    Dim liTemp As Long
    Dim liSign As Long
    With SysTime
        liSign = Sgn(Millisecs)
        Select Case liSign
            Case 1
                liTemp = 1000 - .wMilliseconds
            Case -1
                liTemp = -.wMilliseconds
        End Select
        
        If Abs(Millisecs) < Abs(liTemp) Then
            .wMilliseconds = .wMilliseconds + Millisecs
        Else
            Millisecs = Millisecs - liTemp
            liTemp = Millisecs \ 1000
            Millisecs = Millisecs - liTemp * 1000
            Date1 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
            If liSign = 1 Then liTemp = liTemp + 1
            Date1 = DateAdd("s", liTemp, Date1)
            
            
            .wYear = Year(Date1)
            .wMonth = Month(Date1)
            .wDay = Day(Date1)
            .wHour = Hour(Date1)
            .wMinute = Minute(Date1)
            .wSecond = Second(Date1)
            .wMilliseconds = Millisecs
        End If
            
    End With
End Sub

'Even object references are more fun when they're illegal
Private Property Get ObjectFromPtr(ByVal lPtr As Long) As iTimer
   Dim loTemp As iTimer
   CopyMemory loTemp, lPtr, 4
   Set ObjectFromPtr = loTemp
   CopyMemory loTemp, 0&, 4
End Property

'
'Private Sub ShowDebug(s As SYSTEMTIME)
'    Dim Date1 As Date
'    With s
'        Date1 = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
'    End With
'    Debug.Print vbTab & "Time: " & Date1
'End Sub

'This sub tests the Systime Add, Comp, Diff functions I wrote
'Public Sub test()
'    Dim t1 As SYSTEMTIME, t2 As SYSTEMTIME
'    With t1
'        .wYear = 2003
'        .wMonth = 1
'        .wDay = 8
'        .wHour = 12
'        .wMinute = 56
'        .wSecond = 55
'        .wMilliseconds = 754
'        GetSystemTime t1
'    End With
'    LSet t2 = t1
'    SysTimeAdd t2, 1001
'    Debug.Assert SysTimeDiff(t1, t2) = 1001
'    Debug.Assert SysTimeComp(t1, t2) = 1
'    SysTimeAdd t2, -2002
'    Debug.Assert SysTimeComp(t1, t2) = -1
'    Debug.Assert SysTimeDiff(t1, t2) = -1001
'End Sub

'Public Sub test()
'    Dim Testing As cTest
'
'    Set Testing = New cTest
'    Testing.SetTimer 10000, 10
'    Testing.SetTimer 7000, 7
'    Testing.SetTimer 15000, 15
'    Do
'        DoEvents
'    Loop
'    Testing.KillTimer 7
'    Set Testing = Nothing
'End Sub
'
'This was a shortcut I was using during development to avoid random timers firing.
'You don't want to use it normally because it is the subscriber's responsibility
'to decide when it is done with the timers, and the timer events should not just
'stop inexplicibly.
'But if you're modifying the code you may find it useful
'Public Sub KillMyTimer()
'    KillIt
'    Erase Data
'End Sub
