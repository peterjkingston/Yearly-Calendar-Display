Dim DisplayCalendar as new CDisplayCalendar

Private m_displayLines as Collection
Private WithEvents m_calEventBuilder as CCalEventBuilder

Public Sub Constructor()
    '// add declarations
    On Error GoTo catchError
    Set m_displayLines = new Collection
    m_displayLines.Add new CDisplayLine
exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub

Public Sub Schedule(subject as String, start as Integer, end as Integer)
    Dim i as Integer

    Do Until DisplayLine(i).TrySchedule(subject,start,end)
        if LineCount = i Then
            m_displayLines.Add new CDisplayLine
            m_displayLines(m_displayLines.Count).Constructor
        end if 
        i = i + 1
    Loop

End Sub

Public Sub Listen(listenObj as Object)
    if typeof listenObj is CCalEventBuilder then: set m_calEventBuilder = listenObj
End Sub

Public Property Get DisplayLine(indexBaseZero as Integer) as CDisplayLine
    set DisplayLine = m_displayLines(indexBaseZero + 1)
End Property

Public Property Get LineCount() as Integer
    LineCount = m_displayLines.Count
End Property