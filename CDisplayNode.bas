Dim DisplayNode as new CDisplayNode

Private m_location As Integer, m_calEvent As String, m_startEvent as Integer, m_endEvent as Integer, m_isScheduled as Boolean

Public Sub Constructor(index as Integer)
    m_location = Integer
    m_isScheduled = false
    m_calEvent = ""
End Sub

Public Sub AssignValue(calendarEvent as String, start as Integer, end as Integer)
    m_calEvent = calendarEvent
    m_startEvent = start
    m_endEvent = end
    m_isScheduled = true
End Sub

Public Property Get IsScheduled() as Boolean
    IsScheduled = m_isScheduled
End Property

Public Property Get EventStart() as Integer
    EventStart = m_startEvent
End Property

Public Property Get EventEnd() as Integer
    EventEnd = m_endEvent
End Property