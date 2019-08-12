Dim CalEvent as new CCalEvent

Private m_startDate as Date, m_endDate as Date, m_subject as String

Public Sub Constructor(start as Date, end as Date, subject as String)
    m_startDate = start
    m_endDate = end
    m_subject = subject
End Sub

Public Property Get StartDate() as Date
    StartDate = m_startDate
End Sub

Public Property Get EndDate() as Date
    EndDate = m_endDate
End Sub

Public Property Get Subject() as String
    Subject = m_subject
End Sub