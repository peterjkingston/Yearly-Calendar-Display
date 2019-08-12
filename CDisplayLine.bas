Dim DisplayLine as new DisplayLine

Private m_idRow as Integer, m_displayNodes as Collection

Public Property Get DisplayNode(index as Integer) as CDisplayNode
    set DisplayNode = m_displayNodes(index)
End Property

Public Sub Constructor(displayRow as Integer)
    set m_displayNodes = new Collection

    for i = 0 to 47
        set displayNode = new CDisplayNode
        displayNode.Constructor i
        set m_displayNodes.Add displayNode 
    next i
End Sub

Public Function TrySchedule(subject as String, start as Date, end as Date) As Boolean
    dim result as Boolean, i as Integer, j as Integer, convertedStart as Integer, convertedEnd as Integer
    '// add declarations
    On Error GoTo catchError
    convertedStart = MDateNodeConverter(start)
    convertedEnd = MDateNodeConverter(end)

    for i = convertedStart to convetertedEnd
        if m_displayNodes(i).IsScheduled then
            result = false
            exit for
        end if
        if i = convertedEnd then
            for j = convertedStart to convertedEnd
                m_displayNodes(j).Schedule(subject, convertedStart, convertedEnd)
            next j
            result = true
        end if
    next i
    
exitFunction:
    TrySchedule = result
    Exit Function
catchError:
    '// add error handling
    GoTo exitFunction
End Function