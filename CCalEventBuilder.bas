Dim CalEventBuilder as new CCalEventBuilder

Private m_worksheet as Excel.Worksheet, m_calName as String
Private const M_COLUMN_STARTDATE as Integer = 0
Private const M_COLUMN_ENDDATE as Integer = 0
Private const M_COLUMN_SUBJECT as Integer = 0
Private const M_COLUMN_CALNAME as Integer = 0

Public Event CalEventBuilt(e as CCalEvent)

Public Sub Constructor(dataSourceWorksheet as Excel.Worksheet, olCalendarName as String)
    set m_worksheet = ws
    m_calName = olCalendarName
End Sub

Public Sub Refresh()
End Sub

Public Function BuildEvents(startRowData as Integer) as Collection
    Dim result as Collection, calEvent as CCalEvent, i as Integer, xlCell as Range
    set result = new Collection
    i = startRowData

    set xlCell = m_Worksheet.Cells(i,M_COLUMN_CALNAME)
    Do until xlCell.Value = ""
        if(xlCell.Value) = m_calName then
            set calEvent = new CCalEvent
            with m_worksheet
                calEvent.Constructor(.Cells(i, M_COLUMN_STARTDATE), _
                                     .Cells(i, M_COLUMN_ENDDATE) _
                                     .Cells(i, M_COLUMN_SUBJECT))
            end with
            RaiseEvent CalEventBuilt(calEvent)
            result.Add calEvent
        end if
        i = i + 1
    Loop

    set BuildEvents = result
End Sub