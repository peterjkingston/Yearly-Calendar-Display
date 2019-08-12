Dim ProgramMain as new CProgramMain

Private const DATASOURCE_WORKSHEET_NAME = ""

Public Sub Run()
    Dim calEventBuilder as CCalEventBuilder, calDisplayCalendar as CCalDisplayCalendar, names as Variant

    set calEventBuilder = new CCalEventBuilder
    names = Array("\Calendar\KLT HR Events", _
                 "")
    
    set calDisplayCalendar = new CCalDisplayCalendar
    calDisplayCalendar.Constuctor
    
    for n = 0 to Ubound(names)
        calEventBuilder.Constructor(DATASOURCE_WORKSHEET_NAME,name(n))

    next n
End Sub