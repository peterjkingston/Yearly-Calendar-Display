Public Function ToNode(dateToConvert as Date) as Integer
    Dim nodeMonth as Integer, nodeDay as Integer, result as Integer
    nodeMonth = Month(dateToConvert)
    nodeDay = Day(dateToConvert)
    if (nodeDay/4) < 0 Then
        result = 1 + (nodeMonth * 4)
    else
        result = (nodeDay/4) + (nodeMonth * 4)
    End if
    ToNode = result
End Function