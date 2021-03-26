Private Type cPoint
    x As Double
    y As Double
End Type

Public Sub inicio()
    Dim thePoint As cPoint
    Dim polygon() As cPoint
    Dim aLine() As cPoint
    Dim linea_b() As cPoint
    Dim points_in_polygon() As cPoint

    linea_b = get_points("rm.txt")
    If ArrayLen(linea_b) < 2 Then
        MsgBox "Estructura de buffer no es correcta. Favor revise."
        Exit Sub
    End If

    polygon = get_points("m.txt")
    If ArrayLen(polygon) < 3 Then
        MsgBox "Mascara de la voladura no es correcta. Favor revise."
        Exit Sub
    End If
    points_in_polygon = get_points_inside_polygon(linea_b, polygon)

End Sub

Public Function get_points_inside_polygon(ByRef points() As cPoint, ByRef polygon() As cPoint) As cPoint()
    Dim thePoint As cPoint
    Dim result As Boolean
    Dim pointsInPolygon() As cPoint
    Dim j As Integer
   
    j = 0
    For i = LBound(points) To UBound(points)
        thePoint = points(i)
        result = inside_polygon(thePoint, polygon)
        
        If result = True Then
            ReDim Preserve pointsInPolygon(j)
            pointsInPolygon(j) = thePoint
            j = j + 1
        End If
    Next i

    get_points_inside_polygon = pointsInPolygon
End Function

Public Function get_line_intersection(p0 As cPoint, p1 As cPoint, p2 As cPoint, p3 As cPoint) As cPoint
    Dim s1_x As Double, s1_y As Double, s2_x As Double, s2_y As Double
    Dim s As Double, t As Double
    Dim intersectionPoint As cPoint

    s1_x = p1.x - p0.x
    s1_y = p1.y - p0.y
    s2_x = p3.x - p2.x
    s2_y = p3.y - p2.y

    s = (-s1_y * (p0.x - p2.x) + s1_x * (p0.y - p2.y)) / (-s2_x * s1_y + s1_x * s2_y)
    t = (s2_x * (p0.y - p2.y) - s2_y * (p0.x - p2.x)) / (-s2_x * s1_y + s1_x * s2_y)

    If s >= 0 And s <= 1 And t >= 0 And t <= 1 Then
        'Collision detected
        intersectionPoint.x = p0.x + (t * s1_x)
        intersectionPoint.y = p0.y + (t * s1_y)
        get_line_intersection = intersectionPoint
    End If

End Function

Public Function get_points(filename As String) As cPoint()
    Dim fileContent As String
    Dim polygon() As cPoint
    Dim p As cPoint

    fileContent = LoadFileStr(ThisWorkbook.Path & "/" & filename)

    Set objMatches = apply_regex("\s*(\d{7}\.\d*)\s+(\d{7}\.\d*)", fileContent)
    If objMatches.Count = 0 Then GoTo notFound

    ReDim polygon(0 To objMatches.Count - 1) As cPoint

    i = 0
    For Each aMatch In objMatches
        p.x = aMatch.SubMatches.Item(0)
        p.y = aMatch.SubMatches.Item(1)
        polygon(i) = p
        i = i + 1
    Next

    get_points = polygon
    Exit Function

notFound:
    MsgBox "Not Matches Found!!!!"
End Function

Public Function apply_regex(strPattern As String, strContent As String) As Object
    Dim regEx As New regExp

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = strPattern
    End With

    Set objMatches = regEx.Execute(strContent)
    Set apply_regex = objMatches
End Function

Public Function inside_polygon(p As cPoint, ByRef polygon() As cPoint) As Boolean
    'ray-casting algorithm based on
    'https://wrf.ecse.rpi.edu/Research/Short_Notes/pnpoly.html/pnpoly.html

    Dim x As Double, y As Double
    Dim inside As Boolean, lineIntersect As Boolean, i As Integer, j As Integer

    x = p.x
    y = p.y

    inside = False

    j = UBound(polygon) - LBound(polygon)
    For i = LBound(polygon) To UBound(polygon)
        xi = polygon(i).x
        yi = polygon(i).y
        xj = polygon(j).x
        yj = polygon(j).y

        If yj - yi = 0 Then
            deno = 0
        Else
            deno = (xj - xi) * (y - yi) / (yj - yi)
        End If

        lineIntersect = ((yi > y) <> (yj > y)) And (x < deno + xi)

        If lineIntersect Then
            inside = Not inside
        End If

        j = i
    Next i

    inside_polygon = inside
End Function

Public Function LoadFileStr(FN As String) As String
    With CreateObject("Scripting.FileSystemObject")
        LoadFileStr = .OpenTextFile(FN, 1).ReadAll
    End With
End Function

Public Function ArrayLen(arr() As cPoint) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function