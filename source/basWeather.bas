Attribute VB_Name = "basWeather"
'------------------------------------------------------------------------
' Description  : getting results vom similarity calculation
'------------------------------------------------------------------------
'
'Declarations

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : create report with temperature forcast for the next 5 days
' Parameter     :
'-------------------------------------------------------------
Public Sub runWeatherReportQuery()
    
    Dim colLocation As Collection
    Dim colLocationList As New Collection
    Dim colTemperatures As Collection
    Dim colForecastData As New Collection
    Dim domXmlResponse As DOMDocument60
                     
    On Error GoTo error_handler
    'setup locations
    Set colLocation = New Collection
    colLocation.Add "New York", "city"
    colLocation.Add "40.71", "lat"
    colLocation.Add "-74.00", "lon"
    colLocationList.Add colLocation
    Set colLocation = New Collection
    colLocation.Add "Los Angeles", "city"
    colLocation.Add "34.05", "lat"
    colLocation.Add "-118.25", "lon"
    colLocationList.Add colLocation
    Set colLocation = Nothing
    
    'qery data for each location
    For Each colLocation In colLocationList
        basSystem.log "request data for city " & colLocation("city")
        Set domXmlResponse = basWeather.getWeatherData(colLocation("lat"), colLocation("lon"))
        Set colTemperatures = basWeather.extractTemperaturesFromXml(domXmlResponse)
        colTemperatures.Add colLocation("city"), "city"
        colForecastData.Add colTemperatures
    Next
    basWeather.createReport colForecastData
    Exit Sub
                     
error_handler:
    basSystem.log_error "basWeather.createWeatherReport"
End Sub
'-------------------------------------------------------------
' Description   : extract temperatures from xml server response
' Parameter     : pdomXmlResponse   -
'-------------------------------------------------------------
Public Function extractTemperaturesFromXml(pdomXmlResponse As DOMDocument60) As Collection
    
    Dim colCityResult As New Collection
    Dim colTemperatures As New Collection
    Dim nodTemperatureList As IXMLDOMNodeList
    Dim nodTemperature As IXMLDOMNode
    Dim intNodeCount As Integer
                     
    On Error GoTo error_handler
    'get max. temperatures
    basSystem.log "read max temperatures"
    Set nodTemperatureList = pdomXmlResponse.SelectNodes("//*/temperature[@type='maximum']/value")
    For intNodeCount = 0 To nodTemperatureList.Length - 1
        Set nodTemperature = nodTemperatureList.Item(intNodeCount)
        colTemperatures.Add CLng(nodTemperature.nodeTypedValue), CStr(intNodeCount)
    Next
    colCityResult.Add colTemperatures, "maxTemperatures"
    Set colTemperatures = Nothing
    Set colTemperatures = New Collection
    'get min temperatures
    basSystem.log "read min temperatures"
    Set nodTemperatureList = pdomXmlResponse.SelectNodes("//*/temperature[@type='minimum']/value")
    For intNodeCount = 0 To nodTemperatureList.Length - 1
        Set nodTemperature = nodTemperatureList.Item(intNodeCount)
        colTemperatures.Add CLng(nodTemperature.nodeTypedValue), CStr(intNodeCount)
    Next
    colCityResult.Add colTemperatures, "minTemperatures"
    Set extractTemperaturesFromXml = colCityResult
    Exit Function
                     
error_handler:
    basSystem.log_error "basWeather.extractTemperaturesFromXml"
End Function
'-------------------------------------------------------------
' Description   : requests temperature forecast for a given location (US only)
' Parameter     : plngLatitude     -
'                 plngLongitude
' Returnvalue   : xml DOM object
'-------------------------------------------------------------
Public Function getWeatherData(ByVal plngLatitude As String, ByVal plngLongitude As String) As DOMDocument60
    
    Dim objServerXml As New MSXML2.ServerXMLHTTP60
    Dim strRestRequest As String
                     
    On Error GoTo error_handler
    basSystem.log "call weather api", cLogInfo
    
    strRestRequest = "http://graphical.weather.gov/xml/sample_products/browser_interface/ndfdXMLclient.php" & _
                        "?lat=" & plngLatitude & "&lon=" & plngLongitude & _
                        "&product=time-series" & _
                        "&begin=" & CStr(Format(Date, "YYYY-MM-DD")) & "T00:00:00" & _
                        "&end=" & CStr(Format(Date + 5, "YYYY-MM-DD")) & "T00:00:00" & _
                        "&maxt=maxt&mint=mint&Unit=m"
                        
    basSystem.log "http call:" & strRestRequest, cLogDebug
    With objServerXml
        .Open "GET", strRestRequest, False
        .setRequestHeader "Accept", "text/xml"
        .setRequestHeader "Content-Type", "text/xml"
        .send
        .waitForResponse 60
    End With
    If objServerXml.Status = 200 Then
        basSystem.log "got weather data", cLogInfo
        Set getWeatherData = objServerXml.responseXML
    Else
        basSystem.log_error "basWeather.getWeatherData", "server failed to answer request with return code " & objServerXml.Status
        Set getWeatherData = Nothing
    End If
    Exit Function
                     
error_handler:
    basSystem.log_error "basWeather.getWeatherData"
End Function
'-------------------------------------------------------------
' Description   : create a new xl file with weather forecast data
' Parameter     : pcolForecastData
'-------------------------------------------------------------
Public Sub createReport(pcolForecastData As Collection)
    
    Dim wbkReport As Workbook
    Dim wshReport As Worksheet
    Dim rngCurrent As Range
    Dim intDateCount As Integer
    Dim intCityCount As Integer
    Dim colCityData As Collection
    Dim varTemperature As Variant
    Dim intTemperatureCount As Integer
                     
    On Error GoTo error_handler
    intCityCount = 1
    'create a new workbook
    Set wbkReport = Application.Workbooks.Add
    'keep only a single worksheet
    Application.DisplayAlerts = False
    While wbkReport.Worksheets.Count > 1
        wbkReport.Worksheets(1).Delete
    Wend
    Application.DisplayAlerts = True
    Set wshReport = wbkReport.Worksheets(1)
    wshReport.Name = "weather forecast"
    Set rngCurrent = wshReport.Range("A1")
    'write dates
    rngCurrent.Value = "date"
    For intDateCount = 0 To 4
        rngCurrent.Offset(intDateCount + 1).Value = Date + intDateCount
    Next
    'write temperature data for each city
    For Each colCityData In pcolForecastData
        rngCurrent.Offset(, intCityCount * 2 - 1).Value = colCityData("city") & " high"
        rngCurrent.Offset(, intCityCount * 2 - 1).Columns.AutoFit
        intTemperatureCount = 1
        For Each varTemperature In colCityData("maxTemperatures")
            rngCurrent.Offset(intTemperatureCount, intCityCount * 2 - 1).Value = CLng(varTemperature)
            intTemperatureCount = intTemperatureCount + 1
        Next
        rngCurrent.Offset(, intCityCount * 2).Value = colCityData("city") & " low"
        rngCurrent.Offset(, intCityCount * 2).Columns.AutoFit
        intTemperatureCount = 1
        For Each varTemperature In colCityData("minTemperatures")
            rngCurrent.Offset(intTemperatureCount, intCityCount * 2).Value = CLng(varTemperature)
            intTemperatureCount = intTemperatureCount + 1
        Next
        intCityCount = intCityCount + 1
    Next
Exit Sub
                     
error_handler:
    basSystem.log_error "basWeather.createReport"
End Sub
