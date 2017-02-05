# get weather
This code sample gets current temparture from weather.gov
(http://graphical.weather.gov/xml/rest.php).

## REST query

http://graphical.weather.gov/xml/sample_products/browser_interface/ndfdXMLclient.php?lat=38.99&lon=-77.01&product=time-series&begin=2016-09-25T00:00:00&end=2016-09-30T00:00:00&maxt=maxt&mint=mint&Unit=m

returns temperature date for the given time period. (Hint: the end date has to be a future date!). Result format is xml.

### more info
call XML server
https://msdn.microsoft.com/en-us/library/ms535874(v=vs.85).aspx

read DOM document
https://msdn.microsoft.com/library/ms757878(v=vs.85).aspx

## another example with jmeter
The file weather_us.jmx is am alternative sample this time
using jmeter.
