
public Function ClearErrorInVbs_fn()
	Err.Clear
End Function




Function TimeStamp_fn()
   Dim sHour,sMinute,sSecond
	sHour = Hour(Now)
	sMinute = Minute(Now)
	sSecond = Second(Now)
	TimeStamp_fn = sHour & sMinute & sSecond
End Function
