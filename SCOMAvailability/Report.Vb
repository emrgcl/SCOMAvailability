
Const LocTableStringQuery As String = "R/S[@N='{0}']/text()"

Const TimeZoneParameterName As String = "TimeZone"
Const SD_BaseTypeParameterName As String = "StartDate_BaseType"
Const SD_BaseValueParameterName As String = "StartDate_BaseValue"
Const SD_OffsetTypeParameterName As String = "StartDate_OffsetType"
Const SD_OffsetValueParameterName As String = "StartDate_OffsetValue"
Const ED_BaseTypeParameterName As String = "EndDate_BaseType"
Const ED_BaseValueParameterName As String = "EndDate_BaseValue"
Const ED_OffsetTypeParameterName As String = "EndDate_OffsetType"
Const ED_OffsetValueParameterName As String = "EndDate_OffsetValue"
Const IsRelativeTimeSupported As Boolean = True
Const TimeTypeParameterName As String = "TimeType"
Const TimeWeekMapParameterName As String = "TimeWeekMap"
Const DataAggregationParameterName As String = "DataAggregation"

Dim LocTables As System.Collections.Generic.Dictionary(Of String, Microsoft.EnterpriseManagement.Reporting.XmlStringTable)
Dim ReportTimeZone As Microsoft.EnterpriseManagement.Reporting.TimeZoneCoreInformation
Dim ReportStartDate As DateTime
Dim ReportEndDate As DateTime
Dim ReportTime As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime
Dim ReportCulture As System.Globalization.CultureInfo
Dim ParameterProcessor As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor

Protected Overrides Sub OnInit()
  LocTables = new System.Collections.Generic.Dictionary(Of String, Microsoft.EnterpriseManagement.Reporting.XmlStringTable)()
  ReportTimeZone = Nothing
  ReportStartDate = DateTime.MinValue
  ReportEndDate = DateTime.MinValue
  ReportTime = Nothing
  ReportCulture = System.Globalization.CultureInfo.GetCultureInfo(Report.User("Language"))
  ParameterProcessor = New Microsoft.EnterpriseManagement.Reporting.ParameterProcessor(ReportCulture)
End Sub

Public Function GetCallingManagementGroupId() As String
   Return Microsoft.EnterpriseManagement.Reporting.ReportingConfiguration.ManagementGroupId
End Function

Public Function GetReportLocLanguageCode() As String
   Return ReportCulture.ThreeLetterWindowsLanguageName
End Function

Public Function GetLocTable(Name As String) As Microsoft.EnterpriseManagement.Reporting.XmlStringTable
   Dim LocTable As Microsoft.EnterpriseManagement.Reporting.XmlStringTable

   If Not LocTables.TryGetValue(Name, LocTable) Then
      LocTable = New Microsoft.EnterpriseManagement.Reporting.XmlStringTable(LocTableStringQuery, Report.Parameters(Name).Value)
      LocTables.Add(Name, LocTable)
   End If

   Return LocTable
End Function

Public Function GetReportTimeZone() As Microsoft.EnterpriseManagement.Reporting.TimeZoneCoreInformation
   If IsNothing(ReportTimeZone) Then ReportTimeZone = Microsoft.EnterpriseManagement.Reporting.TimeZoneCoreInformation.FromValueString(Report.Parameters(TimeZoneParameterName).Value)
   Return ReportTimeZone
End Function

Public Function ToDbDate(ByVal DateValue As DateTime) As DateTime
   return GetReportTimeZone.ToUniversalTime(DateValue)
End Function

Public Function ToReportDate(ByVal DateValue As DateTime) As DateTime
   return GetReportTimeZone.ToLocalTime(DateValue)
End Function

Public Function GetReportStartDate() As DateTime
  If (ReportStartDate = DateTime.MinValue) Then
    If (IsRelativeTimeSupported) Then
      ReportStartDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(SD_BaseTypeParameterName).Value, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(SD_OffsetTypeParameterName).Value, Report.Parameters(SD_OffsetValueParameterName).Value, Report.Parameters(TimeTypeParameterName).Value)
    Else
      ReportStartDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(SD_BaseTypeParameterName).Value, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(SD_OffsetTypeParameterName).Value, Report.Parameters(SD_OffsetValueParameterName).Value)
    End if
    
    If (CInt(Report.Parameters(DataAggregationParameterName).Value) > 0) Then
       ReportStartDate = ReportStartDate.Date
    End If
  End If
  return ReportStartDate
End Function

Public Function GetReportEndDate() As DateTime
  If (ReportEndDate = DateTime.MinValue) Then
    If (IsRelativeTimeSupported) Then
      ReportEndDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(ED_BaseTypeParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, Report.Parameters(ED_OffsetTypeParameterName).Value, Report.Parameters(ED_OffsetValueParameterName).Value, Report.Parameters(TimeTypeParameterName).Value)
      If IsBusinessHours(GetReportTimeFilter()) Then ReportEndDate = ReportCulture.Calendar.AddDays(ReportEndDate, 1)
    Else
      ReportEndDate = ParameterProcessor.GetDateTime(ToReportDate(DateTime.UtcNow), Report.Parameters(ED_BaseTypeParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, Report.Parameters(ED_OffsetTypeParameterName).Value, Report.Parameters(ED_OffsetValueParameterName).Value)
    End if

    If (CInt(Report.Parameters(DataAggregationParameterName).Value) > 0) Then
       ReportEndDate = ReportEndDate.Date
    End If
  End If
  return ReportEndDate
End Function

Public Function GetReportTimeFilter() As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime
  If IsNothing(ReportTime) Then
    If (CInt(Report.Parameters(DataAggregationParameterName).Value) = 0)
		ReportTime = New Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime(Report.Parameters(TimeTypeParameterName).Value, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, CStr(Join(Report.Parameters(TimeWeekMapParameterName).Value, ",")))
	Else
		ReportTime = New Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime(Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTimeType.Regular, Report.Parameters(SD_BaseValueParameterName).Value, Report.Parameters(ED_BaseValueParameterName).Value, CStr(Join(Report.Parameters(TimeWeekMapParameterName).Value, ",")))
	End if
  End If
  return ReportTime
End Function

Public Function IsBusinessHours(Value As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime) As Boolean
  return (Not IsNothing(Value)) And (Value.TimeType = Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTimeType.Business)
End Function

Public Function FormatDateTime(Format As String, Value As DateTime) As String
  return Value.ToString(Format, ReportCulture)
End Function

Public Function FormatNumber(Format As String, Value As Decimal) As String
  return Value.ToString(Format, ReportCulture)
End Function

Public Function FormatString(Format As String, ParamArray Values() as  Object) As String
  return String.Format(ReportCulture, Format, Values)
End Function

Public Function NullFormatString(Format As String, Value as  String) As String
  return IIF(String.IsNullOrEmpty(Value), String.Empty, String.Format(ReportCulture, Format, Value))
End Function

Public Function FormatBusinessHours(Format As String, Value As Microsoft.EnterpriseManagement.Reporting.ParameterProcessor.RelativeTime) As String
  Dim result As String

  If IsBusinessHours(Value) Then
 
    Dim firstDay As DayOfWeek
    Dim days As System.Collections.Generic.List(Of String)

    firstDay = ReportCulture.DateTimeFormat.FirstDayOfWeek
    days = new System.Collections.Generic.List(Of String)()

    For loopDay As DayOfWeek = DayOfWeek.Sunday To DayOfWeek.Saturday
      Dim day As DayOfWeek
      day = CType((CInt(loopDay) + CInt(firstDay)) Mod 7, DayOfWeek)

      If value.WeekMap.Contains(day) Then days.Add(ReportCulture.DateTimeFormat.GetAbbreviatedDayName(day))
    Next loopDay

     result = FormatString(Format, DateTime.Today.Add(Value.StartTime).ToString(ReportCulture.DateTimeFormat.ShortTimePattern), DateTime.Today.Add(Value.EndTime).ToString(ReportCulture.DateTimeFormat.ShortTimePattern), String.Join(",", days.ToArray()))

  Else
     result = String.Empty
  End if
  
  return result
End Function

REM ---------------------------------

Public Enum DataAggregationType
    Hourly = 0
    Daily = 1
End Enum

REM ---------------------------------

Const DownTimeParameterName As String = "DownTime"
Dim DownTimeTable As System.Collections.Generic.List(Of Integer)

Public Enum StateIntervalType
	Red = 0
	Green = 1
	Yellow = 2
	White = 3
	Gray = 4
	Black = 5
	Blue = 6
	ServiceGray = 7
End Enum

Public Function IsDownTime(time As StateIntervalType)
	Select Case time
		Case StateIntervalType.Red
			Return True

		Case StateIntervalType.Green
			Return False

		Case Else
			If IsNothing(DownTimeTable) Then
				DownTimeTable = New System.Collections.Generic.List(Of Integer)()
				Dim val As String
				For Each val in Report.Parameters(DownTimeParameterName).Value
					DownTimeTable.Add(CInt(val))
				Next
			End If
			
			Return DownTimeTable.Contains(CInt(time))
	End Select
End Function


REM ---------------------------------

    Public Function FormatTimeInterval(ByVal TimeInSeconds As Integer) As String
        Dim Hours, Minutes, Seconds As Integer

        Hours = Math.Floor(TimeInSeconds / 3600)
        Minutes = Math.Floor((TimeInSeconds Mod 3600) / 60)
        Seconds = TimeInSeconds Mod 60

        Return FormatString("{0}:{1}:{2}", Hours.ToString(), Minutes.ToString("00"), Seconds.ToString("00"))
    End Function

REM ---------------------------------

    Public Enum AggregationType
        Hourly = 0
        Daily = 1
        Monthly = 2
    End Enum

    Public Function GetAggregationType(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As AggregationType
        Dim Result As AggregationType
        Dim TimeDiff As TimeSpan

        TimeDiff = EndDate - StartDate

        If (TimeDiff.TotalDays > 32) Then
            Result = AggregationType.Monthly
        ElseIf (TimeDiff.TotalDays > 1) Then
            Result = AggregationType.Daily
        Else
            If (CInt(Report.Parameters(DataAggregationParameterName).Value) = 0) Then
               Result = AggregationType.Hourly
            Else
               Result = AggregationType.Daily
            End If
        End If
        Return Result
    End Function
