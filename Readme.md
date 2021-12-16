# Documentation 

The Reporting services availability report proceses data in two pillars.

- SQL Story: The group membership and the aggregated data is gathered. for group membership the group object has to be passed as an xml string to sql referred in details below.
- Reporting Services Story: Availability is calculated at this stage

# Reporting Services Story
## Percent Uptime
=Code.FormatNumber("P2", Sum(Fields!UpTimeMilliseconds.Value) / Sum(Fields!IntervalDurationMilliseconds.Value))

## UpTimeMilliSeconds
```vb
=Fields!InGreenStateMilliseconds.Value +
 IIF(Code.IsDownTime(Code.StateIntervalType.Yellow), 0, Fields!InYellowStateMilliseconds.Value) +
 IIF(Code.IsDownTime(Code.StateIntervalType.White), 0, Fields!InWhiteStateMilliseconds.Value) +
 IIF(Code.IsDownTime(Code.StateIntervalType.Gray), 0, Fields!InDisabledStateMilliseconds.Value) +
 IIF(Code.IsDownTime(Code.StateIntervalType.ServiceGray), 0, Fields!HealthServiceUnavailableMilliseconds.Value) +
 IIF(Code.IsDownTime(Code.StateIntervalType.Blue), 0, Fields!InPlannedMaintenanceMilliseconds.Value) +
 IIF(Code.IsDownTime(Code.StateIntervalType.Black), 0, Fields!InUnplannedMaintenanceMilliseconds.Value)
```

## DowntimeMilliSeconds
```vb
=Fields!InRedStateMilliseconds.Value + 
 IIF(Code.IsDownTime(Code.StateIntervalType.Yellow), Fields!InYellowStateMilliseconds.Value, 0) + 
 IIF(Code.IsDownTime(Code.StateIntervalType.White), Fields!InWhiteStateMilliseconds.Value, 0) + 
 IIF(Code.IsDownTime(Code.StateIntervalType.Gray), Fields!InDisabledStateMilliseconds.Value, 0) + 
 IIF(Code.IsDownTime(Code.StateIntervalType.ServiceGray), Fields!HealthServiceUnavailableMilliseconds.Value, 0) +
 IIF(Code.IsDownTime(Code.StateIntervalType.Blue), Fields!InPlannedMaintenanceMilliseconds.Value, 0) + 
 IIF(Code.IsDownTime(Code.StateIntervalType.Black), Fields!InUnplannedMaintenanceMilliseconds.Value, 0)
```



> Note: IIF (Clause, return when true, return when False)

Enum and Function used to calculate the above expresion.
```vb
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
```
# SQL Story

## SP execute

## How to run the sp 
Microsoft_SystemCenter_DataWarehouse_Report_Library_AvailabilityReportDataGet:

```sql
exec Microsoft_SystemCenter_DataWarehouse_Report_Library_AvailabilityReportDataGet @ObjectList=N'<Data><Objects><Object Use="Containment">137</Object></Objects></Data>',@MonitorName=N'System.Health.AvailabilityState',@StartDate='2021-12-01 08:54:00',@EndDate='2021-12-16 08:54:00',@DataAggregation=0,@LanguageCode=N'ENU'
```

The Xml used in the SP for Group Containment.

```xml
<Data><Objects><Object Use="Containment">137</Object></Objects></Data>
```

The number  used represent ManagedEntityRowId in OperationsManagerDW