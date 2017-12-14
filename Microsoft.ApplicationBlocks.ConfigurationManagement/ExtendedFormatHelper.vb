' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' ExtendedFormatHelper.vb
'
' Provides an extended format helper to specify concurrent expirations.
' 
'
' For more information see the Configuration Management Application Block Implementation Overview. 
' 
' ===============================================================================
' Copyright (C) 2000-2001 Microsoft Corporation
' All rights reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
' FITNESS FOR A PARTICULAR PURPOSE.
' ==============================================================================

Imports System
Imports System.Collections
Imports [SC] = System.Configuration

#Region "Extended Format class"

' <summary>
' This class represents a extended format 
' </summary>

Friend Class ExtendedFormat
    Private Shared ARGUMENT_DELIMITER As Char = _
                            Convert.ToChar(",", System.Globalization.CultureInfo.CurrentUICulture)
    Private Shared WILDCARD_ALL As Char = _
                            Convert.ToChar("*", System.Globalization.CultureInfo.CurrentUICulture)
    Private Shared REFRESH_DELIMITER As Char = _
                            Convert.ToChar(" ", System.Globalization.CultureInfo.CurrentUICulture)

    Private _minutes() As Integer
    Private _hours() As Integer
    Private _days() As Integer
    Private _months() As Integer
    Private _daysOfWeek() As Integer


    Public Sub New(ByVal format As String)
        Dim parsedFormat As String() = format.Trim().Split(REFRESH_DELIMITER)

        If parsedFormat.Length <> 5 Then
            Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidExtendedFormatArguments"))
        End If

        _minutes = ParseValueToInt(parsedFormat(0))
        Dim minute As Integer
        For Each minute In _minutes
            If (minute > 59) Then
                Throw New ArgumentOutOfRangeException("format", Resource.ResourceManager("RES_ExceptionExtendedFormatIncorrectMinutePart"))
            End If
        Next

        _hours = ParseValueToInt(parsedFormat(1))
        Dim hour As Integer
        For Each hour In _hours
            If (hour > 23) Then
                Throw New ArgumentOutOfRangeException("format", Resource.ResourceManager("RES_ExceptionExtendedFormatIncorrectHourPart"))
            End If
        Next

        _days = ParseValueToInt(parsedFormat(2))
        Dim day As Integer
        For Each day In _days
            If (day > 31) Then
                Throw New ArgumentOutOfRangeException("format", Resource.ResourceManager("RES_ExceptionExtendedFormatIncorrectDayPart"))
            End If
        Next

        _months = ParseValueToInt(parsedFormat(3))
        Dim month As Integer
        For Each month In _months
            If (month > 12) Then
                Throw New ArgumentOutOfRangeException("format", Resource.ResourceManager("RES_ExceptionExtendedFormatIncorrectMonthPart"))
            End If
        Next

        _daysOfWeek = ParseValueToInt(parsedFormat(4))
        Dim dayOfWeek As Integer
        For Each dayOfWeek In _daysOfWeek
            If (dayOfWeek > 6) Then
                Throw New ArgumentOutOfRangeException("format", Resource.ResourceManager("RES_ExceptionExtendedFormatIncorrectDayOfWeekPart"))
            End If
        Next
    End Sub 'New


    Public ReadOnly Property Minutes() As Integer()
        Get
            Return _minutes
        End Get
    End Property

    Public ReadOnly Property Hours() As Integer()
        Get
            Return _hours
        End Get
    End Property

    Public ReadOnly Property Days() As Integer()
        Get
            Return _days
        End Get
    End Property

    Public ReadOnly Property Months() As Integer()
        Get
            Return _months
        End Get
    End Property

    Public ReadOnly Property DaysOfWeek() As Integer()
        Get
            Return _daysOfWeek
        End Get
    End Property

    Public ReadOnly Property ExpireEveryMinute() As Boolean
        Get
            Return _minutes(0) = -1
        End Get
    End Property

    Public ReadOnly Property ExpireEveryDay() As Boolean
        Get
            Return _days(0) = -1
        End Get
    End Property

    Public ReadOnly Property ExpireEveryHour() As Boolean
        Get
            Return _hours(0) = -1
        End Get
    End Property

    Public ReadOnly Property ExpireEveryMonth() As Boolean
        Get
            Return _months(0) = -1
        End Get
    End Property

    Public ReadOnly Property ExpireEveryDayOfWeek() As Boolean
        Get
            Return _daysOfWeek(0) = -1
        End Get
    End Property

    Private Function ParseValueToInt(ByVal value As String) As Integer()
        Dim result() As Integer

        If value.IndexOf(WILDCARD_ALL) <> -1 Then
            result = New Integer(0) {}
            result(0) = -1
        Else
            Dim values As String() = value.Split(ARGUMENT_DELIMITER)
            result = New Integer(values.Length) {}
            Dim i As Integer
            For i = 0 To values.Length - 1
                result(i) = Integer.Parse(values(i), System.Globalization.CultureInfo.CurrentUICulture)
            Next i
        End If

        Return result
    End Function 'ParseValueToInt
End Class 'ExtendedFormat

#End Region

#Region "Extended Format Helper class"

' <summary>
' This class tests if a item was expired using a extended format
' </summary>
' <remarks>
' Extended format syntax : <br/><br/>
' 
' Minute       - 0-59 <br/>
' Hour         - 0-23 <br/>
' Day of month - 1-31 <br/>
' Month        - 1-12 <br/>
' Day of week  - 0-6 (Sunday is 0 ) <br/>
' Wildcards    - * means run every <br/>
' Examples: <br/>
' * * * * *    - always expires <br/>
' 5 * * * *    - expire 5th minute of every hour <br/>
' * 21 * * *   - expire every minute of the 21st hour of every day <br/>
' 31 15 * * *  - expire 3:31 PM every day <br/>
' 7 4 * * 6    - expire Saturday 4:07 AM <br/>
' 15 21 4 7 *  - expire 9:15 PM on 4 July <br/>
'    Therefore 6 6 6 6 1 means:
'    •	have we crossed/entered the 6th minute AND
'    •	have we crossed/entered the 6th hour AND 
'    •	have we crossed/entered the 6th day AND
'    •	have we crossed/entered the 6th month AND
'    •	have we crossed/entered A MONDAY?'

'    Therefore these cases should exhibit these behaviors:
'
'			    getTime = DateTime.Parse( "02/20/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/07/2003 07:07:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 1", getTime, nowTime );
 '   TRUE, ALL CROSSED/ENTERED
  '  			
'			    getTime = DateTime.Parse( "02/20/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/07/2003 07:07:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 5", getTime, nowTime );
 '   TRUE
'    			
'			    getTime = DateTime.Parse( "02/20/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/06/2003 06:06:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 *", getTime, nowTime );
 '   TRUE
'    	
'    			
'			    getTime = DateTime.Parse( "06/05/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/06/2003 06:06:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 5", getTime, nowTime );
'    TRUE
'    						
'			    getTime = DateTime.Parse( "06/05/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/06/2005 05:06:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 1", getTime, nowTime );
'    TRUE
'    						
'			    getTime = DateTime.Parse( "06/05/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/06/2003 05:06:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 1", getTime, nowTime );
'    FALSE:  we did not cross 6th hour, nor did we cross Monday
'    						
'			    getTime = DateTime.Parse( "06/05/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/06/2003 06:06:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 5", getTime, nowTime );
'    TRUE, we cross/enter Friday
'
'
'			    getTime = DateTime.Parse( "06/05/2003 04:06:55 AM" );
'			    nowTime = DateTime.Parse( "06/06/2003 06:06:00 AM" );
'			    isExpired = ExtendedFormatHelper.IsExtendedExpired( "6 6 6 6 1", getTime, nowTime );
'    FALSE:  we don’t cross Monday but all other conditions satisfied
' </remarks>

Friend Class ExtendedFormatHelper
    Private Shared REFRESH_DELIMITER As Char = _
                    Convert.ToChar(" ", System.Globalization.CultureInfo.CurrentUICulture)
    Private Shared WILDCARD_ALL As Char = _
                    Convert.ToChar("*", System.Globalization.CultureInfo.CurrentUICulture)

    Private Shared _parsedFormatCache As New Hashtable


    ' <summary>
    ' Test the extended format with a given date.
    ' </summary>
    ' <param name="format">The extended format arguments.</param>
    ' <param name="getTime">The time when the item has been refreshed.</param>
    ' <param name="nowTime">Always DateTime.Now, or the date to test with.</param>
    ' <returns>true if the item was expired, otherwise false</returns>
    Public Shared Function IsExtendedExpired(ByVal format As String, ByVal getTime As DateTime, _
                                ByVal nowTime As DateTime) As Boolean
        'Validate arguments
        If format Is Nothing Then
            Throw New ArgumentNullException("format")
        End If

        'Remove the seconds to provide better precission on calculations
        getTime = getTime.AddSeconds(getTime.Second * -1)
        nowTime = nowTime.AddSeconds(nowTime.Second * -1)

        Dim parsedFormat As ExtendedFormat = CType(_parsedFormatCache(format), ExtendedFormat)
        If parsedFormat Is Nothing Then
            parsedFormat = New ExtendedFormat(format)
            SyncLock _parsedFormatCache.SyncRoot
                _parsedFormatCache(format) = parsedFormat
            End SyncLock
        End If
        'Validate the format arguments

        If (nowTime.Subtract(getTime).TotalMinutes < 1) Then Return False

        ' Validate the format arguments
        Dim minute As Integer
        For Each minute In parsedFormat.Minutes
            Dim hour As Integer
            For Each hour In parsedFormat.Hours
                Dim day As Integer
                For Each day In parsedFormat.Days
                    Dim month As Integer
                    For Each month In parsedFormat.Months

                        ' Set the expiration date parts
                        Dim expirMinute As Integer = CInt(IIf(minute = -1, getTime.Minute, minute))
                        Dim expirHour As Integer = CInt(IIf(hour = -1, getTime.Hour, hour))
                        Dim expirDay As Integer = CInt(IIf(day = -1, getTime.Day, day))
                        Dim expirMonth As Integer = CInt(IIf(month = -1, getTime.Month, month))
                        Dim expirYear As Integer = getTime.Year

                        ' Adjust when wildcards are set
                        If (minute = -1 AndAlso hour <> -1) Then expirMinute = 0
                        If (hour = -1 AndAlso day <> -1) Then expirHour = 0
                        If (minute = -1 AndAlso day <> -1) Then expirMinute = 0
                        If (day = -1 AndAlso month <> -1) Then expirDay = 1
                        If (hour = -1 AndAlso month <> -1) Then expirHour = 0
                        If (minute = -1 AndAlso month <> -1) Then expirMinute = 0

                        If (DateTime.DaysInMonth(expirYear, expirMonth) < expirDay) Then
                            If (expirMonth = 12) Then
                                expirMonth = 1
                                expirYear = expirYear + 1
                            End If
                        Else
                            expirMonth = expirMonth + 1
                            expirDay = 1
                        End If

                        ' See http://www.dotnet247.com/247reference/msgs/34/173671.aspx
                        '----fix for Month > 12)
                        If (expirMonth > 12) Then
                            expirMonth = 1
                            expirDay = 1
                            expirYear = expirYear + 1
                        End If
                        '----end fix

                        ' Create the date with the adjusted parts
                        Dim expTime As DateTime = _
                                    New DateTime(expirYear, expirMonth, expirDay, expirHour, expirMinute, 0)

                        ' Adjust when expTime is before getTime
                        If (expTime < getTime) Then
                            If (month <> -1 AndAlso getTime.Month >= month) Then
                                expTime = expTime.AddYears(1)
                            ElseIf (day <> -1 AndAlso getTime.Day >= day) Then
                                expTime = expTime.AddMonths(1)
                            ElseIf (hour <> -1 AndAlso getTime.Hour >= hour) Then
                                expTime = expTime.AddDays(1)
                            ElseIf (minute <> -1 AndAlso getTime.Minute >= minute) Then
                                expTime = expTime.AddHours(1)
                            End If
                        End If
                        ' Is Expired?
                        If (parsedFormat.ExpireEveryDayOfWeek) Then
                            If (nowTime >= expTime) Then Return True
                        Else
                            ' Validate WeekDay
                            Dim weekDay As Integer
                            For Each weekDay In parsedFormat.DaysOfWeek
                                Dim tmpTime As DateTime = getTime
                                tmpTime = tmpTime.AddHours(-1 * tmpTime.Hour)
                                tmpTime = tmpTime.AddMinutes(-1 * tmpTime.Minute)
                                Do While (Int(tmpTime.DayOfWeek) <> weekDay)
                                    tmpTime = tmpTime.AddDays(1)
                                Loop
                                If (nowTime >= tmpTime AndAlso nowTime >= expTime) Then Return True
                            Next
                        End If
                    Next
                Next
            Next
        Next
        Return False
    End Function 'IsExtendedExpired
End Class 'ExtendedFormatHelper
#End Region
