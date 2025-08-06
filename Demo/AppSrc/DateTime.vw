Use Windows.pkg
Use DFClient.pkg
Use cCJGrid.pkg
Use cCJGridColumn.pkg
Use dfTabDlg.pkg
Use cSplitButton.pkg
Use cCJGridColumnRowIndicator.pkg

Use Classes\cDateTimeHandler.pkg
Use Classes\cLocaleInfoHandler.pkg

// From conversions library
Use cTimeHandler.pkg

// Modified DateTime form that can display milliseconds
Use cDateTimeFormEx.pkg

Use EnumDynamicTimeZoneInformation.h.pkg
Use cSplitterContainer.pkg
Use dfSpnFrm.pkg

Use MonthCalendarPrompt.dg

Activate_View Activate_oDateTimeView for oDateTimeView
Object oDateTimeView is a dbView
    Set Label To "DateTime Functions"
    Set Border_Style to Border_Thick
    Set Size to 219 523
    Set Location to 3 17
    Set Icon to "DateTime.Ico"
    Set pbAutoActivate to True

    Object oDateTimeHandler is a cDateTimeHandler
    End_Object

    Object oLocaleHandler is a cLocaleInfoHandler
    End_Object

    Object oDateTimeTabDialog is a dbTabDialog
        Set Size to 210 514
        Set Location to 4 5

        Set Rotate_Mode to RM_Rotate
        Set MultiLine_State to True
        Set peAnchors to anAll

        Object oSystemTimeTabPage is a dbTabPage
            Set Label to "System Time"

            Object oSystemTimeForm is a Form
                Set Size to 13 80
                Set Location to 5 102
                Set Label to "SystemTime:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set psToolTip to "System time (date and time)"
            End_Object

            Object oSystemTimeAsFileTimeForm is a Form
                Set Size to 13 80
                Set Location to 21 102
                Set psToolTip to "System time (as filetime)"
                Set Label to "FileTime:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
            End_Object

            Object oGetSystemTimeButton is a Button
                Set Size to 13 50
                Set Location to 5 5
                Set Label to "Get"
                Set psToolTip to "Gets the Windows system time and converts this to a DateTime value"

                Procedure OnClick
                    tWinSystemTime NowAsSystemTime
                    tWinFileTime NowAsFileTime
                    DateTime dtNow
                    UBigInt ubiFileTime

                    Get SystemTime of oDateTimeHandler to NowAsSystemTime
                    Get SystemTimeToDateTime of oDateTimeHandler NowAsSystemTime to dtNow
                    Set Value of oSystemTimeForm to dtNow

                    Get SystemTimeAsFileTime of oDateTimeHandler to NowAsFileTime
                    Move ((NowAsFileTime.dwHighDateTime * (2^32)) + NowAsFileTime.dwLowDateTime) to ubiFileTime
                    Set Value of oSystemTimeAsFileTimeForm to ubiFileTime
                End_Procedure
            End_Object

            Object oSetSystemTimeButton is a Button
                Set Size to 13 50
                Set Location to 5 186
                Set Label to "Set"
                Set psToolTip to "Sets the Windows system time if you have appropriate rights. Start the application with 'Run as Administrator' to try."

                Procedure Activating
                    Forward Send Activating

                    Set Enabled_State to (IsAdministrator ())
                End_Procedure

                Procedure OnClick
                    tWinSystemTime NewSystemDateTime
                    DateTime dtNewDateTime

                    Get Value of oSystemTimeForm to dtNewDateTime
                    Get DateTimeToSystemTime of oDateTimeHandler dtNewDateTime to NewSystemDateTime
                    Send SetSystemTime of oDateTimeHandler NewSystemDateTime
                End_Procedure
            End_Object
        End_Object

        Object oPerformanceTabPage is a dbTabPage
            Set Label to "Performance"

            Object oCounterForm is a Form
                Set Size to 13 60
                Set Location to 5 55
                Set Label to "Counter:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
                Set psToolTip to "The current value of the performance counter, which is a high resolution (<1us) time stamp that can be used for time-interval measurements."
            End_Object

            Object oPercentageForm is a Form
                Set Size to 13 60
                Set Location to 20 55
                Set Label to "Percentage:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
                Set psToolTip to "The frequency of the performance counter. The frequency of the performance counter is fixed at system boot and is consistent across all processors. Therefore, the frequency need only be queried upon application initialization, and the result can be cached."
            End_Object

            Object oQueryButton is a Button
                Set Location to 35 55
                Set Label to 'Get'

                Procedure OnClick
                    UBigInt ubiValue

                    Get QueryPerformanceCounter of oDateTimeHandler to ubiValue
                    Set Value of oCounterForm to ubiValue

                    Get QueryPerformancePercentage of oDateTimeHandler to ubiValue
                    Set Value of oPercentageForm to ubiValue
                End_Procedure
            End_Object
        End_Object

        Object oUnbiasedTimeTabPage is a dbTabPage
            Set Label to "Unbiased time"

            Object oUnbiasedTimeButton is a Button
                Set Size to 13 50
                Set Location to 110 55
                Set Label to 'Get'
                Set psToolTip to "Gets the current unbiased interrupt-time count. The unbiased interrupt-time count does not include time the system spends in sleep or hibernation."

                Procedure OnClick
                    UBigInt ubiTime ubiConverted
                    Handle hoTimeHandler

                    Get Create (RefClass (cTimeHandler)) to hoTimeHandler

                    Get QueryUnbiasedInterruptTime of oDateTimeHandler to ubiTime
                    Set Value of oUnbiasedNanoSecondsForm to ubiTime

                    Get NanoToMicroSecond of hoTimeHandler ubiTime to ubiConverted
                    Set Value of oUnbiasedMicroSecondsForm to ubiConverted

                    Get NanoToMilliSecond of hoTimeHandler ubiTime to ubiConverted
                    Set Value of oUnbiasedMilliSecondsForm to ubiConverted

                    Get NanoToSecond of hoTimeHandler ubiTime to ubiConverted
                    Set Value of oUnbiasedSecondsForm to ubiConverted

                    Get NanoSecondToMinute of hoTimeHandler ubiTime to ubiConverted
                    Set Value of oUnbiasedMinutesForm to ubiConverted

                    Get NanoSecondToHour of hoTimeHandler ubiTime to ubiConverted
                    Set Value of oUnbiasedHoursForm to ubiConverted

                    Get NanoSecondToDay of hoTimeHandler ubiTime to ubiConverted
                    Set Value of oUnbiasedDaysForm to ubiConverted

                    Send Destroy of hoTimeHandler
                End_Procedure
            End_Object

            Object oUnbiasedNanoSecondsForm is a Form
                Set Size to 13 60
                Set Location to 5 55
                Set Label to "Nano-seconds:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oUnbiasedMicroSecondsForm is a Form
                Set Size to 13 60
                Set Location to 20 55
                Set Label to "Micro-seconds:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oUnbiasedMilliSecondsForm is a Form
                Set Size to 13 60
                Set Location to 35 55
                Set Label to "Milli-Seconds:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oUnbiasedSecondsForm is a Form
                Set Size to 13 60
                Set Location to 50 55
                Set Label to "Seconds:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oUnbiasedMinutesForm is a Form
                Set Size to 13 60
                Set Location to 65 55
                Set Label to "Minutes:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oUnbiasedHoursForm is a Form
                Set Size to 13 60
                Set Location to 80 55
                Set Label to "Hours:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oUnbiasedDaysForm is a Form
                Set Size to 13 60
                Set Location to 95 55
                Set Label to "Days:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object
        End_Object

        Object oSystemTimesTabPage is a dbTabPage
            Set Label to "System times"

            Object oIdleForm is a Form
                Set Size to 13 48
                Set Location to 5 55
                Set Label to "Idle:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to Mask_Time
                Set psToolTip to "The amount of time that the system has been idle"
            End_Object

            Object oUserForm is a Form
                Set Size to 13 48
                Set Location to 20 55
                Set Label to "User:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to Mask_Time
                Set psToolTip to "The amount of time that the system has spent executing in User mode (including all threads in all processes, on all processors)"
            End_Object

            Object oKernelForm is a Form
                Set Size to 13 48
                Set Location to 35 55
                Set Label to "Kernel:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to Mask_Time
                Set psToolTip to "The amount of time that the system has spent executing in Kernel mode (including all threads in all processes, on all processors). This time value also includes the amount of time the system has been idle."
            End_Object

            Object oPreciseSystemTimeForm is a Form
                Set Size to 13 70
                Set Location to 50 55
                Set Label to "Precise:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to Mask_Datetime_Window
                Set psToolTip to "The GetSystemTimePreciseAsFileTime function retrieves the current system date and time with the highest possible level of precision (<1us)"
            End_Object

            Object oSystemTimesButton is a Button
                Set Size to 13 50
                Set Location to 65 55
                Set Label to 'Get'

                Procedure OnClick
                    Time tIdle tKernel tUser
                    DateTime dtSytemTime
                    Boolean bOk

                    Get SystemTimes of oDateTimeHandler (&tIdle) (&tKernel) (&tUser) to bOk
                    If (bOk) Begin
                        Set Value of oIdleForm to tIdle
                        Set Value of oKernelForm to tKernel
                        Set Value of oUserForm to tUser
                    End
                    Get PreciseSytemTime of oDateTimeHandler to dtSytemTime
                    Set Value of oPreciseSystemTimeForm to dtSytemTime
                End_Procedure
            End_Object
        End_Object

        Object oTimeAdjustmentsTabPage is a dbTabPage
            Set Label to "Time Adjustments"

            Object oGetResultsButton is a Button
                Set Size to 13 50
                Set Location to 5 50
                Set Label to 'Get'

                Procedure OnClick
                    tTimeAdjustment TimeAdjustments

                    Get SystemTimeAdjustment of oDateTimeHandler to TimeAdjustments
                    Set Value of oAdjustmentForm to TimeAdjustments.uiAdjustment
                    Set Value of oIncrementForm to TimeAdjustments.uiIncrement
                    Set Checked_State of oDisabledCheckbox to TimeAdjustments.bDisabled
                End_Procedure
            End_Object

            Object oAdjustmentForm is a Form
                Set Size to 13 50
                Set Location to 21 50
                Set Label to "Adjustment:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to 2
                Set psToolTip to "If the checkbox below is checked this value is not used"
            End_Object

            Object oIncrementForm is a Form
                Set Size to 13 50
                Set Location to 36 50
                Set Label to "Increment:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to 8
                Set psToolTip to "If the checkbox below is checked this value is not used"
            End_Object

            Object oDisabledCheckbox is a Checkbox
                Set Size to 13 100
                Set Location to 51 50
                Set Label to "Disabled"
                Set psToolTip to "Tells if the system uses a time adjustment or not"
            End_Object
        End_Object

        Object oISOWeekNumberTabPage is a dbTabPage
            Set Label to "Weeknumber"

            Object oDateForm is a Form
                Set Size to 13 50
                Set Location to 5 50
                Set Label to "Date:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to Date_Window

                Procedure Activating
                    Date dToday

                    Forward Send Activating

                    Sysdate dToday
                    Set Value to dToday
                End_Procedure
            End_Object

            Object oISOWeekNumberButton is a Button
                Set Size to 13 50
                Set Location to 20 50
                Set Label to 'Get'

                Procedure OnClick
                    Date dTheDate
                    Integer iWeekNumber

                    Get Value of oDateForm to dTheDate
                    Get ISO8601WeekNumber of oDateTimeHandler dTheDate to iWeekNumber
                    Set Value of oWeekNumberForm to iWeekNumber
                End_Procedure
            End_Object

            Object oWeekNumberForm is a Form
                Set Size to 13 50
                Set Location to 36 50
                Set Label to "Week No.:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to 0
            End_Object

            Object oYearNumberSpinForm is a SpinForm
                Set Size to 13 40
                Set Location to 5 194
                Set Label to "Year:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to 0
                Set Maximum_Position to 2500
                Set Minimum_Position to 1
            End_Object

            Object oFirstWeekDayInMonthOfYearButton is a Button
                Set Size to 13 50
                Set Location to 20 194
                Set psToolTip to 'First Weekday Year/WeekNumber/Month'
                Set Label to "Get"

                Procedure OnClick
                    Integer iYear
                    DateTime dtStart

                    Get Value of oYearNumberSpinForm to iYear

                    Get StartDateFirstWeekOfYear of oDateTimeHandler iYear to dtStart
                    Set Value of oDateResultForm to dtStart
                End_Procedure
            End_Object

            Object oDateResultForm is a Form
                Set Size to 13 50
                Set Location to 36 194
                Set Label to "Result:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Enabled_State to False
                Set Form_Datatype to Mask_Date_Window
            End_Object
        End_Object

        Object oTickCountsTabPage is a dbTabPage
            Set Label to "Tick Counts"

            Object oTickCountForm is a Form
                Set Size to 13 70
                Set Location to 5 50
                Set Label to "Small:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oTickCount64Form is a Form
                Set Size to 13 70
                Set Location to 20 50
                Set Label to "Big:"
                Set Label_Justification_Mode to JMode_Right
                Set Label_Col_Offset to 2
                Set Entry_State to False
                Set Form_Datatype to 0
            End_Object

            Object oRefreshButton is a Button
                Set Size to 13 50
                Set Location to 12 125
                Set Label to 'Refresh'

                Procedure OnClick
                    Send ShowValues
                End_Procedure
            End_Object

            Procedure Activating
                Forward Send Activating

                Send ShowValues
            End_Procedure

            Procedure ShowValues
                UInteger uiTickCount
                UBigInt ubiTickCount

                Get TickCount of oDateTimeHandler to uiTickCount
                Get TickCount64 of oDateTimeHandler to ubiTickCount

                Set Value of oTickCountForm to uiTickCount
                Set Value of oTickCount64Form to ubiTickCount
            End_Procedure
        End_Object


        Object oConvertTabPage is a dbTabPage
            Set Label to "DateTime -> FileTime"

            Object oDateTimeForm is a Form
                Set Size to 13 100
                Set Location to 5 60
                Set Label to "DateTime:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to Mask_Datetime_Window

                Procedure Entering Returns Integer
                    Forward Send Entering

                    Set Value to (CurrentDateTime ())
                End_Procedure
            End_Object

            Object oFileTimeForm is a Form
                Set Size to 13 100
                Set Location to 20 60
                Set Label to "FileTime:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Enabled_State to False
            End_Object

            Object oConvertedBackDateTimeForm is a Form
                Set Size to 13 100
                Set Location to 35 60
                Set Label to "Conv. DateTime:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Form_Datatype to Mask_Datetime_Window
                Set Enabled_State to False
            End_Object

            Object oConvertToFileTimeButton is a Button
                Set Size to 13 88
                Set Location to 5 167
                Set Label to 'Convert to FileTime'

                Procedure OnClick
                    DateTime dtIn
                    tWinFileTime ftResult
                    UBigInt ftInteger

                    Get Value of oDateTimeForm to dtIn
                    Get DateTimeToFileTime of oDateTimeHandler dtIn to ftResult
                    Move ((ftResult.dwHighDateTime * (2^32)) + ftResult.dwLowDateTime) to ftInteger
                    Set Value of oFileTimeForm to ftInteger
                End_Procedure
            End_Object

            Object oConvertToDateTimeButton is a Button
                Set Size to 13 88
                Set Location to 20 167
                Set Label to 'Convert to DateTime'

                Procedure OnClick
                    UBigInt ftInteger
                    tWinFileTime ftValue
                    DateTime dtBack

                    Get Value of oFileTimeForm to ftInteger
                    Move (ftInteger / (2^32)) to ftValue.dwHighDateTime
                    Move (ftInteger - (ftValue.dwHighDateTime * (2^32))) to ftValue.dwLowDateTime
                    Get FileTimeToDateTime of oDateTimeHandler ftValue to dtBack
                    Set Value of oConvertedBackDateTimeForm to dtBack
                End_Procedure
            End_Object

            Object oCompareFileTimeButton is a Button
                Set Size to 13 88
                Set Location to 35 167
                Set Label to 'Compare'

                Procedure OnClick
                    DateTime dtFirst dtSecond
                    tWinFileTime ftFirst ftSecond
                    ftCompareResult eResult

                    Get Value of oDateTimeForm to dtFirst
                    Get DateTimeToFileTime of oDateTimeHandler dtFirst to ftFirst

                    Get Value of oConvertedBackDateTimeForm to dtSecond
                    Get DateTimeToFileTime of oDateTimeHandler dtSecond to ftSecond

                    Get CompareFileTimes of oDateTimeHandler ftFirst ftSecond to eResult
                    Case Begin
                        Case (eResult = ftCompareResultFirstEarlier)
                            Send Info_Box (SFormat ("%1 before than %2", dtFirst, dtSecond))
                            Case Break
                        Case (eResult = ftCompareResultEqual)
                            Send Info_Box (SFormat ("%1 same as %2", dtFirst, dtSecond))
                            Case Break
                        Case (eResult = ftCompareResultFirstLater)
                            Send Info_Box (SFormat ("%1 later than %2", dtFirst, dtSecond))
                            Case Break
                    Case End
                End_Procedure
            End_Object
        End_Object

        Object oLocalTimeTabPage is a dbTabPage
            Set Label to "Local Time"

            Object oLocalTimeForm is a Form
                Set Size to 13 100
                Set Location to 5 50
                Set Label to "LocalTime:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
            End_Object

            Object oSystemTimeToSpecificLocalTimeButton is a Button
                Set Size to 13 50
                Set Location to 5 155
                Set Label to "Get"
                Set psToolTip to "Gets the Windows system time and via TimeZone information converts this to a DateTime value"

                Procedure OnClick
                    tWinSystemTime NowAsSystemTime stLocal
                    tTimeZoneInformation tzInfo
                    DateTime dtLocal

                    Get SystemTime of oDateTimeHandler to NowAsSystemTime
                    Get TimeZoneInformation of oDateTimeHandler to tzInfo
                    Get SystemTimeToTzSpecificLocalTime of oDateTimeHandler tzInfo NowAsSystemTime to stLocal
                    Get SystemTimeToDateTime of oDateTimeHandler stLocal to dtLocal
                    Set Value of oLocalTimeForm to dtLocal
                End_Procedure
            End_Object

            Object oSetNewLocalTimeButton is a Button
                Set Size to 13 50
                Set Location to 5 210
                Set Label to "Set"
                Set psToolTip to "Sets the Windows system time"
                Set Enabled_State to (IsAdministrator ())

                Procedure OnClick
                    DateTime dtNewDateTime
                    Boolean bErr

                    Get Value of oLocalTimeForm to dtNewDateTime
                    Get SetLocalTime of oDateTimeHandler dtNewDateTime to bErr
                End_Procedure
            End_Object

            Object oGetTimeZoneInfoButton is a Button
                Set Size to 13 50
                Set Location to 151 123
                Set Label to "Get"

                Procedure OnClick
                    tTimeZoneInformation tzInfo

                    Get TimeZoneInformation of oDateTimeHandler to tzInfo

                    Send DisplayInfo of oTimeZoneInfoGrid tzInfo
                End_Procedure
            End_Object

            Object oTimeZoneInfoGrid is a cCJGrid
                Set Size to 127 218
                Set Location to 22 51
                Set pbHideSelection to True
                Set psNoItemsText to "Click the 'get' button below the grid to get the info"
                Set pbAutoAppend to False
                Set pbAllowInsertRow to False
                Set pbAllowEdit to False
                Set pbAllowDeleteRow to False
                Set pbAllowColumnResize to False
                Set pbAllowColumnReorder to False
                Set pbAllowColumnRemove to False
                Set pbAllowAppendRow to False
                Set pbAutoSave to False
                Set pbReadOnly to True

                Object oMemberColumn is a cCJGridColumn
                    Set piWidth to 156
                    Set psCaption to "Member"
                End_Object

                Object oValueColumn is a cCJGridColumn
                    Set piWidth to 171
                    Set psCaption to "Value"
                End_Object

                Procedure DisplayInfo tTimeZoneInformation tzInfo
                    DateTime dtDate
                    Integer iItems iItem
                    tDataSourceRow[] TimeZoneData

                    Move (CurrentDateTime ()) to dtDate
                    Move (DateGetYear (dtDate)) to tzInfo.StandardDate.wYear
                    Move tzInfo.StandardDate.wYear to tzInfo.DayLightDate.wYear

                    Move "Bias" to TimeZoneData[0].sValue[0]
                    Move tzInfo.Bias to TimeZoneData[0].sValue[1]

                    Move "Standard Name" to TimeZoneData[1].sValue[0]
                    Move (PointerToWString (AddressOf (tzInfo.StandardName))) to TimeZoneData[1].sValue[1]
                    Move (Trim (TimeZoneData[1].sValue[1])) to TimeZoneData[1].sValue[1]

                    Move "Standard Date" to TimeZoneData[2].sValue[0]
                    Get TransitionDateTimeForYear of oDateTimeHandler tzInfo.StandardDate tzInfo.StandardDate.wYear to TimeZoneData[2].sValue[1]

                    Move "Standard Bias" to TimeZoneData[3].sValue[0]
                    Move tzInfo.StandardBias to TimeZoneData[3].sValue[1]

                    Move "Daylight Name" to TimeZoneData[4].sValue[0]
                    Move (PointerToWString (AddressOf (tzInfo.DaylightName))) to TimeZoneData[4].sValue[1]
                    Move (Trim (TimeZoneData[4].sValue[1])) to TimeZoneData[4].sValue[1]

                    Move "Daylight Date" to TimeZoneData[5].sValue[0]
                    Get TransitionDateTimeForYear of oDateTimeHandler tzInfo.DaylightDate tzInfo.DaylightDate.wYear to TimeZoneData[5].sValue[1]

                    Move "Daylight Bias" to TimeZoneData[6].sValue[0]
                    Move tzInfo.DaylightBias to TimeZoneData[6].sValue[1]

                    Send InitializeData TimeZoneData
                End_Procedure
            End_Object

            Object oTimeZoneInfoTextBox is a TextBox
                Set Size to 1 1
                Set Location to 54 7
                Set Label to "TimeZone:"
            End_Object
        End_Object

        Object oUTCTabPage is a dbTabPage
            Set Label to "UTC"

            Object oGetUTCTimeButton is a Button
                Set Size to 13 50
                Set Location to 5 5
                Set Label to "Get"

                Procedure OnClick
                    DateTime dtUTC

                    Get CurrentUTCDateTime of oDateTimeHandler to dtUTC
                    Set Value of oUTCDateTimeForm to dtUTC
                End_Procedure
            End_Object

            Object oUTCDateTimeForm is a Form
                Set Size to 13 80
                Set Location to 5 60
                Set Entry_State to False
                Set Form_Datatype to Mask_Datetime_Window
            End_Object
        End_Object

        Object oTimeZonesTabPage is a dbTabPage
            Set Label to "TimeZones"

            Object oTimeZonesContainer is a cSplitterContainer
                Set pbSplitVertical to False
                Set piSplitterLocation to 92

                Object oTimeZonesContainer is a cSplitterContainerChild
                    Object oTimeZonesGrid is a cCJGrid
                        Set Size to 68 499
                        Set Location to 5 5
                        Set peAnchors to anAll
                        Set piHeaderHeightMultiplier to 3
                        Set psNoItemsText to "No Time Zones Read / Found"
                        Set pbUseAlternateRowBackgroundColor to True
                        Set pbAllowAppendRow to False
                        Set pbAllowColumnRemove to False
                        Set pbAllowColumnReorder to False
                        Set pbAllowDeleteRow to False
                        Set pbAllowEdit to False
                        Set pbAllowInsertRow to False
                        Set pbAutoAppend to False
                        Set pbAutoSave to False
                        Set pbEditOnTyping to False

                        Object oCJGridColumnRowIndicator is a cCJGridColumnRowIndicator
                        End_Object

                        Object oBiasColumn is a cCJGridColumn
                            Set piWidth to 50
                            Set psCaption to "Bias"
                            Set peDataType to 0
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oFirstYearColumn is a cCJGridColumn
                            Set piWidth to 40
                            Set psCaption to "First Year"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set peDataType to Mask_Numeric_Window
                            Set psMask to "Z*"
                        End_Object

                        Object oLastYearColumn is a cCJGridColumn
                            Set piWidth to 40
                            Set psCaption to "Last Year"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set peDataType to Mask_Numeric_Window
                            Set psMask to "Z*"
                        End_Object

                        Object oStandardNameColumn is a cCJGridColumn
                            Set piWidth to 100
                            Set psCaption to "Standard Name"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oStandardDateColumn is a cCJGridColumn
                            Set piWidth to 70
                            Set psCaption to "To Standard Date"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oStandardBiasColumn is a cCJGridColumn
                            Set piWidth to 50
                            Set psCaption to "Standard Bias"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set peDataType to 0
                        End_Object

                        Object oDaylightNameColumn is a cCJGridColumn
                            Set piWidth to 100
                            Set psCaption to "Daylight Name"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oDaylightDateColumn is a cCJGridColumn
                            Set piWidth to 70
                            Set psCaption to "To Daylight Date"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oDaylightBiasColumn is a cCJGridColumn
                            Set piWidth to 50
                            Set psCaption to "Daylight Bias"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set peDataType to 0
                        End_Object

                        Object oTimeZoneKeyNameColumn is a cCJGridColumn
                            Set piWidth to 100
                            Set psCaption to "TimeZone Key Name"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oDynamicDaylightDisabledColumn is a cCJGridColumn
                            Set piWidth to 50
                            Set psCaption to "Dynamic Daylight Disabled"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set pbCheckbox to True
                        End_Object

                        Procedure EnumerateTimeZoneInformation
                            tDataSourceRow[] TimeZoneData
                            tDynamicTimeZoneInformation tzDynamicInfo
                            Integer iIndex iRetval iRow  iFirstYear iLastYear
                            Integer iBiasColumnId iFirstYearColumnId iLastYearColumnId
                            Integer iStandardNameColumnId iStandardDateColumnId iStandardBiasColumnId
                            Integer iDaylightNameColumnId iDaylightDateColumnId iDaylightBiasColumnId
                            Integer iTimeZoneKeyNameColumnId iDynamicTimeZoneDisabledColumnId

                            Get piColumnId of oBiasColumn to iBiasColumnId
                            Get piColumnId of oFirstYearColumn to iFirstYearColumnId
                            Get piColumnId of oLastYearColumn to iLastYearColumnId
                            Get piColumnId of oStandardNameColumn to iStandardNameColumnId
                            Get piColumnId of oStandardDateColumn to iStandardDateColumnId
                            Get piColumnId of oStandardBiasColumn to iStandardBiasColumnId
                            Get piColumnId of oDaylightNameColumn to iDaylightNameColumnId
                            Get piColumnId of oDaylightDateColumn to iDaylightDateColumnId
                            Get piColumnId of oDaylightBiasColumn to iDaylightBiasColumnId
                            Get piColumnId of oTimeZoneKeyNameColumn to iTimeZoneKeyNameColumnId
                            Get piColumnId of oDynamicDaylightDisabledColumn to iDynamicTimeZoneDisabledColumnId

                            Repeat
                                Increment iIndex
                                Move (WinAPI_EnumDynamicTimeZoneInformation (iIndex, AddressOf (tzDynamicInfo))) to iRetval
                                If (iRetval = ERROR_SUCCESS) Begin
                                    Move 0 to iFirstYear
                                    Move 0 to iLastYear
                                    Move (WinAPI_GetDynamicTimeZoneInformationEffectiveYears (AddressOf (tzDynamicInfo), AddressOf (iFirstYear), AddressOf (iLastYear))) to iRetval
                                    Move iFirstYear to TimeZoneData[iRow].sValue[iFirstYearColumnId]
                                    Move iLastYear to TimeZoneData[iRow].sValue[iLastYearColumnId]

                                    Move iIndex to TimeZoneData[iRow].vTag
                                    Move tzDynamicInfo.Bias to TimeZoneData[iRow].sValue[iBiasColumnId]

                                    Move (PointerToWString (AddressOf (tzDynamicInfo.StandardName))) to TimeZoneData[iRow].sValue[iStandardNameColumnId]
                                    Get TransitionDateTime of oDateTimeHandler tzDynamicInfo.StandardDate to TimeZoneData[iRow].sValue[iStandardDateColumnId]
                                    Move tzDynamicInfo.StandardBias to TimeZoneData[iRow].sValue[iStandardBiasColumnId]

                                    Move (PointerToWString (AddressOf (tzDynamicInfo.DaylightName))) to TimeZoneData[iRow].sValue[iDaylightNameColumnId]
                                    Get TransitionDateTime of oDateTimeHandler tzDynamicInfo.DaylightDate to TimeZoneData[iRow].sValue[iDaylightDateColumnId]
                                    Move tzDynamicInfo.DaylightBias to TimeZoneData[iRow].sValue[iDaylightBiasColumnId]

                                    Move (PointerToWString (AddressOf (tzDynamicInfo.TimeZoneKeyName))) to TimeZoneData[iRow].sValue[iTimeZoneKeyNameColumnId]
                                    Move tzDynamicInfo.DynamicDaylightTimeDisabled to TimeZoneData[iRow].sValue[iDynamicTimeZoneDisabledColumnId]

                                    Increment iRow
                                End
                            Until (iRetval = ERROR_NO_MORE_ITEMS)

                            Send InitializeData TimeZoneData
                        End_Procedure

                        Procedure OnRowChanged Integer iOldRow Integer iNewSelectedRow
                            Handle hoDataSource
                            tDataSourceRow CurrentTimeZoneData
                            Integer iFirstYearColumnId iLastYearColumnId

                            Forward Send OnRowChanged iOldRow iNewSelectedRow

                            Get phoDataSource to hoDataSource
                            Get pCurrentDataSourceRow of hoDataSource to CurrentTimeZoneData
                            Get piColumnId of oFirstYearColumn to iFirstYearColumnId
                            Get piColumnId of oLastYearColumn to iLastYearColumnId
                            Send ShowTimeZoneInformationForYears of oTimeZoneYearsGrid CurrentTimeZoneData.vTag ;
                                CurrentTimeZoneData.sValue[iFirstYearColumnId] CurrentTimeZoneData.sValue[iLastYearColumnId]
                        End_Procedure
                    End_Object

                    Object oReadTimeZonesButton is a Button
                        Set Size to 13 50
                        Set Location to 75 454
                        Set Label to "Read"
                        Set peAnchors to anBottomRight

                        Procedure OnClick
                            Send EnumerateTimeZoneInformation of oTimeZonesGrid
                        End_Procedure
                    End_Object
                End_Object

                Object oTimeZoneYearsContainer is a cSplitterContainerChild
                    Object oTimeZoneYearsGrid is a cCJGrid
                        Set Size to 79 499
                        Set Location to 5 5
                        Set peAnchors to anAll
                        Set piHeaderHeightMultiplier to 3
                        Set psNoItemsText to "No TimeZone Data for a series of years available"
                        Set pbUseAlternateRowBackgroundColor to True
                        Set pbAllowAppendRow to False
                        Set pbAllowColumnRemove to False
                        Set pbAllowColumnReorder to False
                        Set pbAllowDeleteRow to False
                        Set pbAllowEdit to False
                        Set pbAllowInsertRow to False
                        Set pbAutoAppend to False
                        Set pbAutoSave to False
                        Set pbEditOnTyping to False

                        Object oCJGridColumnRowIndicator is a cCJGridColumnRowIndicator
                            Set piWidth to 21
                        End_Object

                        Object oYearColumn is a cCJGridColumn
                            Set piWidth to 40
                            Set psCaption to "Year"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oBiasColumn is a cCJGridColumn
                            Set piWidth to 60
                            Set psCaption to "Bias"
                            Set peDataType to 0
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oStandardDateColumn is a cCJGridColumn
                            Set piWidth to 200
                            Set psCaption to "To Standard Date"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oStandardBiasColumn is a cCJGridColumn
                            Set piWidth to 60
                            Set psCaption to "Standard Bias"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set peDataType to 0
                        End_Object

                        Object oDaylightDateColumn is a cCJGridColumn
                            Set piWidth to 200
                            Set psCaption to "To Daylight Date"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                        End_Object

                        Object oDaylightBiasColumn is a cCJGridColumn
                            Set piWidth to 60
                            Set psCaption to "Daylight Bias"
                            Set peHeaderAlignment to (xtpAlignmentWordBreak + xtpAlignmentVCenter)
                            Set peDataType to 0
                        End_Object

                        Procedure ShowTimeZoneInformationForYears Integer iIndex Integer iFirstYear Integer iLastYear
                            Integer iYear iRetval iRow
                            Integer iYearColumnId iBiasColumnId iStandardDateColumnId
                            Integer iStandardBiasColumnId iDaylightDateColumnId iDaylightBiasColumnId
                            tDynamicTimeZoneInformation tzDynamicInfo
                            tTimeZoneInformation tzInfo
                            String sValue
                            tDataSourceRow[] TimeZoneData
                            Handle hoDataSource

                            If (iIndex > 0) Begin
                                Get piColumnId of oYearColumn to iYearColumnId
                                Get piColumnId of oBiasColumn to iBiasColumnId
                                Get piColumnId of oStandardDateColumn to iStandardDateColumnId
                                Get piColumnId of oStandardBiasColumn to iStandardBiasColumnId
                                Get piColumnId of oDaylightDateColumn to iDaylightDateColumnId
                                Get piColumnId of oDaylightBiasColumn to iDaylightBiasColumnId
                                Move (WinAPI_EnumDynamicTimeZoneInformation (iIndex, AddressOf (tzDynamicInfo))) to iRetval
                                If (iRetval = ERROR_SUCCESS) Begin
                                    If (iFirstYear > 0) Begin
                                        For iYear from iFirstYear to iLastYear
                                            Move (WinAPI_GetTimeZoneInformationForYear (iYear, AddressOf (tzDynamicInfo), AddressOf (tzInfo))) to iRetval
                                            If (iRetval <> 0) Begin
                                                Move iYear to TimeZoneData[iRow].sValue[iYearColumnId]
                                                Move tzInfo.Bias to TimeZoneData[iRow].sValue[iBiasColumnId]
                                                Move tzInfo.DaylightBias to TimeZoneData[iRow].sValue[iDaylightBiasColumnId]
                                                Move tzInfo.StandardBias to TimeZoneData[iRow].sValue[iStandardBiasColumnId]
                                                Get TransitionDateTimeForYear of oDateTimeHandler tzInfo.StandardDate iYear to TimeZoneData[iRow].sValue[iStandardDateColumnId]
                                                Get TransitionDateTimeForYear of oDateTimeHandler tzInfo.DaylightDate iYear to TimeZoneData[iRow].sValue[iDaylightDateColumnId]
                                                Increment iRow
                                            End
                                        Loop
                                    End
                                    Else Begin
                                        If (tzDynamicInfo.DaylightDate.wMonth > 0 or tzDynamicInfo.StandardDate.wMonth > 0) Begin
                                            Move (DateGetYear (CurrentDateTime ()) - 50) to iFirstYear
                                            Move (iFirstYear + 60) to iLastYear
                                            For iYear from iFirstYear to iLastYear
                                                Move iYear to TimeZoneData[iRow].sValue[iYearColumnId]
                                                Move tzDynamicInfo.Bias to TimeZoneData[iRow].sValue[iBiasColumnId]
                                                Move tzDynamicInfo.DaylightBias to TimeZoneData[iRow].sValue[iDaylightBiasColumnId]
                                                Move tzDynamicInfo.StandardBias to TimeZoneData[iRow].sValue[iStandardBiasColumnId]
                                                Get TransitionDateTimeForYear of oDateTimeHandler tzDynamicInfo.StandardDate iYear to TimeZoneData[iRow].sValue[iStandardDateColumnId]
                                                Get TransitionDateTimeForYear of oDateTimeHandler tzDynamicInfo.DaylightDate iYear to TimeZoneData[iRow].sValue[iDaylightDateColumnId]
                                                Increment iRow
                                            Loop
                                        End
                                    End
                                End
                            End

                            Send InitializeData TimeZoneData

                            Get phoDataSource to hoDataSource
                            Get FindColumnValue of hoDataSource iYearColumnId (DateGetYear (CurrentDateTime ())) True 0 False to iRow
                            Send MoveToRow iRow
                        End_Procedure
                    End_Object
                End_Object
            End_Object
        End_Object

        Object oUnixTimeTabPage is a dbTabPage
            Set Label to "Unix"

            Object oNowAsUnixTimeButton is a Button
                Set Size to 13 98
                Set Location to 5 5
                Set Label to 'Now as UNIX Epoch time'

                Procedure OnClick
                    Integer iUnixTime

                    Get DateTimeToUnixTime of oDateTimeHandler (CurrentDateTime ()) to iUnixTime
                    Set Value of oUnixEpochTimeForm to iUnixTime
                End_Procedure
            End_Object

            Object oUnixEpochTimeForm is a Form
                Set Size to 13 100
                Set Location to 5 108
                Set Form_Datatype to 0
                Set Entry_State to False
            End_Object

            Object oTimeStampToLocalDateTimeButton is a Button
                Set Size to 13 92
                Set Location to 5 213
                Set Label to 'To local datetime'
                Set psToolTip to 'Epoch timestamp to local datetime'

                Procedure OnClick
                    Integer iUnixTime
                    DateTime dtValue

                    Get Value of oUnixEpochTimeForm to iUnixTime
                    Get UnixTimeToLocalDateTime of oDateTimeHandler iUnixTime to dtValue
                    Set Value of oLocalDateTimeForm to dtValue
                End_Procedure
            End_Object

            Object oLocalDateTimeForm is a Form
                Set Size to 13 80
                Set Location to 5 332
                Set Form_Datatype to Mask_Datetime_Window
                Set Label to "Local:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Entry_State to False
            End_Object

            Object oTimeStampToUTCDateTimeButton is a Button
                Set Size to 13 92
                Set Location to 20 213
                Set Label to 'To UTC datetime'
                Set psToolTip to 'Epoch timestamp to UTC datetime'

                Procedure OnClick
                    Integer iUnixTime
                    DateTime dtValue

                    Get Value of oUnixEpochTimeForm to iUnixTime
                    Get UnixTimeToUTCDateTime of oDateTimeHandler iUnixTime to dtValue
                    Set Value of oUTCDateTimeForm to dtValue
                End_Procedure
            End_Object

            Object oUTCDateTimeForm is a Form
                Set Size to 13 80
                Set Location to 20 332
                Set Form_Datatype to Mask_Datetime_Window
                Set Label to "UTC:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Entry_State to False
            End_Object
        End_Object

        Object oJavaScriptTimeTabPage is a dbTabPage
            Set Label to "JavaScript"

            Object oNowAsJavaScriptTimeButton is a Button
                Set Size to 13 109
                Set Location to 5 5
                Set Label to 'Now as JavaScript Epoch time'

                Procedure OnClick
                    BigInt biJavaScriptTime

                    Get DateTimeToJavaScriptTime of oDateTimeHandler (CurrentDateTime ()) to biJavaScriptTime
                    Set Value of oJavaScriptEpochTimeForm to biJavaScriptTime
                End_Procedure
            End_Object

            Object oJavaScriptEpochTimeForm is a Form
                Set Size to 13 100
                Set Location to 5 118
                Set Form_Datatype to 0
                Set Entry_State to False
            End_Object

            Object oTimeStampToLocalDateTimeButton is a Button
                Set Size to 13 92
                Set Location to 5 223
                Set Label to 'To local datetime'
                Set psToolTip to 'Epoch timestamp to local datetime'

                Procedure OnClick
                    BigInt biJavaScriptTime
                    DateTime dtValue

                    Get Value of oJavaScriptEpochTimeForm to biJavaScriptTime
                    Get JavaScriptTimeToLocalDateTime of oDateTimeHandler biJavaScriptTime to dtValue
                    Set Value of oLocalDateTimeForm to dtValue
                End_Procedure
            End_Object

            Object oLocalDateTimeForm is a cDateTimeFormEx
                Set Size to 13 115
                Set Location to 5 342
                Set Label to "Local:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Entry_State to False
            End_Object

            Object oTimeStampToUTCDateTimeButton is a Button
                Set Size to 13 92
                Set Location to 20 223
                Set Label to 'To UTC datetime'
                Set psToolTip to 'Epoch timestamp to UTC datetime'

                Procedure OnClick
                    BigInt biJavaScriptTime
                    DateTime dtValue

                    Get Value of oJavaScriptEpochTimeForm to biJavaScriptTime
                    Get JavaScriptTimeToUTCDateTime of oDateTimeHandler biJavaScriptTime to dtValue
                    Set Value of oUTCDateTimeForm to dtValue
                End_Procedure
            End_Object

            Object oUTCDateTimeForm is a cDateTimeFormEx
                Set Size to 13 114
                Set Location to 20 342
                Set Label to "UTC:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Entry_State to False
            End_Object
        End_Object

        Object oLocaleInfoTabPage is a dbTabPage
            Set Label to "Locale Info"

            Object oLocalesComboForm is a ComboForm
                Set Size to 13 100
                Set Location to 5 50
                Set Label to "Select Locale:"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set Entry_State to False
                Set Combo_Sort_State to False

                Procedure Combo_Fill_List
                    String[] sCodes
                    Integer iElements iElement

                    Send Combo_Add_Item "Select a Locale"

                    Get AllLocaleCodes of oLocaleHandler to sCodes
                    Move (SizeOfArray (sCodes) - 1) to iElements
                    For iElement from 0 to iElements
                        Send Combo_Add_Item sCodes[iElement]
                    Loop
                End_Procedure

                Procedure OnChange
                    String sValue

                    Get Value to sValue
                    If (sValue <> "Select a Locale") Begin
                        Set psLocaleName of oLocaleHandler to sValue
                        Send UpdateGrid of oLocaleValuesGrid
                    End
                End_Procedure
            End_Object

            Object oLocaleValuesGrid is a cCJGrid
                Set Size to 160 456
                Set Location to 20 50
                Set pbAllowAppendRow to False
                Set pbAllowDeleteRow to False
                Set pbAllowInsertRow to False

                Object oLocaleTypeColumn is a cCJGridColumn
                    Set piWidth to 170
                    Set psCaption to "Type"
                    Set pbEditable to False
                End_Object

                Object oLocaleValueColumn is a cCJGridColumn
                    Set piWidth to 170
                    Set psCaption to "Value"
                    Set pbEditable to False
                End_Object

                Object oCurrentDateTimeColumn is a cCJGridColumn
                    Set piWidth to 170
                    Set psCaption to "Current"
                    Set pbEditable to False
                End_Object

                Procedure UpdateGrid
                    tDataSourceRow[] GridData
                    Integer iTypeColumnId iValueColumnId iCurrentColumnId iRow
                    DateTime dtNow
                    String sLocaleName

                    Get piColumnId of oLocaleTypeColumn to iTypeColumnId
                    Get piColumnId of oLocaleValueColumn to iValueColumnId
                    Get piColumnId of oCurrentDateTimeColumn to iCurrentColumnId

                    Get psLocaleName of oLocaleHandler to sLocaleName

                    Move (CurrentDateTime ()) to dtNow

                    Move "Short Date Format" to GridData[iRow].sValue[iTypeColumnId]
                    Get LocaleShortDateFormat of oLocaleHandler to GridData[iRow].sValue[iValueColumnId]
                    Get LocaleFormatDateTime of oDateTimeHandler dtNow sLocaleName GridData[iRow].sValue[iValueColumnId] to GridData[iRow].sValue[iCurrentColumnId]
                    Increment iRow
                    Move "Long Date Format" to GridData[iRow].sValue[iTypeColumnId]
                    Get LocaleLongDateFormat of oLocaleHandler to GridData[iRow].sValue[iValueColumnId]
                    Get LocaleFormatDateTime of oDateTimeHandler dtNow sLocaleName GridData[iRow].sValue[iValueColumnId] to GridData[iRow].sValue[iCurrentColumnId]
                    Increment iRow
                    Move "Time Format" to GridData[iRow].sValue[iTypeColumnId]
                    Get LocaleTimeFormat of oLocaleHandler to GridData[iRow].sValue[iValueColumnId]
                    Get LocaleFormatTime of oDateTimeHandler dtNow sLocaleName GridData[iRow].sValue[iValueColumnId] to GridData[iRow].sValue[iCurrentColumnId]
                    Increment iRow
                    Move "Decimal Separator" to GridData[iRow].sValue[iTypeColumnId]
                    Get LocaleDecimalSeparator of oLocaleHandler to GridData[iRow].sValue[iValueColumnId]
                    Increment iRow
                    Move "Thousands Separator" to GridData[iRow].sValue[iTypeColumnId]
                    Get LocaleThousandsSeparator of oLocaleHandler to GridData[iRow].sValue[iValueColumnId]
                    Increment iRow
                    Move "Currency Symbol" to GridData[iRow].sValue[iTypeColumnId]
                    Get LocaleCurrencySymbol of oLocaleHandler to GridData[iRow].sValue[iValueColumnId]

                    Send InitializeData GridData
                End_Procedure
            End_Object
        End_Object

        Object oTimeSpanTabPage is a dbTabPage
            Set Label to "TimeSpan"

            Object oStartDateForm is a Form
                Set Label to "Start Date"
                Set Prompt_Button_Mode to PB_PromptOn
                Set Prompt_Object to oMonthCalendarPrompt
                Set Form_Datatype to Date_Window
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set psToolTip to "Enter or select a start date for the tests"
                Set Size to 13 60
                Set Location to 5 50
            End_Object

            Object oEndDateForm is a Form
                Set Label to "End Date"
                Set Prompt_Button_Mode to PB_PromptOn
                Set Form_Datatype to Date_Window
                Set Prompt_Object to oMonthCalendarPrompt
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
                Set psToolTip to "Enter or select an end date for the tests"
                Set Size to 13 60
                Set Location to 20 50
            End_Object

            Object oCalculateSpans is a Button
                Set Size to 14 60
                Set Location to 15 115
                Set Label to "Calculate"
                Set psToolTip to "Calculate Time Span"

                Procedure OnClick
                    Integer iYearResult iMonthResult iWeekResult
                    DateTime dtVarStart dtVarEnd

                    Get Value of oStartDateForm to dtVarStart
                    Get Value of oEndDateForm to dtVarEnd

                    If (dtVarEnd < dtVarStart) Begin
                        Send Info_Box "Start Date must preceed End Date. Correct your entries and try again."
                    End
                    Else Begin
                        Get TotalYearsBetween of oDateTimeHandler dtVarStart dtVarEnd to iYearResult
                        Get TotalMonthsBetween of oDateTimeHandler dtVarStart dtVarEnd to iMonthResult
                        Get TotalWeeksBetween of oDateTimeHandler dtVarStart dtVarEnd to iWeekResult

                        Set Value of oNumMonthsForm to iMonthResult
                        Set Value of oNumYearsForm to iYearResult
                        Set Value of oNumWeeksForm to iWeekResult
                    End
                End_Procedure
            End_Object

            Object oNumYearsForm is a Form
                Set Location to 5 209
                Set Size to 13 60
                Set Form_Datatype to 0
                Set Label to "Years"
                Set psToolTip to "Number of Years"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
            End_Object

            Object oNumMonthsForm is a Form
                Set Location to 20 209
                Set Size to 13 60
                Set Form_Datatype to 0
                Set psToolTip to "Number of Months"
                Set Label to "Months"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
            End_Object

            Object oNumWeeksForm is a Form
                Set Location to 35 209
                Set Size to 13 60
                Set Form_Datatype to 0
                Set Label to "Weeks"
                Set psToolTip to "Number of Weeks"
                Set Label_Col_Offset to 2
                Set Label_Justification_Mode to JMode_Right
            End_Object
        End_Object
    End_Object
End_Object
