VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "Select Date"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2745
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum calDayOfWeek
    Sunday = 1
    Monday = 2
    Tuesday = 3
    Wednesday = 4
    Thursday = 5
    Friday = 6
    Saturday = 7
End Enum

Public Enum calFirstWeekOfYear
    FirstJan1 = 1
    FirstFourDays = 2
    FirstFullWeek = 3
                                                                                
End Enum

Private UserformEventsEnabled As Boolean
Private DateOut As Date
Private SelectedDateIn As Date
                                            
Private OkayEnabled As Boolean
Private TodayEnabled As Boolean
Private MinDate As Date
Private MaxDate As Date
Private cmbYearMin As Long
Private cmbYearMax As Long
Private StartWeek As VbDayOfWeek
Private WeekOneOfYear As VbFirstWeekOfYear
Private HoverControlName As String
                                            
Private HoverControlColor As Long
Private RatioToResize As Double
                                            
Private bgDateColor As Long
Private bgDateHoverColor As Long
Private bgDateSelectedColor As Long
Private lblDateColor As Long
Private lblDatePrevMonthColor As Long
Private lblDateTodayColor As Long
Private lblDateSatColor As Long
Private lblDateSunColor As Long

Public Function GetDate(Optional SelectedDate As Date = 0, _
    Optional FirstDayOfWeek As calDayOfWeek = Sunday, _
    Optional MinimumDate As Date = 0, _
    Optional MaximumDate As Date = 0, _
    Optional RangeOfYears As Long = 10, _
    Optional DateFontSize As Long = 9, _
    Optional TodayButton As Boolean = False, Optional OkayButton As Boolean = False, _
    Optional ShowWeekNumbers As Boolean = False, Optional FirstWeekOfYear As calFirstWeekOfYear = FirstJan1, _
    Optional PositionTop As Long = -5, Optional PositionLeft As Long = -5, _
    Optional BackgroundColor As Long = 16777215, _
    Optional HeaderColor As Long = 15658734, _
    Optional HeaderFontColor As Long = 0, _
    Optional SubHeaderColor As Long = 16448250, _
    Optional SubHeaderFontColor As Long = 8553090, _
    Optional DateColor As Long = 16777215, _
    Optional DateFontColor As Long = 0, _
    Optional SaturdayFontColor As Long = 0, _
    Optional SundayFontColor As Long = 0, _
    Optional DateBorder As Boolean = False, Optional DateBorderColor As Long = 15658734, _
    Optional DateSpecialEffect As fmSpecialEffect = fmSpecialEffectFlat, _
    Optional DateHoverColor As Long = 15658734, _
    Optional DateSelectedColor As Long = 14277081, _
    Optional TrailingMonthFontColor As Long = 12566463, _
    Optional TodayFontColor As Long = 15773696) As Date
    
    DateFontSize = Max(DateFontSize, 9)
    OkayEnabled = OkayButton
    TodayEnabled = TodayButton
    RatioToResize = DateFontSize / 9
    bgDateColor = DateColor
    lblDateColor = DateFontColor
    lblDateSatColor = SaturdayFontColor
    lblDateSunColor = SundayFontColor
    bgDateHoverColor = DateHoverColor
    bgDateSelectedColor = DateSelectedColor
    lblDatePrevMonthColor = TrailingMonthFontColor
    lblDateTodayColor = TodayFontColor
    StartWeek = FirstDayOfWeek
    WeekOneOfYear = FirstWeekOfYear
    
    UserformEventsEnabled = False
    Call InitializeUserform(SelectedDate, MinimumDate, MaximumDate, RangeOfYears, PositionTop, PositionLeft, _
        DateFontSize, ShowWeekNumbers, BackgroundColor, HeaderColor, HeaderFontColor, SubHeaderColor, _
        SubHeaderFontColor, DateBorder, DateBorderColor, DateSpecialEffect)
    UserformEventsEnabled = True
    
    Me.Show
    GetDate = DateOut
    Unload Me
End Function

Private Sub InitializeUserform(SelectedDate As Date, MinimumDate As Date, MaximumDate As Date, _
    RangeOfYears As Long, _
    PositionTop As Long, PositionLeft As Long, _
    SizeFont As Long, bWeekNumbers As Boolean, _
    BackgroundColor As Long, _
    HeaderColor As Long, _
    HeaderFontColor As Long, _
    SubHeaderColor As Long, _
    SubHeaderFontColor As Long, _
    DateBorder As Boolean, DateBorderColor As Long, _
    DateSpecialEffect As fmSpecialEffect)
    
    Dim TempDate As Date
    Dim SelectedYear As Long
    Dim SelectedMonth As Long
    Dim SelectedDay As Long
    Dim TempDayOfWeek As Long
    Dim BorderSpacing As Double
    Dim HeaderDefaultFontSize As Long
    Dim bgHeaderDefaultHeight As Double
    Dim lblMonthYearDefaultHeight As Double
    Dim scrlMonthDefaultHeight As Double
    Dim bgDayLabelsDefaultHeight As Double
    Dim bgDateDefaultHeight As Double
    Dim bgDateDefaultWidth As Double
    Dim lblDateDefaultHeight As Double
    Dim cmdButtonDefaultHeight As Double
    Dim cmdButtonDefaultWidth As Double
    Dim cmdButtonsCombinedWidth As Double
    Dim cmdButtonsMaxHeight As Double
    Dim cmdButtonsMaxWidth As Double
    Dim cmdButtonsMaxFontSize As Long
    Dim bgControl As MSForms.Control
    Dim lblControl As MSForms.Control
    Dim HeightOffset As Double
    Dim i As Long
    Dim j As Long
    
    BorderSpacing = 6 * RatioToResize
    HeaderDefaultFontSize = 11
    bgHeaderDefaultHeight = 30
    lblMonthYearDefaultHeight = 13.5
    scrlMonthDefaultHeight = 18
    bgDayLabelsDefaultHeight = 18
    bgDateDefaultHeight = 18
    bgDateDefaultWidth = 18
    lblDateDefaultHeight = 10.5
    cmdButtonDefaultHeight = 24
    cmdButtonDefaultWidth = 60
    cmdButtonsMaxHeight = 36
    cmdButtonsMaxWidth = 90
    cmdButtonsMaxFontSize = 14

    
    If MinimumDate <= 0 Then
        MinDate = CDate("1/1/1900")
    Else
        MinDate = MinimumDate
    End If
    If MaximumDate = 0 Then
        MaxDate = CDate("12/31/9999")
    Else
        MaxDate = MaximumDate
    End If
    If MaxDate < MinDate Then MaxDate = MinDate
    
    If Date < MinDate Or Date > MaxDate Then TodayEnabled = False

    If PositionTop <> -5 And PositionLeft <> -5 Then
        Me.StartUpPosition = 0
        Me.Top = PositionTop
        Me.Left = PositionLeft
    Else
        Me.StartUpPosition = 1
    End If
    
    With bgHeader
        .Height = bgHeaderDefaultHeight * RatioToResize
        If bWeekNumbers Then
            .Width = 8 * (bgDateDefaultWidth * RatioToResize) + BorderSpacing
        Else
            .Width = 7 * (bgDateDefaultWidth * RatioToResize)
        End If
        .Left = BorderSpacing
        .Top = BorderSpacing
    End With
    With scrlMonth
        .Width = bgHeader.Width - (2 * BorderSpacing)
        .Left = bgHeader.Left + BorderSpacing
        .Height = scrlMonthDefaultHeight * RatioToResize
        If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
        .Top = bgHeader.Top + ((bgHeader.Height - .Height) / 2)
    End With
    With bgScrollCover
        .Height = scrlMonth.Height
        .Width = scrlMonth.Width - 25
                                     
        .Left = scrlMonth.Left + 12.5
        .Top = scrlMonth.Top
    End With
    With lblMonth
        .AutoSize = False
        .Height = lblMonthYearDefaultHeight * RatioToResize
        .Font.size = HeaderDefaultFontSize * RatioToResize
        .Top = bgScrollCover.Top + ((bgScrollCover.Height - .Height) / 2)
    End With
    With lblYear
        .AutoSize = False
        .Height = lblMonthYearDefaultHeight * RatioToResize
        .Font.size = HeaderDefaultFontSize * RatioToResize
        .Top = bgScrollCover.Top + ((bgScrollCover.Height - .Height) / 2)
    End With
    cmbMonth.Top = lblMonth.Top + (lblMonth.Height - cmbMonth.Height)
    cmbYear.Top = lblYear.Top + (lblYear.Height - cmbYear.Height)

    With bgDayLabels
        .Height = bgDayLabelsDefaultHeight * RatioToResize
        If bWeekNumbers Then
            .Width = 8 * (bgDateDefaultWidth * RatioToResize) + BorderSpacing
        Else
            .Width = 7 * (bgDateDefaultWidth * RatioToResize)
        End If
        .Left = BorderSpacing
        .Top = bgHeader.Top + bgHeader.Height
    End With
    If Not bWeekNumbers Then
        lblWk.Visible = False
    Else
        With lblWk
            .AutoSize = False
            .Height = lblDateDefaultHeight * RatioToResize
            .Font.size = SizeFont
            .Width = bgDateDefaultWidth * RatioToResize
            .Top = bgDayLabels.Top + ((bgDayLabels.Height - .Height) / 2)
            .Left = BorderSpacing
        End With
    End If
    For i = 1 To 7
        With Me("lblDay" & CStr(i))
            .AutoSize = False
            .Height = lblDateDefaultHeight * RatioToResize
            .Font.size = SizeFont
            .Width = bgDateDefaultWidth * RatioToResize
            .Top = bgDayLabels.Top + ((bgDayLabels.Height - .Height) / 2)
            If i = 1 Then
                If bWeekNumbers Then
                    .Left = lblWk.Left + lblWk.Width + BorderSpacing
                Else
                    .Left = BorderSpacing
                End If
            Else
                .Left = Me("lblDay" & CStr(i - 1)).Left + Me("lblDay" & CStr(i - 1)).Width
            End If
        End With
    Next i
        
    For i = 1 To 6
        If Not bWeekNumbers Then
            Me("lblWeek" & CStr(i)).Visible = False
        Else
            With Me("lblWeek" & CStr(i))
                .AutoSize = False
                .Height = lblDateDefaultHeight * RatioToResize
                .Font.size = SizeFont
                .Width = bgDateDefaultWidth * RatioToResize
                .Left = BorderSpacing
                If i = 1 Then
                    .Top = bgDayLabels.Top + bgDayLabels.Height + (((bgDateDefaultHeight * RatioToResize) - .Height) / 2)
                Else
                    .Top = Me("bgDate" & CStr(i - 1) & "1").Top + Me("bgDate" & CStr(i - 1) & "1").Height + (((bgDateDefaultHeight * RatioToResize) - .Height) / 2)
                End If
            End With
        End If
                
        For j = 1 To 7
            Set bgControl = Me("bgDate" & CStr(i) & CStr(j))
            Set lblControl = Me("lblDate" & CStr(i) & CStr(j))
            With bgControl
                .Height = bgDateDefaultHeight * RatioToResize
                .Width = bgDateDefaultWidth * RatioToResize
                If j = 1 Then
                    
                    If bWeekNumbers Then
                        .Left = Me("lblWeek" & CStr(i)).Left + Me("lblWeek" & CStr(i)).Width + BorderSpacing
                    Else
                        .Left = BorderSpacing
                    End If
                Else
                    .Left = Me("bgDate" & CStr(i) & CStr(j - 1)).Left + Me("bgDate" & CStr(i) & CStr(j - 1)).Width
                End If
                If i = 1 Then
                    .Top = bgDayLabels.Top + bgDayLabels.Height
                Else
                    .Top = Me("bgDate" & CStr(i - 1) & CStr(j)).Top + Me("bgDate" & CStr(i - 1) & CStr(j)).Height
                End If
            End With
            
            With lblControl
                .AutoSize = False
                .Height = lblDateDefaultHeight * RatioToResize
                .Font.size = SizeFont
                .Width = bgControl.Width
                .Left = bgControl.Left
                .Top = bgControl.Top + ((bgControl.Height - .Height) / 2)
            End With
        Next j
    Next i
    
    
    frameCalendar.Width = bgDate67.Left + bgDate67.Width + BorderSpacing
    
    If Me.InsideWidth < (frameCalendar.Left + frameCalendar.Width) Then
        Me.Width = Me.Width + ((frameCalendar.Left + frameCalendar.Width) - Me.InsideWidth)
    End If
    
    If Not OkayEnabled Then
        cmdOkay.Visible = False
        lblSelection.Visible = False
        lblSelectionDate.Visible = False
    Else
        With cmdOkay
            .Visible = True
            .Height = cmdButtonDefaultHeight * RatioToResize
            If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
            .Width = cmdButtonDefaultWidth * RatioToResize
            If .Width > cmdButtonsMaxWidth Then .Width = cmdButtonsMaxWidth
            If SizeFont > cmdButtonsMaxFontSize Then
                .Font.size = cmdButtonsMaxFontSize
            Else
                .Font.size = SizeFont
            End If
            .Top = bgDate61.Top + bgDate61.Height + bgDayLabels.Height + BorderSpacing
        End With
        
        With lblSelection
            .Visible = True
            .AutoSize = False
            .Height = lblMonthYearDefaultHeight * RatioToResize
            .Width = frameCalendar.Width
            .Font.size = HeaderDefaultFontSize * RatioToResize
            .AutoSize = True
            .Top = (bgDate61.Top + bgDate61.Height) + ((bgDayLabels.Height + BorderSpacing - .Height) / 2)
        End With
        
        With lblSelectionDate
            .Visible = True
            .AutoSize = False
            .Height = lblMonthYearDefaultHeight * RatioToResize
            .Width = frameCalendar.Width - lblSelection.Width
            .Font.size = HeaderDefaultFontSize * RatioToResize
            .Top = lblSelection.Top
        End With
    End If
    
    If Not TodayEnabled Then
        cmdToday.Visible = False
    Else
        With cmdToday
            .Visible = True
            .Height = cmdButtonDefaultHeight * RatioToResize
            If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
            .Width = cmdButtonDefaultWidth * RatioToResize
            If .Width > cmdButtonsMaxWidth Then .Width = cmdButtonsMaxWidth
            If SizeFont > cmdButtonsMaxFontSize Then
                .Font.size = cmdButtonsMaxFontSize
            Else
                .Font.size = SizeFont
            End If
        End With
    End If
    
    If OkayEnabled And TodayEnabled Then
        cmdToday.Top = cmdOkay.Top
        cmdButtonsCombinedWidth = cmdToday.Width + cmdOkay.Width
        cmdToday.Left = ((frameCalendar.Width - cmdButtonsCombinedWidth) / 2) - (BorderSpacing / 2)
        cmdOkay.Left = cmdToday.Left + cmdToday.Width + BorderSpacing
    ElseIf OkayEnabled Then
        cmdOkay.Left = (frameCalendar.Width - cmdOkay.Width) / 2
    ElseIf TodayEnabled Then
        cmdToday.Top = bgDate61.Top + bgDate61.Height + BorderSpacing
        cmdToday.Left = (frameCalendar.Width - cmdToday.Width) / 2
    End If
        
    HeightOffset = Me.Height - Me.InsideHeight
    If OkayEnabled Then
        frameCalendar.Height = cmdOkay.Top + cmdOkay.Height + HeightOffset + BorderSpacing
    ElseIf TodayEnabled Then
        frameCalendar.Height = cmdToday.Top + cmdToday.Height + HeightOffset + BorderSpacing
    Else
        frameCalendar.Height = bgDate61.Top + bgDate61.Height + HeightOffset + BorderSpacing
    End If
    
    If Me.InsideHeight < (frameCalendar.Top + frameCalendar.Height) Then
        Me.Height = Me.Height + ((frameCalendar.Top + frameCalendar.Height) - Me.InsideHeight - HeightOffset)
    End If
    
    If SelectedDate > 0 Then
        If SelectedDate < MinDate Then
            SelectedDate = MinDate
        ElseIf SelectedDate > MaxDate Then
            SelectedDate = MaxDate
        End If
        SelectedDateIn = SelectedDate
        SelectedYear = Year(SelectedDateIn)
        SelectedMonth = Month(SelectedDateIn)
        SelectedDay = Day(SelectedDateIn)
        Call SetSelectionLabel(SelectedDateIn)
    Else
        cmdOkay.Enabled = False
        TempDate = Date
        If TempDate < MinDate Then
            TempDate = MinDate
        ElseIf TempDate > MaxDate Then
            TempDate = MaxDate
        End If
        SelectedYear = Year(TempDate)
        SelectedMonth = Month(TempDate)
        SelectedDay = 0
        Call SetSelectionLabel(Empty)
    End If
    
    Call SetMonthCombobox(SelectedYear, SelectedMonth)
    scrlMonth.value = SelectedMonth
    cmbYearMin = SelectedYear - RangeOfYears
    cmbYearMax = SelectedYear + RangeOfYears
    If cmbYearMin < Year(MinDate) Then
        cmbYearMin = Year(MinDate)
    End If
    If cmbYearMax > Year(MaxDate) Then
        cmbYearMax = Year(MaxDate)
    End If
    For i = cmbYearMin To cmbYearMax
        cmbYear.AddItem i
    Next i
    cmbYear.value = SelectedYear
    
    Me.BackColor = BackgroundColor
    frameCalendar.BackColor = BackgroundColor
    bgHeader.BackColor = HeaderColor
    bgScrollCover.BackColor = HeaderColor
    lblMonth.ForeColor = HeaderFontColor
    lblYear.ForeColor = HeaderFontColor
    lblSelection.ForeColor = SubHeaderFontColor
    lblSelectionDate.ForeColor = SubHeaderFontColor
    bgDayLabels.BackColor = SubHeaderColor
    For i = 1 To 7
        Me("lblDay" & CStr(i)).ForeColor = SubHeaderFontColor
    Next i
    If bWeekNumbers Then
        lblWk.ForeColor = SubHeaderFontColor
        For i = 1 To 6
            Me("lblWeek" & CStr(i)).ForeColor = SubHeaderFontColor
        Next i
    End If
    For i = 1 To 6
        For j = 1 To 7
            With Me("bgDate" & CStr(i) & CStr(j))
                If DateBorder Then
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DateBorderColor
                End If
                .SpecialEffect = DateSpecialEffect
            End With
        Next j
    Next i
    
    TempDayOfWeek = StartWeek
    For i = 1 To 7
        Me("lblDay" & CStr(i)).Caption = Choose(TempDayOfWeek, "Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")
        TempDayOfWeek = TempDayOfWeek + 1
        If TempDayOfWeek = 8 Then TempDayOfWeek = 1
    Next i
            
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
End Sub

Private Sub cmdOkay_Click()
    DateOut = SelectedDateIn
    Me.Hide
End Sub

Private Sub cmdToday_Click()
    Dim SelectedMonth As Long
    Dim SelectedYear As Long
    Dim SelectedDay As Long
    Dim TodayDate As Date
    
    UserformEventsEnabled = False
    SelectedDay = 0
    TodayDate = Date
    
    If OkayEnabled Then
        cmdOkay.Enabled = True
        SelectedDateIn = TodayDate
        Call SetSelectionLabel(TodayDate)
        SelectedDay = Day(TodayDate)
    End If
    
    SelectedMonth = Month(TodayDate)
    SelectedYear = Year(TodayDate)
    SelectedDay = GetSelectedDay(SelectedMonth, SelectedYear)
    scrlMonth.value = SelectedMonth
    
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    
    UserformEventsEnabled = True
End Sub


Private Sub ClickControl(ctrl As MSForms.Control)
    Dim SelectedMonth As Long
    Dim SelectedYear As Long
    Dim SelectedDay As Long
    Dim SelectedDate As Date
    Dim RowIndex As Long
    Dim ColumnIndex As Long
    
    SelectedMonth = scrlMonth.value
    SelectedYear = cmbYear.value
    
    RowIndex = CLng(Left(Right(ctrl.Name, 2), 1))
    ColumnIndex = CLng(Right(ctrl.Name, 1))
    SelectedDay = CLng(ctrl.Caption)
        
    If RowIndex = 1 And SelectedDay > 7 Then
        SelectedMonth = SelectedMonth - 1
    
        If SelectedMonth = 0 Then
            SelectedYear = SelectedYear - 1
            SelectedMonth = 12
        End If
    
    
    ElseIf RowIndex >= 5 And SelectedDay < 20 Then
        SelectedMonth = SelectedMonth + 1
    
        If SelectedMonth = 13 Then
            SelectedYear = SelectedYear + 1
            SelectedMonth = 1
        End If
    End If
    
    SelectedDate = DateSerial(SelectedYear, SelectedMonth, SelectedDay)
        
    If Not OkayEnabled Then
        DateOut = SelectedDate
        Me.Hide
    Else
        UserformEventsEnabled = False
            cmdOkay.Enabled = True
            SelectedDateIn = SelectedDate
            scrlMonth.value = SelectedMonth
            Call SetSelectionLabel(SelectedDate)
            Call SetMonthYear(SelectedMonth, SelectedYear)
            Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
        UserformEventsEnabled = True
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HoverControl
'
' This sub handles the event of hovering over one of the date label controls. Every date
' label has a MouseMove event which passes that label to this sub.
'
' This sub returns the last hovered date label to its original color, sets the currently
' hovered date label to the bgDateHoverColor, and stores its name and original color
' to global variables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HoverControl(ctrl As MSForms.Control)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
    HoverControlName = ctrl.Name
    HoverControlColor = ctrl.BackColor
    ctrl.BackColor = bgDateHoverColor
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' lblMonth_Click / lblYear_Click
'
' The month and year labels in the header have invisible comboboxes behind them. These
' two subs show the combobox drop downs when you click on the labels.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblMonth_Click()
    cmbMonth.DropDown
End Sub
Private Sub lblYear_Click()
    cmbYear.DropDown
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cmbMonth_Change / cmbYear_Change
'
' The month and year comboboxes both call the cmbMonthYearChange sub when the user makes
' a selection. The year combobox also resets the month combobox, in case the user
' selects a year that is limited by a minimum or maximum date, to make sure the month
' combobox doesn't end up with selections that shouldn't be available.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbMonth_Change()
    Call cmbMonthYearChange
End Sub
Private Sub cmbYear_Change()
    If Not UserformEventsEnabled Then Exit Sub
    
    UserformEventsEnabled = False
    Call SetMonthCombobox(cmbYear.value, scrlMonth.value)
    UserformEventsEnabled = True
    
    Call cmbMonthYearChange
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cmbMonthYearChange
'
' This sub handles the user making a selection from either the month or year combobox.
' It gets the selected month and year from the comboboxes, sets the value of the month
' scroll bar to match, and resets the calendar date labels.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbMonthYearChange()
    Dim SelectedMonth As Long           'Month of selected date
    Dim SelectedYear As Long            'Year of selected date
    Dim SelectedDay As Long             'Day of selected date
    
    If Not UserformEventsEnabled Then Exit Sub
    UserformEventsEnabled = False
    

    SelectedYear = cmbYear.value
    If SelectedYear = Year(MinDate) Then
        SelectedMonth = cmbMonth.ListIndex + Month(MinDate)
    Else
        SelectedMonth = cmbMonth.ListIndex + 1
    End If
    
    'Get selected day, set the value of the month scroll bar, and reset all
    'date labels on the userform
    SelectedDay = GetSelectedDay(SelectedMonth, SelectedYear)
    scrlMonth.value = SelectedMonth
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    
    UserformEventsEnabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' scrlMonth_Change
'
' This sub handles the user clicking the scroll bar to increment or decrement the month.
' It checks to keep the month within the bounds set by the minimum or maximum date,
' and resets all the labels of the userform to the new month.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub scrlMonth_Change()
    Dim TempYear As Long        'Temporarily store selected year to test min and max dates
    Dim MinMonth As Long        'Sets lower limit of scroll bar
    Dim MaxMonth As Long        'Sets upper limit of scroll bar
    Dim SelectedMonth As Long   'Month of selected date
    Dim SelectedYear As Long    'Year of selected date
    Dim SelectedDay As Long     'Day of selected date
    
    If Not UserformEventsEnabled Then Exit Sub
    UserformEventsEnabled = False
    
    'Default lower and upper limit of scroll bar to allow full range of months
    MinMonth = 0
    MaxMonth = 13
    
    'If the current year is the min or max year, set min or max months
    TempYear = cmbYear.value
    If TempYear = Year(MinDate) Then MinMonth = Month(MinDate)
    If TempYear = Year(MaxDate) Then MaxMonth = Month(MaxDate)
    
    'Keep scroll bar within range of min and max dates
    If scrlMonth.value < MinMonth Then scrlMonth.value = scrlMonth.value + 1
    If scrlMonth.value > MaxMonth Then scrlMonth.value = scrlMonth.value - 1
    
    'If user goes down one month from January, scroll bar will have value of
    '0. In this case, reset scroll bar back to December and decrement year
    'by 1.
    If scrlMonth.value = 0 Then
        scrlMonth.value = 12
        cmbYear.value = cmbYear.value - 1
        'If new year is outside range of combobox, add it to combobox
        If cmbYear.value < cmbYearMin Then
            cmbYear.AddItem cmbYear.value, 0
            cmbYearMin = cmbYear.value
        End If
        Call SetMonthCombobox(cmbYear.value, scrlMonth.value)
    'If user goes up one month from December, scroll bar will have value of
    '13. Reset to January and increment year.
    ElseIf scrlMonth.value = 13 Then
        scrlMonth.value = 1
        cmbYear.value = cmbYear.value + 1
        'If new year is outside range of combobox, add it to combobox
        If cmbYear.value > cmbYearMax Then
            cmbYear.AddItem cmbYear.value, cmbYear.ListCount
            cmbYearMax = cmbYear.value
        End If
        Call SetMonthCombobox(cmbYear.value, scrlMonth.value)
    End If
    
    'Get selected month, year, and day, and reset all userform labels
    SelectedMonth = scrlMonth.value
    SelectedYear = cmbYear.value
    SelectedDay = GetSelectedDay(SelectedMonth, SelectedYear)
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    
    UserformEventsEnabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetMonthCombobox
'
' This sub clears the list in the month combobox and resets it. This is done every time
' the month changes to make sure the months displayed in the combobox don't ever fall
' outside the bounds set by the minimum or maximum date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMonthCombobox(YearIn As Long, MonthIn As Long)
    Dim YearMinDate As Long             'Year of the minimum date
    Dim YearMaxDate As Long             'Year of the maximum date
    Dim MonthMinDate As Long            'Month of the minimum date
    Dim MonthMaxDate As Long            'Month of the maximum date
    Dim i As Long                       'Used for looping
    
    'Get month and year of minimum and maximum dates and clear combobox
    YearMinDate = Year(MinDate)
    YearMaxDate = Year(MaxDate)
    MonthMinDate = Month(MinDate)
    MonthMaxDate = Month(MaxDate)
    cmbMonth.Clear

    'Both minimum and maximum dates occur in selected year
    If YearIn = YearMinDate And YearIn = YearMaxDate Then
        For i = MonthMinDate To MonthMaxDate
            cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        Next i
        If MonthIn < MonthMinDate Then MonthIn = MonthMinDate
        If MonthIn > MonthMaxDate Then MonthIn = MonthMaxDate
        cmbMonth.ListIndex = MonthIn - MonthMinDate
    
    'Only minimum date occurs in selected year
    ElseIf YearIn = YearMinDate Then
        For i = MonthMinDate To 12
            cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        Next i
        If MonthIn < MonthMinDate Then MonthIn = MonthMinDate
        cmbMonth.ListIndex = MonthIn - MonthMinDate
    
    'Only maximum date occurs in selected year
    ElseIf YearIn = YearMaxDate Then
        For i = 1 To MonthMaxDate
            cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        Next i
        If MonthIn > MonthMaxDate Then MonthIn = MonthMaxDate
        cmbMonth.ListIndex = MonthIn - 1
    
    'No minimum or maximum date in selected year. Add all months to combobox
    Else
        cmbMonth.List = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        cmbMonth.ListIndex = MonthIn - 1
    End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetMonthYear
'
' This sub sets the month and year comboboxes to keep them in sync with any changes
' made to the selected month or year. It also sets the month and year labels in the
' header, and positions them in the center of the month scroll bar.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMonthYear(MonthIn As Long, YearIn As Long)
    Dim ExtraSpace As Double                'Space between month and year labels
    Dim CombinedLabelWidth As Double        'Combined width of both month and year labels
    
    ExtraSpace = 4 * RatioToResize
    
    'Set value of comboboxes
    If YearIn = Year(MinDate) Then
        cmbMonth.ListIndex = MonthIn - Month(MinDate)
    Else
        cmbMonth.ListIndex = MonthIn - 1
    End If
    cmbYear.value = YearIn
    
    'Set labels and position to center of scroll buttons. Labels are first
    'set to the width of the userform to avoid overflow, and then autosized
    'to fit to the text before being centered
    With lblMonth
        .AutoSize = False
        .Width = frameCalendar.Width
        .Caption = cmbMonth.value
        .AutoSize = True
    End With
    With lblYear
        .AutoSize = False
        .Width = frameCalendar.Width
        .Caption = cmbYear.value
        .AutoSize = True
    End With
    
    'Get combined width of labels and center to scroll bar
    CombinedLabelWidth = lblMonth.Width + lblYear.Width
    With lblMonth
        .Left = ((frameCalendar.Width - CombinedLabelWidth) / 2) - (ExtraSpace / 2)
    End With
    With lblYear
        .Left = lblMonth.Left + lblMonth.Width + ExtraSpace
    End With
    
    'Reposition comboboxes to line up with labels
    cmbMonth.Left = lblMonth.Left - (cmbMonth.Width - lblMonth.Width) - ExtraSpace - 2
    cmbYear.Left = lblYear.Left
    
    'Clear hover control name, so labels in new month don't revert to
    'colors from previously selected month
    HoverControlName = vbNullString
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetDays
'
' This sub sets the caption, visibility, and colors of all the date labels on the
' userform, as well as the week number labels. If a selected day is passed to the
' sub, it will highlight that date accordingly. Otherwise, no selected date will be
' highlighted.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetDays(MonthIn As Long, YearIn As Long, Optional DayIn As Long)
    Dim PrevMonth As Long               'Month preceding selected month. Used for trailing dates
    Dim NextMonth As Long               'Month following selected month. Used for trailing dates
    Dim Today As Date                   'Today's date
    Dim TodayDay As Long                'Day number of today's date
    Dim StartDayOfWeek  As Long         'Stores the weekday number of the first day in selected month
    Dim LastDayOfMonth As Long          'Last day of the month
    Dim LastDayOfPrevMonth As Long      'Last day of preceding month. Used for trailing dates
    Dim CurrentDay As Long              'Tracks current day in the month while setting labels
    Dim TempCurrentDay As Long          'Tracks the current day for previous month without incrementing actual CurrentDay
    Dim WeekNumber As Long              'Stores week number for week number labels
    Dim StartDayOfWeekDate As Date      'Stores first date in the week. Used to calculate week numbers
    Dim SaturdayIndex As Long           'Column index of Saturdays. Used to set color of Saturday labels, if applicable
    Dim SundayIndex As Long             'Column index of Sundays
    Dim MinDay As Long                  'Stores lower limit of days if minimum date falls in selected month
    Dim MaxDay As Long                  'Stores upper limit of days if maximum date falls in selected month
    Dim PrevMonthMinDay As Long         'Stores lower limit of days if minimum date falls in preceding month
    Dim NextMonthMaxDay As Long         'Stores upper limit of days if maximum date falls in next month
    Dim lblControl As MSForms.Control   'Stores current date label while changing settings
    Dim bgControl As MSForms.Control    'Stores current date label background while changing settings
    Dim i As Long                       'Used for looping
    Dim j As Long                       'Used for looping
    
    'Set min and max day, if applicable. If not, min and max day are set to 0 and 32,
    'respectively, since dates will never fall outside those bounds
    MinDay = 0
    MaxDay = 32
    If YearIn = Year(MinDate) And MonthIn = Month(MinDate) Then MinDay = Day(MinDate)
    If YearIn = Year(MaxDate) And MonthIn = Month(MaxDate) Then MaxDay = Day(MaxDate)
    
    'Find previous month and next month. Handle January
    'and December appropriately
    PrevMonth = MonthIn - 1
    If PrevMonth = 0 Then PrevMonth = 12
    NextMonth = MonthIn + 1
    If NextMonth = 13 Then NextMonth = 1
    
    'Set min and max days for previous month and next month, if applicable
    PrevMonthMinDay = 0
    NextMonthMaxDay = 32
    If YearIn = Year(MinDate) And PrevMonth = Month(MinDate) Then PrevMonthMinDay = Day(MinDate)
    If YearIn = Year(MaxDate) And NextMonth = Month(MaxDate) Then NextMonthMaxDay = Day(MaxDate)

    'Find last day of selected month and previous month. Find first weekday
    'in current month, and index of Saturday and Sunday relative to first weekday
    LastDayOfMonth = Day(DateSerial(YearIn, MonthIn + 1, 0))
    LastDayOfPrevMonth = Day(DateSerial(YearIn, MonthIn, 0))
    StartDayOfWeek = Weekday(DateSerial(YearIn, MonthIn, 1), StartWeek)
    If StartWeek = 1 Then SundayIndex = 1 Else SundayIndex = 9 - StartWeek
    SaturdayIndex = 8 - StartWeek

    'If user is viewing current month/year, we want to highlight today's date. If
    'not, TodayDay is set to 0, since that value will never be encountered
    Today = Date
    If YearIn = Year(Today) And MonthIn = Month(Today) Then
        TodayDay = Day(Today)
    Else
        TodayDay = 0
    End If
    
    CurrentDay = 1
    For i = 1 To 6
    
        If StartDayOfWeek = 1 And i = 1 Then
            TempCurrentDay = CLng(LastDayOfPrevMonth - (StartDayOfWeek + 5))
            If PrevMonth <> 12 Then
                StartDayOfWeekDate = DateSerial(YearIn, PrevMonth, TempCurrentDay)
            Else
                StartDayOfWeekDate = DateSerial(YearIn - 1, PrevMonth, TempCurrentDay)
            End If
            
        ElseIf i = 1 Then
            StartDayOfWeekDate = DateSerial(YearIn, MonthIn, 1)
        
        Else
            If CurrentDay <= LastDayOfMonth Then
                TempCurrentDay = CurrentDay
                StartDayOfWeekDate = DateSerial(YearIn, MonthIn, TempCurrentDay)
            
            Else
                TempCurrentDay = CLng(CurrentDay - LastDayOfMonth)
                If NextMonth <> 1 Then
                    StartDayOfWeekDate = DateSerial(YearIn, NextMonth, TempCurrentDay)
                Else
                    StartDayOfWeekDate = DateSerial(YearIn + 1, NextMonth, TempCurrentDay)
                End If
            End If
        End If
        WeekNumber = DatePart("ww", StartDayOfWeekDate, StartWeek, WeekOneOfYear)
        
        If WeekNumber > 52 And TempCurrentDay > 25 Then
            WeekNumber = DatePart("ww", DateSerial(YearIn + 1, 1, 1), StartWeek, WeekOneOfYear)
        End If
        Me("lblWeek" & CStr(i)).Caption = WeekNumber
        
        For j = 1 To 7
            Set lblControl = Me("lblDate" & CStr(i) & CStr(j))
            Set bgControl = Me("bgDate" & CStr(i) & CStr(j))
            With lblControl
                
                If StartDayOfWeek = 1 And i = 1 Then
                    If MinDay <> 0 Then
                        .Visible = False
                        bgControl.Visible = False
                    Else
                        TempCurrentDay = CLng(LastDayOfPrevMonth - (StartDayOfWeek + 6 - j))
                        If TempCurrentDay < PrevMonthMinDay Then
                            .Visible = False
                            bgControl.Visible = False
                        Else
                            .Visible = True
                            bgControl.Visible = True
                            .ForeColor = lblDatePrevMonthColor
                            .Caption = CStr(TempCurrentDay)
                            bgControl.BackColor = bgDateColor
                        End If
                    End If
                    
                ElseIf i = 1 And j < StartDayOfWeek Then
                    If MinDay <> 0 Then
                        .Visible = False
                        bgControl.Visible = False
                    Else
                        TempCurrentDay = CLng(LastDayOfPrevMonth - (StartDayOfWeek - 1 - j))
                        If TempCurrentDay < PrevMonthMinDay Then
                            .Visible = False
                            bgControl.Visible = False
                        Else
                            .Visible = True
                            .Enabled = True
                            bgControl.Visible = True
                            .ForeColor = lblDatePrevMonthColor
                            .Caption = CStr(TempCurrentDay)
                            bgControl.BackColor = bgDateColor
                        End If
                    End If

                ElseIf CurrentDay > LastDayOfMonth Then
                    If MaxDay <> 32 Then
                        .Visible = False
                        bgControl.Visible = False
                    Else
                        TempCurrentDay = CLng(CurrentDay - LastDayOfMonth)
                        If TempCurrentDay > NextMonthMaxDay Then
                            .Visible = False
                            bgControl.Visible = False
                        Else
                            .Visible = True
                            .Enabled = True
                            bgControl.Visible = True
                            .ForeColor = lblDatePrevMonthColor
                            .Caption = CStr(TempCurrentDay)
                            bgControl.BackColor = bgDateColor
                        End If
                    End If
                    CurrentDay = CurrentDay + 1
                    
                Else
                    If CurrentDay < MinDay Or CurrentDay > MaxDay Then
                        .Visible = True
                        .Enabled = False
                        bgControl.Visible = False
                    Else
                        .Visible = True
                        .Enabled = True
                        bgControl.Visible = True
                        
                        If CurrentDay = TodayDay Then
                            .ForeColor = lblDateTodayColor
                        ElseIf j = SaturdayIndex Then
                            .ForeColor = lblDateSatColor
                        ElseIf j = SundayIndex Then
                            .ForeColor = lblDateSunColor
                        Else
                            .ForeColor = lblDateColor
                        End If
                        
                        If CurrentDay = DayIn Then
                            bgControl.BackColor = bgDateSelectedColor
                        Else
                            bgControl.BackColor = bgDateColor
                        End If
                    End If
                    .Caption = CStr(CurrentDay)
                    CurrentDay = CurrentDay + 1
                End If
            End With
        Next j
    Next i
End Sub


Private Sub SetSelectionLabel(DateIn As Date)
    Dim CombinedLabelWidth As Double
    Dim ExtraSpace As Double
    
    ExtraSpace = 3 * RatioToResize
    
    If DateIn = 0 Then
        lblSelectionDate.Caption = vbNullString
        lblSelection.Left = frameCalendar.Left + ((frameCalendar.Width - lblSelection.Width) / 2)
    Else
        With lblSelectionDate
            .AutoSize = False
            .Width = frameCalendar.Width
            .Caption = Format(DateIn, "mm/dd/yyyy")
            .AutoSize = True
        End With
    
        CombinedLabelWidth = lblSelection.Width + lblSelectionDate.Width
        lblSelection.Left = ((frameCalendar.Width - CombinedLabelWidth) / 2) - (ExtraSpace / 2)
        lblSelectionDate.Left = lblSelection.Left + lblSelection.Width + ExtraSpace
    End If
End Sub


Private Function GetSelectedDay(MonthIn As Long, YearIn As Long) As Long
    GetSelectedDay = 0
    
    If SelectedDateIn <> 0 Then
        If MonthIn = Month(SelectedDateIn) And YearIn = Year(SelectedDateIn) Then
            GetSelectedDay = Day(SelectedDateIn)
        End If
    End If
End Function

Private Function Min(ParamArray values() As Variant) As Variant
   Dim minValue As Variant
   Dim value As Variant
   minValue = values(0)
   For Each value In values
       If value < minValue Then minValue = value
   Next
   Min = minValue
End Function
Private Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue As Variant
   Dim value As Variant
   maxValue = values(0)
   For Each value In values
       If value > maxValue Then maxValue = value
   Next
   Max = maxValue
End Function

Private Sub bgDate11_Click(): ClickControl lblDate11: End Sub
Private Sub bgDate12_Click(): ClickControl lblDate12: End Sub
Private Sub bgDate13_Click(): ClickControl lblDate13: End Sub
Private Sub bgDate14_Click(): ClickControl lblDate14: End Sub
Private Sub bgDate15_Click(): ClickControl lblDate15: End Sub
Private Sub bgDate16_Click(): ClickControl lblDate16: End Sub
Private Sub bgDate17_Click(): ClickControl lblDate17: End Sub
Private Sub bgDate21_Click(): ClickControl lblDate21: End Sub
Private Sub bgDate22_Click(): ClickControl lblDate22: End Sub
Private Sub bgDate23_Click(): ClickControl lblDate23: End Sub
Private Sub bgDate24_Click(): ClickControl lblDate24: End Sub
Private Sub bgDate25_Click(): ClickControl lblDate25: End Sub
Private Sub bgDate26_Click(): ClickControl lblDate26: End Sub
Private Sub bgDate27_Click(): ClickControl lblDate27: End Sub
Private Sub bgDate31_Click(): ClickControl lblDate31: End Sub
Private Sub bgDate32_Click(): ClickControl lblDate32: End Sub
Private Sub bgDate33_Click(): ClickControl lblDate33: End Sub
Private Sub bgDate34_Click(): ClickControl lblDate34: End Sub
Private Sub bgDate35_Click(): ClickControl lblDate35: End Sub
Private Sub bgDate36_Click(): ClickControl lblDate36: End Sub
Private Sub bgDate37_Click(): ClickControl lblDate37: End Sub
Private Sub bgDate41_Click(): ClickControl lblDate41: End Sub
Private Sub bgDate42_Click(): ClickControl lblDate42: End Sub
Private Sub bgDate43_Click(): ClickControl lblDate43: End Sub
Private Sub bgDate44_Click(): ClickControl lblDate44: End Sub
Private Sub bgDate45_Click(): ClickControl lblDate45: End Sub
Private Sub bgDate46_Click(): ClickControl lblDate46: End Sub
Private Sub bgDate47_Click(): ClickControl lblDate47: End Sub
Private Sub bgDate51_Click(): ClickControl lblDate51: End Sub
Private Sub bgDate52_Click(): ClickControl lblDate52: End Sub
Private Sub bgDate53_Click(): ClickControl lblDate53: End Sub
Private Sub bgDate54_Click(): ClickControl lblDate54: End Sub
Private Sub bgDate55_Click(): ClickControl lblDate55: End Sub
Private Sub bgDate56_Click(): ClickControl lblDate56: End Sub
Private Sub bgDate57_Click(): ClickControl lblDate57: End Sub
Private Sub bgDate61_Click(): ClickControl lblDate61: End Sub
Private Sub bgDate62_Click(): ClickControl lblDate62: End Sub
Private Sub bgDate63_Click(): ClickControl lblDate63: End Sub
Private Sub bgDate64_Click(): ClickControl lblDate64: End Sub
Private Sub bgDate65_Click(): ClickControl lblDate65: End Sub
Private Sub bgDate66_Click(): ClickControl lblDate66: End Sub
Private Sub bgDate67_Click(): ClickControl lblDate67: End Sub

Private Sub lblDate11_Click(): ClickControl lblDate11: End Sub
Private Sub lblDate12_Click(): ClickControl lblDate12: End Sub
Private Sub lblDate13_Click(): ClickControl lblDate13: End Sub
Private Sub lblDate14_Click(): ClickControl lblDate14: End Sub
Private Sub lblDate15_Click(): ClickControl lblDate15: End Sub
Private Sub lblDate16_Click(): ClickControl lblDate16: End Sub
Private Sub lblDate17_Click(): ClickControl lblDate17: End Sub
Private Sub lblDate21_Click(): ClickControl lblDate21: End Sub
Private Sub lblDate22_Click(): ClickControl lblDate22: End Sub
Private Sub lblDate23_Click(): ClickControl lblDate23: End Sub
Private Sub lblDate24_Click(): ClickControl lblDate24: End Sub
Private Sub lblDate25_Click(): ClickControl lblDate25: End Sub
Private Sub lblDate26_Click(): ClickControl lblDate26: End Sub
Private Sub lblDate27_Click(): ClickControl lblDate27: End Sub
Private Sub lblDate31_Click(): ClickControl lblDate31: End Sub
Private Sub lblDate32_Click(): ClickControl lblDate32: End Sub
Private Sub lblDate33_Click(): ClickControl lblDate33: End Sub
Private Sub lblDate34_Click(): ClickControl lblDate34: End Sub
Private Sub lblDate35_Click(): ClickControl lblDate35: End Sub
Private Sub lblDate36_Click(): ClickControl lblDate36: End Sub
Private Sub lblDate37_Click(): ClickControl lblDate37: End Sub
Private Sub lblDate41_Click(): ClickControl lblDate41: End Sub
Private Sub lblDate42_Click(): ClickControl lblDate42: End Sub
Private Sub lblDate43_Click(): ClickControl lblDate43: End Sub
Private Sub lblDate44_Click(): ClickControl lblDate44: End Sub
Private Sub lblDate45_Click(): ClickControl lblDate45: End Sub
Private Sub lblDate46_Click(): ClickControl lblDate46: End Sub
Private Sub lblDate47_Click(): ClickControl lblDate47: End Sub
Private Sub lblDate51_Click(): ClickControl lblDate51: End Sub
Private Sub lblDate52_Click(): ClickControl lblDate52: End Sub
Private Sub lblDate53_Click(): ClickControl lblDate53: End Sub
Private Sub lblDate54_Click(): ClickControl lblDate54: End Sub
Private Sub lblDate55_Click(): ClickControl lblDate55: End Sub
Private Sub lblDate56_Click(): ClickControl lblDate56: End Sub
Private Sub lblDate57_Click(): ClickControl lblDate57: End Sub
Private Sub lblDate61_Click(): ClickControl lblDate61: End Sub
Private Sub lblDate62_Click(): ClickControl lblDate62: End Sub
Private Sub lblDate63_Click(): ClickControl lblDate63: End Sub
Private Sub lblDate64_Click(): ClickControl lblDate64: End Sub
Private Sub lblDate65_Click(): ClickControl lblDate65: End Sub
Private Sub lblDate66_Click(): ClickControl lblDate66: End Sub
Private Sub lblDate67_Click(): ClickControl lblDate67: End Sub


Private Sub bgDate11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate11: End Sub
Private Sub bgDate12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate12: End Sub
Private Sub bgDate13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate13: End Sub
Private Sub bgDate14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate14: End Sub
Private Sub bgDate15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate15: End Sub
Private Sub bgDate16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate16: End Sub
Private Sub bgDate17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate17: End Sub
Private Sub bgDate21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate21: End Sub
Private Sub bgDate22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate22: End Sub
Private Sub bgDate23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate23: End Sub
Private Sub bgDate24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate24: End Sub
Private Sub bgDate25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate25: End Sub
Private Sub bgDate26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate26: End Sub
Private Sub bgDate27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate27: End Sub
Private Sub bgDate31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate31: End Sub
Private Sub bgDate32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate32: End Sub
Private Sub bgDate33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate33: End Sub
Private Sub bgDate34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate34: End Sub
Private Sub bgDate35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate35: End Sub
Private Sub bgDate36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate36: End Sub
Private Sub bgDate37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate37: End Sub
Private Sub bgDate41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate41: End Sub
Private Sub bgDate42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate42: End Sub
Private Sub bgDate43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate43: End Sub
Private Sub bgDate44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate44: End Sub
Private Sub bgDate45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate45: End Sub
Private Sub bgDate46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate46: End Sub
Private Sub bgDate47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate47: End Sub
Private Sub bgDate51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate51: End Sub
Private Sub bgDate52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate52: End Sub
Private Sub bgDate53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate53: End Sub
Private Sub bgDate54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate54: End Sub
Private Sub bgDate55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate55: End Sub
Private Sub bgDate56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate56: End Sub
Private Sub bgDate57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate57: End Sub
Private Sub bgDate61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate61: End Sub
Private Sub bgDate62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate62: End Sub
Private Sub bgDate63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate63: End Sub
Private Sub bgDate64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate64: End Sub
Private Sub bgDate65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate65: End Sub
Private Sub bgDate66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate66: End Sub
Private Sub bgDate67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate67: End Sub

Private Sub lblDate11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate11: End Sub
Private Sub lblDate12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate12: End Sub
Private Sub lblDate13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate13: End Sub
Private Sub lblDate14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate14: End Sub
Private Sub lblDate15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate15: End Sub
Private Sub lblDate16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate16: End Sub
Private Sub lblDate17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate17: End Sub
Private Sub lblDate21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate21: End Sub
Private Sub lblDate22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate22: End Sub
Private Sub lblDate23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate23: End Sub
Private Sub lblDate24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate24: End Sub
Private Sub lblDate25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate25: End Sub
Private Sub lblDate26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate26: End Sub
Private Sub lblDate27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate27: End Sub
Private Sub lblDate31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate31: End Sub
Private Sub lblDate32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate32: End Sub
Private Sub lblDate33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate33: End Sub
Private Sub lblDate34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate34: End Sub
Private Sub lblDate35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate35: End Sub
Private Sub lblDate36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate36: End Sub
Private Sub lblDate37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate37: End Sub
Private Sub lblDate41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate41: End Sub
Private Sub lblDate42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate42: End Sub
Private Sub lblDate43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate43: End Sub
Private Sub lblDate44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate44: End Sub
Private Sub lblDate45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate45: End Sub
Private Sub lblDate46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate46: End Sub
Private Sub lblDate47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate47: End Sub
Private Sub lblDate51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate51: End Sub
Private Sub lblDate52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate52: End Sub
Private Sub lblDate53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate53: End Sub
Private Sub lblDate54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate54: End Sub
Private Sub lblDate55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate55: End Sub
Private Sub lblDate56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate56: End Sub
Private Sub lblDate57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate57: End Sub
Private Sub lblDate61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate61: End Sub
Private Sub lblDate62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate62: End Sub
Private Sub lblDate63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate63: End Sub
Private Sub lblDate64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate64: End Sub
Private Sub lblDate65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate65: End Sub
Private Sub lblDate66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate66: End Sub
Private Sub lblDate67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate67: End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
End Sub
Private Sub frameCalendar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
End Sub
Private Sub bgDayLabels_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
End Sub
