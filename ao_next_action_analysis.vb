Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module ao_next_action_analysis_module
    'ao_next_action_analysis 25Jan11 http://creativecommons.org/licenses/by-nc-nd/3.0/
    '"Next Action Analysis" (TM) ActivityOwner.com
    '
    '#uses "ao_common.mmbas"
    'recent changes
    'add overdue project information
    'error trap hyperlink follow
    'avoid identifying completed tasks as overdue
    'break out common functions to ao_common.mmbas
    'add configuration map and version checking
    'trap error
    'try to avoid naalog overwrite
    'getmap fixes
    'avoid crash on missint callout topic
    'avoid divide by 0 issues and missing callout topic issue (finally?)
    'change default to not show web advice
    'trap error on most work report
    'report number of tasks in biggest context
    'change weighting to favor fixing the oldest task
    '19Apr change version check strategy
    '02May09 Add a metric for setting time estimates for tasks, trim off empty items, add 10/10 advice
    '03May09 fix version number
    '17May09 change to sigmoidal scoring
    '07Jul09 fix bug related to "unimplemented feature" -- looking for central topic callout
    '07Sep09 tweak scoring to penalize too many and too old daily capture
    '28Feb10 add total time estimate
    '09Mar10 fix bug preventing perfect scores and re-shape scoring curve
    '14Mar10 gather up in-tray items
    '20Mar10 fix bug in "no dates" metric, add support for 2m action
    '21Mar10 improve area search
    '22Mar10 loosen project count target, allow a few no context/time estimate items
    '07Apr10 allow a grace period on no-context item penalty
    '12Apr10 lower weight on project count
    '17Oct10 Don't count projects with start date= end date as overdue (likely due to mm9 task issue)
    '25Jan11 fix bug missing some projects in summary list
    Sub next_action_analysis(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Const ProgramVersion = "20110125"
        Const VersionCheckLink = "http://activityowner.com/installers/versioncheck.php"
        Dim showadvice As Boolean 'set this to false if you don't want to be prompted for advice.
        Dim askadvice As Boolean  'set this to false if you want to show advice automatically without being prompted.
        Dim logprompt As Boolean 'prompt before saving score to log
        Dim checkempty As Boolean
        Const maxactions = 500
        Const maxprojects = 200
        Const maxmaps = 600
        Const maxareas = 30
        Const mapcentralcat = "mc*"
        Const dcmapname = "Daily Capture Map"
        Dim addtolog As Boolean
        Dim numprojects As Integer
        Dim nummainprojects As Integer
        Dim numactions As Integer
        Dim nummaps As Integer
        Dim mostactions As Integer
        Dim mostmap As Integer
        Dim numdatedprojects As Integer
        Dim numnopriorityprojects As Integer
        Dim numdatedactions As Integer
        Dim numdatedintrayactions As Integer
        Dim numnocontext As Integer
        Dim numneednextaction As Integer
        Dim numoverdue As Integer
        Dim numoverdueprojects As Integer
        Dim numpasttarget As Integer
        Dim numoverduewaiting As Integer
        Dim numwaiting As Integer
        Dim numDailyCapture As Integer
        Dim contexttotal As Double
        Dim contextcount As Double
        Dim oldavg As Double

        Dim actions(maxactions) As mm.Topic
        Dim allactions(maxactions) As mm.Topic
        Dim oldestintrayitemxml As String
        Dim oldestintrayage As Integer
        Dim age(maxactions) As Integer
        Dim InTrayAge(maxactions) As Integer
        Dim projectxml(maxprojects) As String
        Dim nextactionxml(maxactions) As String
        Dim nextactioncat(maxactions) As String
        Dim repeating(maxactions) As Boolean
        Dim taskarea(maxactions) As String
        Dim projectarea(maxprojects) As String
        Dim nextactionpriority(maxactions) As Integer
        Dim maplinks(maxmaps) As String
        Dim mapcount(maxmaps) As Integer
        Dim projecttext(maxprojects) As String
        Dim projectpriority(maxprojects) As Integer
        Dim projectdated(maxprojects) As Boolean
        Dim projectoverdue(maxprojects) As Boolean
        Dim actioncount(maxprojects) As Integer
        Dim nextactions(maxactions) As String
        Dim isproject(maxprojects) As Boolean
        Dim isprioritized(maxprojects) As Boolean
        Dim area(maxareas) As String
        Dim numareas As Integer
        Dim cnt As Integer

        'report branches
        Dim reporttopic As mm.Topic
        Dim ratingtopic As mm.Topic
        Dim datedprojects As mm.Topic
        Dim undatedprojects As mm.Topic
        Dim undatedsubprojects As mm.Topic
        Dim deltopic As mm.Topic
        Dim nocontextactions As mm.Topic
        Dim neednextaction As mm.Topic
        Dim oldcontexttopic As mm.Topic
        Dim timecategories As mm.Topic
        Dim taskprioritytopic As mm.Topic
        Dim c2m As mm.Topic
        Dim c15m As mm.Topic
        Dim c1h As mm.Topic
        Dim c2h As mm.Topic
        Dim noest As mm.Topic
        Dim ptopic(5) As mm.Topic

        Dim t As mm.Topic
        Dim tt As mm.Topic
        Dim ttt As mm.Topic
        Dim tttt As mm.Topic
        Dim persontopic As mm.Topic
        Dim swap As mm.Topic
        Dim swapa As Integer
        Dim dashboard As mm.Document
        Dim reportdoc As mm.Document
        Dim NextActionDoc As mm.Document
        Dim NoContextDoc As mm.Document
        Dim NoActionDoc As mm.Document
        Dim DeadlinesDoc As mm.Document
        Dim RelationshipDoc As mm.Document
        Dim projectsdoc As mm.Document
        Dim overduedoc As mm.Document
        Dim LogDoc As mm.Document
        '	Dim IntrayDoc as mm.document

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim avgage As Double
        Dim avInTrayAge As Double
        Dim oldest As Double
        Dim found As Boolean
        Dim pct_complete As Double
        Dim done As Double
        Dim rating As Double
        Const nm = 15
        Dim metric(nm) As Double
        Dim weight(nm) As Double
        Dim goal(nm) As Double
        Dim half(nm) As Double
        Dim mname(nm) As String
        Dim advice(nm) As String
        Dim ratings(nm) As Double
        Dim opportunity(nm) As Double
        Dim hlink(nm) As String
        Dim maxopportunity As Double
        Dim secondopportunity As Double
        Dim thirdopportunity As Double
        Dim firstitem As Integer
        Dim seconditem As Integer
        Dim thirditem As Integer

        Dim FocusTopic As mm.Topic
        Dim FinishTopic As mm.Topic
        Dim FreshnessTopic As mm.Topic
        Dim FeasibilityTopic As mm.Topic
        Dim ForesightTopic As mm.Topic
        Dim AdviceTopic As mm.Topic
        Dim AdviceTopic1 As mm.Topic

        Dim reporttitle As String
        Dim logdocname As String
        Dim configdocname As String
        Dim ConfigDoc As mm.Document
        Dim DueInSeven As Integer
        Dim graceperiod As Integer

        Dim dummy As mm.Topic

        graceperiod = 2
        DueInSeven = 0
        configdocname = "AO\NAAconfig.mmap"
        ConfigDoc = getmap(m_app, configdocname)
        Upgrade3(m_app, ConfigDoc)
        'VersionCheck(VersionCheckLink, "NAA", ProgramVersion, ConfigDoc)
        showadvice = optiontrue("showadvice", ConfigDoc, Nothing)
        askadvice = optiontrue("askadvice", ConfigDoc, Nothing)
        logprompt = optiontrue("logprompt", ConfigDoc, Nothing)
        checkempty = optiontrue("checkempty", ConfigDoc, Nothing)
        logdocname = getoption("logdocname", ConfigDoc, Nothing)
        'get rid of previous report------------------------------
        reporttitle = "NEXT" & Chr(10) & "ACTION" & Chr(10) & "ANALYSIS"
        dashboard = m_app.ActiveDocument
        If dashboard Is Nothing Then Exit Sub
        If Not InStr(LCase(dashboard.CentralTopic.Text), "action") > 0 Then
            MsgBox("Please open your daily action dashboard")
            Exit Sub
        End If
        reportdoc = m_app.Documents.Add(True)
        reportdoc.CentralTopic.Text = "Next Action Analysis In Progress"
        deltopic = reportdoc.CentralTopic.AddSubTopic("Status")
        For Each t In dashboard.CentralTopic.AllSubTopics
            If t.Text = reporttitle Then t.Delete()
        Next
        reporttopic = dashboard.CentralTopic.AddBalancedSubTopic(reporttitle)

        deltopic.AddSubTopic("Creating temporary maps")
        NextActionDoc = m_app.Documents.Add(False)
        NoContextDoc = m_app.Documents.Add(False)
        NoActionDoc = m_app.Documents.Add(False)
        DeadlinesDoc = m_app.Documents.Add(False)
        RelationshipDoc = m_app.Documents.Add(False)
        projectsdoc = m_app.Documents.Add(False)
        overduedoc = m_app.Documents.Add(False)
        'Set IntrayDoc = m_app.documents.Add(False)

        '---------------------------------------------------------
        'It is easier to do stats on branches if they live on their own temporary map
        deltopic.AddSubTopic("Copying branches to temporary maps")
        copybranchcontainingtomap(dashboard.CentralTopic, "committed next actions", NextActionDoc) 'next actions branch
        copybranchcontainingtomap(NextActionDoc.CentralTopic, "no context", NoContextDoc) 'find no context branch
        copybranchcontainingtomap(dashboard.CentralTopic, "needing next actions", NoActionDoc) 'projects needing actions
        copybranchcontainingtomap(dashboard.CentralTopic, "deadlines", DeadlinesDoc) 'find overdue branch
        copybranchcontainingtomap(dashboard.CentralTopic, "waiting", RelationshipDoc) 'find branch with word relationship
        copybranchcontainingtomap(dashboard.CentralTopic, "committed projects by area", projectsdoc)
        copybranchcontainingtomap(dashboard.CentralTopic, "overdue", overduedoc)
        'copybranchcontainingtomap dashboard.CentralTopic,"In-trays",IntrayDoc 'find branch with intrays
        '----------------------------------------------------------
        deltopic.AddSubTopic("Adding up totals")
        numnocontext = TotalActivities(NoContextDoc) 'add up actions in no-context branch
        numneednextaction = TotalActivities(NoActionDoc) 'add up projects in no-next action branch
        numoverdue = TotalActivities(overduedoc) ' add up overdue activities
        numpasttarget = TotalRedActivities(NextActionDoc) - numoverdue ' add up number past target date
        'numoverduewaiting = TotalRedActivitiesWithParentContaining(RelationshipDoc, "waiting") 'add up overdue waiting for task
        numwaiting = TotalActivities(RelationshipDoc) 'add up tasks being waited for
        numprojects = projectsdoc.CentralTopic.SubTopics.Count

        '---------------------------------------------------------
        'compile parent project info
        numdatedprojects = 0
        numnopriorityprojects = 0
        numoverdueprojects = 0
        nummainprojects = 0
        deltopic.AddSubTopic("Reviewing projects/priortities")
        For j = 1 To maxprojects
            actioncount(j) = 0
        Next
        For j = 1 To maxmaps
            mapcount(j) = 0
        Next

     

        For Each t In NextActionDoc.Range(mm.MmRange.mmRangeAllTopics)
            If t.IsCalloutTopic And Not t.Task.Complete = 100 Then
                If Not (t.ParentTopic.IsCalloutTopic And t.ParentTopic.Task.Complete < 100) And Not InStr(t.Xml, mapcentralcat) > 0 And Not InStr(LCase(t.Xml), "in-tray*") > 0 Then  'what if text happens to have in-tray*?
                    found = False
                    For j = 1 To numprojects
                        If projecttext(j) = t.Text Then
                            found = True
                            actioncount(j) = actioncount(j) + 1
                            Exit For
                        End If
                    Next
                    If Not found Then
                        numprojects = numprojects + 1
                        projecttext(numprojects) = t.Text
                        projectarea(numprojects) = getdashboardarea(t)
                        projectxml(numprojects) = t.Xml
                        projectpriority(numprojects) = t.Task.Priority
                        If Not isdate0(t.Task.DueDate) Then projectdated(numprojects) = True Else projectdated(numprojects) = False
                        If Not isdate0(t.Task.DueDate) Or (Not isdate0(t.ParentTopic.Task.DueDate)) Then numdatedprojects = numdatedprojects + 1
                        If (Not isdate0(t.Task.DueDate) And t.Task.DueDate < Now) Then
                            If Not Math.Round(DateDiff(DateInterval.Day, t.Task.DueDate, t.Task.StartDate)) = 0 Then
                                projectoverdue(numprojects) = True
                                numoverdueprojects = numoverdueprojects + 1
                            End If
                        Else
                            projectoverdue(numprojects) = False
                        End If
                        If t.Icons.Item(1).Name = "CustomIcon--1181845906" Then
                            isproject(numprojects) = False
                        Else
                            isproject(numprojects) = True
                            nummainprojects = nummainprojects + 1
                            If t.Task.Priority = 0 Then
                                numnopriorityprojects = numnopriorityprojects + 1 'only count projects
                                isprioritized(numprojects) = False
                            Else
                                isprioritized(numprojects) = True
                            End If
                        End If
                        If Not isdate0(t.Task.DueDate) Then
                            If DateDiff(DateInterval.Day, Today, t.Task.DueDate) < 7 Then 'due - today
                                DueInSeven = DueInSeven + 1
                            End If
                        End If
                    End If
                End If
            End If
            If Not t.IsCalloutTopic Then
                found = False
                If t.HasHyperlink Then
                    For j = 1 To nummaps
                        If maplinks(j) = t.Hyperlink.Address Then
                            found = True
                            mapcount(j) = mapcount(j) + 1
                            Exit For
                        End If
                    Next
                    If Not found Then
                        nummaps = nummaps + 1
                        maplinks(nummaps) = t.Hyperlink.Address
                    End If
                End If
            End If
        Next

        'calculate number of items in daily capture map
        numDailyCapture = 0
        For j = 1 To nummaps
            If InStr(maplinks(j), dcmapname) > 0 Then
                numDailyCapture = mapcount(j)
                Exit For
            End If
        Next

        mostactions = 1
        For j = 1 To numprojects
            If actioncount(j) > actioncount(mostactions) Then mostactions = j
        Next
        mostmap = 1
        For j = 1 To nummaps
            If mapcount(j) > mapcount(mostmap) Then mostmap = j
        Next

        '---------------------------------------------------------
        'calculate percent complete
        done = 0
        numactions = 0
        deltopic.AddSubTopic("Calculating percentage completion")
        For Each t In NextActionDoc.Range(mm.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If Not t.IsCalloutTopic Then
                    found = False
                    For j = 1 To numactions
                        If nextactions(j) = t.Text Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        numactions = numactions + 1
                        nextactions(numactions) = t.Text
                        nextactionxml(numactions) = t.Xml
                        nextactioncat(numactions) = cat(t)
                        nextactionpriority(numactions) = 0
                        taskarea(numactions) = getdashboardarea(t)
                        allactions(numactions) = t

                        If t.Task.Priority > 0 Then
                            nextactionpriority(numactions) = t.Task.Priority
                        ElseIf t.CalloutTopics.Count > 0 Then
                            If t.CalloutTopics.Item(1).Task.Complete < 100 Then
                                If t.CalloutTopics.Item(1).Task.Priority > 0 Then
                                    nextactionpriority(numactions) = t.CalloutTopics.Item(1).Task.Priority
                                ElseIf t.CalloutTopics.Item(1).CalloutTopics.Count > 0 Then
                                    If t.CalloutTopics.Item(1).CalloutTopics.Item(1).Task.Complete < 100 Then
                                        nextactionpriority(numactions) = t.CalloutTopics.Item(1).CalloutTopics.Item(1).Task.Priority
                                    End If
                                End If
                            End If
                        End If
                        repeating(numactions) = isrepeating(t)
                        If t.Task.IsDone Or t.Icons.HasStockIcon(mm.MmStockIcon.mmStockIconCheck) Then
                            done = done + 1
                        ElseIf t.Task.Complete > 0 Then
                            done = done + t.Task.Complete / 100
                        End If
                    End If
                End If
            End If
        Next
        If numactions > 0 Then
            pct_complete = Math.Round(100 * done / numactions, 2)
        Else
            pct_complete = 0
        End If
        '
        'analyze dated tasks
        deltopic.AddSubTopic("Analyzing task aging")
        numdatedactions = 0
        For Each t In NextActionDoc.Range(mm.MmRange.mmRangeAllTopics)
            If Not t.IsCalloutTopic And t.Task.IsValid Then
                If Not isdate0(t.Task.StartDate) And t.Task.Complete < 100 Then
                    numdatedactions = numdatedactions + 1
                    actions(numdatedactions) = t
                    age(numdatedactions) = CInt(DateDiff(DateInterval.Day, t.Task.StartDate, Today))
                    If age(numdatedactions) < 0 Then age(numdatedactions) = 0
                End If
            End If
        Next

        For i = 1 To numdatedactions - 1
            For j = i + 1 To numdatedactions
                If age(i) < age(j) Then
                    swap = actions(i)
                    actions(i) = actions(j)
                    actions(j) = swap
                    swapa = age(i)
                    age(i) = age(j)
                    age(j) = swapa
                End If
            Next
        Next

        avgage = arrayaverage(age, numdatedactions)

        oldest = age(1)

        oldavg = 0
        For Each t In NextActionDoc.CentralTopic.AllSubTopics
            contexttotal = 0
            contextcount = 0
            For Each tt In t.AllSubTopics
                If Not isdate0(tt.Task.StartDate) Then
                    contextcount = contextcount + 1
                    contexttotal = contexttotal + DateDiff(DateInterval.Day, tt.Task.StartDate, Today)
                End If
            Next
            If contextcount > 0 Then
                If contexttotal / contextcount > oldavg Then
                    oldavg = contexttotal / contextcount
                    oldcontexttopic = t
                End If
            End If
        Next

        'count up young noncontext entries

        For Each t In NextActionDoc.CentralTopic.AllSubTopics
            If InStr(LCase(t.Text), "no context") > 0 Then
                For Each tt In t.AllSubTopics
                    If Not isdate0(tt.Task.StartDate) Then
                        If DateDiff(DateInterval.Day, tt.Task.StartDate, Today) < graceperiod Then
                            numnocontext = numnocontext - 1
                        End If
                    End If
                Next
                Exit For
            End If
        Next


        deltopic.AddSubTopic("Analyzing in-tray task aging")
        numdatedintrayactions = 0
        oldestintrayage = 0
        For Each t In NextActionDoc.Range(mm.MmRange.mmRangeAllTopics)
            If Not t.IsCalloutTopic And t.Task.IsValid Then
                If Not isdate0(t.Task.StartDate) And t.Task.Complete < 100 Then
                    If t.CalloutTopics.Count > 0 Then
                        If InStr(LCase(t.CalloutTopics(True).Item(1).Text), "in-tray") > 0 Then
                            numdatedintrayactions = numdatedintrayactions + 1
                            InTrayAge(numdatedintrayactions) = CInt(DateDiff(DateInterval.Day, t.Task.StartDate, Today))
                            If InTrayAge(numdatedintrayactions) > oldestintrayage Then
                                oldestintrayitemxml = t.Xml
                                oldestintrayage = InTrayAge(numdatedintrayactions)
                            End If
                            If InTrayAge(numdatedintrayactions) < 0 Then InTrayAge(numdatedintrayactions) = 0
                        End If
                    End If
                End If
            End If
        Next
        avInTrayAge = arrayaverage(InTrayAge, numdatedintrayactions)
        deltopic.AddSubTopic("Calculating total time estimate")
        Dim numNoEstimate As Integer
        numNoEstimate = 0
        Dim TotalTimeEst As Double
        TotalTimeEst = 0
        For i = 1 To numactions
            If InStr(nextactioncat(i), "2m") > 0 Then
                TotalTimeEst = TotalTimeEst + 0.033
            ElseIf InStr(nextactioncat(i), "15m") > 0 Then
                TotalTimeEst = TotalTimeEst + 0.25
            ElseIf InStr(nextactioncat(i), "1h") > 0 Then
                TotalTimeEst = TotalTimeEst + 1
            ElseIf InStr(nextactioncat(i), "2h") > 0 Then
                TotalTimeEst = TotalTimeEst + 0.25
            Else
                If age(i) > graceperiod Then
                    numNoEstimate = numNoEstimate + 1
                End If
            End If
        Next


        '---------------------------------------------------------
        'Generate Report

        Const freshness = 1
        Const focus = 2
        Const feasibility = 3
        Const finishing = 4
        Const foresight = 5

        'Dashboard Scoring

        deltopic.AddSubTopic("Calculating Metrics")

        Dim mclass(15) As Integer
        'freshness metrics
        mclass(1) = freshness : metric(1) = avgage : weight(1) = 0.1 : goal(1) = 14 : half(1) = 28 : mname(1) = "avg age "
        mclass(2) = freshness : metric(2) = avInTrayAge : weight(2) = 0.1 : goal(2) = 14 : half(2) = 30 : mname(2) = "In-Tray age"

        'focus metrics
        mclass(3) = focus : metric(3) = numprojects : weight(3) = 0.02 : goal(3) = 40 : half(3) = 80 : mname(3) = "# projects "
        mclass(4) = focus
        If numprojects > 0 Then
            metric(4) = (numprojects - numdatedprojects) / numprojects : weight(4) = 0.09 : goal(4) = 0.5 : half(4) = 0.75 : mname(4) = "no dates   "
        Else
            metric(4) = 1 : weight(4) = 0.09 : goal(4) = 0.5 : half(4) = 0 : mname(4) = "no dates   "
        End If
        mclass(11) = focus
        If nummainprojects > 0 Then
            metric(11) = numnopriorityprojects / nummainprojects : weight(11) = 0.08 : goal(11) = 0 : half(11) = 0.1 : mname(11) = "no priority"
        Else
            metric(11) = 1 : weight(11) = 0.08 : goal(11) = 0 : half(11) = 0.1 : mname(11) = "no priority"
        End If

        'finishing metrics
        mclass(6) = finishing : metric(6) = numoverdue : weight(6) = 0.06 : goal(6) = 0 : half(6) = 2 : mname(6) = "overdue    "
        mclass(7) = finishing : metric(7) = numoverduewaiting : weight(7) = 0.04 : goal(7) = 0 : half(7) = 2 : mname(7) = "waiting    "
        mclass(8) = finishing : metric(8) = numpasttarget : weight(8) = 0.03 : goal(8) = 0 : half(8) = 4 : mname(8) = "past target"
        mclass(12) = finishing : metric(12) = numwaiting : weight(12) = 0.02 : goal(12) = 15 : half(12) = 20 : mname(12) = "waiting on"
        mclass(13) = finishing : metric(13) = numoverdueprojects : weight(13) = 0.05 : goal(13) = 0 : half(13) = 2 : mname(13) = "late projects"

        'foresight metrics
        mclass(9) = foresight : metric(9) = numnocontext : weight(9) = 0.08 : goal(9) = 3 : half(9) = 8 : mname(9) = "aging no context "
        mclass(10) = foresight : metric(10) = numneednextaction : weight(10) = 0.08 : goal(10) = 0 : half(10) = 2 : mname(10) = "next step"
        mclass(14) = foresight : metric(14) = numNoEstimate : weight(14) = 0.04 : goal(14) = 3 : half(14) = 10 : mname(14) = "no time estimate"

        'fesibility metrics
        mclass(5) = feasibility : metric(5) = numactions : weight(5) = 0.1 : goal(5) = 60 : half(5) = 90 : mname(5) = "# actions   "
        mclass(15) = feasibility : metric(15) = numDailyCapture : weight(15) = 0.1 : goal(15) = 0 : half(15) = 15 : mname(15) = "# in Daily Capture"


        advice(1) = "Complete aging actions or put on someday list, delegate, or insert do-able predecessor"
        advice(2) = "Deal with old items still sitting in in-trays"
        advice(3) = "Reduce your " & numprojects & " project by completing, delegating, deferring, or putting on someday list"
        advice(4) = "Add target dates to some of your projects"
        advice(5) = "Reduce your " & numactions & " Next Actions by doing, delegating, deferring, or putting on someday list"
        advice(6) = "Renegotiate your deadlines or meet them"
        advice(7) = "Follow-up with people whose deadlines have slipped"
        advice(8) = "Review the tasks that have slipped past their target date"
        advice(9) = "Add contexts to actions or replace with better 'physical' next actions"
        advice(10) = "Add next steps to projects that need them"
        advice(11) = "Add priority to your projects and subprojects"
        advice(12) = "Set target dates for items you are waiting for"
        advice(13) = "Finish or renegotiate your overdue projects"
        advice(14) = "Estimate time required for tasks (set category=2m,15m,1h,2h)"
        advice(15) = "Reduce the number of unprocessed items in your daily capture map"

        hlink(1) = "http://wiki.activityowner.com/index.php?title=NAA_Average_Age"
        hlink(2) = "http://wiki.activityowner.com/index.php?title=NAA_Aging_InTray_Items"
        hlink(3) = "http://wiki.activityowner.com/index.php?title=NAA_Many_Projects"
        hlink(4) = "http://wiki.activityowner.com/index.php?title=NAA_Doable_Projects"
        hlink(5) = "http://wiki.activityowner.com/index.php?title=NAA_Number_of_Actions"
        hlink(6) = "http://wiki.activityowner.com/index.php?title=NAA_Overdue"
        hlink(7) = "http://wiki.activityowner.com/index.php?title=NAA_Overdue_Waiting"
        hlink(8) = "http://wiki.activityowner.com/index.php?title=NAA_Past_Target"
        hlink(9) = "http://wiki.activityowner.com/index.php?title=NAA_Context"
        hlink(10) = "http://wiki.activityowner.com/index.php?title=NAA_Need_Next_Action"
        hlink(11) = "http://wiki.activityowner.com/index.php?title=NAA_Need_Priority"
        hlink(12) = "http://wiki.activityowner.com/index.php?title=NAA_Waiting_On"
        hlink(13) = "http://wiki.activityowner.com/index.php?title=NAA_Overdue_Projects"
        hlink(14) = "http://wiki.activityowner.com/index.php?title=NAA_Time_Estimates"
        hlink(15) = "http://wiki.activityowner.com/index.php?title=NAA_Too_Many_Daily_Capture"

        'ratings add up to maximum 1.0 (multipled by 10 later)
        For i = 1 To nm
            ratings(i) = score(metric(i), goal(i), half(i), weight(i))
        Next

        Dim F(5) As Double
        F(freshness) = ratings(1) + ratings(2)     '20% freshness
        F(focus) = ratings(3) + ratings(4) + ratings(11)   '20% focus
        F(feasibility) = ratings(5) + ratings(15)             '20% feasibility
        F(finishing) = ratings(6) + ratings(7) + ratings(8) + ratings(12) + ratings(13)  '20% finishing
        F(foresight) = ratings(9) + ratings(10) + ratings(14)   '20% foresight

        '

        rating = F(1) + F(2) + F(3) + F(4) + F(5)
        deltopic.AddSubTopic("Generating Report")
        AdviceTopic = reporttopic.AddSubTopic("Advice:")
        'scoring
        ratingtopic = reporttopic.AddSubTopic("Overall NAA Rating: " & Math.Round(rating * 10, 2) & "/10")
        For i = 1 To nm
            ratingtopic.Notes.Text = ratingtopic.Notes.Text & vbCrLf & "Metric " & Str(i) & " : " & mname(i) & "   " & Str(Math.Round(metric(i), 2)) & "  " & Str(Math.Round(ratings(i), 2)) & " of " & Str(Math.Round(weight(i), 2))
            ratingtopic.Notes.Commit()
        Next
        '
        'show in sorted order
        Dim allshown As Boolean
        allshown = False
        Dim minrating As Double
        Dim mintopic As Integer
        Dim shown(5) As Boolean
        For i = 1 To 5
            shown(i) = False
        Next
        deltopic.AddSubTopic("Sorting 5F scores")
        While Not allshown
            minrating = 10
            For i = 1 To 5
                If F(i) <= minrating And shown(i) = False Then
                    minrating = F(i)
                    mintopic = i
                End If
            Next
            shown(mintopic) = True
            If mintopic = foresight Then ForesightTopic = ratingtopic.AddSubTopic("Foresight: " & Math.Round(F(foresight) * 5 * 10, 2) & "/10")
            If mintopic = freshness Then FreshnessTopic = ratingtopic.AddSubTopic("Freshness: " & Math.Round(F(freshness) * 5 * 10, 2) & "/10")
            If mintopic = focus Then FocusTopic = ratingtopic.AddSubTopic("Focus: " & Math.Round(F(focus) * 5 * 10, 2) & "/10")
            If mintopic = feasibility Then FeasibilityTopic = ratingtopic.AddSubTopic("Feasibility: " & Math.Round(F(feasibility) * 5 * 10, 2) & "/10")
            If mintopic = finishing Then FinishTopic = ratingtopic.AddSubTopic("Finishing: " & Math.Round(F(finishing) * 5 * 10, 2) & "/10")
            allshown = True
            For i = 1 To 5
                If shown(i) = False Then allshown = False
            Next
        End While
        'flag good/bad dimensions
        If rating < 0.66 Then ratingtopic.TextColor.SetARGB(255, 255, 0, 0)
        If rating > 0.89 Then ratingtopic.TextColor.SetARGB(255, 0, 255, 0)
        If F(foresight) < 0.65 * 0.2 Then ForesightTopic.TextColor.SetARGB(255, 255, 0, 0)
        If F(freshness) < 0.65 * 0.2 Then FreshnessTopic.TextColor.SetARGB(255, 255, 0, 0)
        If F(focus) < 0.65 * 0.2 Then FocusTopic.TextColor.SetARGB(255, 255, 0, 0)
        If F(finishing) < 0.65 * 0.2 Then FinishTopic.TextColor.SetARGB(255, 255, 0, 0)
        If F(feasibility) < 0.65 * 0.2 Then FeasibilityTopic.TextColor.SetARGB(255, 255, 0, 0)


        FeasibilityTopic.AddSubTopic("Your tasks with time estimates add up to " & Str(Math.Round(TotalTimeEst, 1)) & " hours")
        'add advice for area(s) of most improvement opportunity
        maxopportunity = 0
        firstitem = 1
        seconditem = 1
        thirditem = 1
        secondopportunity = 0
        thirdopportunity = 0
        deltopic.AddSubTopic("Creating Recommendations")
        For i = 1 To nm
            opportunity(i) = weight(i) - ratings(i)
            If opportunity(i) > maxopportunity Then
                maxopportunity = opportunity(i)
                firstitem = i
            End If
        Next
        For i = 1 To nm
            If opportunity(i) > secondopportunity Then
                If Not (i = firstitem) Then
                    secondopportunity = opportunity(i)
                    seconditem = i
                End If
            End If
        Next
        For i = 1 To nm
            If opportunity(i) > thirdopportunity Then
                If Not (i = firstitem) And Not (i = seconditem) Then
                    thirdopportunity = opportunity(i)
                    thirditem = i
                End If
            End If
        Next
        If rating > 0.89 Then
            AdviceTopic.AddSubTopic("Your system is in great shape -- Go pick a context list and get things done on it!").Icons.AddStockIcon(mm.MmStockIcon.mmStockIconSmileyHappy)
        End If
        If maxopportunity > 0 Then
            AdviceTopic1 = AdviceTopic.AddSubTopic(advice(firstitem) & " for " & Math.Round(opportunity(firstitem) * 10, 2) & " points")
            AdviceTopic1.Task.Complete = 0
            AdviceTopic1.Task.Priority = mm.MmTaskPriority.mmTaskPriority1
            AdviceTopic1.CreateHyperlink(hlink(firstitem))
        End If
        If secondopportunity > 0 Then
            t = AdviceTopic.AddSubTopic(advice(seconditem) & " for " & Math.Round(opportunity(seconditem) * 10, 2) & " points")
            t.Task.Complete = 0
            t.Task.Priority = mm.MmTaskPriority.mmTaskPriority2
            t.CreateHyperlink(hlink(seconditem))
        End If
        If thirdopportunity > 0 Then
            t = AdviceTopic.AddSubTopic(advice(thirditem) & " for " & Math.Round(opportunity(thirditem) * 10, 2) & " points")
            t.Task.Complete = 0
            t.Task.Priority = mm.MmTaskPriority.mmTaskPriority3
            t.CreateHyperlink(hlink(thirditem))
        End If
        If AdviceTopic.AllSubTopics.Count > 0 Then
            AdviceTopic.AddSubTopic("Visit Links for more targeted advice and resources for addressing the challenges above from activityowner.com")
        Else
            AdviceTopic.AddSubTopic("Perfect Score!")
            AdviceTopic.AddSubTopic("go to review dashboard and consider your someday maybes")
            AdviceTopic.AddSubTopic("Do a mind sweep/office sweep/house sweep and make sure you have captured all")
        End If
        '----------------------------------------------------------------------------------------------------------------------------
        'add supporting information
        deltopic.AddSubTopic("Adding Supporting Information for ...")
        '--------------------FEASIBILITY------------------------------------------------------------------------
        deltopic.AddSubTopic(" * Feasibility")
        t = FeasibilityTopic.AddSubTopic("You have " & numactions & " next actions")
        If numactions > 100 Then t.TextColor.SetARGB(255, 255, 0, 0)
        t = maxbranch(NextActionDoc)
        If Not t Is Nothing Then
            tt = FeasibilityTopic.AddSubTopic("You have " & Str(t.AllSubTopics.Count) & " actions (" & Str(Math.Round(t.AllSubTopics.Count * 100 / numactions, 0)) & "%) in the " & t.Text & " context")
            For Each ttt In t.AllSubTopics
                tt.AddSubTopic("").Xml = ttt.Xml
            Next
            tt.SetLevelOfDetail(0)
        End If
        If numprojects > 0 And actioncount(mostactions) > 0 Then
            FeasibilityTopic.AddSubTopic("You have " & actioncount(mostactions) & " actions in...").AddSubTopic("").Xml = projectxml(mostactions)
        End If
        FeasibilityTopic.AddSubTopic("You have " & mapcount(mostmap) & " actions in " & Mid(maplinks(mostmap), InStrRev(maplinks(mostmap), "\") + 1, Len(maplinks(mostmap)) - InStrRev(maplinks(mostmap), "\") - 4)).CreateHyperlink(maplinks(mostmap))
        tt = FeasibilityTopic.AddSubTopic("Repeating Actions")
        numareas = 0
        For i = 1 To numactions
            If repeating(i) Then tt.AddSubTopic("").Xml = nextactionxml(i)
            found = False
            If numareas > 0 Then
                For j = 1 To numareas
                    If LCase(taskarea(i)) = LCase(area(j)) Then
                        found = True
                        Exit For
                    End If
                Next
            End If
            If Not found Then
                numareas = numareas + 1
                area(numareas) = taskarea(i)
            End If
        Next
        tt = Nothing
        ttt = Nothing
        '----FINISHING-----------------------------------------------------------------------------------------------------------------------
        deltopic.AddSubTopic(" * Finishing")
        If metric(13) > 0 Then
            If metric(13) = 1 Then
                t = FinishTopic.AddSubTopic(metric(13) & " project is overdue (see note)")
                t.Notes.Text = "Note others may actually be overdue if they happen to have start date = due date. These are ignored as a MM9 workaround."
                t.Notes.Commit()
            Else
                t = FinishTopic.AddSubTopic(metric(13) & " projects are overdue (see note)")
                t.Notes.Text = "Note others may actually be overdue if they happen to have start date = due date. These are ignored as a MM9 workaround."
                t.Notes.Commit()
            End If
            For i = 1 To numprojects
                If projectoverdue(i) = True Then
                    t.AddSubTopic("").Xml = projectxml(i)
                End If
            Next
            t.SetLevelOfDetail(0)
        End If
        If metric(6) > 0 Then
            t = FinishTopic.AddSubTopic(metric(6) & " actions are past hard deadline")
            If numoverdue > 0 Then t.TextColor.SetARGB(255, 255, 0, 0)
            For Each tt In DeadlinesDoc.Range(mm.MmRange.mmRangeAllTopics)
                If isred(tt) And tt.Task.Complete < 100 And tt.Task.Complete >= 0 Then t.AddSubTopic("").Xml = tt.Xml
            Next
            t.SetLevelOfDetail(0)
        End If

        If metric(8) > 0 Then
            t = FinishTopic.AddSubTopic(metric(8) & " next actions are past target")
            If numoverdue > 0 Then t.TextColor.SetARGB(255, 255, 0, 0)
            For Each tt In NextActionDoc.Range(mm.MmRange.mmRangeAllTopics)
                'don't enumerate an item from an overdue project
                If isred(tt) Then
                    If Not tt.IsCalloutTopic Then
                        If Not tt.Icons.HasStockIcon(mm.MmStockIcon.mmStockIconExclamationMark) And Not tt.IsCentralTopic Then
                            If tt.CalloutTopics.Count > 0 Then
                                If isdate0(tt.CalloutTopics.Item(1).Task.DueDate) And tt.CalloutTopics.Item(1).Task.DueDate < Now Then
                                    t.AddSubTopic("").Xml = tt.Xml
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            t.SetLevelOfDetail(0)
        End If
        If metric(7) > 0 Then
            t = FinishTopic.AddSubTopic(metric(7) & " things you are waiting for are overdue")
            If numoverduewaiting > 0 Then
                t.TextColor.SetARGB(255, 255, 0, 0)
            Else
                t.Delete()
            End If
            For Each tt In RelationshipDoc.Range(mm.MmRange.mmRangeAllTopics)
                If isred(tt) Then
                    If parentcontains(tt, "waiting") Then
                        t.AddSubTopic("").Xml = tt.Xml
                    End If
                End If
            Next
            t.SetLevelOfDetail(0)
        End If

        t = FinishTopic.AddSubTopic("Targeted for Today or tomorrow")
        For i = 1 To numactions
            If allactions(i).Task.DueDate = Today Or allactions(i).Task.DueDate = DateAdd(DateInterval.Day, 1, Today) Then
                t.AddSubTopic("").Xml = nextactionxml(i)
            ElseIf allactions(i).AllCalloutTopics.Count > 0 Then
                If allactions(i).AllCalloutTopics.Item(1).Task.DueDate = Today Or allactions(i).AllCalloutTopics.Item(1).Task.DueDate = DateAdd(DateInterval.Day, 1, Today) Then
                    t.AddSubTopic("").Xml = nextactionxml(i)
                End If
            End If
        Next
        t.SetLevelOfDetail(0)
        t = FinishTopic.AddSubTopic(pct_complete & " percent complete overall so far")

        If numwaiting > 0 Then
            t = FinishTopic.AddSubTopic("Waiting for " & numwaiting & " tasks with no target date")
        End If
        For Each tt In RelationshipDoc.CentralTopic.AllSubTopics
            For Each ttt In tt.AllSubTopics
                If InStr(ttt.Text, "waiting") > 0 Then
                    persontopic = t.AddSubTopic(tt.Text)
                    For Each tttt In ttt.AllSubTopics
                        If isdate0(tttt.Task.DueDate) Then
                            If tttt.CalloutTopics.Count > 0 Then
                                If isdate0(tttt.CalloutTopics.Item(1).Task.DueDate) Then
                                    persontopic.AddSubTopic("").Xml = tttt.Xml
                                End If
                            End If
                        End If
                    Next
                    If persontopic.AllSubTopics.Count = 0 Then persontopic.Delete()
                End If
            Next
        Next
        t.SetLevelOfDetail(0)

        '----FORESIGHT--------------------------------------------------------------------------------------------------------------------------
        deltopic.AddSubTopic(" * Foresight")
        nocontextactions = ForesightTopic.AddSubTopic(Math.Round(metric(9)) & " of your next actions do not have a context")
        If metric(9) = 0 Then
            nocontextactions.Delete()
        Else
            If numactions > 0 Then
                If numnocontext / numactions > 0.2 Then nocontextactions.TextColor.SetARGB(255, 255, 0, 0)
            End If
            If numnocontext > 0 Then
                For Each t In NoContextDoc.CentralTopic.AllSubTopics
                    nocontextactions.AddSubTopic("").Xml = t.Xml
                Next
            End If
            nocontextactions.SetLevelOfDetail(0)
        End If
        If numNoEstimate > 0 Then
            timecategories = ForesightTopic.AddSubTopic("tasks by time (" & numNoEstimate & " need estimates)")
        Else
            timecategories = ForesightTopic.AddSubTopic("tasks by time")
        End If
        noest = timecategories.AddSubTopic("no time estimate")
        c2m = timecategories.AddSubTopic("2m")
        c15m = timecategories.AddSubTopic("15m")
        c1h = timecategories.AddSubTopic("1h")
        c2h = timecategories.AddSubTopic("2h")
        For i = 1 To numactions
            If InStr(nextactioncat(i), "2m") > 0 Then
                c2m.AddSubTopic("").Xml = nextactionxml(i)
            ElseIf InStr(nextactioncat(i), "15m") > 0 Then
                c15m.AddSubTopic("").Xml = nextactionxml(i)
            ElseIf InStr(nextactioncat(i), "1h") > 0 Then
                c1h.AddSubTopic("").Xml = nextactionxml(i)
            ElseIf InStr(nextactioncat(i), "2h") > 0 Then
                c2h.AddSubTopic("").Xml = nextactionxml(i)
            Else
                noest.AddSubTopic("").Xml = nextactionxml(i)
            End If
        Next
        For Each t In timecategories.AllSubTopics
            If t.AllSubTopics.Count = 0 Then t.Delete()
        Next
        'gather up in-tray items
        Dim intrayitems As mm.Topic
        intrayitems = ForesightTopic.AddSubTopic("Items in Intrays")
        For Each t In NextActionDoc.Range(mm.MmRange.mmRangeAllTopics)
            Debug.Print("intrays not being checked in naa")
            '  If Not t.IsCalloutTopic And t.Task.IsValid Then
            'If Not isdate0(t.Task.StartDate) And t.Task.Complete < 100 Then
            'If InStr(LCase(t.CalloutTopics(True).Item(1).Text), "in-tray") > 0 Then
            'intrayitems.AddSubTopic("").Xml = t.Xml
            'End If
            'End If
            'End If
        Next
        intrayitems = Nothing
        '---FRESHNESS-----------------------------------------------------------------------------------------------------------------------
        deltopic.AddSubTopic(" * Freshness")
        t = FreshnessTopic.AddSubTopic("Your " & numdatedactions & " dated next actions average " & Math.Round(avgage) & " days old")
        If avgage > 40 Then t.Text = t.Text & "!"
        If avgage > 40 Then t.TextColor.SetARGB(255, 255, 0, 0)

        t = FreshnessTopic.AddSubTopic("They have been around for a grand total of " & Math.Round(numdatedactions * avgage / 365, 1) & " years")
        If numdatedactions * avgage / 365 > 7 Then t.Text = t.Text & "!"
        If numdatedactions * avgage / 365 > 7 Then t.TextColor.SetARGB(255, 255, 0, 0)
        t = FreshnessTopic.AddSubTopic("Your oldest in-tray item is " & oldestintrayage & " days old")
        If Not oldestintrayitemxml = "" Then
            dummy = t.AddSubTopic("")
            dummy.Xml = oldestintrayitemxml
        End If

        Try
            If numactions > 0 Then
                t = reporttopic.AddSubTopic("Random Activities for the Day:")
                t.AddSubTopic("").Xml = actions(CInt(Math.Round(Rnd() * (numactions - 1)) + 1)).Xml
                t.AddSubTopic("").Xml = actions(CInt(Math.Round(Rnd() * (numactions - 1)) + 1)).Xml
                t.AddSubTopic("").Xml = actions(CInt(Math.Round(Rnd() * (numactions - 1)) + 1)).Xml
            End If
        Catch
        End Try
        If numdatedactions > 0 Then
            t = FreshnessTopic.AddSubTopic("Oldest Next Actions")
            If 0.1 * numdatedactions > 1 Then
                For i = 1 To CInt(Math.Round(0.1 * numdatedactions))
                    t.AddSubTopic(Str(age(i))).AddSubTopic("").Xml = actions(i).Xml
                Next
            End If
            t.SetLevelOfDetail(0)
            t = FreshnessTopic.AddSubTopic("Youngest Next Actions")
            If 0.9 * numdatedactions > 1 Then
                For i = CInt(Math.Round(0.9 * numdatedactions)) To numdatedactions
                    t.AddSubTopic(Str(age(i))).AddSubTopic("").Xml = actions(i).Xml
                Next
            End If
            t.SetLevelOfDetail(0)
            t = FreshnessTopic.AddSubTopic("Undated Next Actions")
            If numactions > 0 Then
                For i = 1 To numactions
                    If isdate0(allactions(i).Task.StartDate) Then t.AddSubTopic("").Xml = allactions(i).Xml
                Next
            End If
            t.SetLevelOfDetail(0)
        End If
        If Not oldcontexttopic Is Nothing Then
            t = FreshnessTopic.AddSubTopic("Items in your " & oldcontexttopic.Text & " average " & Math.Round(oldavg) & " days old")
            For Each tt In oldcontexttopic.AllSubTopics
                t.AddSubTopic("").Xml = tt.Xml
            Next
        End If
        t.SetLevelOfDetail(0)




        '---Focus----------------------------------------------------------------------------------------------------------------
        deltopic.AddSubTopic(" * Focus")
        t = FocusTopic.AddSubTopic("You are trying to advance " & Str(numprojects) & " projects and Sub-projects this week")
        If numprojects > 50 Then t.Text = t.Text & "!"
        If numprojects > 50 Then t.TextColor.SetARGB(255, 255, 0, 0)

        FocusTopic.AddSubTopic("Your next actions are derived from " & nummaps & " maps")
        FocusTopic.AddSubTopic(Str(DueInSeven) & " of your next actions are targeted for completion in next 7 days")
        If metric(11) > 0 Then t = FocusTopic.AddSubTopic(Math.Round(metric(11) * 100, 0) & "% of your main projects have not been prioritized.")
        If numnopriorityprojects > 0 Then
            For i = 1 To numprojects
                If isproject(i) And projectpriority(i) = 0 Then t.AddSubTopic("").Xml = projectxml(i)
            Next
            t.SetLevelOfDetail(0)
        End If
        If numneednextaction > 0 Then
            neednextaction = ForesightTopic.AddSubTopic(numneednextaction & " committed projects need next actions")
            If NoActionDoc.CentralTopic.AllSubTopics.Count > 0 Then
                For Each t In NoActionDoc.CentralTopic.AllSubTopics
                    neednextaction.AddSubTopic("").Xml = t.Xml
                Next
                neednextaction.TextColor.SetARGB(255, 255, 0, 0)
            End If
        End If
        taskprioritytopic = FocusTopic.AddSubTopic("tasks by priority")
        For i = 2 To 5
            ptopic(i) = taskprioritytopic.AddSubTopic(Str(i - 1))
        Next
        ptopic(1) = taskprioritytopic.AddSubTopic("Unprioritized")
        For i = 1 To 5
            For j = 1 To numactions
                If nextactionpriority(j) = (i - 1) Then
                    ptopic(i).AddSubTopic("").Xml = nextactionxml(j)
                End If
            Next
        Next
        taskprioritytopic.SetLevelOfDetail(0)
        For i = 1 To 5
            ptopic(i).SetLevelOfDetail(0)
        Next
        ForesightTopic.SetLevelOfDetail(1)
        If numactions > 0 Then
            If numnocontext / numactions > 0.2 Then nocontextactions.TextColor.SetARGB(255, 255, 0, 0)
        End If

        If numprojects > 0 Then
            t = FocusTopic.AddSubTopic(Math.Round(numdatedprojects / numprojects * 100, 0) & "% of projects/sub-projects have target dates")
            If numdatedprojects / numprojects < 0.5 Then t.Text = "Only " & t.Text & "!"
            If numdatedprojects / numprojects < 0.5 Then t.TextColor.SetARGB(255, 255, 0, 0)
        End If
        datedprojects = t.AddSubTopic("Target Date Projects")
        undatedprojects = t.AddSubTopic("Undated Projects")
        undatedsubprojects = t.AddSubTopic("Undated Sub-Projects")

        If numareas > 0 Then
            t = FocusTopic.AddSubTopic("Tasks by Area")
            For i = 1 To numareas
                tt = t.AddSubTopic(area(i))
                If area(i) = "" Then tt.Text = "undefined"
                For k = 0 To 4
                    ttt = tt.AddSubTopic(Str(k))
                    If k = 0 Then ttt.Text = "No Priority"
                    For j = 1 To numactions
                        If LCase(taskarea(j)) = LCase(area(i)) And nextactionpriority(j) = k Then
                            ttt.AddSubTopic("").Xml = nextactionxml(j)
                        End If
                    Next
                    If ttt.AllSubTopics.Count = 0 Then ttt.Delete()
                Next
            Next
            For Each tt In t.AllSubTopics
                cnt = 0
                For Each ttt In tt.AllSubTopics
                    cnt = cnt + ttt.AllSubTopics.Count
                Next
                If numactions > 0 Then
                    tt.Text = tt.Text & " (" & LTrim(Str(Math.Round(cnt / numactions * 100, 0))) & "%)"
                End If
                tt.SetLevelOfDetail(0)
                If tt.AllSubTopics.Count = 0 Then tt.Delete()
            Next
            t.SetLevelOfDetail(0)
            t = FocusTopic.AddSubTopic("Projects by Area")
            For i = 1 To numareas
                tt = t.AddSubTopic(area(i))
                If area(i) = "" Then tt.Text = "undefined"
                For k = 0 To 4
                    ttt = tt.AddSubTopic(Str(k))
                    If k = 0 Then ttt.Text = "No Priority"
                    For j = 1 To numprojects
                        If LCase(projectarea(j)) = LCase(area(i)) And projectpriority(j) = k Then
                            Try
                                ttt.AddSubTopic("").Xml = projectxml(j)
                            Catch
                                Debug.Print("error adding no priority xml")
                            End Try

                        End If
                    Next
                    If ttt.AllSubTopics.Count = 0 Then ttt.Delete()
                Next
            Next
            For Each tt In t.AllSubTopics
                cnt = 0
                For Each ttt In tt.AllSubTopics
                    cnt = cnt + ttt.AllSubTopics.Count
                Next
                If numprojects > 0 Then
                    tt.Text = tt.Text & " (" & LTrim(Str(Math.Round(cnt / numprojects * 100, 0))) & "%)"
                End If
                tt.SetLevelOfDetail(0)
                If tt.AllSubTopics.Count = 0 Then tt.Delete()
            Next
            t.SetLevelOfDetail(0)
        End If


        For j = 0 To 9
            For i = 1 To numprojects
                If projectpriority(i) = j Then
                    If projectdated(i) Then
                        datedprojects.AddSubTopic("").Xml = projectxml(i)
                    Else
                        If isproject(i) Then
                            undatedprojects.AddSubTopic("").Xml = projectxml(i)
                        Else
                            Try
                                undatedsubprojects.AddSubTopic("").Xml = projectxml(i)
                            Catch
                                Debug.Print("error adding undated project xml")
                            End Try

                        End If
                    End If
                End If
            Next
        Next
        addtolog = True
        If logprompt Then
            addtolog = MsgBox("Do you want to save your scores to the log?", vbYesNo) = vbYes
        End If
        If addtolog Then
            deltopic.AddSubTopic("Updating Log map")
            Try
                LogDoc = getmap(m_app, logdocname)
            Catch
            End Try
            If Not LogDoc Is Nothing Then
                LogDoc.CentralTopic.Notes.Text = LogDoc.CentralTopic.Notes.Text & vbCrLf & Now & "," & Str(Math.Round(rating * 10, 2))
                LogDoc.CentralTopic.Notes.Commit()

                For i = 1 To 12
                    LogDoc.CentralTopic.Notes.Text = LogDoc.CentralTopic.Notes.Text & "," & Str(Math.Round(ratings(i) * 10, 2))
                    t = createmainbranch(mname(i), LogDoc, "")
                    t.Notes.Text = t.Notes.Text & vbCrLf & Now & "," & Str(Math.Round(metric(i), 2))
                    t.Notes.Commit()
                Next
                t = createmainbranch("Overall Score", LogDoc, "")
                t.Notes.Text = t.Notes.Text & vbCrLf & Now & "," & Str(Math.Round(rating, 2))
                t.Notes.Commit()
                LogDoc.SaveAs(m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps) & logdocname)
                LogDoc.Close()
                LogDoc = Nothing
            End If
        End If
        datedprojects.SetLevelOfDetail(0)
        undatedprojects.SetLevelOfDetail(0)
        undatedsubprojects.SetLevelOfDetail(0)
        FeasibilityTopic.SetLevelOfDetail(1)
        If showadvice Then
            Try
                If Not AdviceTopic1 Is Nothing Then
                    AdviceTopic1.Hyperlink.Follow()
                End If
            Catch

            End Try
        End If
        deltopic.AddSubTopic("Closing temporary files")

        NextActionDoc.Close()
        NoContextDoc.Close()
        NoActionDoc.Close()
        DeadlinesDoc.Close()
        projectsdoc.Close()
        RelationshipDoc.Close()
        deltopic.AddSubTopic("Cleaning Up")
        NextActionDoc = Nothing
        NoContextDoc = Nothing
        projectsdoc = Nothing
        NoActionDoc = Nothing
        RelationshipDoc = Nothing
        For i = 1 To 500
            actions(i) = Nothing
            allactions(i) = Nothing
        Next
        t = Nothing
        tt = Nothing
        ttt = Nothing
        tttt = Nothing
        persontopic = Nothing
        AdviceTopic = Nothing
        swap = Nothing
        deltopic = Nothing
        dashboard.Activate()
        reportdoc.Close()
        reportdoc = Nothing
        reporttopic = Nothing
        oldcontexttopic = Nothing
        c15m = Nothing
        c1h = Nothing
        c2h = Nothing
        For i = 1 To 5
            ptopic(i) = Nothing
        Next
        taskprioritytopic = Nothing
        'IntrayDoc.Close
        dashboard = Nothing

    End Sub
    Function score(ByVal ovalue As Object, ByVal goal As Double, ByVal half As Double, ByVal weight As Double) As Double
        Dim value As Double
        value = CDbl(ovalue)
        If value <= goal Then
            score = weight
        ElseIf value < half Then
            score = weight * (0.5 + 0.5 * Math.Cos((value - goal) / (goal - half) * 0.5 * 3.1416))
        Else
            score = weight * 0.5 * 1 / Math.Exp(-(value - half) / (goal - half))
        End If
        If score < 0 Then score = 0
    End Function
    Sub Upgrade3(ByRef m_app As Mindjet.MindManager.Interop.Application, ByRef ConfigDoc As mm.Document)
        Dim a As mm.Topic
        Dim lastupgrade As String
        Dim RunUpgrade As Boolean
        Const currentversion = "20090109"
        lastupgrade = getoption("lastupgrade", ConfigDoc, Nothing)
        RunUpgrade = False
        Try
            If Len(lastupgrade) = 0 Then
                RunUpgrade = True
            ElseIf CInt(lastupgrade) < CInt(currentversion) Then 'do not want to eval("")
                RunUpgrade = True
            End If
        Catch
            RunUpgrade = True
        End Try
        If RunUpgrade Then
            'OPTIONS-----------------------------------------------------------
            createoption("showadvice", "0", ConfigDoc)
            createoption("askadvice", "0", ConfigDoc)
            createoption("logprompt", "0", ConfigDoc)
            createoption("lastversioncheck", Today, ConfigDoc)
            createoption("versioncheckfrequency", "30", ConfigDoc)
            createoption("checkempty", "0", ConfigDoc)
            createoption("logdocname", "AO\NAAlog.mmap", ConfigDoc)
            checkforduplicates(ConfigDoc)
            'Mark map as upgraded
            setoption("lastupgrade", currentversion, ConfigDoc)
            If ConfigDoc.IsModified Then ConfigDoc.Save()
            a = Nothing
        End If
    End Sub
    Function getdashboardarea(ByRef t As mm.Topic) As String
        If Len(getfirstarea(t)) > 0 Then
            getdashboardarea = getfirstarea(t)
        Else
            If t.CalloutTopics.Count > 0 Then
                getdashboardarea = getfirstarea(t.CalloutTopics.Item(1))
            Else
                getdashboardarea = ""
            End If
        End If
        If getdashboardarea = "" And t.CalloutTopics.Count > 0 Then
            If t.CalloutTopics.Item(1).CalloutTopics.Count > 0 Then
                If t.CalloutTopics.Item(1).CalloutTopics.Item(1).Task.Complete < 100 Then
                    getdashboardarea = getfirstarea(t.CalloutTopics.Item(1).CalloutTopics.Item(1))
                End If
            End If
        End If
    End Function

End Module
