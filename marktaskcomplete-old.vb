Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module marktaskcompleteModule
    'mark_task_complete.mmbas:  mark and log tasks done and advance repeating tasks
    '04Feb11 http://creativecommons.org/licenses/by-nc-nd/3.0/
    'http://wiki.activityowner.com/index.php?Title=Mark_Task_Complete
    'recent changes
    '07Feb09 -- change version check frequency to 30 by default
    '10Feb09 -- error trapping, organize options
    '11Feb09 -- allow completed and reference text to be international, move version check to end, error trap
    '13Feb09 -- make sure configuration map has new completed and reference text
    '13Feb09 -- fix bug on floating complete topic
    '15Feb09 -- speed up option loading
    '16Feb09 -- save maps when not using dashboard
    '25Feb09 -- delete relationships from completed tasks moved to completed topic
    '26Feb09 -- Fix bug -- not marking non-repeat items done from dashboard if not moved (speeded up with code change in ao_common.mmbas)
    '03Mar09 -- take out english-specific hard coding, rely on configuration map
    '14Apr09 -- Add "ongoing" category
    '18Apr09 -- Add deleteoriginal configuration option
    '19Apr09 -- change version check strategy
    '03May09 -- add each3weeks
    '01Aug09 -- make map saving optional (not recommended, but improves compatibility with some other add-ins)
    '30Dec09 -- close hidden maps to avoid multicomputer conflicts
    '30Dec09 -- debug
    '05Feb10 -- fix run time error when used in catalyst
    '07Feb10 -- add check icon to items marked complete
    '04Feb11 -- adapt to MM9 date handling bugs
    '#uses "ao_common.mmbas"
    Structure myoptiontype
        Dim logbasename As String
        Dim completedtext As String
        Dim referencetext As String
        Dim logname As String
        Dim movetocomplete As Boolean
        Dim storecompleteinproject As Boolean
        Dim storeinresult As Boolean
        Dim mapcalendar As Boolean
        Dim setduetoday As Boolean
        Dim logdone As Boolean
        Dim savedb As Boolean
        Dim deleteoriginal As Boolean
        Dim savemaps As Boolean
    End Structure
    Public Const advancecodetext = "Advance Codes"
    Public Const daycodetext = "Day Codes"
    Sub Marktaskcomplete(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Const ProgramVersion = "20100207"
        Const VersionCheckLink = "http://activityowner.com/installers/versioncheck.php"
        Const configmapname = "ao\CompletedConfig.mmap"
        Const maxselected = 100
        Dim DefaultOptions As myoptiontype
        Dim opt As myoptiontype
        Dim localopt As myoptiontype
        Dim configdoc As mm.Document
        Dim isdb As Boolean
        Dim t As mm.Topic
        Dim tt As mm.Topic
        Dim redo As Boolean
        Dim selected(maxselected) As mm.Topic
        Dim doccurrent As mm.Document
        Dim i As Integer
        Dim scount As Integer
        Dim tmap As mm.Document
        Dim rawmap As mm.Document
        doccurrent = m_app.ActiveDocument
        configdoc = getmap(m_app, configmapname)
        tmap = Nothing
        If configdoc Is Nothing Then
            MsgBox("Configuration Map not created. Contact ActivityOwner. Exiting Program.")
            Exit Sub
        End If
        ' set tasks to be completed
        scount = doccurrent.Selection.Count
        isdb = f_IsADashboardMap(doccurrent)
        For i = 1 To scount
            selected(i) = doccurrent.Selection.Item(i)
        Next
        Upgrade(m_app, configdoc)
        DefaultOptions = LoadDefaultOptions(configdoc)
        doccurrent.Activate()
        If Not isdb Then opt = LoadLocalOptions(DefaultOptions, doccurrent)
        '
        'loop through tasks to be completed
        For i = 1 To scount
            t = selected(i)
            redo = advancetask(t, False, configdoc) 'determine if it is a repeating task
            If Not isdb Then
                If opt.logdone Then
                    tmap = getmap(m_app, opt.logname)
                    copytocalendarlog(tmap, t)  'copy to completed Log
                End If
                If opt.mapcalendar Then copytoRefcalendarlog(m_app.ActiveDocument, t, opt.referencetext)
                If Not redo Then
                    t.Task.Complete = 100
                    If opt.setduetoday Then t.Task.DueDate = Today
                End If
                If opt.movetocomplete Then
                    If redo Then
                        copytocompletedtopic(t, doccurrent, opt)
                    Else
                        movetocompletedtopic(t, doccurrent, opt)
                    End If
                End If
                If redo Then advancetask(t, True, configdoc)
                If Not redo Then If opt.deleteoriginal Then t.Delete()
            Else 'isdb
                For Each tt In doccurrent.Range(mm.MmRange.mmRangeAllTopics)
                    If isclone(tt, t) Then
                        If redo Then
                            tt.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconCheck)
                        Else
                            tt.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconCheck)
                            tt.Task.Complete = 100
                            If opt.setduetoday Then t.Task.DueDate = Today
                        End If
                    End If
                Next
                t = followedhyperlink(m_app, t, doccurrent)
                If Not (t Is Nothing) Then
                    rawmap = t.Document
                    opt = LoadLocalOptions(DefaultOptions, rawmap)
                    If opt.logdone Then
                        tmap = getmap(m_app, opt.logname)
                        copytocalendarlog(tmap, t)  'copy to completed Log
                    End If
                    If opt.mapcalendar Then copytoRefcalendarlog(t.Document, t, opt.referencetext)
                    If Not redo Then t.Task.Complete = 100
                    If opt.movetocomplete Then
                        If redo Then
                            copytocompletedtopic(t, t.Document, opt)
                        Else
                            If opt.setduetoday Then t.Task.DueDate = Today
                            movetocompletedtopic(t, t.Document, opt)
                        End If
                    End If
                    If redo Then advancetask(t, True, configdoc)
                    If Not redo Then If opt.deleteoriginal Then t.Delete()
                    If Not rawmap.ExternalDocument.IsExternal And opt.savemaps Then rawmap.Save()
                    rawmap = Nothing
                End If
            End If
        Next
        '
        If isdb Then doccurrent.Activate()
        'On Error Resume Next
        If (Not doccurrent.ExternalDocument.IsExternal) And (opt.savemaps Or (isdb And opt.savedb)) Then doccurrent.Save()
        If opt.logdone And opt.savemaps Then tmap.Save()

        On Error GoTo 0
        doccurrent = Nothing
        tmap = Nothing
        t = Nothing
        For i = 1 To scount
            selected(i) = Nothing
        Next
        VersionCheck(VersionCheckLink, "Mark_Task_Complete", ProgramVersion, configdoc)
        configdoc = Nothing
        CloseHiddenMaps(m_app)
        PlaySoundchirp()
    End Sub
    Function LoadDefaultOptions(ByRef configdoc As mm.Document) As myoptiontype
        Dim optionbranch As mm.Topic
        optionbranch = createmainbranch("options", configdoc, "")
        With LoadDefaultOptions
            .logbasename = getoption("log-map-base-name", configdoc, optionbranch)
            .logname = .logbasename & Microsoft.VisualBasic.Year(Today) & "-" & Microsoft.VisualBasic.Month(Today) & ".mmap"
            .referencetext = getoption("referencetext", configdoc, optionbranch)
            .completedtext = getoption("completedtext", configdoc, optionbranch)
            .movetocomplete = optiontrue("move-complete-to-branch", configdoc, optionbranch) 'move completed tasks to "complete" branch or floating Topic
            .storecompleteinproject = optiontrue("store-complete-in-project", configdoc, optionbranch)     'store completed tasks in each project instead of floating topic
            .storeinresult = optiontrue("store-in-result", configdoc, optionbranch)   'store in result if possible
            .logdone = optiontrue("copy-completed-to-log-map", configdoc, optionbranch)
            .mapcalendar = optiontrue("copy-completed-to-calendar-branch", configdoc, optionbranch)
            .savedb = optiontrue("save-dashboards", configdoc, optionbranch) 'save time by not saving dashboard after each use
            .setduetoday = optiontrue("setduetoday", configdoc, optionbranch)   'set duedate to today completed tasks
            .deleteoriginal = optiontrue("delete-original", configdoc, optionbranch)
            .savemaps = optiontrue("save-maps", configdoc, optionbranch)
        End With
        optionbranch = Nothing
    End Function
    Function LoadLocalOptions(ByRef DefaultOptions As myoptiontype, ByRef currentdoc As mm.Document) As myoptiontype
        Dim rnote As String
        Dim rbranch As mm.topic
        rbranch = findmainbranch(DefaultOptions.referencetext, currentdoc)
        LoadLocalOptions = DefaultOptions
        If Not rbranch Is Nothing Then
            rnote = rbranch.Notes.Text
            If Not rnote = "" Then
                With LoadLocalOptions
                    .movetocomplete = optionlocal("move-complete-to-branch", rnote, DefaultOptions.movetocomplete)
                    .storecompleteinproject = optionlocal("store-complete-in-project", rnote, DefaultOptions.storecompleteinproject)
                    .storeinresult = optionlocal("store-in-result", rnote, DefaultOptions.storeinresult)
                    .logdone = optionlocal("copy-completed-to-log-map", rnote, DefaultOptions.logdone)
                    .mapcalendar = optionlocal("copy-completed-to-calendar-branch", rnote, DefaultOptions.mapcalendar)
                    .savedb = optionlocal("save-dashboards", rnote, DefaultOptions.savedb)
                    .setduetoday = optionlocal("setduetoday", rnote, DefaultOptions.setduetoday)
                    .deleteoriginal = optionlocal("delete-original", rnote, DefaultOptions.deleteoriginal)
                End With
            End If
        End If
    End Function
    Function optionlocal(ByRef setting As String, ByRef rnotes As String, ByVal oDefault As Boolean) As Boolean
        'future use to override default settings by looking at text in a note in map
        If InStr(rnotes, setting) > 0 Then
            If InStr(rnotes, setting & "=0") > 0 Then
                Debug.Print("false it")
                optionlocal = False
            ElseIf InStr(rnotes, setting & "=1") > 0 Then
                optionlocal = True
            Else
                optionlocal = True
            End If
        Else
            optionlocal = oDefault
        End If
    End Function
    Function advancetask(ByRef t As mm.Topic, ByRef changedates As Boolean, ByVal configdoc As mm.Document) As Boolean
        Const sdd = False       'set due date even if no due date in place
        Dim codetopic As mm.topic
        Dim units As String
        Dim lead As Integer
        Dim inc As Integer
        Dim datefixed As Boolean
        Dim hasstart As Boolean
        Dim hasdue As Boolean
        Dim i As Integer
        Dim m As Integer
        Dim d As Integer
        Dim cats As String
        Dim a As mm.topic
        cats = LCase(t.Task.Categories)
        advancetask = False
        If Len(cats) > 0 Then 'don't bother with searching for repeat codes if categories are blank
            a = createmainbranch(advancecodetext, configdoc, "")
            For Each codetopic In a.AllSubTopics
                If InStr(cats, codetopic.Text) > 0 Then
                    advancetask = True
                    If changedates Then
                        m = i
                        datefixed = (codetopic.AllSubTopics(3).Text = "-1")
                        hasstart = Not isdate0(t.Task.StartDate)
                        hasdue = Not isdate0(t.Task.DueDate)
                        units = codetopic.AllSubTopics.Item(1).Text
                        inc = CInt(Val(codetopic.AllSubTopics.Item(2).Text))
                        lead = CInt(Val(codetopic.AllSubTopics.Item(3).Text))

                        If datefixed And hasdue Then t.Task.DueDate = DateAdd(units, inc, t.Task.DueDate)
                        If Not datefixed And hasdue Then t.Task.DueDate = DateAdd(units, inc, Today)
                        If sdd And Not hasdue Then t.Task.DueDate = DateAdd(units, inc, Today)

                        If datefixed And hasstart Then setstartdate(t, DateAdd(units, inc, t.Task.StartDate))
                        If Not datefixed And hasstart Then setstartdate(t, DateAdd(DateInterval.Day, -lead, DateAdd(units, inc, Today)))
                        If sdd And Not hasstart Then setstartdate(t, DateAdd(units, inc - lead, Today))

                        If Not hasstart And Not hasdue And Not sdd Then MsgBox("You should Set either an initial start Date Or an initial due Date")
                        t.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconRedo)
                    End If
                    Exit Function
                End If
            Next

            'If advance codes not found, look for day of week code
            a = createmainbranch(daycodetext, configdoc, "")
            For Each codetopic In a.AllSubTopics
                If InStr(LCase(cats), LCase(codetopic.Text)) > 0 Then
                    advancetask = True
                    If changedates Then
                        d = CInt(Val(Trim(codetopic.Notes.Text)))
                        If d > 0 Then
                            If Microsoft.VisualBasic.Weekday(Today) <= 1 Then
                                t.Task.DueDate = DateAdd("d", d - Microsoft.VisualBasic.Weekday(Today), Today)
                            Else
                                t.Task.DueDate = DateAdd("d", d - Microsoft.VisualBasic.Weekday(Today) + 7, Today)
                            End If
                            setstartdate(t, DateAdd("d", -3, t.Task.DueDate))
                            t.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconRedo)
                        End If
                    End If
                    Exit Function
                End If
            Next

            'if day of week code not found, look for end of codes
            If InStr(LCase(cats), getoption("endofmonthcode", configdoc, Nothing)) > 0 Or InStr(LCase(cats), getoption("endofquartercode", configdoc, Nothing)) > 0 Then
                t.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconRedo)
                advancetask = True
                If changedates Then
                    If InStr(cats, getoption("endofmonthcode", configdoc, Nothing)) > 0 Then
                        If Not isdate0(t.Task.DueDate) Then
                            t.Task.DueDate = DateAdd("m", 1 - Microsoft.VisualBasic.Day(t.Task.DueDate), t.Task.DueDate)
                        Else
                            If sdd Then t.Task.DueDate = DateAdd("m", 1 - Microsoft.VisualBasic.Day(Today), Today)
                        End If
                        If Not isdate0(t.Task.StartDate) Then
                            setstartdate(t, DateAdd("m", 1 - Microsoft.VisualBasic.Day(t.Task.StartDate), t.Task.StartDate))
                        End If
                    ElseIf InStr(cats, getoption("endofquartercode", configdoc, Nothing)) > 0 Then
                        If Not isdate0(t.Task.DueDate) Then
                            t.Task.DueDate = DateAdd("m", (4 - (Month(t.Task.DueDate) Mod 3) Mod 12) - Microsoft.VisualBasic.Day(t.Task.DueDate), t.Task.DueDate)
                        Else
                            If sdd Then t.Task.DueDate = DateAdd("m", (4 - (Month(Now) Mod 3) Mod 12) - Microsoft.VisualBasic.Day(Now), Today)
                        End If
                        If Not isdate0(t.Task.StartDate) Then
                            setstartdate(t, DateAdd("m", (4 - (Month(t.Task.StartDate) Mod 3) Mod 12) - Microsoft.VisualBasic.Day(t.Task.StartDate), t.Task.StartDate))
                        End If
                    End If
                End If
            End If
        End If
    End Function
    Sub movetocompletedtopic(ByRef t As mm.Topic, ByRef doccurrent As mm.Document, ByRef opt As myoptiontype)
        Dim r As mm.Relationship
        Dim completedtopic As mm.topic
        If Not t Is Nothing Then
            If Not t.IsCentralTopic Then
                completedtopic = createcompletedtopic(doccurrent, t, opt)
                If Not completedtopic Is Nothing Then
                    completedtopic.SubTopics(True).Insert(t)
                    For Each r In t.AllRelationships
                        r.Delete()
                    Next
                    completedtopic.Collapsed = True
                    completedtopic = Nothing
                Else
                    MsgBox("Completed topic not found/created")
                End If
            End If
        Else
            MsgBox("topic is empty")
        End If
    End Sub
    Sub copytocompletedtopic(ByRef t As mm.Topic, ByRef doccurrent As mm.Document, ByVal opt As myoptiontype)
        Dim completedtopic As mm.topic
        Dim tt As mm.topic
        If Not t.IsCentralTopic Then
            completedtopic = createcompletedtopic(doccurrent, t, opt)
            tt = completedtopic.AddSubTopic("")
            tt.Xml = t.Xml
            tt.Task.DueDate = Today
            tt.Task.Complete = 100
            completedtopic.Collapsed = True
            completedtopic = Nothing
        End If
    End Sub
    Function createcompletedtopic(ByRef doccurrent As mm.Document, ByRef donetask As mm.Topic, ByRef opt As myoptiontype) As mm.Topic
        'this function finds or creates a floating completed topic or a topic under the parent result or project depending on the options chosen
        Dim t As mm.Topic
        Dim p As mm.Topic
        Dim picon As mm.Icon
        createcompletedtopic = Nothing
        If Not opt.storecompleteinproject Then
            For Each t In doccurrent.AllFloatingTopics
                If LCase(t.Text) = LCase(opt.completedtext) Then
                    t.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconNoEntry)
                    createcompletedtopic = t
                End If
            Next
            If createcompletedtopic Is Nothing Then
                createcompletedtopic = doccurrent.AllFloatingTopics.Add
                createcompletedtopic.Text = opt.completedtext
                createcompletedtopic.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconNoEntry)
            End If
        Else
            p = donetask
            While Not p.IsCentralTopic 'jump out when found or at central
                p = p.ParentTopic
                For Each picon In p.Icons
                    If picon.Name = "CustomIcon-2051436099" Then
                        Exit While
                    End If
                    If picon.Name = "CustomIcon--1181845906" And opt.storeinresult Then
                        Exit While
                    End If
                Next
            End While
            For Each t In p.AllSubTopics
                If LCase(t.Text) = LCase(opt.completedtext) Then 'how to avoid false positives here?
                    createcompletedtopic = t
                    Exit Function
                End If
            Next
            If createcompletedtopic Is Nothing Then
                createcompletedtopic = p.AddSubTopic(opt.completedtext)
                createcompletedtopic.Icons.AddStockIcon(mm.MmStockIcon.mmStockIconNoEntry)
            End If
        End If
        t = Nothing
        p = Nothing
        picon = Nothing
    End Function
    Sub Upgrade(ByRef m_app As Mindjet.MindManager.Interop.Application, ByRef configdoc As mm.Document)
        'Adds new branches and keywords to existing mark_task_complete_config.mmap.  Change "lastupgrade" entry to avoid doing twice.
        Const currentversion = 20090801
        Dim a As mm.topic
        Dim lastupgrade As String
        Dim RunUpgrade As Boolean
        lastupgrade = getoption("lastupgrade", configdoc, Nothing)
        If lastupgrade = "" Then
            lastupgrade = "0"
            RunUpgrade = True
        ElseIf inteval(m_app, lastupgrade) < currentversion Then 'do not want to eval("")
            RunUpgrade = True
        End If
        If RunUpgrade Then
            '
            configdoc.CentralTopic.Text = "Mark Task Complete Configuration Map"
            If MsgBox("Mark_task_complete needs to make some upgrades to your Configuration Map. This will take a minute.", vbOKCancel) = vbCancel Then Exit Sub
            'OPTIONS-----------------------------------------------------------
            createoption("log-map-base-name", "ao\Completed", configdoc)
            createoption("completedtext", "Completed", configdoc)
            createoption("referencetext", "Reference", configdoc)
            createoption("move-complete-to-branch", "1", configdoc)
            createoption("store-complete-in-project", "1", configdoc)
            createoption("store-in-result", "1", configdoc)
            createoption("copy-completed-to-log-map", "1", configdoc)
            createoption("copy-completed-to-calendar-branch", "1", configdoc)
            createoption("save-dashboards", "1", configdoc)
            createoption("setduetoday", "0", configdoc)
            createoption("versioncheckfrequency", "30", configdoc) 'check version weekly
            createoption("lastversioncheck", Str(Today), configdoc)
            createoption("endofmonthcode", "endofmonth", configdoc)
            createoption("endofquartercode", "endofquarter", configdoc)
            createoption("delete-original", "0", configdoc)
            createoption("save-maps", "1", configdoc)

            '
            a = createmainbranch("Visit http://wiki.activityowner.com/index.php?title=Mark_Task_Complete for explanation of configuration options and how to setup repeating tasks", configdoc, "")
            a.CreateHyperlink("http://wiki.activityowner.com/index.php?title=Mark_Task_Complete")

            a = createmainbranch(advancecodetext, configdoc, "")
            a.Notes.Text = "Categories keywords used to advance tasks. 1st entry is units, 2nd entry tells how much to advance, 3rd entry tells how many units to start before due date.  -1 = advance from duedate instead of done date (e.g. for mortgage)"
            addtriplet(a, "daily", "d", "1", "0", "20090105.1", lastupgrade)  'keyword, units, increment, lead time, added with upgrade, lastupgrade
            addtriplet(a, "everytwo", "d", "2", "1", "20090105.1", lastupgrade)
            addtriplet(a, "weekly", "d", "7", "3", "20090105.1", lastupgrade)
            addtriplet(a, "eachweek", "d", "7", "-1", "20090105.1", lastupgrade)
            addtriplet(a, "monthly", "m", "1", "10", "20090105.1", lastupgrade)
            addtriplet(a, "eachmonth", "m", "1", "-1", "20090105.1", lastupgrade)
            addtriplet(a, "biannual", "m", "6", "30", "20090105.1", lastupgrade)
            addtriplet(a, "yearly", "m", "12", "30", "20090105.1", lastupgrade)
            addtriplet(a, "eachyear", "m", "12", "-1", "20090105.1", lastupgrade)
            addtriplet(a, "quarterly", "m", "3", "30", "20090105.1", lastupgrade)
            addtriplet(a, "eachquarter", "m", "3", "-1", "20090105.1", lastupgrade)
            addtriplet(a, "everyotherweek", "d", "14", "7", "20090105.1", lastupgrade)
            addtriplet(a, "everythreedays", "d", "3", "1", "20090105.1", lastupgrade)
            addtriplet(a, "everyfivedays", "d", "5", "2", "20090105.1", lastupgrade)
            addtriplet(a, "everyothermonth", "m", "2", "7", "20090105.1", lastupgrade)
            addtriplet(a, "every3weeks", "d", "21", "7", "20090105.1", lastupgrade)
            addtriplet(a, "every2weeks", "d", "14", "7", "20090105.1", lastupgrade)
            addtriplet(a, "each2weeks", "d", "14", "-1", "20090105.1", lastupgrade)
            addtriplet(a, "each3weeks", "d", "21", "-1", "20090503", lastupgrade)
            addtriplet(a, "fortnightly", "d", "14", "7", "20090105.1", lastupgrade)
            addtriplet(a, "eachfortnight", "d", "14", "-1", "20090105.1", lastupgrade)
            addtriplet(a, "ongoing", "d", "0", "0", "20090414", lastupgrade)

            a = createmainbranch(daycodetext, configdoc, "")
            addkeyword(a, "sunday", "1", "20090303", lastupgrade)
            addkeyword(a, "monday", "2", "20090303", lastupgrade)
            addkeyword(a, "tuesday", "3", "20090303", lastupgrade)
            addkeyword(a, "wednesday", "4", "20090303", lastupgrade)
            addkeyword(a, "thursday", "5", "20090303", lastupgrade)
            addkeyword(a, "friday", "6", "20090303", lastupgrade)
            addkeyword(a, "saturday", "7", "20090303", lastupgrade)

            '---------------------------------------------------------------
            checkforduplicates(configdoc)
            'Mark map as upgraded
            setoption("lastupgrade", Str(currentversion), configdoc)
            If configdoc.IsModified Then configdoc.Save()
            a = Nothing
            MsgBox("Configuration Update complete")
        End If
    End Sub

End Module
