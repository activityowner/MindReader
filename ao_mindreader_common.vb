Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
'bugs add the blob commented out
Module ao_mindreader_common
    'ao_mindreader_common
    'Copyright: http://creativecommons.org/licenses/by-nc-nd/3.0/
    'Information: http://wiki.activityowner.com/index.php?Title=MindReader
    '
    'primary routines are mindreaderNLP and mindreaderopen
    'these are called by ao_mindreaderNLP.mmba and ao_mindreaderOpen.mmba
    '29Jan09 -- added use of optiontrue to reduce configuration typos
    '11Feb09 -- move version checks to after program runs.
    '12Feb09 -- eliminate need for rm:sendchanges, update gyroq.ini as well
    '12Feb09 -- fix bugs! (log as 13Feb version)
    '15Feb09 -- missed some items in dashboard m tag mode
    '17Feb09 -- fix bug in speedup, don't close delete me maps if they have main topics
    '22Feb09 -- fix historical issues with %me instead of %me% in configuration
    '29Feb09 -- Trap hyperlink error in outlinker transfer
    '03Mar09 -- remove unsupported togglelist code
    '05Mar09 -- add keyword to attach a file [attach:fname] -- initially used by outlinker, fix bug
    '13Mar09 -- revise category keywords to eliminate overlap with due date keywords
    '17Mar09 -- add a language setting to options so upgrade can differentiate between maps in future
    '03Apr09 -- avoid missing resources -- some risk of false positives
    '11Apr09 -- fix naming of mindreaderconfig.mmap in error messages
    '19Apr09 -- change version check strategy
    '26Apr09 -- adding funnel delimiters
    '03May09 -- add parallel funnel delimiter, add error trapping, isstalled, issnagged keywords, handling of .. in names
    '17Jul09 -- add support for onenote hyperlinks
    '15Nov09 -- update error checking
    '16Nov09 -- deal with 8.2 bug
    '25Nov09 -- deal with notes character substition issue in 8.2
    '26Nov09 -- make outlook hyperlinks absolute (8.2 issue?)
    '27Nov09 -- don't try to save mjc documents
    '27Dec09 -- trap internal browser error with o command
    '30Dec09 -- close hidden files
    '30Dec09 -- trap error when underlying topic has been moved for m tag2
    '17Jan10 -- try to deal with 64bit upgrades or dual use on own rather than failing
    '21Jan10 -- add option to not save maps so much
    '23Jan10 -- ao_common updated to address bug in multiline processing (don't close hidden config map)
    '31Oct10 -- adjust code for setting start date to deal with MM9 behavior
    '01Nov10 -- Don't send due dates to underlying maps if they are the same as start date
    '20Mar11 -- changed source of sample config map
    '22Mar11 -- don't add start date to central topics by default
    '02Apr11 -- fix new map bug
    '07Apr11 -- use chr(32) for space character instead of " "
    '#uses "ao_common.mmba"
    Structure mroptionstype
        Dim confirmmarkup As Boolean        'true = play sound when finished (set in options)
        Dim AddStart As Boolean             'True = add start dates
        Dim AddBlob As Boolean              'Add a "blob" icon to task that don't have an identified context
        Dim AutoDelete As Boolean 'delete keywords like "next week" from strings
        Dim rmMe As String
        Dim EnglishSpeedUp As Boolean 'skip branches unless primary text is present -- will break international config maps
        Dim autosave As Boolean   'turn off frequent map saving during mindreader
    End Structure
    'functions geared toward centralizing path and file names
    Function MindReaderFolderPath(ByRef m_app As Mindjet.MindManager.Interop.Application) As String
        MindReaderFolderPath = m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps) & "ao\"
    End Function
    Function MindReaderConfigMapFullName(ByRef m_app As Mindjet.MindManager.Interop.Application) As String
        Const ConfigName = "mindreaderconfig.mmap"
        'if const has a full path then use it, otherwise add leading path
        If isabsolute(ConfigName) Then
            MindReaderConfigMapFullName = ConfigName
        Else
            MindReaderConfigMapFullName = MindReaderFolderPath(m_app) & ConfigName
        End If
    End Function

    Function getmroptions(ByRef configdoc As mm.Document) As mroptionstype
        Dim optionbranch As mm.Topic
        optionbranch = createmainbranch("options", configdoc, "")
        With getmroptions
            .confirmmarkup = optiontrue("confirmmarkup", configdoc, optionbranch)
            .AddBlob = optiontrue("addblob", configdoc, optionbranch)
            .AddStart = optiontrue("addstart", configdoc, optionbranch)
            .AutoDelete = optiontrue("autodelete", configdoc, optionbranch)
            .rmMe = getoption("me", configdoc, optionbranch)
            .EnglishSpeedUp = optiontrue("englishspeedup", configdoc, optionbranch)
            .autosave = optiontrue("autosave", configdoc, optionbranch)
        End With
    End Function

    Sub MindReaderNLP(ByVal m_app As Mindjet.MindManager.Interop.Application, ByVal cmd As String)
        'sw("start")
        Const ProgramVersion = "20101101"
        Const VersionCheckLink = "http://activityowner.com/installers/versioncheck.php"
        Dim opt As mroptionstype
        Dim isdb As Boolean
        Dim usingMtag As Boolean                'if processing text with m tag

        Dim ActiveDoc As mm.Document           'Document being marked up
        Dim DocCurrent As mm.Document          'one of open documents
        Dim configdoc As mm.Document  'mindreader configuration file

        Dim BranchTopic As mm.Topic            'one of main topics of configuration map
        Dim ActiveTopic As mm.Topic            'one of selected topics
        Dim UnderlyingTopic As mm.Topic        'topic dashboard topic is pointing to
        Dim resourcelist As mm.Topic  'try leaving this in place for session
        Dim t As mm.Topic
        Dim fname As String

        Dim ParseText As String             'Text to mindread (either from selected topic or passed by command line)
        'Dim cat As String                   'Category text of selected topic
        Dim mrmapStr As String              'Filename of configuration file

        'change this line if you want to store your mindreader.mmap file in a location other than default "My Maps" directory
        mrmapStr = MindReaderConfigMapFullName(m_app)
        usingMtag = Not cmd = "" 'text passed in by "m" tag is "Command" variable to mark up selected tasks
        If usingMtag Then Debug.Print("using m tag")
        ActiveDoc = m_app.ActiveDocument
        isdb = f_IsADashboardTopic(ActiveDoc.Selection.PrimaryTopic)
        On Error Resume Next
        configdoc = OpenMapHidden(m_app, mrmapStr)
        If Err.Number > 0 Then Err.Clear()
        On Error GoTo 0
        If configdoc Is Nothing Then
            Install_or_Migrate_MindReader_Config(m_app)
            configdoc = OpenMapHidden(m_app, mrmapStr)
        End If
        Upgrade(m_app, configdoc)
        '
        '
        'sw("loading options")
        opt = getmroptions(configdoc)
        'sw("options loaded")
        ActiveDoc.Activate()
        If ActiveDoc.IsReadOnly Then MsgBox("This map is read only.  Your changes will not be saved!")
        resourcelist = createmainbranch("resourcelist", configdoc, "")
        '
        'set and get options from mindreader, exit when done
        If usingMtag Then
            If InStr(ParseText, "setoption:") = 1 Then
                usersetoption(configdoc, ParseText)
                Exit Sub
            ElseIf InStr(ParseText, "getoption:") = 1 Then
                usergetoption(configdoc, ParseText)
                Exit Sub
            ElseIf InStr(ParseText, "listlinks") = 1 Then
                t = createmainbranch("links", configdoc, "")
                configdoc.Selection.Set(t)
                configdoc.Selection.Copy()
                MsgBox(Clipboard.GetText)
                Exit Sub
            End If
        End If
        'experimental code to allow funnel entry
        Const funneldelimFollows = "<<"
        Const FunnelDelimPrecedes = ">>"
        Const FunnelDelimParallel = "&&"
        Dim temptext As String
        Dim temptopic As mm.Topic
        If Not usingMtag Then
            For Each ActiveTopic In ActiveDoc.Selection
                temptopic = ActiveTopic
                While InStr(temptopic.Text, funneldelimFollows) > 0
                    temptext = temptopic.Text
                    temptopic.Text = Left(temptext, InStr(temptext, funneldelimFollows) - 1)
                    temptopic = temptopic.AddSubTopic(Right(temptext, Len(temptext) - InStr(temptext, funneldelimFollows) - Len(funneldelimFollows) + 1))
                    ActiveDoc.Selection.Add(temptopic)
                End While
            Next
            For Each ActiveTopic In ActiveDoc.Selection
                temptopic = ActiveTopic
                While InStr(temptopic.Text, FunnelDelimPrecedes) > 0
                    temptext = temptopic.Text
                    temptopic.Text = Right(temptext, Len(temptext) - InStrRev(temptext, FunnelDelimPrecedes) - 1)
                    temptopic = temptopic.AddSubTopic(Left(temptext, InStrRev(temptext, FunnelDelimPrecedes) - 1))
                    ActiveDoc.Selection.Add(temptopic)
                End While
            Next
            For Each ActiveTopic In ActiveDoc.Selection
                temptopic = ActiveTopic
                While InStr(temptopic.Text, FunnelDelimParallel) > 0
                    temptext = temptopic.Text
                    temptopic.Text = Left(temptext, InStr(temptext, FunnelDelimParallel) - 1)
                    temptopic = temptopic.ParentTopic.AddSubTopic(Right(temptext, Len(temptext) - InStr(temptext, FunnelDelimParallel) - Len(FunnelDelimParallel) + 1))
                    ActiveDoc.Selection.Add(temptopic)
                End While
            Next
        End If
        '
        For Each ActiveTopic In ActiveDoc.Selection
            If InStr(ActiveTopic.Text, "[attach:") > 1 Then
                On Error Resume Next
                fname = ActiveTopic.Text
                fname = Right(fname, Len(fname) - InStr(fname, "[attach:") - 9)
                fname = Left(fname, InStr(fname, "]") - 1)
                Debug.Print(fname)
                ActiveTopic.Attachments.Add(fname)
                If Not Err.Number = 0 Then
                    MsgBox("file not attached")
                    Err.Clear()
                Else
                    'remove hyperlink perhaps make optional
                    If ActiveTopic.HasHyperlink Then
                        If InStr(LCase(ActiveTopic.Hyperlink.Address), "outlook:") > 0 Then
                            ActiveTopic.Hyperlink.Delete()
                        End If
                    End If
                    'should delete at some point with successful transfer
                End If
                On Error GoTo 0
            End If
            If isdb Then
                sw("following hyperlink")
                UnderlyingTopic = followedhyperlink(m_app, ActiveTopic, ActiveDoc)
                If UnderlyingTopic Is Nothing Then Debug.Print("underlyingtopic is nothing")
            End If
            'cat = ActiveTopic.Task.Categories  'avoid loss of existing categories
            If InStr(ActiveTopic.Text, "(Jott to Self) ") > 0 Then ActiveTopic.Text = Replace(ActiveTopic.Text, "(Jott to Self) ", "")

            If usingMtag Then
                ParseText = makereplacements(cmd, configdoc)  'Allow mark-up of topics based on a entered string instead of topic content ("m" tag)
            Else
                ActiveTopic.Text = makereplacements(ActiveTopic.Text, configdoc)
                ParseText = ActiveTopic.Text
            End If

            SetDueDate(ActiveTopic, ParseText, configdoc, opt.AutoDelete, usingMtag)
            If Not usingMtag Then
                If ((Not isdate0(ActiveTopic.Task.DueDate)) Or opt.AddStart) And isdate0(ActiveTopic.Task.StartDate) Then
                    If Not ActiveTopic.IsCentralTopic Then
                        setstartdate(ActiveTopic, Today)
                    End If
                End If
            End If
            If Not usingMtag And opt.AddBlob Then ActiveTopic.Icons.RemoveStockIcon(Mindjet.MindManager.Interop.MmStockIcon.mmStockIconMarker7) 'Remove the "Blob" icon when reprocessing

            If ActiveTopic.Task.Complete = -1 Then ActiveTopic.Task.Complete = 0
            'sw("entering branch loop")
            For Each BranchTopic In configdoc.CentralTopic.AllSubTopics
                mindreadtopic(m_app, ParseText, BranchTopic, ActiveTopic, UnderlyingTopic, resourcelist, usingMtag, configdoc, opt)
            Next
            'sw("leaving branch loop")

            'MsgBox("before add the blob")
            'If opt.AddBlob And Not usingMtag Then addtheblob(ActiveTopic, cat, configdoc, "")
            RemoveBracketedText(ActiveTopic)
            If Not usingMtag Then autodeletekeywords(ActiveTopic, configdoc)
            If isdb Then
                sw("sending to underlying map")
                If Not UnderlyingTopic Is Nothing Then
                    With UnderlyingTopic.Task
                        .Resources = ActiveTopic.Task.Resources
                        If Not (ActiveTopic.Task.DueDate = ActiveTopic.Task.StartDate) Then 'avoid issue with mm9 false due dates on dashboards
                            .DueDate = ActiveTopic.Task.DueDate
                        End If
                        setstartdate(UnderlyingTopic, ActiveTopic.Task.StartDate)
                        .Complete = ActiveTopic.Task.Complete
                    End With
                    UnderlyingTopic.Text = ActiveTopic.Text
                    If Not isworkspacemap(UnderlyingTopic.Document) Then
                        UnderlyingTopic.Document.Save()
                    End If
                Else
                    Debug.Print("topic missing")
                End If
                ActiveDoc.Activate()
            End If
            'testing approach of adding outlook attachments

            'If InStr(LCase(ParseText), "aey") > 0 Then Call MacroRun(getoption("resultsmanagerpath", configdoc) & "ResultManager-X5-Edit.MMBas") 'supplement with edit dialog
            If Not isdb Then
                On Error Resume Next
                If ActiveTopic.Document.IsModified And opt.autosave Then ActiveTopic.Document.Save()
                If Err.Number > 0 Then Err.Clear()
                On Error GoTo 0
            End If
        Next
        If configdoc.IsModified Then
            On Error Resume Next
            configdoc.Save()
            If Err.Number > 0 Then Err.Clear()
            On Error GoTo 0
        End If
        If Not usingMtag Then
            For Each DocCurrent In m_app.Documents
                If DocCurrent.CentralTopic.Text = "DeleteMe" Then
                    If DocCurrent.CentralTopic.AllSubTopics.Count = 0 Then
                        DocCurrent.Close()  'guard against issues
                    Else
                        MsgBox("Programming Error" & DocCurrent.Name & " map has delete me in central topic.")
                    End If
                End If
            Next
        End If
        ActiveDoc = Nothing
        ActiveTopic = Nothing
        UnderlyingTopic = Nothing
        BranchTopic = Nothing
        resourcelist = Nothing
        DocCurrent = Nothing
        'sw("done with main")
        If opt.confirmmarkup Then PlaySoundchirp()
        If usingMtag Then VersionCheck(VersionCheckLink, "MindReader", ProgramVersion, configdoc)
        configdoc = Nothing
        CloseHiddenMaps(m_app)
    End Sub
    Sub mindreaderopen(ByRef m_app As mm.Application, ByVal odoc As mm.Document, ByVal cmd As String)
        '**********************
        'We communicate from GyroQ to SAX basic by creating a temporary map
        ' as the activedocument and using its centraltopic and notes
        'notes=blank if being used by "o" to open a map based on keyword
        '     ="1" if used by "q" to open a destination map for a queued item
        '     ="2" if used by "s" to cut/paste topics to a destination map
        'Keyword comparisons are done in lower case
        '***********************
        'Code below allows function to be called directly by MindReaderCall.mmba instead of GyroQ
        'This enables the non-queued "o" and "s" commands to work under v7 and to work faster in v6
        '
        'Booleans
        Dim issend As Boolean   'being used by "s" macro
        Dim isQueue As Boolean  'being used by "q"
        Dim usetextgrab As Boolean 'use textgrab instead of clipboard
        Dim SplitText As Boolean 'split multi-line text into separate topics
        Dim found As Boolean    'true if in-tray and later link keyword is found
        Dim ReturnOnSend As Boolean  'Return to source document after send
        'strings
        Dim aStr As String      'text passed by command line
        Dim restStr As String   'rest of multiline string
        Dim currentmapname As String 'used in mjc hack
        'Documents
        Dim msgdoc As mm.Document  'temp map used to pass in text
        Dim Doc As mm.Document     'mindreader.mmap file
        Dim OriginalDoc As mm.Document
        Dim Destdoc As mm.Document 'Destination document
        Dim DocCurrent As mm.Document 'One of currently open documents: used to search for deleteme map
        'topics
        Dim linktopic As mm.Topic  'topic identified as maintopic of "links" branch
        Dim itopic As mm.Topic     'in-tray topic
        Dim t As mm.Topic
        Dim tt As mm.Topic
        Dim mtopic As mm.Topic     'one of main topics in destination map
        'Links
        Dim defaultlink As mm.Hyperlink
        Dim link As mm.Hyperlink
        '
        Dim i As Integer        'counter used in linktopic search
        Dim d As Integer 'index of defaultmap branch
        Dim sw As Double
        '

        ReturnOnSend = True
        SplitText = False
        On Error Resume Next
        Doc = OpenMapHidden(m_app, MindReaderConfigMapFullName(m_app))
        If Err.Number > 0 Then Err.Clear()
        On Error GoTo 0
        If Doc Is Nothing Then
            Install_or_Migrate_MindReader_Config(m_app)
            Doc = OpenMapHidden(m_app, MindReaderConfigMapFullName(m_app))
        End If

        If Not cmd = "" Then
            'if called by macrorun, figure out what to do based on 1st 5 characters
            If InStr(cmd, "/send") = 1 Then
                issend = True
                isQueue = False
                OriginalDoc = m_app.ActiveDocument
            ElseIf InStr(cmd, "/open") = 1 Then
                issend = False
                isQueue = False
            ElseIf InStr(cmd, "/queu") = 1 Then
                issend = False
                isQueue = True
            End If
            'trim off the command code
            aStr = Right(cmd, Len(cmd) - 5)
        Else
            'if called from GyroQ, read the temp map for information
            msgdoc = m_app.ActiveDocument
            If msgdoc Is Nothing Then
                MsgBox("no active document")
                Exit Sub
            End If
            aStr = msgdoc.CentralTopic.Text
            isQueue = msgdoc.CentralTopic.Notes.Text = "1"
            issend = msgdoc.CentralTopic.Notes.Text = "2"
            ReturnOnSend = False  'not possible with slower version of send
            '***********************
            'mark the message map for later deletion. If we delete it now, v7 gets confused
            '***********************
            msgdoc.CentralTopic.Text = "DeleteMe"
        End If
        restStr = aStr
        While Len(restStr) > 0 'process more than one line of text at a time
            If InStr(restStr, Chr(10)) > 0 Then
                aStr = Left(restStr, InStr(restStr, Chr(10)) - 1)
                restStr = Right(restStr, Len(restStr) - InStr(restStr, Chr(10)))
                SplitText = True
            Else
                aStr = restStr
                restStr = ""
                SplitText = False
            End If
            '**************************
            'Open the mindreader.mmap file to match keyword to link
            'Handle situation where mindreader.mmap is missing
            '******************************

            usetextgrab = (Not getoption("usetextgrab", Doc, Nothing) = "0") 'default to use textgrab unless explicitly set not to
            '***********************
            'Search for for map-keyword-link branch of mindreader.mmap
            '***********************
            linktopic = createmainbranch("links", Doc, "")
            link = destinationlink(linktopic, makereplacements(aStr, Doc))
            '**************************
            'if mindreader.mmap has a v7 topic hyperlink, put it there instead of in in-tray
            '**************************
            found = InStr(link.Address, "mj-map") = 1 Or ismindjetconnectlink(link)
            '**************************
            'Catch error if hyperlink doesn't work
            '**************************
            On Error Resume Next 'disable error checking when following mindmanager hyperlink
            If found Then
                link.Follow()
                'the following code is a hack to deal with the delay in link.follow bringing up
                'mindjet connect workspaces
                If ismindjetconnectlink(link) Then
                    currentmapname = m_app.ActiveDocument.Name
                    For i = 1 To 150
                        If Not m_app.ActiveDocument.Name = currentmapname Then Exit For
                        Pause(0.1)
                    Next
                End If
                Destdoc = m_app.ActiveDocument
                If Destdoc.IsReadOnly Then GoTo E4
            End If
            '
            On Error GoTo E2
            If Not found Then
                If InStr(link.Address, "mmap") > 0 Then
                    If link.Absolute Or isabsolute(link.Address) Then
                        Destdoc = m_app.Documents.Open(link.Address)
                    Else
                        Destdoc = m_app.Documents.Open(MindReaderFolderPath(m_app) & link.Address, "", True)
                    End If
                Else
                    link.Follow()
                    If isQueue Or issend Then 'catch in a new map
                        Destdoc = m_app.Documents.Add
                    End If
                End If
            End If

            '**************************
            'Error Handling
            'If can't open mindreaderconfig.mmap or follow hyperlink, create a temp
            'map to "catch" the incoming task from gyroQ
            '**************************
            If Not Err.Number = 0 Then
E1:             MsgBox("mindreaderconfig.mmap Open Error:" & Err.Description) : GoTo X
E2:             'Debug.Print(Err.Number)
                If Not Err.Number = -2147467259 Then
                    MsgBox("Error trying to follow hyperlink in mindreaderconfig.mmap:" & Err.Description)
                End If
                GoTo X
E4:             MsgBox("Destination map is read only") : GoTo X
X:              If Not Err.Number = -2147467259 Then
                    Destdoc = m_app.Documents.Add
                End If
                Err.Clear()
            End If
            On Error GoTo 0
            '*********************************
            'If "o" in use, then we are done.
            'If "q" or "s" in use (add or add2 true), select topic for subsequent macro to add topics to
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            If isQueue Or issend Then
                '*******************************
                'If topic hyperlink was found set that to destination topic
                '*******************************
                If found Then
                    itopic = Destdoc.Selection.PrimaryTopic
                Else
                    '************************
                    'if not a topic hyperlink, search for in-tray
                    '*************************
                    For Each mtopic In Destdoc.CentralTopic.AllSubTopics
                        If mtopic.TextLabels.ContainsTextLabel("In-tray") Or mtopic.TextLabels.ContainsTextLabel("in-tray") Then
                            itopic = mtopic
                            found = True
                            Exit For
                        End If
                    Next
                End If
                '****************************
                'If in-tray not found, create one
                If Not found Then
                    itopic = Destdoc.CentralTopic.AddSubTopic("in-Tray")
                    addmarker(itopic, Nothing, "Areas", "in-tray", False)
                    'itopic.Task.Complete = 0
                End If
                '************************
                'If "q" in use select intray for subsequent macro to add items to
                If isQueue Then
                    t = itopic.AddSubTopic(aStr)
                    If Len(Clipboard.GetText) > 0 Then 'q and fq tags load clipboard with textgrab
                        If (InStr(Clipboard.GetText, "http") = 1 Or InStr(LCase(Clipboard.GetText), "onenote:") = 1) And Not InStr(Clipboard.GetText, " ") > 0 Then
                            t.CreateHyperlink(Clipboard.GetText)
                        Else
                            t.Notes.Text = Clipboard.GetText
                        End If
                    End If
                    Destdoc.Selection.Set(t)
                End If
                '*************************
                'If "s" in use, paste items to in-tray
                If issend Then
                    If Not odoc Is Nothing Then copymarkers(odoc, Destdoc)
                    Destdoc.Selection.Set(itopic)
                    Destdoc.Selection.Paste()
                End If
                On Error GoTo Y
                'Destdoc.Save
                If Err.Number > 0 Then
Y:                  MsgBox("Error Trying to Save Destination Document" & Err.Description)
                End If
                'end of "if isQueue or isSend" section
            End If
            On Error GoTo 0
            If isQueue Then 'isqueue will delete deleteme map after mindreadernlp.mmba and also save destdoc so don't do it now
                'Bring destination map to foreground if necessary
                Destdoc.Activate()
                If SplitText Then MindReaderNLP(m_app, "")
            Else
                For Each DocCurrent In m_app.Documents
                    If DocCurrent.CentralTopic.Text = "DeleteMe" Then
                        DocCurrent.Close()
                        Exit For
                    End If
                Next
            End If
            '
            If issend Then
                'OriginalDoc.Save
                'Destdoc.Save
                If ReturnOnSend Then
                    OriginalDoc.Activate()
                Else
                    Destdoc.Activate()
                End If
                On Error Resume Next
                For Each Doc In m_app.Documents
                    If Doc.IsModified And Not isworkspacemap(Doc) And issaved(Doc) Then
                        Doc.Save()
                    End If
                Next
                On Error GoTo 0
            End If
        End While
        'clean up pointers
        msgdoc = Nothing
        Doc = Nothing   'line causes problems?
        linktopic = Nothing
        defaultlink = Nothing
        link = Nothing
        t = Nothing
        tt = Nothing
        itopic = Nothing
        Destdoc = Nothing
        mtopic = Nothing
        DocCurrent = Nothing
        link = Nothing
        OriginalDoc = Nothing
    End Sub

    '*********************************************************************************************************************************
    Sub mindreadtopic(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal ParseText As String, ByRef BranchTopic As mm.Topic, ByRef ActiveTopic As mm.Topic, ByRef UnderlyingTopic As mm.Topic, ByRef resourcelist As mm.Topic, ByRef usingMtag As Boolean, ByRef configdoc As mm.Document, ByRef opt As mroptionstype)
        'This is the heart of the mindreaderNLP program.  If keyword in topic matches, it uses notes associated with it to mark up topic
        Dim match As Integer    'location of keyword in text being processed
        Dim mStr As String      'suBranchTopic notes
        Dim remaining As String
        Dim BracketText As String
        Dim NonBracketText As String
        Dim lcaseparsetext As String
        Dim cboard As String
        Dim r As mm.Topic          'Resources
        Dim found As Boolean
        Dim resource As String
        Dim matchbracket As Integer
        Dim KeywordTopic As mm.Topic
        Dim keyword As String
        Dim finalresource As String
        Dim delim As String
        Dim cstart As Integer
        Dim cend As Integer
        Dim inbracket As Boolean
        Dim completion As Integer 'use -1 For non Task
        Dim i As Integer
        Dim branchname As String
        Dim firsttry As Boolean
        Dim isdb As Boolean
        '-----------------------------
        branchname = BranchTopic.Text
        lcaseparsetext = LCase(ParseText)
        isdb = f_IsADashboardTopic(ActiveTopic)
        If opt.EnglishSpeedUp Then
            If branchname = "extend" And InStr(lcaseparsetext, "extend") = 0 Then Exit Sub
            If branchname = "advance" And InStr(lcaseparsetext, "advance") = 0 Then Exit Sub
            If branchname = "start" And InStr(lcaseparsetext, "start in") = 0 Then Exit Sub
            If branchname = "priority" And InStr(lcaseparsetext, "p") = 0 Then Exit Sub
            If branchname = "delay" And InStr(lcaseparsetext, "delay") = 0 Then Exit Sub
        End If
        If branchname = "links" Then Exit Sub
        If branchname = "options" Then Exit Sub
        If branchname = "resourcelist" Then Exit Sub
        If Len(Trim(branchname)) = 0 Then Exit Sub
        '-----------------
        finalresource = ""
        'separate out bracketed and nonbracketed text
        BracketText = ""
        NonBracketText = ""
        inbracket = False
        For i = 1 To Len(ParseText)
            If Mid(ParseText, i, 1) = "[" Then
                inbracket = True
            ElseIf Mid(ParseText, i, 1) = "]" Then
                inbracket = False
                BracketText = BracketText & " "
                If Len(NonBracketText) > 0 Then NonBracketText = NonBracketText & " "
            ElseIf inbracket Then
                BracketText = BracketText & Mid(ParseText, i, 1)
            Else
                NonBracketText = NonBracketText & Mid(ParseText, i, 1)
            End If
        Next
        '
        For Each KeywordTopic In BranchTopic.AllSubTopics
            keyword = LCase(KeywordTopic.Text)
            If InStr(keyword, ":") > 0 Then
                match = InStr(LCase(LTrim(NonBracketText)), keyword)
                matchbracket = InStr(LCase(LTrim(BracketText)), keyword)
            ElseIf InStr(LCase(BranchTopic.Text), "resourceverbs") > 0 Then
                match = InStr(LCase(LTrim(NonBracketText)), keyword & " ") 'avoid false positives
                matchbracket = InStr(LCase(LTrim(BracketText)), keyword & " ")
            Else
                match = InStr(LCase(LTrim(NonBracketText)), keyword) 'avoid false positives
                matchbracket = InStr(LCase(LTrim(BracketText)), keyword)
            End If
            If match > 0 Or matchbracket > 0 Then 'require resource verbs to lead main or bracketed text
                mStr = KeywordTopic.Notes.Text
                Select Case LCase(BranchTopic.Text)
                    Case "contexts"
                        If match = 1 Or matchbracket > 0 Then
                            addmarker(ActiveTopic, UnderlyingTopic, "Contexts", mStr, False)
                        End If
                        'cat = cat & ",@" & mStr
                    Case "resourceverbs"
                        If Not InStr(keyword, ":") > 0 Then
                            If match > 0 Then
                                remaining = Mid(NonBracketText, InStr(LCase(NonBracketText), keyword) + Len(keyword & " "))
                            ElseIf matchbracket > 0 Then
                                remaining = Mid(BracketText, InStr(LCase(BracketText), keyword) + Len(keyword & " "))
                            End If
                        Else
                            If match > 0 Then
                                remaining = Mid(NonBracketText, InStr(LCase(NonBracketText), keyword) + Len(keyword))
                            ElseIf matchbracket > 0 Then
                                remaining = Mid(BracketText, InStr(LCase(BracketText), keyword) + Len(keyword))
                            End If
                        End If
                        found = False
                        For Each r In resourcelist.AllSubTopics
                            If InStr(LCase(remaining), LCase(r.Text)) = 1 Then
                                found = True
                                resource = r.Notes.Text
                            End If
                        Next
                        If Not found Then
                            If match = 1 Or matchbracket > 0 Or (match > 0 And InStr(keyword, ":") > 0) Then
                                resource = FirstWord(remaining)
                                found = True
                            End If
                        End If
                        If found Then
                            If opt.rmMe = "" Then opt.rmMe = "%me%"
                            If opt.rmMe = "%me" Then opt.rmMe = "%me%"
                            If finalresource = "" Then delim = "" Else delim = ","
                            If mStr = "partner" Then
                                If Not InStr(finalresource, "@" & resource) > 0 Then
                                    If optiontrue("atresource", configdoc, Nothing) Then
                                        finalresource = "@" & resource & delim & finalresource
                                    Else
                                        finalresource = resource & "@" & delim & finalresource
                                    End If
                                End If
                            End If
                            If mStr = "waiting" Then finalresource = finalresource & delim & resource
                            If mStr = "delegated" Then finalresource = finalresource & delim & opt.rmMe & "," & resource
                            If mStr = "owe" Then finalresource = resource & "," & opt.rmMe & delim & finalresource
                        End If

                    Case "dates"
                        Try
                            ActiveTopic.Task.DueDate = DateValue(eval(m_app, mStr))
                            If isdb Then UnderlyingTopic.Task.DueDate = ActiveTopic.Task.DueDate
                        Catch
                            MsgBox("error parsing due date info from configuration map:" & mStr)
                        End Try
                    Case "delay"
                        If usingMtag Then
                            Try
                                If Not isdate0(ActiveTopic.Task.DueDate) Then
                                    ActiveTopic.Task.DueDate = DateAdd(DateInterval.Day, inteval(m_app, mStr), ActiveTopic.Task.DueDate)

                                    If isdb Then UnderlyingTopic.Task.DueDate = ActiveTopic.Task.DueDate
                                End If
                                If Not isdate0(ActiveTopic.Task.StartDate) Then setstartdate(ActiveTopic, DateAdd(DateInterval.Day, inteval(m_app, mStr), ActiveTopic.Task.StartDate))
                            Catch
                                MsgBox("error parsing delay info from configuration map")
                            End Try
                        End If
                    Case "advance"
                        If usingMtag Then
                            Try
                                If Not isdate0(ActiveTopic.Task.StartDate) Then setstartdate(ActiveTopic, DateAdd(DateInterval.Day, -inteval(m_app, mStr), ActiveTopic.Task.StartDate))
                                If Not isdate0(ActiveTopic.Task.DueDate) Then
                                    ActiveTopic.Task.DueDate = DateAdd(DateInterval.Day, -inteval(m_app, mStr), ActiveTopic.Task.DueDate)
                                    If isdb Then UnderlyingTopic.Task.DueDate = ActiveTopic.Task.DueDate
                                End If
                            Catch
                                MsgBox("error parsing advance info from configuration map")
                            End Try
                        End If

                    Case "extend"
                        Try
                            If Not isdate0(ActiveTopic.Task.DueDate) Then
                                ActiveTopic.Task.DueDate = DateAdd(DateInterval.Day, inteval(m_app, mStr), ActiveTopic.Task.DueDate)
                                If isdb Then UnderlyingTopic.Task.DueDate = ActiveTopic.Task.DueDate
                            End If
                        Catch
                            MsgBox("error parsing due date info from configuration map")
                        End Try

                    Case "start in"
                        Try
                            If isdate0(ActiveTopic.Task.DueDate) Or DateDiff(DateInterval.Day, Today, ActiveTopic.Task.DueDate) >= inteval(m_app, mStr) Then
                                setstartdate(ActiveTopic, DateAdd(DateInterval.Day, inteval(m_app, mStr), Today))
                            Else
                                MsgBox("Can not set Start date after due date")
                            End If
                        Catch
                            MsgBox("error parsing start in info from configuration map")
                            Err.Clear()
                        End Try
                    Case "starting"
                        Try
                            If inteval(m_app, mStr) < 1000 Then 'then assume it is relative to due date 
                                If Not isdate0(ActiveTopic.Task.DueDate) Then
                                    setstartdate(ActiveTopic, DateAdd(DateInterval.Day, -inteval(m_app, mStr), ActiveTopic.Task.DueDate))
                                End If
                                'If inteval(m_app, mStr) = -1 Then ' -1 means elminate start date
                                'setstartdate(ActiveTopic, DateValue("12/31/1899"))
                                'End If
                            Else  'allow code in mStr to set the actual start date
                                setstartdate(ActiveTopic, DateValue(eval(m_app, mStr)))
                            End If
                        Catch
                            MsgBox("error parsing starting date info from configuration map")
                        End Try
                    Case "icons"
                        Try
                            ActiveTopic.Icons.AddStockIcon(CType(inteval(m_app, mStr), Mindjet.MindManager.Interop.MmStockIcon))
                            If Not UnderlyingTopic Is Nothing Then UnderlyingTopic.Icons.AddStockIcon(CType(inteval(m_app, mStr), Mindjet.MindManager.Interop.MmStockIcon))
                        Catch
                            MsgBox("error parsing icon info from configuration map" & inteval(m_app, mStr))
                        End Try
                    Case "completion"
                        ActiveTopic.Task.Complete = inteval(m_app, mStr)
                    Case "customicons"
                        firsttry = True
                        Try
                            mStr = Replace(mStr, Chr(160), Chr(32))
                            Trace.WriteLine(mStr)
                            ActiveTopic.Icons.AddCustomIcon(mStr)
                            Trace.WriteLine("disabled adding customicon")
                            If Not UnderlyingTopic Is Nothing Then
                                UnderlyingTopic.Icons.AddCustomIcon(mStr)
                            End If
                        Catch
                            Try
                                If firsttry Then
                                    firsttry = False
                                    If InStr(mStr, "(x86)") > 0 Then
                                        mStr = Replace(mStr, " (x86)", "")
                                    Else
                                        mStr = Replace(mStr, "iles\", "iles (x86)")
                                    End If
                                    mStr = Replace(mStr, Chr(160), Chr(32))
                                    ActiveTopic.Icons.AddCustomIcon(mStr)
                                    If Not UnderlyingTopic Is Nothing Then
                                        UnderlyingTopic.Icons.AddCustomIcon(mStr)
                                    End If
                                End If
                            Catch
                                Trace.WriteLine("error with icons add")
                            End Try
                        End Try
                    Case "priority"
                        Try
                            ActiveTopic.Task.Priority = CType(inteval(m_app, mStr), Mindjet.MindManager.Interop.MmTaskPriority)
                            If Not UnderlyingTopic Is Nothing Then UnderlyingTopic.Task.Priority = ActiveTopic.Task.Priority
                        Catch
                            MsgBox("error parsing priority info from configuration map")
                            Err.Clear()
                        End Try
                    Case "area"
                        If match = 1 Or matchbracket > 0 Then
                            addmarker(ActiveTopic, UnderlyingTopic, "Areas", mStr, False)
                        End If
                    Case "categories"
                        If match = 1 Or matchbracket > 0 Then
                            addmarker(ActiveTopic, UnderlyingTopic, "Categories", mStr, False)
                        End If
                    Case "clips"
                        If Len(Clipboard.GetText) > 0 Then
                            cboard = Replace(Clipboard.GetText, "*nl*", vbCrLf) 'convert *nl* crlf codes from outlinker
                            If mStr = "olmsg" Then
                                cboard = Replace(Clipboard.GetText, "*nl*", vbCrLf)
                                Try
                                    ActiveTopic.CreateHyperlink(Left(cboard, InStr(Clipboard.GetText, "|") - 1))
                                    ActiveTopic.Hyperlink.Absolute = True
                                    ActiveTopic.Notes.Text = Mid(cboard, InStr(cboard, "|") + 1)
                                Catch
                                    MsgBox("error processing hyperlink")
                                    ActiveTopic.Notes.Text = Clipboard.GetText
                                    Err.Clear()
                                End Try
                            End If
                            If Not getoption("usetextgrab", configdoc, Nothing) = "0" Then 'default to usetextgrab if option not set
                                If mStr = "link" And (Not ActiveTopic.HasHyperlink) Then
                                    If InStr(Replace(LCase(ParseText), "outlinker", ""), keyword) > 0 Then 'avoid false postive on "outlinker"
                                        If Not (InStr(LCase(ParseText), "olmsg") > 0) Then 'avoid adding outlinker clipboard
                                            ActiveTopic.CreateHyperlink(fixclipboard(cboard))
                                        End If
                                    End If
                                End If
                                If mStr = "note" And ActiveTopic.Notes.Text = "" Then ActiveTopic.Notes.Text = fixclipboard(cboard)
                            End If

                        End If
                End Select
            End If
        Next

        If LCase(BranchTopic.Text) = "contexts" Then
            addmarker(ActiveTopic, UnderlyingTopic, "Contexts", addhardcoded(ParseText, "@", ""), False)
        End If
        If LCase(BranchTopic.Text) = "area" Then
            addmarker(ActiveTopic, UnderlyingTopic, "Areas", addhardcoded(ParseText, "^", ""), False)
        End If
        If LCase(BranchTopic.Text) = "category" Then
            addmarker(ActiveTopic, UnderlyingTopic, "Categories", addhardcoded(ParseText, "~", ""), False)
        End If
        If LCase(BranchTopic.Text) = "resourceverbs" Then 'look for undefined partner resources and make final resource assignment
            cstart = 0
            If InStr(ParseText, "@") = Len(ParseText) Then 'if at end of sentence
                cend = Len(ParseText) - 1
            ElseIf InStr(ParseText, "@ ") > 0 Then 'if not at beginning of sentence
                cend = InStr(ParseText, "@ ") - 1
            ElseIf InStr(ParseText, "@]") > 0 Then
                cend = InStr(ParseText, "@]") - 1
            ElseIf InStr(ParseText, "@") > 0 And Not InStr(ParseText, " @") > 0 Then
                cend = InStr(ParseText, "@") - 1
            End If
            If cend > 0 Then
                If finalresource = "" Then delim = "" Else delim = ","
                If Len(LastWord(Left(ParseText, cend))) > 0 Then
                    finalresource = finalresource & delim & LastWord(Left(ParseText, cend)) & "@"
                End If
            End If
            If Not InStr(finalresource, opt.rmMe) > 0 Then
                If Not usingMtag Then
                    If optiontrue("defaultownerme", configdoc, Nothing) Then
                        finalresource = opt.rmMe & "," & finalresource
                    End If
                End If
            End If

            'Resource setting strategy
            '1. Only overwrite resources if resourceverbs were identified
            '2. If defaultownerme, add owner if not present in string already (e.g. in delegated or I owe form)
            '3. Don't overwrite resource if only "me," in finalresource
            If Not finalresource = "" Then
                If Not finalresource = opt.rmMe & "," Then
                    ActiveTopic.Task.Resources = finalresource
                Else
                    finalresource = opt.rmMe
                    If ActiveTopic.Task.Resources = "" Then
                        ActiveTopic.Task.Resources = finalresource
                    End If
                End If
            End If
        End If
        r = Nothing
        KeywordTopic = Nothing
        Exit Sub
    End Sub


    

    Function addhardcoded(ByRef ParseText As String, ByRef symbol As String, ByRef cat As String) As String
        Dim cstart As Integer
        Debug.Print("in add hardcoded" & ParseText)
        cstart = 0
        If InStr(ParseText, symbol) = 1 Then
            cstart = 2
        ElseIf InStr(ParseText, " " & symbol) > 0 Then
            cstart = InStr(ParseText, " " & symbol) + 2
        ElseIf InStr(ParseText, "[" & symbol) > 0 Then
            cstart = InStr(ParseText, "[" & symbol) + 2
        End If
        If Len(cat) = 0 Then
            If cstart > 0 Then
                If symbol = "~" Then
                    cat = FirstWord(Mid(ParseText, cstart))
                Else
                    cat = symbol & FirstWord(Mid(ParseText, cstart))
                End If
            End If
        Else
            If cstart > 0 Then
                If symbol = "~" Then
                    cat = cat & ", " & FirstWord(Mid(ParseText, cstart))
                Else
                    cat = cat & ", " & symbol & FirstWord(Mid(ParseText, cstart))
                End If
            End If
        End If
        addhardcoded = cat
        Debug.Print(addhardcoded)
    End Function
    Sub SetDueDate(ByRef ActiveTopic As mm.Topic, ByRef ParseText As String, ByRef configdoc As mm.Document, ByRef AutoDelete As Boolean, ByRef usingMtag As Boolean)
        'Sets due date on active topic based on presence of #date string# in ParseText
        Dim s As Integer
        Dim e As Integer
        Dim delim As String
        Dim found As Boolean
        Dim dvalue As Date
        Dim dstring As String
        '
        'legacy approach to dates
        delim = getoption("datedelimiter", configdoc, Nothing)
        If delim = "" Then delim = "#"
        s = InStr(ParseText, delim)
        e = InStrRev(ParseText, delim)
        On Error Resume Next
        If s > 0 And e > s Then dvalue = DateValue(Mid(ParseText, s + 1, e - s - 1))
        If LCase(ActiveTopic.Text) = LCase(ParseText) Then
            If Err.Number = 0 And s > 0 And e > 0 Then
                ActiveTopic.Task.DueDate = dvalue
                ActiveTopic.Text = Replace(ActiveTopic.Text, Mid(ActiveTopic.Text, s, e - s + 1), "")
            End If
        End If
        Err.Clear()
        '
        'look for dates in / / format
        Dim i As Integer
        Dim firstslash As Integer
        Dim secondslash As Integer
        Dim sfound As Boolean
        Dim dfound As Boolean
        If InStr(ParseText, "/") > 0 And Not (InStr(ParseText, "/") = InStrRev(ParseText, "/")) Then
            i = 1
            sfound = False
            firstslash = 0
            secondslash = 0
            While Not sfound And i <= Len(ParseText)
                If Mid(ParseText, i, 1) = "/" Then
                    If firstslash = 0 Then
                        firstslash = i
                    Else
                        secondslash = i
                    End If
                    If (secondslash - firstslash) = 3 Or (secondslash - firstslash) = 2 Then 'look for  xx/xx/xx or x/x/xx or x/x/xxxx
                        sfound = True
                    Else
                        If secondslash > 0 Then
                            firstslash = secondslash
                            secondslash = 0
                        End If
                    End If
                End If
                i = i + 1
            End While
        End If
        If sfound Then
            'take lazy approach -- try some things until no error
            Err.Clear()
            dfound = False
            dstring = middlestr(ParseText, firstslash - 2, secondslash + 5)
            dvalue = DateValue(dstring) 'look for xx/xx/xxxx
            If Err.Number = 0 Then
                dfound = True
            Else
                Err.Clear()
                dstring = middlestr(ParseText, firstslash - 2, secondslash + 3)
                dvalue = DateValue(dstring) 'look for xx/xx/xx
                If Err.Number = 0 Then
                    dfound = True
                Else
                    Err.Clear()
                    dstring = middlestr(ParseText, firstslash - 1, secondslash + 5)
                    dvalue = DateValue(dstring) 'look for x/x/xxxx
                    If Err.Number = 0 Then
                        dfound = True
                    Else
                        Err.Clear()
                        dstring = middlestr(ParseText, firstslash - 1, secondslash + 3)
                        dvalue = DateValue(dstring) 'look for x/x/xx
                        If Err.Number = 0 Then
                            dfound = True
                        End If
                    End If
                End If
            End If
        End If
        If Err.Number > 0 Then Err.Clear()
        On Error GoTo 0
        If dfound Then
            If Not ActiveTopic.Task.DueDateReadOnly Then
                ActiveTopic.Task.DueDate = dvalue
                If AutoDelete Then ActiveTopic.Text = Replace(ActiveTopic.Text, dstring, "")
            Else
                MsgBox("Duedate readonly")
            End If
        End If
        Err.Clear()
        Exit Sub
    End Sub
    Sub RemoveBracketedText(ByRef ActiveTopic As mm.Topic)
        'Remove bracketed text from the active topic
        Dim BracketText As String
        Dim NonBracketText As String
        Dim inbracket As Boolean
        Dim ParseText As String
        Dim i As Integer
        BracketText = ""
        NonBracketText = ""
        ParseText = ActiveTopic.Text
        inbracket = False
        If InStr(ParseText, "[") > 0 Then
            For i = 1 To Len(ParseText)
                If Mid(ParseText, i, 1) = "[" Then
                    inbracket = True
                ElseIf Mid(ParseText, i, 1) = "]" Then
                    inbracket = False
                    BracketText = BracketText & " "
                    If Len(NonBracketText) > 0 Then NonBracketText = NonBracketText & " "
                ElseIf inbracket Then
                    BracketText = BracketText & Mid(ParseText, i, 1)
                Else
                    NonBracketText = NonBracketText & Mid(ParseText, i, 1)
                End If
            Next
            ActiveTopic.Text = NonBracketText
        End If
    End Sub
    Function expandword(ByVal sometext As String) As String
        'expands a "single.word" into a "double word"
        'replace .. with . for names with dots in them
        'eliminate trailing portion of email addresses being expanded
        If InStr(sometext, "..") > 0 Then
            sometext = Replace(sometext, "..", "." & Chr(32))
        Else
            sometext = Replace(sometext, ".", Chr(32))
        End If
        If InStr(sometext, "@") > 0 Then
            sometext = Left(sometext, InStr(sometext, "@") - 1)
        End If
        expandword = sometext
    End Function
    Function FirstWord(ByVal sometext As String) As String
        sometext = Replace(sometext, ":", "")
        sometext = Replace(sometext, "]", "")
        sometext = Replace(sometext, "[", " ")
        If InStr(sometext, " ") > 0 Then
            sometext = Left(sometext, InStr(sometext, " ") - 1)
        End If
        FirstWord = expandword(sometext)
    End Function
    Function LastWord(ByVal sometext As String) As String
        sometext = Replace(sometext, "[", " ")
        sometext = Replace(sometext, "]", " ")
        If InStr(sometext, " ") > 0 Then
            sometext = Mid(sometext, InStrRev(sometext, " ") + 1)
        End If
        LastWord = expandword(sometext)
    End Function
    Function FirstBracketWord(ByVal sometext As String) As String
        If InStr(sometext, "]") > 0 Then
            sometext = Left(sometext, InStr(sometext, "]") - 1)
        End If
        'sometext = Replace(sometext,"]","")
        's = Mid(s,InStr(s,"[")+1)
        's=Trim(s)
        'If InStr(s, " ")>0 Then
        '	s = Left(s,InStr(s," ")-1)
        'End If
        FirstBracketWord = expandword(sometext)
    End Function
    Sub addtheblob(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal ActiveTopic As mm.Topic, ByVal cat As String, ByVal configdoc As mm.Document)
        'MindReader has an option to add a "blob" to a task that doesn't have a good action verb
        Dim resulticon As String    'results icon
        Dim projecticon As String   'project icon
        Dim somedayicon As Integer  'Someday icon code
        somedayicon = 48
        On Error GoTo iconerror
        projecticon = m_app.Utilities.GetCustomIconSignature(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\resultmanager-projecticon.ico")
        resulticon = m_app.Utilities.GetCustomIconSignature(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\resultmanager-resulticon.ico")
        On Error GoTo 0
        If Err.Number > 0 Then
iconerror:
            Err.Clear()
            If Not (InStr(cat, "@") > 0 _
             Or ActiveTopic.Task.Resources <> "" _
             Or ActiveTopic.Icons.Count > 0 _
               Or ActiveTopic.Task.Complete = -1) Then ActiveTopic.Icons.AddStockIcon(Mindjet.MindManager.Interop.MmStockIcon.mmStockIconMarker7)
            Exit Sub
        End If
        '
        If Not (InStr(cat, "@") > 0 _
          Or ActiveTopic.Task.Resources <> "" _
            Or ActiveTopic.Icons.ContainsCustomIcon(resulticon) _
            Or ActiveTopic.Icons.ContainsCustomIcon(projecticon) _
            Or ActiveTopic.Task.Complete = -1) Then ActiveTopic.Icons.AddStockIcon(Mindjet.MindManager.Interop.MmStockIcon.mmStockIconMarker7)
        Exit Sub
    End Sub
    Function fixclipboard(ByVal clipboardstr As String) As String
        'This function removes the _textgrab_ from the tail of the clipboard for versions of gyroQ prior to implementation of this feature.
        Dim textgrab As String
        textgrab = "_textgrab_"
        If InStrRev(clipboardstr, textgrab) = (Len(clipboardstr) - Len(textgrab) + 1) Then
            fixclipboard = Mid(clipboardstr, 1, Len(clipboardstr) - Len(textgrab))
        Else
            fixclipboard = clipboardstr
        End If
    End Function

    Sub autodeletekeywords(ByRef atopic As mm.Topic, ByRef configdoc As mm.Document)
        Dim BranchTopic As mm.Topic
        Dim t As mm.Topic
        Dim txt As String
        If optiontrue("autodelete", configdoc, Nothing) Then
            For Each BranchTopic In configdoc.CentralTopic.AllSubTopics
                txt = BranchTopic.Text
                If txt = "CustomIcons" Or _
                   txt = "customicons" Or _
                   txt = "dates" Or _
                   txt = "starting" Or _
                   txt = "priority" Or _
                   txt = "completion" Or _
                   txt = "start in" Or _
                   txt = "category" Or _
                   txt = "clips" Then
                    For Each t In BranchTopic.AllSubTopics
                        If Len(atopic.Text) > 0 And Len(t.Text) > 0 Then
                            If InStr(atopic.Text, t.Text) > 0 Then
                                atopic.Text = Replace(atopic.Text, t.Text, "")
                            End If
                        End If
                    Next
                End If
            Next
            If Len(atopic.Text) > 0 Then atopic.Text = Replace(atopic.Text, "!", "")
            If Len(atopic.Text) > 0 Then atopic.Text = Replace(atopic.Text, "someday", "")
            RemoveHardCodedString(atopic, "@")
            RemoveHardCodedString(atopic, "^")
            RemoveHardCodedString(atopic, "~")
            RemoveHardCodedString(atopic, "R:")
            RemoveHardCodedString(atopic, "r:")
        End If
        If Len(atopic.Text) > 0 Then
            atopic.Text = Replace(atopic.Text, "  ", " ") 'remove double spaces
        End If
    End Sub
    Sub RemoveHardCodedString(ByRef atopic As mm.Topic, ByRef symbol As String)
        Dim cstart As Integer
        cstart = 0
        If InStr(atopic.Text, symbol) = 1 Then
            cstart = 2
        ElseIf InStr(atopic.Text, " " & symbol) > 0 Then
            cstart = InStr(atopic.Text, " " & symbol) + 2
        ElseIf InStr(atopic.Text, "[" & symbol) > 0 Then
            cstart = InStr(atopic.Text, "[" & symbol) + 2
        End If
        If cstart > 0 Then
            atopic.Text = Replace(atopic.Text, symbol & FirstWord(Mid(atopic.Text, cstart)), "")
        End If
    End Sub
    Function makereplacements(ByVal sometext As String, ByRef configdoc As mm.Document) As String
        Dim BranchTopic As mm.Topic
        Dim KeywordTopic As mm.Topic
        Dim start As Integer
        BranchTopic = createmainbranch("alias", configdoc, "Substitute longer strings for short aliases")
        For Each KeywordTopic In BranchTopic.AllSubTopics
            start = InStr(LCase(sometext), LCase(KeywordTopic.Text))
            If start > 0 Then
                sometext = Replace(sometext, Mid(sometext, start, Len(KeywordTopic.Text)), KeywordTopic.Notes.Text)
            End If
        Next
        sometext = Replace(sometext, "%27%27", Chr(34))
        sometext = Replace(sometext, "%27", Chr(39))
        makereplacements = sometext
    End Function

    Sub Upgrade(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal configdoc As mm.Document)
        'Adds new branches and keywords to existing mindreader.mmap.  Change "lastupgrade" entry to avoid doing twice.
        Dim a As mm.Topic
        Dim lastupgrade As String
        Dim RunUpgrade As Boolean
        'NOTE currentversion is incremented when configuration map changes need to be made.  It will trail the programversion
        Const currentversion = "20130915"
        lastupgrade = getoption("lastupgrade", configdoc, Nothing)
        If lastupgrade = "" Then
            RunUpgrade = True
        ElseIf Val(lastupgrade) < Val(currentversion) Then 'do not want to eval("")
            RunUpgrade = True
        End If
        If RunUpgrade Then
            If MsgBox("MindReader needs to make some upgrades to your MindReader Configuration Map. This will take a minute.", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
            'OPTIONS-----------------------------------------------------------
            createoption("confirmmarkup", "1", configdoc)
            createoption("datedelimiter", "#", configdoc)
            createoption("addstart", "1", configdoc)
            createoption("addblob", "0", configdoc)
            deleteoption("resultsmanagerpath", configdoc)
            createoption("resultsmanagerpath", "C:\Program Files (x86)\Gyronix\Gyronix ResultsManager v3 for Mindjet 11\", configdoc)
            createoption("gyroqpath", "C:\Program Files\Gyronix\GyroQ\", configdoc)
            createoption("atresource", "0", configdoc)
            createoption("language", "English", configdoc)
            If getoption("me", configdoc, Nothing) = "%me" Then
                deleteoption("me", configdoc)
            End If
            createoption("me", "%me%", configdoc)
            createoption("autodelete", "1", configdoc)
            createoption("usetextgrab", "1", configdoc)
            createoption("defaultownerme", "0", configdoc)
            createoption("lastversioncheck", DateString, configdoc)
            createoption("versioncheckfrequency", "30", configdoc)
            createoption("englishspeedup", "1", configdoc)
            createoption("autosave", "1", configdoc)
            deleteoption("skipdashboards", configdoc)
            deleteoption("listtoggle", configdoc)


            '
            'get rid of old "resources" branch -- not supported.  Implemented with resourceverb and resourcelist
            If lastupgrade < "20090117" Then
                deletemainbranch("resources", configdoc)
            End If
            'alias
            a = createmainbranch("alias", configdoc, "Substitute longer strings in notes for short alias keyword")
            addkeyword(a, "sdpj", "[someday isproject]", "20080128.1", lastupgrade)
            'completion
            a = createmainbranch("completion", configdoc, "Set percent complete or info only")
            addkeyword(a, "info:", "-1", "20081118.2", lastupgrade)
            addkeyword(a, "info only:", "-1", "20081118.2", lastupgrade)
            addkeyword(a, "isinfo", "-1", "20081118.5", lastupgrade)
            addkeyword(a, "complete:", "100", "20081118.2", lastupgrade)
            addkeyword(a, "iscomplete", "100", "20081118.2", lastupgrade)
            addkeyword(a, "half done", "50", "20081118.2", lastupgrade)
            addkeyword(a, "done:", "100", "20080118.2", lastupgrade)
            addkeyword(a, "not done", "0", "20080118.2", lastupgrade)
            addkeyword(a, "isstalled", "85", "20090503", lastupgrade) 'requested last year
            addkeyword(a, "issnagged", "85", "20090503", lastupgrade)


            'clips---------------------------------------------------------------
            a = createmainbranch("clips", configdoc, "")
            'should we delete note and link from old configurations?
            addkeyword(a, "olmsg", "olmsg", "20080118.2", lastupgrade)
            addkeyword(a, "isnote", "note", "20080118.2", lastupgrade)
            WarnFirstDeleteKeyword(a, "note", "20080118.2", lastupgrade)
            addkeyword(a, "see note", "note", "20080118.4", lastupgrade)

            addkeyword(a, "islink", "link", "20080118.2", lastupgrade)
            WarnFirstDeleteKeyword(a, "link", "20080118.2", lastupgrade)
            addkeyword(a, "see link", "link", "20080118.2", lastupgrade)

            WarnFirstDeleteKeyword(a, "isnote", "20080127.1", lastupgrade)
            WarnFirstDeleteKeyword(a, "islink", "20080127.1", lastupgrade)
            WarnFirstDeleteKeyword(a, "see note", "20080127.1", lastupgrade)
            WarnFirstDeleteKeyword(a, "see link", "20080127.1", lastupgrade)

            'icons---------------------------------------------------------------
            a = createmainbranch("icons", configdoc, "")
            addkeyword(a, "olmsg", "mmStockIconLetter", "20080118.2", lastupgrade)
            deletekeyword(a, "find", "20080118.2", lastupgrade)

            'contexts -----------------------------------------------------------
            a = createmainbranch("contexts", configdoc, "")
            createkeyword(a, "find", "home", "20080118.2", lastupgrade)
            'start in------------------------------------------------------------
            a = createmainbranch("start in", configdoc, "Set Start date relative to today") 'can't use "next month", etc here or will match as due date
            Dim quotedm As String
            Dim quotedy As String
            quotedm = Chr(34) & "m" & Chr(34)
            quotedy = Chr(34) & "yyyy" & Chr(34)
            If getoption("language", configdoc, Nothing) = "English" Then
                addkeyword(a, "start in 0 day", "0", "20080430.1", lastupgrade)
                addkeyword(a, "starting now", "0", "20080430.1", lastupgrade)
                addkeyword(a, "start now", "0", "20080430.1", lastupgrade)
                addkeyword(a, "start in 1 day", "1", "20080118.2", lastupgrade)
                addkeyword(a, "start in 2 days", "2", "20080118.2", lastupgrade)
                addkeyword(a, "start in 3 days", "3", "20080118.2", lastupgrade)
                addkeyword(a, "start in 4 days", "4", "20080118.2", lastupgrade)
                addkeyword(a, "start in 5 days", "5", "20080118.2", lastupgrade)
                addkeyword(a, "start in 6 days", "6", "20080118.2", lastupgrade)
                addkeyword(a, "start in 7 days", "7", "20080118.2", lastupgrade)
                addkeyword(a, "start in 1 week", "7", "20080118.2", lastupgrade)
                addkeyword(a, "start in 2 weeks", "14", "20080118.2", lastupgrade)
                addkeyword(a, "start in 3 weeks", "21", "20080118.2", lastupgrade)
                addkeyword(a, "start in 1 month", "dateadd(" & quotedm & ",1,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 2 months", "dateadd(" & quotedm & ",2,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 3 months", "dateadd(" & quotedm & ",3,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 1 quarter", "dateadd(" & quotedm & ",3,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 4 months", "dateadd(" & quotedm & ",4,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 5 months", "dateadd(" & quotedm & ",5,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 6 months", "dateadd(" & quotedm & ",6,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "start in 1 year", "DateAdd(" & quotedy & ",1,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 1 day", "1", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 2 days", "2", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 3 days", "3", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 4 days", "4", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 5 days", "5", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 6 days", "6", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 7 days", "7", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 1 week", "7", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 2 weeks", "14", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 3 weeks", "21", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 1 month", "dateadd(" & quotedm & ",1,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 2 months", "dateadd(" & quotedm & ",2,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 3 months", "dateadd(" & quotedm & ",3,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 1 quarter", "dateadd(" & quotedm & ",3,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 4 months", "dateadd(" & quotedm & ",4,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 5 months", "dateadd(" & quotedm & ",5,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 6 months", "dateadd(" & quotedm & ",6,today)-today", "20080118.2", lastupgrade)
                addkeyword(a, "starting in 1 year", "DateAdd(" & quotedy & ",1,today)-today", "20080118.2", lastupgrade)
            End If
            'customicons---------------------------------------------------------
            a = createmainbranch("CustomIcons", configdoc, "")
            deletekeyword(a, "project", "20080118.2", lastupgrade)
            deletekeyword(a, "result", "20080118.2", lastupgrade)
            deletekeyword(a, "Pject", "20080118.2", lastupgrade)
            deletekeyword(a, "SubPject", "20080118.2", lastupgrade)
            deletekeyword(a, "Rsult", "20080118.2", lastupgrade)
            deletekeyword(a, "isresult", "20130915", lastupgrade)
            deletekeyword(a, "isproject", "20130915", lastupgrade)
            deletekeyword(a, "isresult", "20130915", lastupgrade)
            deletekeyword(a, "project:", "20130915", lastupgrade)
            deletekeyword(a, "result:", "20130915", lastupgrade)
            addkeyword(a, "isproject", Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\" & "images\resultmanager-projecticon.ico", "20130908", lastupgrade)
            addkeyword(a, "isresult", Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\" & "images\resultmanager-resulticon.ico", "20130908", lastupgrade)
            addkeyword(a, "project:", Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\" & "images\resultmanager-projecticon.ico", "20130908", lastupgrade)
            addkeyword(a, "result:", Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\" & "images\resultmanager-resulticon.ico", "20130908", lastupgrade)
            

            'RESOURCELIST--------------------------------------------------------
            a = createmainbranch("resourcelist", configdoc, "")
            addkeyword(a, "ActivityOwner", "ActivityOwner", "20080118.2", lastupgrade)
            addkeyword(a, "Nick", "Nick Duffill", "20080118.2", lastupgrade)
            addkeyword(a, "Nik", "Nik Tipler", "20080118.2", lastupgrade)
            '
            'RESOURCEVERBS------------------------------------------------------
            a = createmainbranch("resourceverbs", configdoc, "")
            addkeyword(a, "assign to", "waiting", "20080118.2", lastupgrade)
            addkeyword(a, "assigned to", "waiting", "20080118.2", lastupgrade)
            addkeyword(a, "r:", "waiting", "20080118.2", lastupgrade)
            addkeyword(a, "contact", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "talk to", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "waiting for", "waiting", "20080118.2", lastupgrade)
            addkeyword(a, "i owe", "owe", "20080118.2", lastupgrade)
            addkeyword(a, "email", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "e-mail", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "ask", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "discuss with", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "remind", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "inform", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "call", "partner", "20080118.2", lastupgrade)
            addkeyword(a, "delegated to", "delegated", "20080118.2", lastupgrade)
            addkeyword(a, "assign to", "waiting", "20080118.2", lastupgrade)
            addkeyword(a, "assigned to", "waiting", "20080118.2", lastupgrade)

            '
            'STARTING-------------------------------------------------------------
            a = createmainbranch("starting", configdoc, "")
            addkeyword(a, "1 day before", "1", "20080118.2", lastupgrade)
            addkeyword(a, "2 days before", "2", "20080118.2", lastupgrade)
            addkeyword(a, "3 days before", "3", "20080118.2", lastupgrade)
            addkeyword(a, "4 days before", "4", "20080118.2", lastupgrade)
            addkeyword(a, "5 days before", "5", "20080118.2", lastupgrade)
            addkeyword(a, "6 days before", "6", "20080118.2", lastupgrade)
            addkeyword(a, "7 days before", "7", "20080118.2", lastupgrade)
            addkeyword(a, "1 week before", "7", "20080118.2", lastupgrade)
            addkeyword(a, "2 weeks before", "14", "20080118.2", lastupgrade)
            addkeyword(a, "3 weeks before", "21", "20080118.2", lastupgrade)
            addkeyword(a, "4 weeks before", "28", "20080118.2", lastupgrade)
            addkeyword(a, "1 month before", "30", "20080118.2", lastupgrade)
            addkeyword(a, "2 months before", "60", "20080118.2", lastupgrade)
            addkeyword(a, "3 months before", "90", "20080118.2", lastupgrade)
            addkeyword(a, "same day", "0", "20080118.2", lastupgrade)
            addkeyword(a, "nsd", "0", "20080118.2", lastupgrade) 'remove start date
            addkeyword(a, "nsd", "-1", "20080118.2", lastupgrade) 'allow start date removal with nsd keyword
            '
            'EXTEND---------------------------------------------------------------
            a = createmainbranch("extend", configdoc, "")
            addkeyword(a, "extend 1 day", "1", "20080118.2", lastupgrade)
            addkeyword(a, "extend 2 days", "2", "20080118.2", lastupgrade)
            addkeyword(a, "extend 3 days", "3", "20080118.2", lastupgrade)
            addkeyword(a, "extend 4 days", "4", "20080118.2", lastupgrade)
            addkeyword(a, "extend 5 days", "5", "20080118.2", lastupgrade)
            addkeyword(a, "extend 6 days", "6", "20080118.2", lastupgrade)
            addkeyword(a, "extend 7 days", "7", "20080118.2", lastupgrade)
            addkeyword(a, "extend 1 week", "7", "20080118.2", lastupgrade)
            addkeyword(a, "extend 1 month", "DateAdd(" & Chr(34) & "m" & Chr(34) & ",1,today)-today", "20080118.2", lastupgrade)
            addkeyword(a, "extend 1 quarter", "DateAdd(" & Chr(34) & "m" & Chr(34) & ",3,today)-today", "20080118.2", lastupgrade)

            '
            'DELAY---------------------------------------------------------------------
            a = createmainbranch("delay", configdoc, "")
            addkeyword(a, "delay 1 day", "1", "20080118.2", lastupgrade)
            addkeyword(a, "delay 2 days", "2", "20080118.2", lastupgrade)
            addkeyword(a, "delay 3 days", "3", "20080118.2", lastupgrade)
            addkeyword(a, "delay 4 days", "4", "20080118.2", lastupgrade)
            addkeyword(a, "delay 5 days", "5", "20080118.2", lastupgrade)
            addkeyword(a, "delay 6 days", "6", "20080118.2", lastupgrade)
            addkeyword(a, "delay 7 days", "7", "20080118.2", lastupgrade)
            addkeyword(a, "delay 1 week", "7", "20080118.2", lastupgrade)
            addkeyword(a, "delay 1 month", "dateadd(" & Chr(34) & "m" & Chr(34) & ",1,today)-today", "20080205.1", lastupgrade)
            addkeyword(a, "delay 1 quarter", "dateadd(" & Chr(34) & "m" & Chr(34) & ",3,today)-today", "20080118.2", lastupgrade)
            addkeyword(a, "delay 1 year", "dateadd(" & Chr(34) & "yyyy" & Chr(34) & ",1,today)-today", "20080118.2", lastupgrade)

            'priority
            a = createmainbranch("priority", configdoc, "")
            addkeyword(a, "P1", "mmTaskPriority1", "20080128.2", lastupgrade)
            addkeyword(a, "P2", "mmTaskPriority2", "20080128.2", lastupgrade)
            addkeyword(a, "P3", "mmTaskPriority3", "20080128.2", lastupgrade)
            addkeyword(a, "P4", "mmTaskPriority4", "20080128.2", lastupgrade)
            addkeyword(a, "p1", "mmTaskPriority1", "20080128.2", lastupgrade)
            addkeyword(a, "p2", "mmTaskPriority2", "20080128.2", lastupgrade)
            addkeyword(a, "p3", "mmTaskPriority3", "20080128.2", lastupgrade)
            addkeyword(a, "p4", "mmTaskPriority4", "20080128.2", lastupgrade)

            'delete "extend" keywords from "delay" branch -- they were put in wrong branch in earlier upgrade
            deletekeyword(a, "extend 1 day", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 2 days", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 3 days", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 4 days", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 5 days", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 6 days", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 7 days", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 1 week", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 1 month", "20080118.2", lastupgrade)
            deletekeyword(a, "extend 1 quarter", "20080118.2", lastupgrade)
            '
            'ADVANCE----------------------------------------------
            a = createmainbranch("advance", configdoc, "")
            addkeyword(a, "advance 1 day", "1", "20080118.2", lastupgrade)
            addkeyword(a, "advance 2 days", "2", "20080118.2", lastupgrade)
            addkeyword(a, "advance 3 days", "3", "20080118.2", lastupgrade)
            addkeyword(a, "advance 4 days", "4", "20080118.2", lastupgrade)
            addkeyword(a, "advance 5 days", "5", "20080118.2", lastupgrade)
            addkeyword(a, "advance 6 days", "6", "20080118.2", lastupgrade)
            addkeyword(a, "advance 7 days", "7", "20080118.2", lastupgrade)
            addkeyword(a, "advance 1 week", "7", "20080118.2", lastupgrade)
            deletekeyword(a, "advance 1 month", "20081111.1", lastupgrade)
            addkeyword(a, "advance 1 month", "dateadd(" & Chr(34) & "m" & Chr(34) & ",1,today)-today", "20081111.1", lastupgrade)
            addkeyword(a, "advance 1 quarter", "dateadd(" & Chr(34) & "m" & Chr(34) & ",3,today)-today", "20080118.2", lastupgrade)

            '
            'CATEGORY------------------------------------------------------------------
            a = createmainbranch("category", configdoc, "")
            deletekeyword(a, "daily", "20090313", lastupgrade)
            addkeyword(a, "rdaily", "daily", "20090313", lastupgrade)
            deletekeyword(a, "monthly", "20090313", lastupgrade)
            addkeyword(a, "rmonthly", "monthly", "20090313", lastupgrade)
            deletekeyword(a, "weekly", "20090313", lastupgrade)
            addkeyword(a, "rweekly", "weekly", "20090313", lastupgrade)
            deletekeyword(a, "quarterly", "20090313", lastupgrade)
            addkeyword(a, "rquarterly", "quarterly", "20090313", lastupgrade)
            deletekeyword(a, "fortnightly", "20090313", lastupgrade)
            addkeyword(a, "rfortnightly", "fortnightly", "20090313", lastupgrade)
            deletekeyword(a, "yearly", "20090313", lastupgrade)
            addkeyword(a, "ryearly", "yearly", "20090313", lastupgrade)
            addkeyword(a, "rbiannually", "biannual", "20090313", lastupgrade)

            addkeyword(a, "eachmonth", "eachmonth", "20080118.2", lastupgrade)
            addkeyword(a, "everytwo", "everytwo", "20080118.2", lastupgrade)
            addkeyword(a, "eachweek", "eachweek", "20080118.2", lastupgrade)
            addkeyword(a, "every2weeks", "every2weeks", "20080118.2", lastupgrade)
            addkeyword(a, "each2weeks", "each2weeks", "20080118.2", lastupgrade)
            addkeyword(a, "eachfortnight", "eachfortnight", "20080118.2", lastupgrade)
            addkeyword(a, "eachquarter", "eachquarter", "20080118.2", lastupgrade)
            addkeyword(a, "eachyear", "eachyear", "20080118.2", lastupgrade)
            addkeyword(a, "endofmonth", "endofmonth", "20080118.2", lastupgrade)
            addkeyword(a, "endofquarter", "endofquarter", "20080118.2", lastupgrade)

            deletekeyword(a, "end of month", "20090313", lastupgrade)
            deletekeyword(a, "end of quarter", "20090313", lastupgrade)

            addkeyword(a, "each month", "eachmonth", "20080118.2", lastupgrade)
            addkeyword(a, "each week", "eachweek", "20080118.2", lastupgrade)
            addkeyword(a, "every 2 weeks", "every2weeks", "20080118.2", lastupgrade)
            addkeyword(a, "each quarter", "eachquarter", "20080118.2", lastupgrade)
            addkeyword(a, "each year", "eachyear", "20080118.2", lastupgrade)

            addkeyword(a, "2m", "2m", "20081220.1", lastupgrade)
            'DATES-------------------------------------------------------------
            a = createmainbranch("dates", configdoc, "")
            createkeyword(a, "ndd", "0", "20080118.2", lastupgrade)
            deletekeyword(a, "nsd", "20080118.2", lastupgrade)  'this belonged in starting branch
            '---------------------------------------------------------------
            checkforduplicates(configdoc)
            'Mark map as upgraded
            setoption("lastupgrade", currentversion, configdoc)
            If configdoc.IsModified Then configdoc.Save()
            a = Nothing
            MsgBox("Configuration Map Upgrade complete")
        End If
    End Sub
    Function destinationlink(ByRef linktopic As mm.Topic, ByRef aStr As String) As mm.Hyperlink
        Dim i As Integer
        Dim defaultlink As mm.Hyperlink
        Dim t As mm.Topic
        destinationlink = Nothing
        i = linktopic.AllSubTopics.Count
        defaultlink = linktopic.AllSubTopics(1).Hyperlink 'default location of default map
        While i > 0 And destinationlink Is Nothing
            t = linktopic.AllSubTopics.Item(i)
            If t.AllSubTopics.Count > 0 Then destinationlink = destinationlinksub(t, aStr)
            If destinationlink Is Nothing Then
                If t.HasHyperlink And Len(Trim(t.Text)) > 0 Then
                    If InStr(LCase(aStr), LCase(t.Text)) > 0 Then destinationlink = t.Hyperlink
                End If
            End If
            If InStr("defaultmap", LCase(linktopic.AllSubTopics(i).Text)) > 0 Then defaultlink = linktopic.AllSubTopics(i).Hyperlink
            i = i - 1
        End While
        If destinationlink Is Nothing Then destinationlink = defaultlink
        defaultlink = Nothing
        t = Nothing
    End Function
    Function destinationlinksub(ByRef linktopic As mm.Topic, ByRef aStr As String) As mm.Hyperlink
        Dim i As Integer
        Dim t As mm.Topic
        i = linktopic.AllSubTopics.Count
        destinationlinksub = Nothing
        While i > 0 And destinationlinksub Is Nothing
            t = linktopic.AllSubTopics.Item(i)
            If t.AllSubTopics.Count > 0 Then destinationlinksub = destinationlinksub(t, aStr)
            If destinationlinksub Is Nothing Then
                If t.HasHyperlink Then
                    If InStr(LCase(aStr), LCase(t.Text)) > 0 Then destinationlinksub = t.Hyperlink
                End If
            End If
            i = i - 1
        End While
        t = Nothing
    End Function
    Sub Install_or_Migrate_MindReader_Config(ByRef m_app As Mindjet.MindManager.Interop.Application)
        'The purpose of this routine is to migrate legacy mindreader configuration maps to new name and location while preserving and fixing its links
        'it may be called by an installer routine or when mindreaderopen or mindreadernlp find the expected configuration map missing
        Dim d As mm.Document
        Dim dd As mm.Document
        Dim oldmapname As String
        Dim startermapname As String
        Dim mymappath As String
        Dim linkbranch As mm.Topic
        Dim legacy As Boolean
        Dim response As Integer
        Dim t As mm.Topic
        mymappath = m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps)
        oldmapname = mymappath & "mindreader.mmap"
        startermapname = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\mindreaderconfigsample.mmap"
        'see if there is a legacy map
        If MsgBox("MindReader Installation/Migration.  If you have already completed migration, Choose cancel It is probably a bug!", vbOKCancel) = vbCancel Then
            MsgBox("Terminating program.  Make sure your configuration map is backed up (rename it mindreaderbackup.mmap and leave in directory)")
            Exit Sub
        End If
        Try
            d = OpenMapHidden(m_app, oldmapname)
        Catch
        End Try
        If Not d Is Nothing Then
            d.Save()
            d.Close()
            legacy = True
            response = MsgBox("A legacy configuration map was found. Choose Yes to migrate, No to start new, or CANCEL if did not just install/upgrade.", vbYesNoCancel)
        Else
            legacy = False
        End If
        If (Not legacy) Or (legacy And response = vbNo) Then 'copy sample map
            d = OpenMapHidden(m_app, startermapname)
            If d Is Nothing Then
                MsgBox("sample map missing. Please reinstall Mindreader")
                Exit Sub
            Else
                If Not My.Computer.FileSystem.FileExists(MindReaderFolderPath(m_app) & "mindreaderconfigsample.mmap") Then
                    'If Not My.Computer.FileSystem.FileExists(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\mindreaderconfigsample.mmap") Then
                    MkDir(MindReaderFolderPath(m_app))
                End If
                d.SaveAs(MindReaderConfigMapFullName(m_app))
                d.Save()
                d.Close()
            End If
        ElseIf legacy And response = vbYes Then
            If isopen(m_app, MindReaderConfigMapFullName(m_app)) Then 'just in case, close open destination map, should not be there or we wouldn't be here
                dd = OpenMapHidden(m_app, MindReaderConfigMapFullName(m_app))
                If Not dd Is Nothing Then
                    dd.Close()
                    dd = Nothing
                End If
            End If
            d = OpenMapHidden(m_app, oldmapname)
            d.SaveAs(MindReaderConfigMapFullName(m_app))
            d.Close()
        Else
            MsgBox("Terminating Program")
            Exit Sub
        End If
        d = OpenMapHidden(m_app, MindReaderConfigMapFullName(m_app))
        If Not d Is Nothing Then
            linkbranch = createmainbranch("links", d, "")
            For Each t In linkbranch.AllSubTopics
                If InStr(LCase(t.Text), "mapmap") = 1 Then
                    t.Hyperlink.Address = MindReaderConfigMapFullName(m_app)
                    Exit For
                End If
            Next
            d.Save()
        End If
        d = Nothing
        dd = Nothing
        t = Nothing
        linkbranch = Nothing
        d = Nothing
        dd = Nothing
        t = Nothing
    End Sub
    Sub newprojectmap(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal npm As String)
        Dim NewMap As mm.Document
        Dim ConfigDoc As mm.Document
        Dim DestinationBranch As mm.Topic
        Dim KeywordTopic As mm.Topic
        Dim DestinationKeyword As String
        Dim t As mm.Topic
        NewMap = m_app.Documents.Add
        NewMap.CentralTopic.Text = npm
        setstartdate(NewMap.CentralTopic, Today)
        NewMap.Selection.Set(NewMap.CentralTopic)
        MindReaderNLP(m_app, "")
        'add in-tray branch
        t = NewMap.CentralTopic.AddSubTopic("in-tray")
        addmarker(t, Nothing, "Areas", "in-tray", False)
        t.CreateBoundary.FillColor.Value = -96
        NewMap.Selection.Set(NewMap.CentralTopic.AllSubTopics(1))
        'MindReaderNLP(m_app, "")
        'add reference branch
        NewMap.CentralTopic.AddSubTopic("Reference").Icons.AddStockIcon(Mindjet.MindManager.Interop.MmStockIcon.mmStockIconNoEntry)
        'add plan branch
        NewMap.Selection.Set(NewMap.CentralTopic.AddBalancedSubTopic("Plan").AddSubTopic(InputBox("What is the next action? Use 1st>>2nd>>3rd to define a sequence of tasks.")))
        MindReaderNLP(m_app, "")
        'add destination keyword
        If MsgBox("Would you like to add a destination keyword pointing to this map's in-tray? Only do this if you anticipate adding many items to map from OutLinker or GyroQ", vbYesNo) = vbYes Then
            DestinationKeyword = InputBox("Enter Destination Keyword")
            If Len(DestinationKeyword) > 0 Then
                ConfigDoc = OpenMapHidden(m_app, MindReaderConfigMapFullName(m_app))
                DestinationBranch = createmainbranch("Links", ConfigDoc, "")
                KeywordTopic = DestinationBranch.AddSubTopic(DestinationKeyword)
                KeywordTopic.CreateHyperlink(LinkToThisTopic(NewMap.CentralTopic.AllSubTopics(1)))
                MsgBox("Destination Keyword sucessfully added to MindReaderConfig.  Make sure you review destination keywords in mindreaderconfig.mmap frequently and remove outdated keywords)")
            End If
        End If
        NewMap = Nothing
        DestinationBranch = Nothing
        ConfigDoc = Nothing
    End Sub
    Sub AddDestinationKeyword(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal destinationkeyword As String)
        'Dim NewMap As mm.Document
        Dim ConfigDoc As mm.Document
        Dim DestinationBranch As mm.Topic
        Dim KeywordTopic As mm.Topic
        If destinationkeyword = "" Then destinationkeyword = InputBox("Enter Destination Keyword")
        If Len(destinationkeyword) > 0 Then
            ConfigDoc = OpenMapHidden(m_app, MindReaderConfigMapFullName(m_app))
            DestinationBranch = createmainbranch("Links", ConfigDoc, "")
            KeywordTopic = DestinationBranch.AddSubTopic(destinationkeyword)
            KeywordTopic.CreateHyperlink(LinkToThisTopic(m_app.ActiveDocument.Selection.PrimaryTopic))
            MsgBox("Destination Keyword sucessfully added to MindReaderConfig.  Make sure you review destination keywords in mindreaderconfig.mmap frequently and remove outdated keywords)")
        End If
    End Sub
    Function eval(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal s As String) As String
        Microsoft.VisualBasic.FileOpen(1, System.IO.Path.GetTempPath() & "mindreader.tmp", OpenMode.Output, OpenAccess.Write)
        Microsoft.VisualBasic.Print(1, s)
        Microsoft.VisualBasic.FileClose(1)
        m_app.RunMacro(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\evalclipboard.mmbas")
        eval = Replace(My.Computer.FileSystem.ReadAllText(System.IO.Path.GetTempPath() & "mindreader.tmp"), Chr(34), "")
    End Function
    Function inteval(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal s As String) As Integer
        Microsoft.VisualBasic.FileOpen(1, System.IO.Path.GetTempPath() & "mindreader.tmp", OpenMode.Output, OpenAccess.Write)

        Microsoft.VisualBasic.Print(1, s)
        Microsoft.VisualBasic.FileClose(1)
        m_app.RunMacro(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\evalclipboard.mmbas")
        s = My.Computer.FileSystem.ReadAllText(System.IO.Path.GetTempPath() & "mindreader.tmp")
        s = Replace(s, Chr(34), "")
        inteval = CInt(s)
    End Function
End Module
