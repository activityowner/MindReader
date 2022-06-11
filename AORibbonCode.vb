
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Namespace MindManagerRibbon

    Class AORibbonGroup

        Implements IDisposable

#Region "Variables"
        Private m_app As mm.Application
        Private WithEvents myCommand4 As mm.Command
        Private WithEvents myCommand5 As mm.Command
        Private WithEvents myCommand6 As mm.Command
        Private WithEvents myCommand7 As mm.Command
        Private WithEvents myCommand8 As mm.Command
        Private WithEvents myCommand9 As mm.Command
        Private WithEvents myCommand10 As mm.Command
        'Private WithEvents myCommand11 As mm.Command

        'Private WithEvents myCommand13 As mm.Command
        'Private WithEvents myCommand14 As mm.Command

        Private WithEvents myCommand16 As mm.Command
        Private WithEvents myCommand17 As mm.Command
        Private WithEvents myCommand18 As mm.Command
        Private WithEvents myCommand19 As mm.Command
        Private WithEvents myCommand20 As mm.Command
        Private WithEvents myCommand21 As mm.Command
        Private WithEvents myCommand22 As mm.Command
        Private WithEvents myCommand23 As mm.Command
        Private WithEvents myCommand24 As mm.Command
        Private WithEvents myCommand25 As mm.Command
        Private WithEvents myCommand26 As mm.Command
        Private WithEvents myCommand27 As mm.Command
        'Private WithEvents myCommand29 As mm.Command
        Private WithEvents myCommand30 As mm.Command

        Private WithEvents myCommand60 As mm.Command
        Private WithEvents myCommand61 As mm.Command
        Private WithEvents myCommand62 As mm.Command
        Private WithEvents myCommand63 As mm.Command
        Private WithEvents myCommand64 As mm.Command
        Private WithEvents myCommand65 As mm.Command

        Private WithEvents myCommand66 As mm.Command
        Private WithEvents myCommand67 As mm.Command
        Private WithEvents myCommand68 As mm.Command

        Private WithEvents myCommand69 As mm.Command
        Private WithEvents mycommand70 As mm.Command
        Private WithEvents mycommand71 As mm.Command

        Private WithEvents mycommand72 As mm.Command
        Private WithEvents mycommand73 As mm.Command
        Private WithEvents mycommand74 As mm.Command
        Private WithEvents mycommand75 As mm.Command
        Private WithEvents mycommand76 As mm.Command
        Private WithEvents mycommand77 As mm.Command
        Private WithEvents mycommand78 As mm.Command
        Private WithEvents mycommand79 As mm.Command
        Private WithEvents mycommand80 As mm.Command
        Private WithEvents mycommand81 As mm.Command
        Private WithEvents mycommand82 As mm.Command
        Private WithEvents mycommand83 As mm.Command
        Private WithEvents mycommand84 As mm.Command
        Private WithEvents mycommand85 As mm.Command
        'Private WithEvents mycommand86 As mm.Command





#End Region

        Public Sub New(ByVal app As Mindjet.MindManager.Interop.Application)
            Try
                m_app = app
                usedit = False
                donated = False
                'Creates the Ribbons
                Dim MindReaderRibbon As Mindjet.MindManager.Interop.ribbonTab = Ribbons.CreateRibbon(m_app, "MindReader", "urn:MRG.Tab")
                Dim MindReaderMarkupRibbon As Mindjet.MindManager.Interop.ribbonTab = Ribbons.CreateRibbon(m_app, "MR-Markup", "urn:MRMG.Tab")

                'Creates the Ribbon Groups
                Dim MindreaderInformationRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderRibbon, "Information", "urn:MRG.Group1")
                Dim MindreaderOpenRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderRibbon, "Open and Close Maps", "urn:MRG.Group2")
                Dim MindReaderNewRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderRibbon, "Create and Capture", "urn:MRG.Group6")

                Dim MindreaderReadRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderMarkupRibbon, "Markup Topic", "urn:MRMG.markup")



                If getmrkey("tablabels", "duelabel") = "" Then setmrkey("tablabels", "duelabel", "Due")
                If getmrkey("tablabels", "contextlabel") = "" Then setmrkey("tablabels", "contextlabel", "context")
                If getmrkey("tablabels", "timelabel") = "" Then setmrkey("tablabels", "timelabel", "time")

                Dim mindreadercontextribbontab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderMarkupRibbon, getmrkey("tablabels", "contextlabel"), "urn:MRMG.context")
                Dim MindreaderDateRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderMarkupRibbon, getmrkey("tablabels", "duelabel"), "urn:MRMG.date")
                Dim mindreadertimeribbontab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderMarkupRibbon, getmrkey("tablabels", "timelabel"), "urn:MRMG.time")

                Dim MindReaderMiscRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderRibbon, "Misc", "urn:MRG.Group7")
                Dim MindReaderAnalysisRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderRibbon, "Analysis", "urn:MRG.Group8")
                Dim MindReaderSendRibbonTab As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(MindReaderRibbon, "Send by Keyword", "urn:MRG.SendGroup")


                'Creates the Ribbon Group Commands
                myCommand4 = m_app.Commands.Add("AORibbon.Connect", "AOCloseUnModified")
                myCommand5 = m_app.Commands.Add("AORibbon.Connect", "AOFilename")

                myCommand7 = m_app.Commands.Add("AORibbon.Connect", "AOOpenMapKeyword")
                myCommand8 = m_app.Commands.Add("AORibbon.Connect", "AONewProjectMap")
                myCommand9 = m_app.Commands.Add("AORibbon.Connect", "AOSendMapKeyword")
                myCommand10 = m_app.Commands.Add("AORibbon.Connect", "AOSaveAllMaps")
                myCommand23 = m_app.Commands.Add("AORibbon.Connect", "AOOpenDailyAction")
                If RMinstalled() Then
                    myCommand6 = m_app.Commands.Add("AORibbon.Connect", "AONextActionAnalysis")
                    'myCommand11 = m_app.Commands.Add("AORibbon.Connect", "AONAATrend")
                    'myCommand13 = m_app.Commands.Add("AORibbon.Connect", "AOMiniMapCentral")
                    'myCommand14 = m_app.Commands.Add("AORibbon.Connect", "AONextActions2Outlook")
                    'myCommand29 = m_app.Commands.Add("AORibbon.Connect", "AOOpenDailyActionTemp")
                End If

                myCommand16 = m_app.Commands.Add("AORibbon.Connect", "AOMindReader")
                myCommand17 = m_app.Commands.Add("AORibbon.Connect", "AOTaskCount")
                myCommand18 = m_app.Commands.Add("AORibbon.Connect", "AOHelp")
                myCommand19 = m_app.Commands.Add("AORibbon.Connect", "AOButtonReader")
                myCommand20 = m_app.Commands.Add("AORibbon.Connect", "AOAddTask")
                myCommand21 = m_app.Commands.Add("AORibbon.Connect", "AONextWeek")
                myCommand22 = m_app.Commands.Add("AORibbon.Connect", "AONextMonth")
                myCommand24 = m_app.Commands.Add("AORibbon.Connect", "AOOpenDailyCapture")
                myCommand25 = m_app.Commands.Add("AORibbon.Connect", "AOBefore")
                myCommand26 = m_app.Commands.Add("AORibbon.Connect", "AOtesting")
                myCommand27 = m_app.Commands.Add("AORibbon.Connect", "AODelete")
                myCommand30 = m_app.Commands.Add("AORibbon.Connect", "AODestinationKeyword")

                myCommand60 = m_app.Commands.Add("AORibbon.Connect", "AOcontext1")
                myCommand61 = m_app.Commands.Add("AORibbon.Connect", "AOcontext2")
                myCommand62 = m_app.Commands.Add("AORibbon.Connect", "AOcontext3")
                myCommand63 = m_app.Commands.Add("AORibbon.Connect", "AOcontext4")
                myCommand64 = m_app.Commands.Add("AORibbon.Connect", "AOcontext5")
                myCommand65 = m_app.Commands.Add("AORibbon.Connect", "AOcontext6")

                myCommand66 = m_app.Commands.Add("AORibbon.Connect", "AO15m")
                myCommand67 = m_app.Commands.Add("AORibbon.Connect", "AO1h")
                myCommand68 = m_app.Commands.Add("AORibbon.Connect", "AO2h")

                myCommand69 = m_app.Commands.Add("AORibbon.Connect", "AOtomorrow")
                mycommand70 = m_app.Commands.Add("AORibbon.Connect", "AOMindreadtopic")

                mycommand71 = m_app.Commands.Add("AORibbon.Connect", "AOcalculator")

                mycommand72 = m_app.Commands.Add("AORibbon.Connect", "AOsend1")
                mycommand73 = m_app.Commands.Add("AORibbon.Connect", "AOsend2")
                mycommand74 = m_app.Commands.Add("AORibbon.Connect", "AOsend3")
                mycommand75 = m_app.Commands.Add("AORibbon.Connect", "AOsend4")
                mycommand76 = m_app.Commands.Add("AORibbon.Connect", "AOsend5")
                mycommand77 = m_app.Commands.Add("AORibbon.Connect", "AOsend6")
                mycommand78 = m_app.Commands.Add("AORibbon.Connect", "AOsend7")
                mycommand79 = m_app.Commands.Add("AORibbon.Connect", "AOsend8")
                mycommand80 = m_app.Commands.Add("AORibbon.Connect", "AOsend9")
                mycommand81 = m_app.Commands.Add("AORibbon.Connect", "AOJustNaa")
                mycommand82 = m_app.Commands.Add("AORibbon.Connect", "AOMarkup")
                mycommand83 = m_app.Commands.Add("AORibbon.Connect", "AOMakeList")

                mycommand84 = m_app.Commands.Add("AORibbon.Connect", "AOsend10")
                mycommand85 = m_app.Commands.Add("AORibbon.Connect", "AOsend11")
                'mycommand86 = m_app.Commands.Add("AORibbon.Connect", "AORefresh")

                'others to add
                'ask2waiting for converter
                'cb -- process a branch of entries with mindreader
                'mna -- mail next actions
                'mtcn 
                'ola -- create and link to outlook appt

                'defer
                'rfd -- refresh dashboard
                'a2l -- move attachment to disk and add hyperlink

                'parallel to selected task
                'enter link keyword
                'mra add new keywords to mindreader
                'lkw list keywords
                'elog -- enter an event into log
                'others


                myCommand4.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand5.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                myCommand7.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\file-open-icon.jpg"
                myCommand8.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\rm-project-icon.jpg"
                myCommand9.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand10.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                myCommand23.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                If RMinstalled() Then
                    myCommand6.LargeImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\next-action-analysis.jpg"
                    'myCommand11.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                    'myCommand13.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                    'myCommand14.LargeImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\outlook-sync.jpg"

                    'myCommand29.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                    'mycommand81.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                End If
                myCommand16.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand16.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\mindreader-icon.jpg"
                myCommand17.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand18.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand19.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand20.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\task-icon.jpg"
                myCommand21.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand22.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                myCommand24.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand25.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\task-icon.jpg"
                myCommand26.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand27.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\delete-icon.jpg"

                myCommand30.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                myCommand60.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand61.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand62.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand63.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand64.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand65.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                myCommand66.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand67.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                myCommand68.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                myCommand69.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand70.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand71.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\calculator.jpg"

                mycommand72.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand73.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand74.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand75.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand76.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand77.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand78.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand79.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand80.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                mycommand84.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand85.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                'mycommand86.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"


                mycommand82.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"
                mycommand83.ImagePath = Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com MindReader2\ao.jpg"

                '***

                myCommand4.BasicCommand = True

                myCommand7.BasicCommand = True
                myCommand8.BasicCommand = True
                myCommand9.BasicCommand = True
                myCommand10.BasicCommand = True

                myCommand23.BasicCommand = True
                If RMinstalled() Then
                    myCommand6.BasicCommand = True
                    'myCommand11.BasicCommand = True
                    'myCommand13.BasicCommand = True
                    'myCommand14.BasicCommand = True
                    'myCommand29.BasicCommand = True
                End If

                myCommand16.BasicCommand = True
                myCommand17.BasicCommand = True
                myCommand18.BasicCommand = True
                myCommand19.BasicCommand = True
                myCommand20.BasicCommand = True
                myCommand21.BasicCommand = True
                myCommand22.BasicCommand = True

                myCommand24.BasicCommand = True
                myCommand25.BasicCommand = True
                myCommand26.BasicCommand = True
                myCommand27.BasicCommand = True

                myCommand30.BasicCommand = True

                myCommand60.BasicCommand = True
                myCommand61.BasicCommand = True
                myCommand62.BasicCommand = True
                myCommand63.BasicCommand = True
                myCommand64.BasicCommand = True
                myCommand65.BasicCommand = True

                myCommand66.BasicCommand = True
                myCommand67.BasicCommand = True
                myCommand68.BasicCommand = True

                myCommand69.BasicCommand = True

                mycommand70.BasicCommand = True
                mycommand71.BasicCommand = True

                mycommand72.BasicCommand = True
                mycommand73.BasicCommand = True
                mycommand74.BasicCommand = True
                mycommand75.BasicCommand = True
                mycommand76.BasicCommand = True
                mycommand77.BasicCommand = True

                mycommand78.BasicCommand = True
                mycommand79.BasicCommand = True
                mycommand80.BasicCommand = True
                mycommand81.BasicCommand = True

                mycommand84.BasicCommand = True
                mycommand85.BasicCommand = True
                'mycommand86.BasicCommand = True

                mycommand82.BasicCommand = True
                mycommand83.BasicCommand = True

                myCommand4.ToolTip = ("Close Unmodified Maps" + (Chr(10)) + "Close Unmodified maps. Same as old clo tag")
                myCommand5.ToolTip = ("Filename of Active Map" + (Chr(10)) + "filename of active map. Same as old fn tag")

                myCommand7.ToolTip = ("Open by keyword" + (Chr(10)) + "Open map on keyword. Same as old o tag")
                myCommand8.ToolTip = ("New Project Map" + (Chr(10)) + "New Project Map.  Same as old nm tag")
                myCommand9.ToolTip = ("Send on Keyword" + (Chr(10)) + "Send selected topic(s) to a map based on a destination keyword. Same as old s tag")
                myCommand10.ToolTip = ("Save all maps" + (Chr(10)) + "Save all maps")

                myCommand23.ToolTip = ("Open DailyAction Map" + (Chr(10)) + "Open My Maps\dashboards\dailyaction.mmap.  You must have created and saved a dailyaction dashboard in this location with this name.")
                If RMinstalled() Then
                    myCommand6.ToolTip = ("Refresh Dashboard and run Next Action Analysis" + (Chr(10)) + "Refresh Dashboard and run Next Action Analysis.")
                    'myCommand11.ToolTip = ("NAA Trend" + (Chr(10)) + "Trend your NAA Score in Excel")
                    'myCommand13.ToolTip = ("Mini Map Central" + (Chr(10)) + "create a temporary mini-map central including only the active maps referenced in a dashboard. Dashboards created from this refresh faster.")
                    'myCommand14.ToolTip = ("Next Actions to Outlook" + (Chr(10)) + "Sync Next Actions to Outlook")
                    'myCommand29.ToolTip = ("open my maps\dailyactions_temp.mmap" + (Chr(10)) + "Open a secondary dailyaction map based on smaller temporary map central.  You must have saved this dashboard here.")
                End If

                myCommand16.ToolTip = ("Mindread Topic" + (Chr(10)) + "Mindread selected topics (same as old m tag)")
                myCommand17.ToolTip = ("Task Count" + (Chr(10)) + "how many open tasks in this map")
                myCommand18.ToolTip = ("Help" + (Chr(10)) + "visit activityowner wiki for help")
                myCommand19.ToolTip = ("ButtonReader" + (Chr(10)) + "Bring up floating form with configurable keyword buttons")
                myCommand20.ToolTip = ("Capture Task" + (Chr(10)) + "Enter a task to be mindread and added to map based on destination keyword. Same as old q tag")
                myCommand21.ToolTip = ("Next Week" + (Chr(10)) + "Set duedate next week")
                myCommand22.ToolTip = ("Next Month" + (Chr(10)) + "Set duedate next month")

                myCommand24.ToolTip = ("Open Daily Capture Map" + (Chr(10)) + "Open My Maps\Daily Capture Map.mmap")
                myCommand26.ToolTip = ("MindReader Options" + (Chr(10)) + "MindReader Options")
                myCommand25.ToolTip = ("Before" + (Chr(10)) + "Add a task before selected task. Same as old b tag")
                myCommand27.ToolTip = ("Cut/Delete" + (Chr(10)) + "Cut selected topic to delete or paste")

                myCommand30.ToolTip = ("Create Destination Keyword" + (Chr(10)) + "Define a keyword used to open map and/or send topics to a topic in it. Same as old k tag")

                mycommand70.ToolTip = ("mindread selected topic(s)" + (Chr(10)) + "mindread selected topic(s)")
                mycommand71.ToolTip = ("Calculator" + (Chr(10)) + "enter an equation like 2+2 and get the result")
                mycommand82.ToolTip = ("Mindread selected topic(s) text - same as old c tag" + (Chr(10)) + "Mindread selected text")
                mycommand83.ToolTip = ("Make sequenced list of selected topics" + (Chr(10)) + "only first is next action")


                myCommand4.Caption = "Close"
                myCommand5.Caption = "Filename"

                myCommand7.Caption = "Open by keyword"
                myCommand8.Caption = "Project Map"
                myCommand9.Caption = "keyword"
                myCommand10.Caption = "Save All"

                myCommand23.Caption = "DailyAction"
                If RMinstalled() Then
                    myCommand6.Caption = "Refresh and NAA"
                    'myCommand11.Caption = "Graph"
                    'myCommand13.Caption = "Temp Central"
                    'myCommand14.Caption = "Outlook Sync"

                    'myCommand29.Caption = "DailyActionTemp"
                    'mycommand81.Caption = "NAA"
                    'mycommand81.ToolTip = "Run Next Action Analysis"
                End If

                myCommand16.Caption = "Use keywords"
                myCommand17.Caption = "Count"
                myCommand18.Caption = "Help"
                myCommand19.Caption = "Popup Menu"
                myCommand20.Caption = "Idea/Task"

                myCommand24.Caption = "DailyCapture"
                myCommand25.Caption = "Prereq"
                myCommand26.Caption = "Options"
                myCommand27.Caption = "Cut/Delete"

                myCommand30.Caption = "Map Keyword"



                mycommand70.Caption = "MindRead topics"
                mycommand71.Caption = "Calculator"

                mycommand82.Caption = "Mindread topic"
                mycommand83.Caption = "Make List"

                setflexiblecommands()

                MindReaderSendRibbonTab.GroupControls.AddButton(myCommand9, 0) ' was misc

                MindreaderOpenRibbonTab.GroupControls.AddButton(myCommand7, 0)

                'If RMinstalled() Then
                MindreaderOpenRibbonTab.GroupControls.AddButton(myCommand23, 0)
                'End If

                If RMinstalled() Then
                    'MindreaderOpenRibbonTab.GroupControls.AddButton(myCommand29, 0)
                End If

                MindreaderOpenRibbonTab.GroupControls.AddButton(myCommand24, 0)

                MindreaderOpenRibbonTab.GroupControls.AddButton(myCommand4, 0)

                MindreaderReadRibbonTab.GroupControls.AddButton(myCommand16, 0)
                MindreaderReadRibbonTab.GroupControls.AddButton(myCommand19, 0)
                MindreaderReadRibbonTab.GroupControls.AddButton(mycommand82, 0)



                MindReaderNewRibbonTab.GroupControls.AddButton(myCommand20, 0)
                MindReaderNewRibbonTab.GroupControls.AddButton(myCommand30, 0)
                MindReaderNewRibbonTab.GroupControls.AddButton(myCommand8, 0)
                MindReaderNewRibbonTab.GroupControls.AddButton(myCommand25, 0)
                MindReaderNewRibbonTab.GroupControls.AddButton(mycommand83, 0)

                MindreaderOpenRibbonTab.GroupControls.AddButton(myCommand10, 0)

                MindReaderMiscRibbonTab.GroupControls.AddButton(myCommand27, 0)
                MindReaderMiscRibbonTab.GroupControls.AddButton(mycommand71, 0)

                If RMinstalled() Then
                    MindReaderAnalysisRibbonTab.GroupControls.AddButton(myCommand6, 0)
                End If
                MindReaderAnalysisRibbonTab.GroupControls.AddButton(myCommand17, 0)
                MindReaderAnalysisRibbonTab.GroupControls.AddButton(myCommand5, 0)

                If RMinstalled() Then
                    'MindReaderAnalysisRibbonTab.GroupControls.AddButton(myCommand13, 0)
                    'MindReaderAnalysisRibbonTab.GroupControls.AddButton(myCommand11, 0)
                    'MindReaderAnalysisRibbonTab.GroupControls.AddButton(mycommand81, 0)
                    'MindReaderMiscRibbonTab.GroupControls.AddButton(myCommand14, 0)
                End If

                MindreaderInformationRibbonTab.GroupControls.AddButton(myCommand18, 0)

                MindreaderInformationRibbonTab.GroupControls.AddButton(myCommand26, 0)


                MindreaderDateRibbonTab.GroupControls.AddButton(myCommand21, 0)
                MindreaderDateRibbonTab.GroupControls.AddButton(myCommand22, 0)
                MindreaderDateRibbonTab.GroupControls.AddButton(myCommand69, 0)

                mindreadercontextribbontab.GroupControls.AddButton(myCommand60, 0)
                mindreadercontextribbontab.GroupControls.AddButton(myCommand61, 0)
                mindreadercontextribbontab.GroupControls.AddButton(myCommand62, 0)
                mindreadercontextribbontab.GroupControls.AddButton(myCommand63, 0)
                mindreadercontextribbontab.GroupControls.AddButton(myCommand64, 0)
                mindreadercontextribbontab.GroupControls.AddButton(myCommand65, 0)

                mindreadertimeribbontab.GroupControls.AddButton(myCommand66, 0)
                mindreadertimeribbontab.GroupControls.AddButton(myCommand67, 0)
                mindreadertimeribbontab.GroupControls.AddButton(myCommand68, 0)

                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand72, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand73, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand74, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand75, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand76, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand77, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand78, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand79, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand80, 0)

                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand84, 0)
                MindReaderSendRibbonTab.GroupControls.AddButton(mycommand85, 0)

                'MindReaderMiscRibbonTab.GroupControls.AddButton(mycommand86, 0)


                'add to context menu
                mycommand70.SetDynamicMenu(Mindjet.MindManager.Interop.MmDynamicMenu.mmDynamicMenuContextTopic, True) 'mindread
                myCommand25.SetDynamicMenu(Mindjet.MindManager.Interop.MmDynamicMenu.mmDynamicMenuContextTopic, True) 'task before
                myCommand9.SetDynamicMenu(Mindjet.MindManager.Interop.MmDynamicMenu.mmDynamicMenuSendTo, True) 'send to by keyword

                setflexiblecommands()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub
        Sub setflexiblecommands()
            If getmrkey("contexts", "context1") = "" Then setmrkey("contexts", "context1", "@desk")
            If getmrkey("contexts", "context2") = "" Then setmrkey("contexts", "context2", "@home")
            If getmrkey("contexts", "context3") = "" Then setmrkey("contexts", "context3", "@errand")
            If getmrkey("contexts", "context4") = "" Then setmrkey("contexts", "context4", "@phone")
            If getmrkey("contexts", "context5") = "" Then setmrkey("contexts", "context5", "@web")
            If getmrkey("contexts", "context6") = "" Then setmrkey("contexts", "context6", "@anywhere")

            If getmrkey("dues", "due1") = "" Then setmrkey("dues", "due1", "tomorrow")
            If getmrkey("dues", "due2") = "" Then setmrkey("dues", "due2", "next week")
            If getmrkey("dues", "due3") = "" Then setmrkey("dues", "due3", "next month")

            If getmrkey("times", "time1") = "" Then setmrkey("times", "time1", "15m")
            If getmrkey("times", "time2") = "" Then setmrkey("times", "time2", "1h")
            If getmrkey("times", "time3") = "" Then setmrkey("times", "time3", "2h")


            myCommand60.Caption = getmrkey("contexts", "context1")
            myCommand60.ToolTip = myCommand60.Caption
            myCommand61.Caption = getmrkey("contexts", "context2")
            myCommand61.ToolTip = myCommand61.Caption
            myCommand62.Caption = getmrkey("contexts", "context3")
            myCommand62.ToolTip = myCommand62.Caption
            myCommand63.Caption = getmrkey("contexts", "context4")
            myCommand63.ToolTip = myCommand63.Caption
            myCommand64.Caption = getmrkey("contexts", "context5")
            myCommand64.ToolTip = myCommand64.Caption
            myCommand65.Caption = getmrkey("contexts", "context6")
            myCommand65.ToolTip = myCommand65.Caption

            myCommand21.Caption = getmrkey("dues", "due1")
            myCommand21.ToolTip = myCommand21.Caption
            myCommand22.Caption = getmrkey("dues", "due2")
            myCommand22.ToolTip = myCommand22.Caption
            myCommand69.Caption = getmrkey("dues", "due3")
            myCommand69.ToolTip = myCommand69.Caption

            myCommand66.Caption = getmrkey("times", "time1")
            myCommand66.ToolTip = myCommand66.Caption
            myCommand67.Caption = getmrkey("times", "time2")
            myCommand67.ToolTip = myCommand67.Caption
            myCommand68.Caption = getmrkey("times", "time3")
            myCommand68.ToolTip = myCommand68.Caption

            'mycommand86.Caption = "Omni2Map"
            'mycommand86.ToolTip = "Omni2Map"

            Try
                mycommand72.Caption = getmrkey("sends", "send1")
                mycommand73.Caption = getmrkey("sends", "send2")
                mycommand74.Caption = getmrkey("sends", "send3")
                mycommand75.Caption = getmrkey("sends", "send4")
                mycommand76.Caption = getmrkey("sends", "send5")
                mycommand77.Caption = getmrkey("sends", "send6")
                mycommand78.Caption = getmrkey("sends", "send7")
                mycommand79.Caption = getmrkey("sends", "send8")
                mycommand80.Caption = getmrkey("sends", "send9")

                mycommand84.Caption = getmrkey("sends", "send10")
                mycommand85.Caption = getmrkey("sends", "send11")


                mycommand72.ToolTip = (getmrkey("sends", "send1"))
                mycommand73.ToolTip = (getmrkey("sends", "send2"))
                mycommand74.ToolTip = (getmrkey("sends", "send3"))
                mycommand75.ToolTip = (getmrkey("sends", "send4"))
                mycommand76.ToolTip = (getmrkey("sends", "send5"))
                mycommand77.ToolTip = (getmrkey("sends", "send6"))
                mycommand78.ToolTip = (getmrkey("sends", "send7"))
                mycommand79.ToolTip = (getmrkey("sends", "send8"))
                mycommand80.ToolTip = (getmrkey("sends", "send9"))
                mycommand84.ToolTip = (getmrkey("sends", "send10"))
                mycommand85.ToolTip = (getmrkey("sends", "send11"))
            Catch
            End Try

        End Sub
        Sub need_two_topics_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles mycommand83.UpdateState
            Try
                pChecked = False
                If Not m_app.ActiveDocument Is Nothing Then
                    If m_app.ActiveDocument.Selection.Count > 1 Then
                        pEnabled = True
                    Else
                        pEnabled = False
                    End If
                Else
                    pEnabled = False
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub need_a_topic_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles _
            myCommand9.UpdateState, myCommand21.UpdateState, myCommand22.UpdateState, _
            myCommand25.UpdateState, myCommand27.UpdateState, myCommand30.UpdateState, myCommand60.UpdateState, _
            myCommand61.UpdateState, myCommand62.UpdateState, myCommand63.UpdateState, myCommand64.UpdateState, _
            myCommand65.UpdateState, myCommand66.UpdateState, myCommand67.UpdateState, myCommand68.UpdateState, myCommand69.UpdateState, mycommand70.UpdateState, _
            mycommand72.updatestate, mycommand73.UpdateState, mycommand74.UpdateState, mycommand75.UpdateState, mycommand76.UpdateState, mycommand77.UpdateState, mycommand78.UpdateState, _
            mycommand79.UpdateState, mycommand80.UpdateState, mycommand84.UpdateState, mycommand85.UpdateState

            Try
                pChecked = False
                If Not m_app.ActiveDocument Is Nothing Then
                    If Not m_app.ActiveDocument.Selection.PrimaryTopic Is Nothing Then
                        pEnabled = True
                    Else
                        pEnabled = False
                    End If
                Else
                    pEnabled = False
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub need_a_business_topic_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) 'Handles mycommand86.UpdateState

            Try
                pChecked = False
                If Not m_app.ActiveDocument Is Nothing Then
                    If Not m_app.ActiveDocument.Selection.PrimaryTopic Is Nothing Then
                        If Len(m_app.ActiveDocument.Selection.PrimaryTopic.BusinessTopic.BusinessTypeUri) > 0 Then
                            pEnabled = True
                        Else
                            pEnabled = False
                        End If
                    Else
                        pEnabled = False
                    End If
                Else
                    pEnabled = False
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub


        Sub need_a_doc_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles myCommand4.UpdateState, myCommand5.UpdateState, myCommand10.UpdateState, myCommand17.UpdateState
            Try
                pChecked = False
                If Not m_app.ActiveDocument Is Nothing Then
                    pEnabled = True
                Else
                    pEnabled = False
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub need_a_dashboard_UpdateStateXXX(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles myCommand6.UpdateState, mycommand81.UpdateState
            'myCommand13.UpdateState, myCommand14.UpdateState,
            Try
                pChecked = False
                If Not m_app.ActiveDocument Is Nothing Then
                    If m_app.ActiveDocument.CentralTopic.Text = "Daily Actions Substitute" Then
                        pEnabled = True
                    Else
                        pEnabled = False
                    End If
                Else
                    pEnabled = False
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub need_nothing_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles myCommand7.UpdateState, myCommand8.UpdateState, _
             myCommand16.UpdateState, mycommand82.UpdateState, myCommand18.UpdateState, myCommand19.UpdateState, myCommand20.UpdateState, _
            myCommand23.UpdateState, myCommand24.UpdateState, myCommand26.UpdateState, _
            myCommand30.UpdateState, mycommand71.UpdateState
            'myCommand11.UpdateState,myCommand29.UpdateState,
            Try
                pChecked = False
                pEnabled = True
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub


        Sub tLock_Click4() Handles myCommand4.Click
            Dim m As mm.Document
            Try
                For Each m In m_app.Documents
                    If Not m.IsModified Then
                        If Not (m_app.ActiveDocument Is m) Then
                            m.Close()
                        End If
                    End If
                Next

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click5() Handles myCommand5.Click
            Try
                Clipboard.SetText(m_app.ActiveDocument.FullName)
                MsgBox("The string: " & m_app.ActiveDocument.FullName & " : has been placed on clipboard")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click6() Handles myCommand6.Click
            Try
                MsgBox("Next Action Analysis is not implemented yet for ResultsManager 3")
                'm_app.RunMacro(RMPath() & "\ResultManager-X5-dashboard.MMBas")
                next_action_analysis(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click81() Handles mycommand81.Click
            Try
                next_action_analysis(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click7() Handles myCommand7.Click
            Try
                Dim a As New mrform
                a.setapp(m_app)
                a.setmode("o")
                a.Show()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click8() Handles myCommand8.Click
            Try
                Dim a As New mrform
                a.setapp(m_app)
                a.setmode("n")
                a.Show()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click9() Handles myCommand9.Click
            Try
                Dim a As New mrform
                a.setapp(m_app)
                a.setmode("s")
                a.Show()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub
        Sub tLock_Click10() Handles myCommand10.Click
            Try
                Dim m As Mindjet.MindManager.Interop.Document
                For Each m In m_app.Documents
                    If m.IsModified Then
                        m.Save()
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click11() 'Handles myCommand11.Click
            Try
                ao_naa_trend(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub


        Sub tLock_Click13() 'Handles myCommand13.Click
            Try
                Minimapcentral(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click14() 'Handles myCommand14.Click
            Try
                next_actions_to_outlook(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click82() Handles mycommand82.Click
            Try
                usedit = True
                'MindReaderNLP(m_app, InputBox("Enter keywords to markup map" & Chr(10) & " OR " & Chr(10) & "Hit enter to read topic itself"))
                MindReaderNLP(m_app, "")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click83() 'Handles mycommand83.Click
            Try
                usedit = True
                makelist(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click16() Handles myCommand16.Click
            Try
                usedit = True
                'MindReaderNLP(m_app, InputBox("Enter keywords to markup map" & Chr(10) & " OR " & Chr(10) & "Hit enter to read topic itself"))
                Dim a As New mrform
                a.setapp(m_app)
                a.setmode("m")
                a.ShowDialog()
                a = Nothing
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click17() Handles myCommand17.Click
            Try
                Dim t As mm.Topic
                Dim c As Integer
                Dim c1 As Integer
                c1 = 0
                c = 0
                For Each t In m_app.ActiveDocument.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
                    If Not t.IsCalloutTopic Then
                        If Not t.Task.Complete < 0 Then
                            If Not t.Task.Complete = 100 Then
                                c = c + 1
                            End If
                        End If
                        c1 = c1 + 1
                    End If
                Next
                MsgBox(c & " tasks " & c1 & " topics in this map")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click18() Handles myCommand18.Click
            Try
                System.Diagnostics.Process.Start("http://wiki.activityowner.com/index.php?title=ActivityOwner_Tools_Add-in")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click19() Handles myCommand19.Click
            Try
                'donated = True
                'System.Diagnostics.Process.Start("http://www.activityowner.com/donate/")
                Dim a As New MarkupForm

                a.setapp(m_app)
                a.Show()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_mindread() Handles myCommand20.Click
            Try
                usedit = True
                Dim a As New mrform
                a.setmode("q")
                a.setapp(m_app)
                a.Show()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click21() Handles myCommand21.Click
            Try
                MindReaderNLP(m_app, myCommand21.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click22() Handles myCommand22.Click
            Try
                MindReaderNLP(m_app, myCommand22.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click23() Handles myCommand23.Click
            Try
                m_app.Documents.Open(m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps) & "Dashboards\dailyactions.mmap")
            Catch ex As Exception
                MsgBox("Note -- This command assumes have have stored your ResultsManager dailyaction dashboard in My Maps\dashboards\dailyactions.mmap")
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click24() Handles myCommand24.Click
            Try
                m_app.Documents.Open(m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps) & "Daily Capture Map.mmap")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click25() Handles myCommand25.Click
            Try

                Dim a As New mrform
                a.setapp(m_app)
                a.setmode("b")
                a.Show()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click26() Handles myCommand26.Click
            Try
                'Dim a As New AboutBox1
                'a.Show()
                'a = Nothing
                Dim a As New mindreaderoptionsform
                a.ShowDialog()
                setflexiblecommands()
                a = Nothing
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click27() Handles myCommand27.Click
            Try
                m_app.ActiveDocument.Selection.Cut()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click29() 'Handles myCommand29.Click
            Try
                m_app.Documents.Open(m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps) & "Dashboards\dailyactions_temp.mmap")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click30() Handles myCommand30.Click
            Try
                usedit = True
                Dim a As New mrform
                a.setapp(m_app)
                a.setmode("k")
                a.Show()
                a = Nothing
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click60() Handles myCommand60.Click
            Try
                MindReaderNLP(m_app, myCommand60.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click61() Handles myCommand61.Click
            Try
                MindReaderNLP(m_app, myCommand61.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click62() Handles myCommand62.Click
            Try
                MindReaderNLP(m_app, myCommand62.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click63() Handles myCommand63.Click
            Try
                MindReaderNLP(m_app, myCommand63.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click64() Handles myCommand64.Click
            Try
                MindReaderNLP(m_app, myCommand64.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click65() Handles myCommand65.Click
            Try
                MindReaderNLP(m_app, myCommand65.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click66() Handles myCommand66.Click
            Try
                MindReaderNLP(m_app, myCommand66.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click67() Handles myCommand67.Click
            Try
                MindReaderNLP(m_app, myCommand67.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click68() Handles myCommand68.Click
            Try
                MindReaderNLP(m_app, myCommand68.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click69() Handles myCommand69.Click
            Try
                MindReaderNLP(m_app, myCommand69.Caption)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click70() Handles mycommand70.Click
            Try
                usedit = True
                MindReaderNLP(m_app, "")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click71() Handles mycommand71.Click
            Dim s As String
            Try
                usedit = True
                s = InputBox("Enter calulation to perform")
                MsgBox(s & "=" & eval(m_app, s), MsgBoxStyle.OkOnly, "MindReader Calculator")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub mindreadersend(ByRef s As String)
            If Len(s) > 0 Then
                If m_app.ActiveDocument.Selection.Count > 0 And Not m_app.ActiveDocument.Selection.PrimaryTopic.IsCentralTopic Then
                    m_app.ActiveDocument.Selection.Cut()
                    mindreaderopen(m_app, m_app.ActiveDocument, "/send" & s)
                End If
            End If
        End Sub
        Sub tLock_Click72() Handles mycommand72.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send1"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click73() Handles mycommand73.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send2"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click74() Handles mycommand74.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send3"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click75() Handles mycommand75.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send4"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click76() Handles mycommand76.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send5"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click77() Handles mycommand77.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send6"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click78() Handles mycommand78.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send7"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click79() Handles mycommand79.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send8"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click80() Handles mycommand80.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send9"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click84() Handles mycommand84.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send10"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click85() Handles mycommand85.Click
            Try
                usedit = True
                mindreadersend(getmrkey("sends", "send11"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        'Sub tLock_Click86() Handles mycommand86.Click
        '   Try
        'usedit = True
        'o2m(m_app)
        'm_app.ActiveDocument.Selection.PrimaryTopic.BusinessTopic.Refresh()
        '  Catch ex As Exception
        '     MsgBox(ex.ToString)
        'End Try
        'End Sub

        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            'usedit = False 'disable
            donatedelay("http://www.activityowner.com/mindreader-trial-user/")
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand4)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand5)
            'If RMinstalled() Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand6)
            'End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand7)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand8)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand9)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand10)
            'If RMinstalled() Then
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand11)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand13)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand14)
            'End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand16)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand17)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand18)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand19)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand20)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand21)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand22)
            'If RMinstalled() Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand23)
            'End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand24)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand25)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand26)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand27)
            'If RMinstalled() Then
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand29)
            'End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand30)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand60)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand61)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand62)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand63)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand64)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand65)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand66)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand67)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand68)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand69)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand70)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand71)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand72)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand73)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand74)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand75)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand76)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand77)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand78)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand79)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand80)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand84)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand85)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand86)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand81)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand82)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand83)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_app)

        End Sub

    End Class

End Namespace


