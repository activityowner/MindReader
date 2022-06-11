Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports outlook = Microsoft.Office.Interop.Outlook
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module aonextactionstooutlookmodule
    'ao_next_actions_to_outlook   http://creativecommons.org/licenses/by-sa/2.5/ http://www.activityowner.com
    '22Nov2009 initial version
    '17Jan2010 debugging
    '09Feb2010 fix bug duplicating import of new tasks
    '10Feb2010 fix bug not removing completed tasks from outlook
    'note that you need to set a reference to microsoft outlook 11 or 12 library!!
    'this macro does the following:
    '1)delete previously exported or transferred items unless marked complete
    '  if marked complete -- run mtc inside of mm on them
    '2)transfer new items into MM via mindreader
    '  leave behind but mark as "ao_transferred"
    '  assume they will come back after a dashboard refresh
    '3) export next actions from daily action dashboard
    '#uses "ao_common.mmbas"
    '#uses "ao_mindreader_common.mmbas"
    Public Const ao_created = "ao_created"
    Public Const aotransferred = "ao_transferred"

    Sub next_actions_to_outlook(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Dim d As mm.Document
        Dim t As mm.Topic
        Dim na As mm.Topic
        Dim st As mm.Topic
        Dim sst As mm.Topic
        Dim markedcomplete As Boolean
        Dim outlooktask As outlook.TaskItem
        Const natext = "My committed Next Actions"
        Const contacttext = "Contact.."
        Dim taskstore As outlook.MAPIFolder
        Dim found As Boolean
        Dim outlookapp As New Microsoft.Office.Interop.Outlook.Application
        If 1 = 1 Then
            MsgBox("Not upgraded for RM3")
        Else
            taskstore = outlookapp.GetNamespace("MAPI").GetDefaultFolder(outlook.OlDefaultFolders.olFolderTasks)
            d = m_app.ActiveDocument
            'If Not f_IsADashboardtopic(mapd) Then
            'MsgBox("Must run this on daily action dashboard")
            'Exit Sub
            'End If
            found = True
            While found
                found = False
                For Each outlooktask In taskstore.Items
                    If InStr(outlooktask.Body, ao_created) > 0 Or InStr(outlooktask.Body, aotransferred) > 0 Then
                        If outlooktask.PercentComplete < 100 Then 'just delete incomplete tasks that will be replaced
                            outlooktask.Delete()
                            found = True
                        Else
                            markedcomplete = False
                            For Each t In m_app.ActiveDocument.Range(mm.MmRange.mmRangeAllTopics)
                                If InStr(t.Text, outlooktask.Subject) = 1 And Len(t.Text) = Len(outlooktask.Subject) Then
                                    found = True
                                    If Not markedcomplete Then
                                        markedcomplete = True
                                        m_app.ActiveDocument.Selection.Set(t)
                                        Marktaskcomplete(m_app)
                                    Else
                                        t.Task.Complete = 100
                                    End If
                                Else
                                    found = False
                                End If
                            Next
                            outlooktask.Delete()
                        End If
                    Else
                        import_task(m_app, outlooktask)
                        outlooktask.Body = aotransferred & vbCrLf & outlooktask.Body
                        outlooktask.Save()
                    End If
                Next
            End While
            d.Activate()
            For Each t In d.CentralTopic.AllSubTopics
                If InStr(LCase(t.Text), LCase(natext)) = 1 Then na = t
            Next
            If Not na Is Nothing Then
                For Each t In na.AllSubTopics
                    If Not InStr(LCase(t.Text), LCase(contacttext)) > 0 Then
                        For Each st In t.AllSubTopics
                            If st.Task.Complete < 100 And Not st.Icons.HasStockIcon(mm.MmStockIcon.mmStockIconCheck) Then addoltask(st, t.Text, taskstore, outlooktask, outlookapp)
                        Next
                    Else
                        For Each st In t.AllSubTopics
                            For Each sst In st.AllSubTopics
                                If sst.Task.Complete < 100 And Not sst.Icons.HasStockIcon(mm.MmStockIcon.mmStockIconCheck) Then addoltask(sst, st.Text, taskstore, outlooktask, outlookapp)
                            Next
                        Next
                    End If
                Next
            Else
                MsgBox("Next Action branch not found")
            End If
        End If
    End Sub
    Sub addoltask(ByRef t As mm.Topic, ByVal outlookcategory As String, ByRef taskstore As outlook.MAPIFolder, ByRef outlooktask As outlook.TaskItem, ByRef olapp As Microsoft.Office.Interop.Outlook.Application)
        outlooktask = olapp.CreateItem(outlook.OlItemType.olTaskItem)
        outlooktask.Subject = t.Text
        If Not isdate0(t.Task.DueDate) Then outlooktask.DueDate = t.Task.DueDate
        If t.Task.Priority = 1 Then outlooktask.Importance = outlook.OlImportance.olImportanceHigh
        outlooktask.Categories = outlookcategory
        outlooktask.Body = ao_created
        outlooktask.Move(taskstore)
    End Sub
    Sub import_task(ByRef m_app As Mindjet.MindManager.Interop.Application, ByRef obj As Outlook.TaskItem)
        Dim hlink As String
        Dim mindreaderstring As String
        If InStr(obj.Body, "Outlook:") = 1 Then 'this came from outlinker
            If InStr(obj.Body, vbCrLf) > 0 Then
                hlink = Left(obj.Body, InStr(obj.Body, vbCrLf) - 1) & "|" & Right(obj.Body, Len(obj.Body) - InStr(obj.Body, vbCrLf) - 1)
            Else
                hlink = obj.Body
            End If
        Else
            hlink = obj.Body 'transfer task notes
        End If
        mindreaderstring = "[" + obj.Companies + obj.Categories + "]" + obj.Subject
        'need better empty due date screen
        If Not Str(obj.DueDate) = "1/1/4501" Then mindreaderstring = mindreaderstring + "[" + Str(obj.DueDate) + "]"
        m_app.Documents.Add()
        m_app.ActiveDocument.Activate()
        m_app.ActiveDocument.CentralTopic.Notes.Text = "1"
        m_app.ActiveDocument.CentralTopic.Notes.Commit()
        m_app.ActiveDocument.CentralTopic.Text = mindreaderstring
        mindreaderopen(m_app, Nothing, "")
        MindReaderNLP(m_app, "")
    End Sub


End Module
