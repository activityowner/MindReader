Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module Modulemindmapcentral
    'ao_minimapcentral 'Copyright: http://creativecommons.org/licenses/by-nc-nd/3.0/
    'Information: http://wiki.activityowner.com
    '02Jan2011
    Sub Minimapcentral(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Dim n As mm.Document
        Dim a As mm.Document
        Dim t As mm.Topic
        Dim s As mm.Topic
        Dim b As mm.Topic
        Dim count As Integer
        Dim found As Boolean
        count = 0
        a = m_app.ActiveDocument
        n = getmap(m_app, m_app.GetPath(mm.MmDirectory.mmDirectoryMyMaps) & "\TempMapCentral.mmap")
        n.Activate()
        For Each t In n.CentralTopic.AllSubTopics
            t.Delete()
        Next
        n.CentralTopic.Text = "Scanning for active maps"
        For Each t In a.Range(mm.MmRange.mmRangeAllTopics)
            found = False

            If t.HasHyperlink Then
                If InStr(LCase(t.Text), "in-tray") = 0 Then
                    t.Hyperlink.Absolute = True
                    For Each s In n.CentralTopic.AllSubTopics
                        If t.Hyperlink.Address = s.Hyperlink.Address Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        b = n.CentralTopic.AddSubTopic(t.Text)
                        Debug.Print(t.Hyperlink.Address)
                        b.CreateHyperlink(t.Hyperlink.Address)
                        b.Notes.Text = t.Hyperlink.Address
                        count = count + 1
                    End If
                End If
            End If
        Next
        n.CentralTopic.Text = Str(count) & " maps with active tasks"
        n.Save()
    End Sub


End Module
