Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module ao_markers
    Sub addmarker(ByRef t As mm.Topic, ByRef u As mm.Topic, ByRef MG As String, ByRef mstr As String, ByRef exclusive As Boolean)
        'AO August 2013
        'This routine adds a text mapmarker to a topic and optionally to a underlying topic it is linked to (used on query topics)
        'It avoids creation of general tags by creating marker groups and tags as necessary
        't is the topic acted on
        'u is an optional underlying topic on another map
        'MG is the marker group t is associated with
        'mstr is the textlabel
        'Sept 13 -- added exclusive flag
        Dim l As mm.MapMarker
        Dim m As mm.MapMarkerGroup
        Dim mmm As mm.MapMarkerGroup
        Dim found As Boolean
        found = False
        mstr = Replace(mstr, "^", "") 'rm2/mindreader indicator of area
        mstr = Replace(mstr, "~", "") 'rm2/mindreader indicator of category
        If Len(mstr) > 0 Then
            For Each m In t.Document.MapMarkerGroups
                If m.Name = MG Then 'marker group found
                    mmm = m
                    found = True
                    Exit For
                End If
            Next
            If Not found Then 'add marker group
                mmm = t.Document.MapMarkerGroups.AddTextLabelMarkerGroup(MG)
                mmm.MutuallyExclusive = exclusive
                Debug.Print("marker group added")
            End If
            found = False
            For Each l In mmm
                If l.Label = mstr Then 'marker found
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                removecategory(t, mstr)
                mmm.AddTextLabelMarker(mstr) 'adding marker
            End If
            t.TextLabels.AddTextLabel(mstr) 'adding textlabel to topic
        End If
        If Not u Is Nothing Then addmarker(u, Nothing, MG, mstr, exclusive) 'add to underlying topic
    End Sub
    Function hastextmarkergroup(ByRef t As mm.Topic, ByRef mgstring As String) As Boolean
        Dim mg As mm.MapMarkerGroup
        Dim m As mm.TextLabel
        Dim tl As mm.TextLabel
        hastextmarkergroup = False
        If t.TextLabels.Count > 0 Then
            For Each mg In t.Document.MapMarkerGroups
                If mg.Name = mgstring Then
                    If mg.Type = Mindjet.MindManager.Interop.MmMapMarkerGroupType.mmMapMarkerGroupTypeTextLabel Then
                        For Each m In mg
                            For Each tl In t.TextLabels
                                If tl.Name = m.Name Then
                                    hastextmarkergroup = True
                                    Exit Function
                                End If
                            Next
                        Next
                    End If
                End If
            Next
        Else
            hastextmarkergroup = False
        End If
    End Function

    Sub removecategory(ByRef t As mm.Topic, ByRef mstr As String)
        Dim l As mm.MapMarker
        Dim m As mm.MapMarkerGroup
        Dim mmm As mm.MapMarkerGroup
        Dim found As Boolean
        found = False
        If Len(mstr) > 0 Then
            For Each m In t.Document.MapMarkerGroups
                If m.Name = "Categories" Then
                    mmm = m
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Exit Sub
            End If
            found = False
            For Each l In mmm
                If l.Label = mstr Then
                    Debug.Print(mstr & " category found and now deleting")
                    l.Delete()
                    Exit For
                End If
            Next

        End If
    End Sub

    Sub copymarkers(ByRef odoc As mm.Document, ByRef ddoc As mm.Document)
        Dim ol As mm.MapMarker
        Dim dl As mm.MapMarker
        Dim om As mm.MapMarkerGroup
        Dim dm As mm.MapMarkerGroup
        Dim dmfound As Boolean
        Dim dlfound As Boolean
        Debug.Print("In copymarkers:")
        dm = Nothing
        For Each om In odoc.MapMarkerGroups
            dmfound = False
            If om.Name = "Contexts" Or om.Name = "Areas" Or om.Name = "Focuses" Or om.Name = "Goals" Or om.Name = "Repeat" Then
                For Each dm In ddoc.MapMarkerGroups
                    If dm.Name = om.Name Then
                        Debug.Print("destination marker group found:" & dm.Name)
                        dmfound = True
                        Exit For
                    End If
                Next
                If Not dmfound Then
                    Debug.Print("creating destination map marker group:" & om.Name)
                    dm = ddoc.MapMarkerGroups.AddTextLabelMarkerGroup(om.Name)
                End If
                Debug.Print("looking for labels in " & om.Name)
                Debug.Print(om.Name & " has " & om.Count & " labels")
                For Each ol In om
                    Debug.Print("looking for " & ol.Label & " in " & dm.Name)
                    dlfound = False
                    For Each dl In dm
                        If dl.Label = ol.Label Then
                            dlfound = True
                            Debug.Print("destination label found:" & dl.Label)
                            Exit For
                        End If
                    Next
                    If Not dlfound Then
                        Debug.Print("adding destination label:" & ol.Label)
                        dm.AddTextLabelMarker(ol.Label)
                    End If
                Next
            End If
        Next
    End Sub
End Module
