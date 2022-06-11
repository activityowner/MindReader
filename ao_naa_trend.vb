Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports ex = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module naa_trend_module
    'ao_naa_trend 17Apr09 http://creativecommons.org/licenses/by-nc-nd/3.0/
    '#uses "ao_common.mmbas"
    Sub ao_naa_trend(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Dim Excelapp As ex.Application
        Dim exceldoc As ex.Workbook
        Dim ConfigDoc As mm.Document
        Dim logdoc As mm.Document
        Dim logdocname As String
        Dim branch As mm.Topic
        Dim ndate As String
        Dim nscore As Double
        Dim configdocname As String
        Dim nlog As String
        Dim i As Integer
        Dim lastlen As Integer
        Dim sheet As Object
        Dim factor As Integer
        factor = 1
        Excelapp = CreateObject("excel.application")
        Excelapp.Visible = True
        configdocname = "AO\NAAconfig.mmap"
        ConfigDoc = getmap(m_app, configdocname)
        logdocname = getoption("logdocname", ConfigDoc, Nothing)
        logdoc = getmap(m_app, logdocname)
        exceldoc = Excelapp.Workbooks.Add
        For i = 1 To exceldoc.Sheets.Count - 1
            exceldoc.Sheets(i).Delete()
        Next
        For Each branch In logdoc.CentralTopic.AllSubTopics
            'Set branch = createmainbranch(branch.Text,logdoc)
            sheet = exceldoc.Sheets.Add
            nlog = branch.Notes.Text
            i = 0
            While Len(nlog) > 0
                i = i + 1
                lastlen = Len(nlog)
                ndate = Left(nlog, InStr(nlog, ",") - 1)
                ndate = Replace(ndate, vbCrLf, "")
                ndate = Replace(ndate, vbCr, "")
                ndate = Replace(ndate, vbLf, "")
                nlog = Right(nlog, Len(nlog) - InStr(nlog, ","))
                If InStr(nlog, vbCrLf) > 0 Then
                    nscore = Val(Left(nlog, InStr(nlog, vbCrLf) - 1)) * factor
                    nlog = Right(nlog, Len(nlog) - InStr(nlog, vbCrLf))
                Else
                    nscore = Val(nlog) * factor
                    nlog = ""
                End If
                sheet.Cells(i, 1).Value = Trim(ndate)
                sheet.Cells(i, 2).Value = nscore
                If Not Err.Number = 0 Then
                    Debug.Print("error parsing line " & Str(i))
                    Err.Clear()
                End If
                If Len(nlog) = lastlen Then
                    MsgBox("stuck on line " & Str(i))
                    Exit While
                End If
            End While
            If i > 1 Then
                With exceldoc.Charts.Add
                    .ChartType = ex.XlChartType.xlXYScatter
                    .SetSourceData(Source:=sheet.Range("A1:B" & Trim(Str(i))))
                    .Name = Trim(branch.Text)
                    .Axes(ex.XlAxisType.xlValue).HasTitle = True
                    .Axes(ex.XlAxisType.xlValue).AxisTitle.Text = branch.Text
                    .Axes(ex.XlAxisType.xlCategory).HasTitle = True
                    .Axes(ex.XlAxisType.xlCategory).AxisTitle.Text = "Date"
                    .Axes(ex.XlAxisType.xlCategory).TickLabels.NumberFormat = "MMM-yyyy"
                    .Legend.Delete()
                End With
            End If
            sheet.name = Trim(branch.Text) & "Data"
            sheet.Visible = True
        Next
        ConfigDoc = Nothing
        logdoc = Nothing
        sheet = Nothing
        exceldoc = Nothing
        Excelapp = Nothing
    End Sub
End Module
