Imports mm = Mindjet.MindManager.Interop
Public Class mrform
    Private fapp As Mindjet.MindManager.Interop.Application
    Private mrmode As String
    Public Sub setapp(ByVal app As Mindjet.MindManager.Interop.Application)
        fapp = app
    End Sub
    Public Sub setmode(ByVal mode As String)
        mrmode = mode
        If mrmode = "m" Then Label1.Text = "Enter keywords to markup topic(m) or ENTER to read(c)"
        If mrmode = "o" Then Label1.Text = "Enter map keyword to open map (o)"
        If mrmode = "q" Then Label1.Text = "Enter task to mindread into a map (q)"
        If mrmode = "b" Then Label1.Text = "Enter task to preceed selected task(b)"
        If mrmode = "s" Then Label1.Text = "Enter destination keyword to send topic(s)"
        If mrmode = "k" Then Label1.Text = "Enter link keyword for topic as destination"
        If mrmode = "n" Then Label1.Text = "Enter Desired outcome for new project map and keywords(n)"
    End Sub
    Public Function getmode() As String
        getmode = mrmode
    End Function
    Public Function getapp() As Mindjet.MindManager.Interop.Application
        getapp = fapp
    End Function
    Private Sub TextBox1_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = Keys.Enter Then
            If mrmode = "" Then mrmode = "m"
            If mrmode = "m" Then
                MindReaderNLP(fapp, TextBox1.Text)
            End If
            If mrmode = "q" Then
                If Len(TextBox1.Text) > 0 Then
                    mindreaderopen(fapp, Nothing, "/queu" & TextBox1.Text)
                    MindReaderNLP(fapp, "")
                End If
            End If
            If mrmode = "b" Then
                Dim n As mm.Topic
                n = fapp.ActiveDocument.Selection.PrimaryTopic.AddSubTopic(TextBox1.Text)
                fapp.ActiveDocument.Selection.Set(n)
                MindReaderNLP(fapp, "")
                'm_app.RunMacro(Environ("ProgramFiles") & "Gyronix\ResultManager\ResultManager-X5-RefreshNow.MMBas'")
            End If
            If mrmode = "o" Then
                mindreaderopen(fapp, Nothing, "/open" & TextBox1.Text)
            End If
            If mrmode = "k" Then
                AddDestinationKeyword(fapp, TextBox1.Text)
            End If
            If mrmode = "s" Then
                If Len(TextBox1.Text) > 0 Then
                    If fapp.ActiveDocument.Selection.Count > 0 And Not fapp.ActiveDocument.Selection.PrimaryTopic.IsCentralTopic Then
                        fapp.ActiveDocument.Selection.Cut()
                        mindreaderopen(fapp, fapp.ActiveDocument, "/send" & TextBox1.Text)
                    End If
                End If
            End If
            If mrmode = "n" Then
                newprojectmap(fapp, TextBox1.Text)
            End If

            e.Handled = True
            Me.TopMost = True
            Me.BringToFront()
            Me.Activate()
            Me.Focus()
            TextBox1.Focus()
            TextBox1.Text = ""
            'Dim i As Integer
            'Dim c As Integer
            'If Not fapp.ActiveDocument Is Nothing Then
            '    c = fapp.ActiveDocument.Selection.Count
            '    Dim s(20) As Mindjet.MindManager.Interop.IDocumentObject
            '    If c > 20 Then
            '        MsgBox("Select less topics")
            '    End If
            '    If c > 0 Then
            '        For i = 1 To c
            '            s(i) = fapp.ActiveDocument.Selection.Item(i)
            '        Next
            '        fapp.ActiveDocument.Selection.RemoveAll()
            '    End If
            '    Me.Show()
            '    Me.Focus()
            '    If c > 0 Then
            '        For i = 1 To c
            '            fapp.ActiveDocument.Selection.Add(s(i))
            '        Next
            '    End If
            '    s = Nothing
            'Else
            '    Me.Show()
            '    Me.Activate()
            'End If
        End If
        If Len(TextBox1.Text) = 1 And Asc(e.KeyChar) = Keys.Space Then
            If TextBox1.Text = "o" Or _
                TextBox1.Text = "m" Or _
                TextBox1.Text = "b" Or _
                TextBox1.Text = "k" Or _
                TextBox1.Text = "q" Or _
                TextBox1.Text = "s" Or _
                TextBox1.Text = "n" Then
                setmode(Mid(TextBox1.Text, 1, 1))
                e.Handled = True
                TextBox1.Text = ""
            End If
        End If

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

  
    Private Sub Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub


    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        If mrmode = "" Then mrmode = "m"
        If mrmode = "m" Then
            MindReaderNLP(fapp, TextBox1.Text)
        End If
        If mrmode = "q" Then
            If Len(TextBox1.Text) > 0 Then
                mindreaderopen(fapp, Nothing, "/queu" & TextBox1.Text)
                MindReaderNLP(fapp, "")
            End If
        End If
        If mrmode = "b" Then
            Dim n As mm.Topic
            n = fapp.ActiveDocument.Selection.PrimaryTopic.AddSubTopic(TextBox1.Text)
            fapp.ActiveDocument.Selection.Set(n)
            MindReaderNLP(fapp, "")
            'm_app.RunMacro(Environ("ProgramFiles") & "Gyronix\ResultManager\ResultManager-X5-RefreshNow.MMBas'")
        End If
        If mrmode = "o" Then
            mindreaderopen(fapp, Nothing, "/open" & TextBox1.Text)
        End If
        If mrmode = "k" Then
            AddDestinationKeyword(fapp, TextBox1.Text)
        End If
        If mrmode = "s" Then
            If Len(TextBox1.Text) > 0 Then
                If fapp.ActiveDocument.Selection.Count > 0 And Not fapp.ActiveDocument.Selection.PrimaryTopic.IsCentralTopic Then
                    fapp.ActiveDocument.Selection.Cut()
                    mindreaderopen(fapp, fapp.ActiveDocument, "/send" & TextBox1.Text)
                End If
            End If
        End If
        If mrmode = "n" Then
            newprojectmap(fapp, TextBox1.Text)
        End If
        Me.Close()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
End Class