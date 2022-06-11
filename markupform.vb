Public Class MarkupForm
    Private fapp As Mindjet.MindManager.Interop.Application
    Private Sub markupform_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Text = getmrkey("buttons", "Button1")
        Button2.Text = getmrkey("buttons", "Button2")
        Button3.Text = getmrkey("buttons", "Button3")
        Button4.Text = getmrkey("buttons", "Button4")
        Button5.Text = getmrkey("buttons", "Button5")
        Button6.Text = getmrkey("buttons", "Button6")
        Button7.Text = getmrkey("buttons", "Button7")
        Button8.Text = getmrkey("buttons", "Button8")
        Button9.Text = getmrkey("buttons", "Button9")
        Button10.Text = getmrkey("buttons", "Button10")
        Button11.Text = getmrkey("Buttons", "Button11")
        Button12.Text = getmrkey("Buttons", "Button12")
        Button13.Text = getmrkey("Buttons", "Button13")
        Button14.Text = getmrkey("Buttons", "Button14")
        Button15.Text = getmrkey("Buttons", "Button15")
        Button16.Text = getmrkey("Buttons", "Button16")
        Button17.Text = getmrkey("Buttons", "Button17")
        Button18.Text = getmrkey("Buttons", "Button18")
        Button19.Text = getmrkey("Buttons", "Button19")
        Button20.Text = getmrkey("Buttons", "Button20")
        Button21.Text = getmrkey("Buttons", "Button21")
        Button22.Text = getmrkey("Buttons", "Button22")
        Button23.Text = getmrkey("Buttons", "Button23")
        Button24.Text = getmrkey("Buttons", "Button24")
        Button25.Text = getmrkey("Buttons", "Button25")
        Button26.Text = getmrkey("Buttons", "Button26")
        Button27.Text = getmrkey("Buttons", "Button27")
        Button28.Text = getmrkey("Buttons", "Button28")
        Button29.Text = getmrkey("Buttons", "Button29")
        Button30.Text = getmrkey("Buttons", "Button30")
        Button31.Text = getmrkey("Buttons", "Button31")
        Button32.Text = getmrkey("Buttons", "Button32")
        Button33.Text = getmrkey("Buttons", "Button33")
        Button34.Text = getmrkey("Buttons", "Button34")
        Button35.Text = getmrkey("Buttons", "Button35")
        Button36.Text = getmrkey("Buttons", "Button36")
        Button37.Text = getmrkey("Buttons", "Button37")
        Button38.Text = getmrkey("Buttons", "Button38")
        Button39.Text = getmrkey("Buttons", "Button39")
        Button40.Text = getmrkey("Buttons", "Button40")
    End Sub
    Public Sub setapp(ByVal app As Mindjet.MindManager.Interop.Application)
        fapp = app
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        runbutton(CType(sender, Button))
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        runbutton(Button2)
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        runbutton(Button3)
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        runbutton(Button4)
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        runbutton(Button5)
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        runbutton(Button6)
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        runbutton(Button7)
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        runbutton(Button8)
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        runbutton(Button9)
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        runbutton(Button10)
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        runbutton(Button11)
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        runbutton(Button12)
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        runbutton(Button13)
    End Sub
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        runbutton(Button14)
    End Sub
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        runbutton(Button15)
    End Sub
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        runbutton(Button16)
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        runbutton(Button17)
    End Sub
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        runbutton(Button18)
    End Sub
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        runbutton(Button19)
    End Sub
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        runbutton(Button20)
    End Sub
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        runbutton(Button21)
    End Sub
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        runbutton(Button22)
    End Sub
    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        runbutton(Button23)
    End Sub
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        runbutton(Button24)
    End Sub
    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        runbutton(Button25)
    End Sub
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        runbutton(Button26)
    End Sub
    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        runbutton(Button27)
    End Sub
    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        runbutton(Button28)
    End Sub
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        runbutton(Button29)
    End Sub
    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        runbutton(Button30)
    End Sub
    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        runbutton(Button31)
    End Sub
    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        runbutton(Button32)
    End Sub
    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        runbutton(Button33)
    End Sub
    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        runbutton(Button34)
    End Sub
    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        runbutton(Button35)
    End Sub
    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        runbutton(Button36)
    End Sub
    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        runbutton(Button37)
    End Sub
    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        runbutton(Button38)
    End Sub
    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        runbutton(Button39)
    End Sub
    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        runbutton(Button40)
    End Sub
    
    Private Sub runbutton(ByRef button As Button)
        If ConfigureCheckBox.Checked Then
            button.Text = InputBox("enter new value", button.Text)
            setmrkey("buttons", button.Name, button.Text)
            ConfigureCheckBox.Checked = False
            button.Refresh()
        Else
            If Not fapp.ActiveDocument Is Nothing Then
                If fapp.ActiveDocument.Selection.Count > 0 Then
                    MindReaderNLP(fapp, button.Text)
                End If
            End If
        End If
    End Sub


    Private Sub ConfigureCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfigureCheckBox.CheckedChanged

    End Sub
End Class