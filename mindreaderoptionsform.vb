Public Class mindreaderoptionsform

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextGroup.Enter

    End Sub
    Private Sub showform(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        ContextBox1.Text = getmrkey("contexts", "context1")
        ContextBox2.Text = getmrkey("contexts", "context2")
        ContextBox3.Text = getmrkey("contexts", "context3")
        ContextBox4.Text = getmrkey("contexts", "context4")
        ContextBox5.Text = getmrkey("contexts", "context5")
        ContextBox6.Text = getmrkey("contexts", "context6")
        ContextTabLabelBox.Text = getmrkey("tablabels", "contextlabel")

        DueBox1.Text = getmrkey("dues", "due1")
        DueBox2.Text = getmrkey("dues", "due2")
        DueBox3.Text = getmrkey("dues", "due3")
        DueTabLabelBox.Text = getmrkey("tablabels", "duelabel")

        TimeBox1.Text = getmrkey("times", "time1")
        TimeBox2.Text = getmrkey("times", "time2")
        TimeBox3.Text = getmrkey("times", "time3")
        TimeTabLabelBox.Text = getmrkey("tablabels", "timelabel")

        SendTextBox1.Text = getmrkey("sends", "send1")
        SendTextBox2.Text = getmrkey("sends", "send2")
        SendTextBox3.Text = getmrkey("sends", "send3")
        SendTextBox4.Text = getmrkey("sends", "send4")
        SendTextBox5.Text = getmrkey("sends", "send5")
        SendTextBox6.Text = getmrkey("sends", "send6")
        SendTextBox7.Text = getmrkey("sends", "send7")
        SendTextBox8.Text = getmrkey("sends", "send8")
        SendTextBox9.Text = getmrkey("sends", "send9")
        SendTextBox10.Text = getmrkey("sends", "send10")
        SendTextBox11.Text = getmrkey("sends", "send11")

    End Sub
    Private Sub ContextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextBox1.TextChanged

    End Sub

    Private Sub ContextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextBox2.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        setmrkey("contexts", "context1", ContextBox1.Text)
        setmrkey("contexts", "context2", ContextBox2.Text)
        setmrkey("contexts", "context3", ContextBox3.Text)
        setmrkey("contexts", "context4", ContextBox4.Text)
        setmrkey("contexts", "context5", ContextBox5.Text)
        setmrkey("contexts", "context6", ContextBox6.Text)
        setmrkey("tablabels", "contextlabel", ContextTabLabelBox.Text)

        setmrkey("dues", "due1", DueBox1.Text)
        setmrkey("dues", "due2", DueBox2.Text)
        setmrkey("dues", "due3", DueBox3.Text)
        setmrkey("tablabels", "duelabel", DueTabLabelBox.Text)

        setmrkey("times", "time1", TimeBox1.Text)
        setmrkey("times", "time2", TimeBox2.Text)
        setmrkey("times", "time3", TimeBox3.Text)
        setmrkey("tablabels", "timelabel", TimeTabLabelBox.Text)

        setmrkey("sends", "send1", SendTextBox1.Text)
        setmrkey("sends", "send2", SendTextBox2.Text)
        setmrkey("sends", "send3", SendTextBox3.Text)
        setmrkey("sends", "send4", SendTextBox4.Text)
        setmrkey("sends", "send5", SendTextBox5.Text)
        setmrkey("sends", "send6", SendTextBox6.Text)
        setmrkey("sends", "send7", SendTextBox7.Text)
        setmrkey("sends", "send8", SendTextBox8.Text)
        setmrkey("sends", "send9", SendTextBox9.Text)
        setmrkey("sends", "send10", SendTextBox10.Text)
        setmrkey("sends", "send11", SendTextBox11.Text)

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub

    Private Sub SendTextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendTextBox9.TextChanged

    End Sub

    Private Sub SendTextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendTextBox8.TextChanged

    End Sub

    Private Sub DueBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DueBox3.TextChanged

    End Sub

    Private Sub DueBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DueBox1.TextChanged

    End Sub
End Class