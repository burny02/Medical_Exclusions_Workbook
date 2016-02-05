Imports System.Threading

Public Class Form1
    Private Dt As New DataTable()
    Private LatestThread As Integer = 0
    Private CondID As Long = 0
    Private AllowNavigate As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUp(Me)

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

        Dim clm1 As New DataColumn
        clm1.ColumnName = "CondName"
        Dt.Columns.Add(clm1)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        ListBox1.Visible = False
        CondID = 0

        If TextBox1.Text.ToString <> "" Then
            Dim trd = New Thread(AddressOf RefreshList)
            LatestThread = trd.ManagedThreadId
            trd.Start()
        End If

    End Sub

    Private Sub RefreshList()

        Dim StringLength As Integer = 20
        Dim TempOverclass As TemplateDB.OverClass = NewOverclass()
        Dim SearchString As String = TextBox1.Text.ToString
        Dim TempDT As DataTable

        If TextBox1.Text.ToString <> "" Then

            TempDT = TempOverclass.TempDataTable("SELECT TOP 4 Cond, ConditionID FROM " &
            "(SELECT Cond, WhichOrder, ConditionID FROM (SELECT TOP 4 iif(len(CondName)>" &
            StringLength & ", '...' & Left(CondName," & StringLength & "),CondName) AS Cond, CondName, ConditionID, 'A' AS WhichOrder " &
            "FROM ConditionTbl WHERE CondName LIKE '" & SearchString & "%'" &
            "UNION SELECT TOP 4 iif(len(CondName)>" &
            StringLength & ",'...' & Left(CondName," & StringLength & "),CondName) AS Cond, CondName, ConditionID, 'B' AS WhichOrder " &
            "FROM ConditionTbl WHERE CondName LIKE '%" & SearchString & "%') ORDER BY WhichOrder ASC)")

            If Thread.CurrentThread.ManagedThreadId = LatestThread And TextBox1.Text.ToString <> "" Then
                Dt = TempDT
            End If

        End If

        TempDT = Nothing
        TempOverclass = Nothing

        If Thread.CurrentThread.ManagedThreadId = LatestThread And TextBox1.Text.ToString <> "" Then
            Me.Invoke(New Action(AddressOf RefreshListBox))
        End If

    End Sub

    Private Sub RefreshListBox()
        ListBox1.DataSource = Dt
        ListBox1.ValueMember = "ConditionID"
        ListBox1.DisplayMember = "Cond"
        If Dt.Rows.Count > 0 Then ListBox1.Visible = True
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick

        TextBox2.Text = ListBox1.Text.ToString
        CondID = ListBox1.SelectedValue.ToString
        Me.TabControl1.SelectedIndex = 1
        Me.TabControl1_Selecting(Me.TabControl1, New TabControlCancelEventArgs(TabPage2, 0, False, TabControlAction.Selecting))

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Select Case e.TabPage.Text

            Case "View"
                If CondID = 0 Then
                    e.Cancel = True
                    Exit Sub
                Else
                    RefreshPage()
                End If

        End Select

    End Sub

    Private Sub RefreshPage()

        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True

        Dim SQLString As String = "SELECT MaleOrFemale, WhatIsIt FROM ConditionTbl WHERE ConditionID=" & CondID
        Dim InfoTbl As DataTable = OverClass.TempDataTable(SQLString)

        TextBox3.Text = InfoTbl.Rows(0).Item("WhatIsit").ToString
        WebBrowser1.Visible = False
        AllowNavigate = True
        WebBrowser1.Navigate("http://www.google.com/images?&q=" & TextBox2.Text)



        If OverClass.ReadOnlyUser = False Then
            TextBox2.ReadOnly = False
            TextBox3.ReadOnly = False
        End If

        CondID = 0

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        WebBrowser1.ScrollBarsEnabled = True
        WebBrowser1.Document.Window.ScrollTo(0, 300)
        WebBrowser1.Visible = True
        AllowNavigate = False
    End Sub

    Private Sub WebBrowser1_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles WebBrowser1.Navigating
        If AllowNavigate = False Then e.Cancel = True
    End Sub
End Class
