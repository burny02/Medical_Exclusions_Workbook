Imports System.Threading

Public Class Form1
    Private Dt As New DataTable()
    Private LatestThread As Integer = 0
    Private CondID As Long = 0
    Private AllowNavigate As Boolean = False
    Private QnAToggle As Boolean = False
    Private CodesToggle As Boolean = False

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

        Dim SQLString As String = "SELECT ImageFileName, WhatIsIt, CondName FROM ConditionTbl WHERE ConditionID=" & CondID
        Dim SQLString2 As String = "SELECT Question, Answer, GivenBy, Reviewed FROM QnA WHERE ConditionID=" & CondID
        Dim InfoTbl As DataTable = OverClass.TempDataTable(SQLString)
        Dim QnATbl As DataTable = OverClass.TempDataTable(SQLString2)

        TextBox2.Text = InfoTbl.Rows(0).Item("CondName")
        Dim ImageFileName As String = ImagePath & InfoTbl.Rows(0).Item("ImageFileName")

        If My.Computer.FileSystem.FileExists(ImageFileName) Then
            PictureBox6.BackgroundImage = Image.FromFile(ImageFileName)
        Else
            PictureBox6.BackgroundImage = Nothing
        End If
        PictureBox6.BackgroundImageLayout = ImageLayout.Zoom


        TextBox3.Text = InfoTbl.Rows(0).Item("WhatIsit").ToString

        TreeView1.Nodes.Clear()
        Dim myImageList As New ImageList()
        myImageList.Images.Add(My.Resources.help)
        myImageList.Images.Add(My.Resources.lightbulb)
        TreeView1.ImageList = myImageList



        For Each row As DataRow In QnATbl.Rows

            Dim AnswerString As String = ""
            Try
                AnswerString = row.Item("Answer")
            Catch ex As Exception
            End Try

            Dim RootNode As TreeNode = Nothing
            Dim RootString As String = row.Item("Question") & "..."
            Dim CutPlace As Long = 50
            Dim ChildString As String = row.Item("Answer") & " - " & row.Item("GivenBy") & "(" & row.Item("Reviewed") & ")"
            Dim SpaceLocation As Long = 0
            Dim i As Integer = 0

            Dim ChildNode() As TreeNode

            If AnswerString <> "" Then


                Do While ChildString <> ""

                    ReDim Preserve ChildNode(i)
                    SpaceLocation = 0
                    SpaceLocation = InStr(CutPlace, ChildString, " ", CompareMethod.Binary)
                    If SpaceLocation <> 0 Then CutPlace = SpaceLocation

                    Dim LeftString As String = Strings.Left(ChildString, CutPlace)

                    ChildNode(i) = New TreeNode(Trim(LeftString))
                    ChildString = Replace(ChildString, LeftString, "")
                    ChildNode(i).ImageIndex = 3
                    ChildNode(i).SelectedImageIndex = 3
                    i += 1

                Loop

                ReDim Preserve ChildNode(i - 1)
                ChildNode(0).ImageIndex = 1
                ChildNode(0).SelectedImageIndex = 1
            End If




            Do While RootString <> ""

                ReDim Preserve ChildNode(i - 1)
                SpaceLocation = 0
                SpaceLocation = InStr(CutPlace, RootString, " ", CompareMethod.Binary)
                If SpaceLocation <> 0 Then CutPlace = SpaceLocation

                Dim LeftString As String = Strings.Left(RootString, CutPlace)

                If Len(RootString) > CutPlace Or AnswerString = "" Then
                    RootNode = New TreeNode(Trim(LeftString))
                Else
                    RootNode = New TreeNode(Trim(LeftString), ChildNode)
                End If

                If RootString = row.Item("Question") & "..." Then
                    RootNode.ImageIndex = 0
                    RootNode.SelectedImageIndex = 0
                Else
                    RootNode.ImageIndex = 3
                    RootNode.SelectedImageIndex = 3
                End If
                RootString = Replace(RootString, LeftString, "")
                TreeView1.Nodes.Add(RootNode)

            Loop

            Dim SpaceNode As New TreeNode("")
            SpaceNode.ImageIndex = 3
            SpaceNode.SelectedImageIndex = 3
            TreeView1.Nodes.Add(SpaceNode)

        Next

        QnAToggle = True
        ToggleByLabel(Panel1, Label1, QnAToggle, "Questions and Answers", 260, 30)
        CodesToggle = True
        ToggleByLabel(Panel2, Label3, CodesToggle, "Codes", 260, 30)

        If OverClass.ReadOnlyUser = False Then
            TextBox2.ReadOnly = False
            TextBox3.ReadOnly = False
        End If

        CondID = 0

    End Sub

    Private Sub Label1_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick
        ToggleByLabel(Panel1, Label1, QnAToggle, "Questions and Answers", 260, 30)
    End Sub


    Private Sub ToggleByLabel(WhichPanel As Panel,
                              WhichLabel As Label,
                              ByRef ToggleBoolean As Boolean,
                              PanelLabel As String,
                              ExpandHeight As Long,
                              RetractHeight As Long)


        If ToggleBoolean = False Then

            Dim HeightChange As Long = WhichPanel.Height
            WhichLabel.Text = "- " & PanelLabel & "..."
            WhichPanel.Height = ExpandHeight
            HeightChange = WhichPanel.Height - HeightChange
            For Each ctl As Control In TabPage2.Controls
                If ctl.Top <= WhichPanel.Top + WhichPanel.Height And
                ctl.Bottom > WhichPanel.Top And
                ctl.Name <> WhichPanel.Name Then
                    ctl.Top = ctl.Top + HeightChange
                End If
            Next

        ElseIf ToggleBoolean = True Then

            Dim HeightChange As Long = WhichPanel.Height
            WhichLabel.Text = "+ " & PanelLabel & "..."
            WhichPanel.Height = RetractHeight
            HeightChange = WhichPanel.Height - HeightChange
            For Each ctl As Control In TabPage2.Controls
                If ctl.Top >= WhichPanel.Top + WhichPanel.Height And
                ctl.Bottom > WhichPanel.Top And
                ctl.Name <> WhichPanel.Name Then
                    ctl.Top = ctl.Top + HeightChange
                End If
            Next

        End If

        ToggleBoolean = Not ToggleBoolean

    End Sub

    Private Sub Label3_DoubleClick(sender As Object, e As EventArgs) Handles Label3.DoubleClick
        ToggleByLabel(Panel2, Label3, CodesToggle, "Codes", 260, 30)
    End Sub
End Class
