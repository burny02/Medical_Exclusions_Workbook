Imports System.Threading

Public Class Form1
    Private Dt As New DataTable()
    Private LatestThread As Integer = 0
    Private CondID As Long = 0
    Private AllowNavigate As Boolean = False
    Private QnAToggle As Boolean = False
    Private CodesToggle As Boolean = False
    Private GoogleToggle As Boolean = False

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

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Select Case e.TabPage.Text

            Case "Conditions"
                RefreshPage()

        End Select

    End Sub

    Private Sub RefreshPage()

        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True

        Dim SQLString As String = "SELECT ImageFileName, WhatIsIt, CondName FROM ConditionTbl WHERE ConditionID=" & CondID
        Dim SQLString2 As String = "SELECT Question, Answer, GivenBy, Reviewed FROM QnA " &
                                    "WHERE ConditionID=" & CondID & " And QType='18-45'"
        Dim SQLString3 As String = "SELECT Question, Answer, GivenBy, Reviewed FROM QnA " &
                                    "WHERE ConditionID=" & CondID & " And QType='46-64'"
        Dim SQLString4 As String = "SELECT Question, Answer, GivenBy, Reviewed FROM QnA " &
                                    "WHERE ConditionID=" & CondID & " And QType='Asthma'"

        Dim InfoTbl As DataTable = OverClass.TempDataTable(SQLString)
        TreeView1.Nodes.Clear()
        TreeView2.Nodes.Clear()
        TreeView3.Nodes.Clear()
        If InfoTbl.Rows.Count <= 0 Then Exit Sub

        Dim QnATbl As DataTable = OverClass.TempDataTable(SQLString2)
        Dim QnATbl2 As DataTable = OverClass.TempDataTable(SQLString3)
        Dim QnATbl3 As DataTable = OverClass.TempDataTable(SQLString4)
        If QnATbl.Rows.Count >= 1 Then SetQnA(QnATbl, TreeView1)
        If QnATbl2.Rows.Count >= 1 Then SetQnA(QnATbl2, TreeView2)
        If QnATbl3.Rows.Count >= 1 Then SetQnA(QnATbl3, TreeView3)

        TextBox2.Text = InfoTbl.Rows(0).Item("CondName")
        Dim ImageFileName As String = ImagePath & InfoTbl.Rows(0).Item("ImageFileName")

        If My.Computer.FileSystem.FileExists(ImageFileName) Then
            PictureBox6.BackgroundImage = Image.FromFile(ImageFileName)
        Else
            PictureBox6.BackgroundImage = Nothing
        End If
        PictureBox6.BackgroundImageLayout = ImageLayout.Zoom

        TextBox3.Text = InfoTbl.Rows(0).Item("WhatIsit").ToString

        QnAToggle = True
        ToggleByLabel(Panel1, Label1, QnAToggle)
        CodesToggle = True
        ToggleByLabel(Panel2, Label3, CodesToggle)

        If OverClass.ReadOnlyUser = False Then
            TextBox2.ReadOnly = False
            TextBox3.ReadOnly = False
        End If


    End Sub

    Private Sub Label1_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick
        ToggleByLabel(Panel1, Label1, QnAToggle)
    End Sub


    Private Sub ToggleByLabel(WhichPanel As Panel,
                              WhichLabel As Label,
                              ByRef ToggleBoolean As Boolean)


        Dim CollapseHeight As Long = 32
        Dim ExpandHeight As Long = 260
        Dim LabelString As String = WhichLabel.Text

        If ToggleBoolean = False Then

            WhichLabel.Text = "-" & Strings.Right(LabelString, Len(LabelString) - 1)
            ResizePanel(WhichPanel)


        ElseIf ToggleBoolean = True Then

            WhichLabel.Text = "+" & Strings.Right(LabelString, Len(LabelString) - 1)
            MoveControls(WhichPanel, -(WhichPanel.Height - CollapseHeight))
            WhichPanel.Height = CollapseHeight

        End If

        ToggleBoolean = Not ToggleBoolean

    End Sub

    Private Sub Label3_DoubleClick(sender As Object, e As EventArgs) Handles Label3.DoubleClick
        ToggleByLabel(Panel2, Label3, CodesToggle)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ListBox1.Visible = False

        If TextBox1.Text.ToString <> "" Then
            Dim trd = New Thread(AddressOf RefreshList)
            LatestThread = trd.ManagedThreadId
            trd.Start()
        End If
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        CondID = ListBox1.SelectedValue.ToString
        Call RefreshPage()
        TextBox1.Text = ""
        ListBox1.Visible = False
    End Sub



    Private Sub ResizePanel(WhichPanel As Panel)

        Dim OldHeight As Long = WhichPanel.Height
        WhichPanel.Height = 0
        Dim Padding As Long = 20
        Dim Lowest As Long = 0

        For Each ctl As Control In WhichPanel.Controls
            Dim ElementHeight As Long = ResizePanelElements(ctl)
            Lowest += ElementHeight
        Next

        Lowest = Lowest + Padding

        MoveControls(WhichPanel, Lowest - OldHeight)
        WhichPanel.Height = Lowest



    End Sub


    Private Sub ResizeTreeView(WhichTree As TreeView)

        Dim Y As Long = 0

        For Each n As TreeNode In WhichTree.Nodes
            Y += WhichTree.ItemHeight

            If n.IsExpanded Then Y += ResizeNode(n, WhichTree.ItemHeight)
        Next

        WhichTree.Height = Y

    End Sub

    Private Function ResizeNode(WhichNode As TreeNode, WhatHeight As Long)

        Dim Y As Long = 0

        For Each n As TreeNode In WhichNode.Nodes
            Y += WhatHeight

            If n.IsExpanded Then Y += ResizeNode(WhichNode, WhatHeight)

        Next

        Return Y

    End Function



    Private Function ResizePanelElements(WhichControl As Control)

        If TypeOf WhichControl Is PictureBox Then
            Return WhichControl.Height
            Exit Function
        End If

        Dim Bottom As Long = 0
        Dim Padding As Long = 10

        WhichControl.Height = 0

        If TypeOf WhichControl Is TabPage Then
            Dim kk As TabControl = WhichControl.Parent
            If kk.SelectedTab IsNot WhichControl Then
                Return 0
                Exit Function
            End If
        End If

        If TypeOf WhichControl Is TreeView Then
            ResizeTreeView(WhichControl)
            Return WhichControl.Height
            Exit Function
        Else

        End If

        For Each ctl As Control In WhichControl.Controls
            Bottom += ResizePanelElements(ctl)
        Next

        Bottom = Bottom + Padding

        WhichControl.Height = Bottom

        Return Bottom

    End Function

    Private Sub SetQnA(QnATbl As DataTable, WhichTree As TreeView)

        WhichTree.Nodes.Clear()
        Dim myImageList As New ImageList()
        myImageList.Images.Add(My.Resources.help)
        myImageList.Images.Add(My.Resources.lightbulb)
        WhichTree.ImageList = myImageList

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
                WhichTree.Nodes.Add(RootNode)

            Loop

            Dim SpaceNode As New TreeNode("")
            SpaceNode.ImageIndex = 3
            SpaceNode.SelectedImageIndex = 3
            WhichTree.Nodes.Add(SpaceNode)

        Next
    End Sub

    Private Sub MoveControls(WhichPanel As Panel, MoveAmount As Long)

        For Each ctl As Control In TabPage2.Controls

            If ctl.Top > WhichPanel.Bottom Then ctl.Top = ctl.Top + MoveAmount

        Next

    End Sub

    Private Sub Label4_DoubleClick(sender As Object, e As EventArgs) Handles Label4.DoubleClick
        ToggleByLabel(Panel3, Label4, GoogleToggle)
    End Sub

    Private Sub TreeView1_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterExpand
        ResizePanel(Panel1)
    End Sub

    Private Sub TreeView1_AfterCollapse(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterCollapse
        ResizePanel(Panel1)
    End Sub

    Private Sub TabControl2_Selected_1(sender As Object, e As TabControlEventArgs) Handles TabControl2.Selected
        ResizePanel(Panel1)
    End Sub

    Private Sub TreeView2_AfterCollapse(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterCollapse
        ResizePanel(Panel1)
    End Sub

    Private Sub TreeView2_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterExpand
        ResizePanel(Panel1)
    End Sub

    Private Sub TreeView3_AfterCollapse(sender As Object, e As TreeViewEventArgs) Handles TreeView3.AfterCollapse
        ResizePanel(Panel1)
    End Sub

    Private Sub TreeView3_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles TreeView3.AfterExpand
        ResizePanel(Panel1)
    End Sub
End Class
