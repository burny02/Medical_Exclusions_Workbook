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

        Dim Blurb As String = "This system is intended as a reference tool for .. ... ." & vbNewLine &
                "... ... ..." & vbNewLine &
                "... ... ..."

        Call StartUp(Me)

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " &
                System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString &
                vbNewLine & vbNewLine & Blurb
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" &
                vbNewLine & vbNewLine & Blurb

        End Try

        Me.Text = SolutionName

        ListBox1.Parent = Me
        ListBox1.BringToFront()

        Dim clm1 As New DataColumn
        clm1.ColumnName = "CondName"
        Dt.Columns.Add(clm1)

        DataGridView3.DataSource = OverClass.TempDataTable("SELECT TOP 12 format(DateTime,'dd-MMM-yyyy')" &
                "& ':    ' & Table & '  -  ' & Action & ' (' & Condition & ')' As Show, ConditionID From History ORDER BY DateTime DESC")
        DataGridView3.Columns("ConditionID").Visible = False

        DataGridView2.DataSource = OverClass.TempDataTable("SELECT format(DateTime,'dd-MMM-yyyy') & ':    Document Added (' & DocName & ')' As Show, " &
                "URL From Docs ORDER BY DateTime DESC")
        DataGridView2.Columns("URL").Visible = False


    End Sub

    Private Sub RefreshList()

        Dim StringLength As Integer = 22
        Dim TempOverclass As TemplateDB.OverClass = NewOverclass()
        Dim SearchString As String = TextBox1.Text.ToString
        Dim TempDT As DataTable

        Dim UnionQry As String = "(SELECT ConditionID, CondName FROM ConditionTbl " &
                                 "UNION ALL Select ConditionID, Alt_Name FROM Synonyms)"

        Dim ExactMatchQry As String = "(SELECT TOP 10 iif(len(CondName)>" & StringLength &
        ", Left(CondName," & StringLength & ")& '...',CondName) AS Cond, ConditionID, 'A' As WhichOrder " &
        "FROM " & UnionQry & " WHERE CondName LIKE '" & SearchString & "%')"

        Dim WithinMatchQry As String = "(SELECT TOP 10 iif(len(CondName)>" & StringLength &
        ", Left(CondName," & StringLength & ") & '...',CondName)  AS Cond, ConditionID, 'B' As WhichOrder " &
        "FROM " & UnionQry & " WHERE CondName LIKE '%" & SearchString & "%')"

        Dim DescMatchQry As String = "(SELECT TOP 10 iif(len(CondName)>" & StringLength &
        ", Left(CondName," & StringLength & ") & '...',CondName)  AS Cond, ConditionID, 'C' As WhichOrder " &
        "FROM ConditionTbl WHERE WhatIsIt LIKE '%" & SearchString & "%')"

        Dim CombineQry As String = "(SELECT * FROM " & ExactMatchQry & " UNION ALL " & WithinMatchQry & " UNION ALL " & DescMatchQry & ")"

        Dim FullSQL = "SELECT TOP 10 Cond, ConditionID FROM " & CombineQry & "GROUP BY Cond, ConditionID ORDER BY first(WhichOrder) ASC, Cond ASC"

        If TextBox1.Text.ToString <> "" Then

            TempDT = TempOverclass.TempDataTable(FullSQL)

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

        ListBox1.Visible = False
        ListBox1.DisplayMember = "Cond"
        ListBox1.ValueMember = "ConditionID"
        ListBox1.DataSource = Dt
        If Dt.Rows.Count > 0 Then ListBox1.Visible = True

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Menu"

                DataGridView3.DataSource = OverClass.TempDataTable("Select TOP 12 format(DateTime,'dd-MMM-yyyy')" &
                "& ':    ' & Table & '  -  ' & Action & ' (' & Condition & ')' As Show, ConditionID From History ORDER BY DateTime DESC")
                DataGridView3.Columns("ConditionID").Visible = False

                DataGridView2.DataSource = OverClass.TempDataTable("SELECT format(DateTime,'dd-MMM-yyyy') & ':    Document Added (' & DocName & ')' As Show, " &
                "URL From Docs ORDER BY DateTime DESC")
                DataGridView2.Columns("URL").Visible = False

            Case "Conditions"
                RefreshPage()

        End Select

    End Sub

    Private Sub RefreshPage()

        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False

        Dim SQLArray(3) As String

        SQLArray(0) = "SELECT ImageFileName, WhatIsIt, CondName FROM ConditionTbl WHERE ConditionID=" & CondID
        SQLArray(1) = "Select Question, Answer, GivenBy, Reviewed FROM QnA " &
                      "WHERE ConditionID=" & CondID & " And QType1=True"
        SQLArray(2) = "Select Question, Answer, GivenBy, Reviewed FROM QnA " &
                      "WHERE ConditionID=" & CondID & " And QType2=True"
        SQLArray(3) = "Select Question, Answer, GivenBy, Reviewed FROM QnA " &
                      "WHERE ConditionID=" & CondID & " And QType3=True"

        Dim Alt_NameString As String = "Select Alt_Name FROM Synonyms WHERE ConditionID=" & CondID & " ORDER By Alt_Name"

        Dim TblArray() As DataTable = OverClass.MultiTempDataTable(SQLArray)
        Dim InfoTbl As DataTable = TblArray(0)
        TreeView1.Nodes.Clear()
        TreeView2.Nodes.Clear()
        TreeView3.Nodes.Clear()
        If InfoTbl.Rows.Count <= 0 Then Exit Sub

        If TblArray(1).Rows.Count >= 1 Then SetQnA(TblArray(1), TreeView1)
        If TblArray(2).Rows.Count >= 1 Then SetQnA(TblArray(2), TreeView2)
        If TblArray(3).Rows.Count >= 1 Then SetQnA(TblArray(3), TreeView3)

        TextBox4.Text = Replace(OverClass.CreateCSVString(Alt_NameString), ", ", ", ")
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
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
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

    Private Sub Label4_DoubleClick(sender As Object, e As EventArgs) Handles Label4.DoubleClick, PictureBox5.DoubleClick
        Dim URL As String = "www.Google.com/search?q=" & Replace(TextBox2.Text.ToString, " ", "+")
        Process.Start(URL)
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

    Private Sub SetUpAZIndex(WhichLabel As Label)
        OverClass.CreateDataSet("Select ConditionID, CondName FROM ConditionTbl WHERE CondName Like '" & WhichLabel.Text & "%'", BindingSource1, DataGridView1)
        DataGridView1.Columns("ConditionID").Visible = False
    End Sub

    Private Sub Label24_DoubleClick_1(sender As Object, e As EventArgs) Handles Label9.DoubleClick, Label8.DoubleClick, Label7.DoubleClick, Label6.DoubleClick, Label33.DoubleClick, Label32.DoubleClick, Label31.DoubleClick, Label30.DoubleClick, Label29.DoubleClick, Label28.DoubleClick, Label27.DoubleClick, Label26.DoubleClick, Label25.DoubleClick, Label24.DoubleClick, Label21.DoubleClick, Label20.DoubleClick, Label19.DoubleClick, Label18.DoubleClick, Label17.DoubleClick, Label16.DoubleClick, Label15.DoubleClick, Label14.DoubleClick, Label13.DoubleClick, Label12.DoubleClick, Label11.DoubleClick, Label10.DoubleClick
        SetUpAZIndex(sender)
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        CondID = sender.item("ConditionID", e.RowIndex).value
        TabControl1.SelectedIndex = 2
        TabControl1_Selecting(TabControl1, New TabControlCancelEventArgs(TabPage2, 0, False, TabControlAction.Selecting))
        Call RefreshPage()
    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        System.Diagnostics.Process.Start(sender.item("URL", e.RowIndex).value)
    End Sub

    Private Sub Label22_DoubleClick(sender As Object, e As EventArgs) Handles Label22.DoubleClick, PictureBox8.DoubleClick
        Dim URL = "https://en.wikipedia.org/w/index.php?search=" & Replace(TextBox2.Text.ToString, " ", "+")
        Process.Start(URL)
    End Sub

    Private Sub Label34_DoubleClick(sender As Object, e As EventArgs) Handles Label34.DoubleClick, PictureBox9.DoubleClick
        Dim URL = "https://vsearch.nlm.nih.gov/vivisimo/cgi-bin/query-meta?query=" & Replace(TextBox2.Text.ToString, " ", "+") & "&v%3Aproject=nlm-main-website"
        Process.Start(URL)
    End Sub
    Private Sub Label35_DoubleClick(sender As Object, e As EventArgs) Handles Label35.DoubleClick, PictureBox10.DoubleClick
        Dim URL = "http://www.webmd.com/search/search_results/default.aspx?query=" & Replace(TextBox2.Text.ToString, " ", "+")
        Process.Start(URL)
    End Sub

    Private Sub DataGridView3_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentDoubleClick
        CondID = sender.item("ConditionID", e.RowIndex).value
        TabControl1.SelectedIndex = 2
        TabControl1_Selecting(TabControl1, New TabControlCancelEventArgs(TabPage2, 0, False, TabControlAction.Selecting))
        Call RefreshPage()
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        CondID = ListBox1.SelectedValue
        TabControl1.SelectedIndex = 2
        TabControl1_Selecting(TabControl1, New TabControlCancelEventArgs(TabPage2, 0, False, TabControlAction.Selecting))
        Call RefreshPage()
        TextBox1.Text = ""
        ListBox1.Visible = False
    End Sub

    Private Sub ListBox1_Leave(sender As Object, e As EventArgs) Handles ListBox1.Leave, TextBox1.Leave
        If ListBox1 Is Me.ActiveControl Then Exit Sub
        If TextBox1 Is Me.ActiveControl Then Exit Sub
        If PictureBox1 Is Me.ActiveControl Then Exit Sub
        ListBox1.Visible = False
    End Sub
End Class
