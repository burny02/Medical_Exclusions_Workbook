Imports TemplateDB

Module Variables
    Public OverClass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Medical_Exclusions_Workbook\Backend.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const AuditTable As String = "[Audit]"
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const ActiveUsersTable As String = "[Users]"
    Private Contact As String = "Elisha Applebaum"
    Public Const SolutionName As String = "Medical Exclusions Workbook"
    Public LabForm As Form1
    Public PickCohort As Long
    Public AppID As Long
    Public Role As String
    Public WhichUser As String


    Public Function GetTheConnection() As String
        GetTheConnection = Connect2
    End Function


    Public Sub StartUp(WhichForm As Form)

        OverClass = New TemplateDB.OverClass
        OverClass.SetPrivate(UserTable,
                           UserField,
                           LockTable,
                           Contact,
                           Connect2,
                           AuditTable)

        OverClass.LockCheck()

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(WhichForm)

        WhichUser = OverClass.GetUserName


        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next


    End Sub

    Public Function NewOverclass() As OverClass

        Dim Whichclass As New OverClass
        Whichclass.SetPrivate(UserTable,
                           UserField,
                           LockTable,
                           Contact,
                           Connect2,
                           AuditTable)
        NewOverclass = Whichclass

    End Function

End Module
