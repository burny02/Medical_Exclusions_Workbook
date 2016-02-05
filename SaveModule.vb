Module SaveModule
    Public Sub Saver(ctl As Object)

        Dim DisplayMessage As Boolean = True

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(OverClass.CurrentDataAdapter)

        Try
            OverClass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name


            'Case "DataGridView1"

            'Dim Person As String = "'" & WhichUser & "'"
            'Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

            'OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " &
            '                                                   "Set Result=@P1, Batch_No=@P2, Entered_Person=" & Person &
            '                                                   ", Entered_Date=" & ThisDate & "WHERE Result_ID=@P3")


            'With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
            '     .Add("@P1", OleDb.OleDbType.Double, 255, "Result")
            '.Add("@P2", OleDb.OleDbType.VarChar, 255, "Batch_No")
            '    .Add("@P3", OleDb.OleDbType.Double, 255, "Result_ID")
            'End With


        End Select


        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl, DisplayMessage)
        If DisplayMessage = False Then MsgBox("Table Updated")


    End Sub


End Module
