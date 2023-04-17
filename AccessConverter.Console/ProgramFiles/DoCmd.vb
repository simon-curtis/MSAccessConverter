Public Class DoCmd
    Friend Shared Sub SetWarnings(WarningsOn As Boolean)
        'Throw New NotImplementedException()
    End Sub

    Friend Shared Sub OpenQuery(QueryName As String, Optional View As Access.AcView = Nothing, Optional DataMode As Access.AcOpenDataMode = Nothing)
        'Throw New NotImplementedException()
    End Sub

    Friend Shared Sub GoToRecord(Optional ObjectType As Access.AcDataObjectType = Nothing, Optional ObjectName As String = Nothing, Optional Record As Access.AcRecord = Nothing, Optional Offset As String = Nothing)
        If TypeOf Form.ActiveForm Is AccessForm Then
            Dim accessForm As AccessForm = CType(Form.ActiveForm, AccessForm)
            accessForm.NewRecord()
        End If
    End Sub

    Friend Shared Sub OpenForm(FormName As String, Optional View As Access.AcFormView = Nothing, Optional FilterName As String = Nothing, Optional WhereCondition As String = Nothing, Optional DataMode As Access.AcFormOpenDataMode = Nothing, Optional WindowMode As Access.AcWindowMode = Nothing, Optional OpenArgs As String = Nothing)
        Dim formType = Reflection.Assembly.GetExecutingAssembly().DefinedTypes.FirstOrDefault(Function(type) type.Name = FormName)
        Dim form As AccessForm = Activator.CreateInstance(formType)
        Select Case WindowMode
            Case Access.AcWindowMode.acDialog
                form.ShowDialog()
                Return

            Case Access.AcWindowMode.acWindowNormal
                form.Show()
                Return
        End Select
    End Sub

    Friend Shared Sub OutputTo(ObjectType As Access.AcOutputObjectType, Optional ObjectName As String = Nothing, Optional OutputFormat As String = Nothing, Optional OutputFile As String = Nothing, Optional AutoStart As Boolean = False, Optional TemplateFile As String = Nothing, Optional Encoding As String = Nothing, Optional OutputQuality As Access.AcExportQuality = Nothing)
        'Throw New NotImplementedException()
    End Sub

    Friend Shared Sub OpenReport(ReportName As String, Optional View As Access.AcView = Nothing, Optional FilterName As String = Nothing, Optional WhereCondition As String = Nothing, Optional WindowMode As Access.AcWindowMode = Nothing, Optional OpenArgs As String = Nothing)
        'Throw New NotImplementedException()
    End Sub

    Friend Shared Sub Quit(Optional Options As Access.AcQuitOption = Nothing)
        'Throw New NotImplementedException()
    End Sub

    Friend Shared Sub Close(Optional ObjectType As Access.AcObjectType = Nothing, Optional ObjectName As String = Nothing, Optional Save As Access.AcCloseSave = Nothing)
        'Throw New NotImplementedException()
    End Sub

    Friend Shared Sub RunCommand(Command As Access.AcCommand)
        'Throw New NotImplementedException()
    End Sub
End Class