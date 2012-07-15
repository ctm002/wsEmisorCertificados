Public Class Logger

    Public Shared Sub Write(ByVal msg As String)
        '  Try
        Dim fic As String = System.Configuration.ConfigurationSettings.AppSettings.Get("archivo_log")

        Dim sw As New System.IO.StreamWriter(fic, True)
        sw.WriteLine(msg & " - " & Now.ToString())
        ' sw.WriteLine("------")
        sw.Close()

        'Catch ex As Exception
        '  MsgBox(ex.Message)
        ' End Try

    End Sub

End Class
