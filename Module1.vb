Module Module1

    Sub Main()

        Console.WriteLine("""oft2mailto"" by KeigoYAMAZAKI (@ymzkei5), 2024.09.27-")
        Console.WriteLine("Converter for Outlook (classic) user template files (*.oft) to 'mailto:' links")
        Console.WriteLine("")

        Dim dir = IO.Path.Combine(Environ("APPDATA"), "Microsoft\Templates")
        If IO.Directory.Exists(dir) = False Then
            OutputError("Templates directory does not exist.", True) : End
        End If

        Dim files = IO.Directory.GetFiles(dir, "*.oft", IO.SearchOption.AllDirectories)
        If files.Length <= 0 Then
            OutputError("User template file does not exist.", True) : End
        End If

        Dim app As New Microsoft.Office.Interop.Outlook.Application
        Dim result As New System.IO.StringWriter

        Dim count = 0
        For Each file In files
            Try
                Console.WriteLine("Processing... " + file)

                Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = CType(app.CreateItemFromTemplate(file), Microsoft.Office.Interop.Outlook.MailItem)

                Dim subject As String = mailItem.Subject
                Dim body As String = mailItem.Body
                Dim rTo As String = String.Join(",", mailItem.Recipients.Cast(Of Microsoft.Office.Interop.Outlook.Recipient)().Where(Function(r) r.Type = 1).Select(Function(r) r.Address)).Replace("""", "%22")
                Dim rCc As String = String.Join(",", mailItem.Recipients.Cast(Of Microsoft.Office.Interop.Outlook.Recipient)().Where(Function(r) r.Type = 2).Select(Function(r) r.Address)).Replace("""", "%22")
                Dim rBcc As String = String.Join(",", mailItem.Recipients.Cast(Of Microsoft.Office.Interop.Outlook.Recipient)().Where(Function(r) r.Type = 3).Select(Function(r) r.Address)).Replace("""", "%22")
                If rCc <> "" Then rCc = "&Cc=" + rCc
                If rBcc <> "" Then rBcc = "&Bcc=" + rBcc

                Dim tname = file.Replace(dir, "").Replace("\", "/").Replace("<", "&lt;").Trim("/").Replace(".oft", "")
                result.WriteLine($"<a href=""mailto:{rTo}?subject={Uri.EscapeDataString(subject)}{rCc}{rBcc}&body={Uri.EscapeDataString(body)}"">{tname}</a><p>")

                count += 1
            Catch ex As Exception
                OutputError("[[ERROR]] " + ex.Message, False)
            End Try
        Next

        If count > 0 Then
            Try
                IO.File.WriteAllText("result.html", "<!DOCTYPE html><html lang=""ja""><head><meta charset=""utf-8""><title>oft2mailto by @ymzkei5</title></head><body><h3>User Template</h3>" + vbCrLf + result.ToString + "</body></html>")
            Catch ex As Exception
                OutputError("[[ERROR]] " + ex.Message, True) : End
            End Try
        End If

        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("The result has been saved to 'result.html'.")
        Console.ResetColor()

        Console.WriteLine("")
        Console.Write("Press Enter to open 'result.html'.")
        Console.ReadLine()

        System.Diagnostics.Process.Start("result.html")

    End Sub

    Sub OutputError(message As String, exit1 As Boolean)

        Console.ForegroundColor = ConsoleColor.Red
        Console.Error.WriteLine(message)
        Console.ResetColor()

        If exit1 Then
            Console.WriteLine("")

            Console.Write("Press Enter to exit...")
            Console.ReadLine()

            Environment.Exit(1)
        End If

    End Sub

End Module
