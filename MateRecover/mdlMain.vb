Imports System.IO
Imports Microsoft.Win32

Module mdlMain

    Sub Main()

        ' Variabelen declareren
        Dim strFilename As String
        Dim strFullFilename As String = ""
        Dim chrKey As Char = ""
        Dim startInfo As New ProcessStartInfo
        Dim lstClubLocations As New List(Of String)
        Dim strBMProLocation As String
        Dim regKey As RegistryKey
        Dim tempKey As RegistryKey

        ' Voorafgaande operaties

        ' Header tonen
        Console.WriteLine()
        Console.WriteLine(" ╔═══════════════════════════════╗")
        Console.WriteLine(" ║ Bridgemate Recovery-programma ║")
        Console.WriteLine(" ║ v1.3.0 - 18 december 2018     ║")
        Console.WriteLine(" ║                               ║")
        Console.WriteLine(" ║ Martijn Verstraelen           ║")
        Console.WriteLine(" ║ martijn@hartenvier.be         ║")
        Console.WriteLine(" ╚═══════════════════════════════╝")
        Console.WriteLine()

        ' Bestandsnaam aanmaken
        strFilename = Today.ToString("MM-dd-yyyy") & ".bws"

        ' Mappen opzoeken van clubs in registry
        Try
            regKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
            regKey = regKey.OpenSubKey("SOFTWARE\Bridge\Clubs")
            For Each strSubkeyName As String In regKey.GetSubKeyNames()
                tempKey = regKey.OpenSubKey(strSubkeyName)
                lstClubLocations.Add(tempKey.GetValue("Destination").ToString())
            Next
        Catch ex As NullReferenceException
            Console.WriteLine(" Kan registersleutel met clubinformatie niet lezen.")
            Console.WriteLine(" Het programma moet worden afgesloten.")
            Console.WriteLine()
            Console.WriteLine(" Druk op een toets om het programma af te sluiten.")
            Console.ReadKey()
            Environment.Exit(0)
        End Try

        ' BMPro-locatie opzoeken
        Try
            regKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
            regKey = regKey.OpenSubKey("SOFTWARE\Bridge Systems\Bridgemate Control")
            strBMProLocation = regKey.GetValue("Install path").ToString()
        Catch ex As NullReferenceException
            Console.WriteLine(" Kan registersleutel van Bridgemate Control niet lezen.")
            Console.WriteLine(" Het programma moet worden afgesloten.")
            Console.WriteLine()
            Console.WriteLine(" Druk op een toets om het programma af te sluiten.")
            Console.ReadKey()
            Environment.Exit(0)
        End Try


        ' Controleer of Bridgemate Control Software afgesloten is...
        If Process.GetProcessesByName("BMPro").Count > 0 Then
            Console.WriteLine(" Bridgemate Control Software draait al!")
            Console.WriteLine(" Sluit eerst Bridgemate Control Software af en voer dan de Recovery opnieuw uit!")
            Console.WriteLine()
            Console.WriteLine(" Druk op een toets om het programma af te sluiten.")
            Console.ReadKey()
            Environment.Exit(0)
        Else
            ' Lopend tornooi opzoeken
            Console.WriteLine(" Bezig met opzoeken van lopend tornooi.")
            Console.WriteLine()
            Console.WriteLine(" - Zoek naar " & Chr(34) & strFilename & Chr(34))


            Console.Write("   .")
            For Each dir As String In lstClubLocations
                Console.Write(".")
                If File.Exists(dir & "\bridgemate\" & strFilename) Then
                    strFullFilename = dir & "\bridgemate\" & strFilename
                    Console.WriteLine()
                    Console.WriteLine(" - Gevonden in " & Chr(34) & dir.Split("\").Last & Chr(34))
                    Exit For
                End If
            Next
            Console.WriteLine()
            Console.WriteLine()
            If strFullFilename = "" Then
                Console.WriteLine(" Jammer; er is geen lopend tornooi teruggevonden.")
                Console.WriteLine()
                Console.WriteLine(" Druk op een toets om het programma af te sluiten.")
                Console.ReadKey()
                Environment.Exit(0)
            Else
                ' Start BMPro
                Console.WriteLine()
                Console.WriteLine(" Als alles goed gaat, dan zal dit programma automatisch afsluiten.")
                Console.WriteLine()
                Console.Write(" Is het spelen al begonnen? Zijn er m.a.w. al scores in de Bridgemates ingegeven? (J/N) ")
                Do Until chrKey = "J" Or chrKey = "N"
                    chrKey = Char.ToUpper(Console.ReadKey().KeyChar)
                Loop

                Console.WriteLine()
                Console.WriteLine()
                Console.WriteLine(" Starten van BMPro...")

                ' Zoeken naar BMPro
                If File.Exists(strBMProLocation & "\BMPro.exe") Then
                    startInfo.FileName = strBMProLocation & "\BMPro.exe"
                    startInfo.Arguments = "/f:[" & strFullFilename & "] /m"
                    If chrKey = "N" Then startInfo.Arguments &= " /s"
                    Process.Start(startInfo)
                End If

                ' Starten van bridge_mate.exe
                If Process.GetProcessesByName("bridge_mate").Count < 1 Then
                    Dim startInfo2 As New ProcessStartInfo
                    startInfo2.FileName = "c:\bridge\bin\bridge_mate.exe"
                    Process.Start(startInfo2)
                    Console.WriteLine(" Starten van MateFlow...")
                End If
            End If
        End If

        ' Afsluiten
    End Sub

End Module
