Imports System.IO	'DO NOT DELETE

Module modStudent

	Public Sub RunProject()
        'Project Header:
        'Name: Contacts2 
        'Purpose: To read the contents of one file into a new file
        'Author: Stephanie Spears
        'Date: 5/24/2016


        'Constants
        Const c_sInputFilePath As String = "Contacts.txt"
        Const c_sLineBreak As String = "---------------------------------------------------------------------" & NL
        Const c_sOutputFilePath As String = "ContactsReport.txt"

        'Variables
        Dim fileInput As StreamReader
        Dim fileOutput As StreamWriter
        Dim sName As String
        Dim sAddress As String
        Dim sPhone As String
        Dim sEmail As String


        'Begin Code
        SetTitle("Contacts")

        fileInput = File.OpenText(c_sInputFilePath)
        fileOutput = File.CreateText(c_sOutputFilePath)


        While fileInput.Peek <> -1
            sName = fileInput.ReadLine
            DisplayLine("Name: " & sName & NL)
            fileOutput.WriteLine("Name: " & sName & NL)

            sAddress = fileInput.ReadLine & NL & fileInput.ReadLine & ", " & fileInput.ReadLine & " " & fileInput.ReadLine
            DisplayLine("Address: " & NL & sAddress & NL)
            fileOutput.WriteLine("Address: " & NL & sAddress & NL)

            sPhone = fileInput.ReadLine
            DisplayLine("Phone: " & sPhone)
            fileOutput.WriteLine("Phone: " & sPhone)

            sEmail = fileInput.ReadLine
            DisplayLine("Email: " & sEmail)
            fileOutput.WriteLine("Email: " & sEmail)

            DisplayLine(c_sLineBreak)
            fileOutput.WriteLine(c_sLineBreak)

        End While


        fileInput.Close()
        fileOutput.Close()

    End Sub

End Module