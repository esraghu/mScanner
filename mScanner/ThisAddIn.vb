Imports System
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports System.IO
Imports System.IO.Compression

Public Class ThisAddIn

    Private finalMsg As String
    Private mScanTesting As Boolean

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        If MsgBox("Do you want to test the mScan functionality in this session?", vbYesNo + vbQuestion) = vbNo Then
            mScanTesting = False
        Else
            mScanTesting = True
        End If


    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_ItemSend(Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend
        If Not mScanTesting Then
            Exit Sub
        End If

        Dim emailExceptions As Boolean = False

        'Initialze final message before anything gets added.
        finalMsg = "Critical Error! Details below." & vbNewLine & vbNewLine

        If IsCardNumberPresent(Item.Subject) Or IsCardNumberPresent(Item.Body) Then
            finalMsg = finalMsg & vbNewLine &
                            "The subject or the body of the email seems to contain Card Number(s)! " &
                              "Consider masking Card number(s)"
            emailExceptions = True
        End If

        If ExtIdPresent(Item) Then
            'Check if restricted attachments present in the email
            If CheckForRestrictedAttachments() Then
                'Restricted attachments present
            End If
            emailExceptions = True
        End If

        finalMsg = finalMsg & vbNewLine &
                    "Do you really want to send this email?"

        If emailExceptions Then
            If MsgBox(finalMsg, vbYesNo + vbCritical, Item.Subject) = vbNo Then
                Cancel = True
            Else
                Cancel = False
            End If
        End If

    End Sub

    Function IsCardNumberPresent(ByVal Message As String) As Boolean

        IsCardNumberPresent = False

        Dim messageLineArray() As String
        Dim messageWordArray() As String

        messageLineArray = Split(Message, vbNewLine)
        For Each Line In messageLineArray
            messageWordArray = Split(Line)
            For Each word In messageWordArray
                If IsNumeric(word) Then
                    If Len(word) > 12 And Len(word) < 20 Then
                        Select Case Left(word, 1)
                            Case "3", "4", "5", "6", "9"
                                IsCardNumberPresent = True
                        End Select
                    End If
                End If
            Next
        Next

    End Function

    Function ExtIdPresent(ByVal Item As Object) As Boolean
        Dim recips As Outlook.Recipients
        Dim recip As Outlook.Recipient
        Dim pa As Outlook.PropertyAccessor
        Dim strMsg As String = ""
        ExtIdPresent = False

        Const PR_STMP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

        'Get all the recipient names in To, CC & BCC fields and check if the domain name isn't example ID.
        recips = Item.Recipients
        For Each recip In recips
            pa = recip.PropertyAccessor
            If InStr(LCase(pa.GetProperty(PR_STMP_ADDRESS)), "@example.com") = 0 Then
                strMsg = strMsg & "  " & pa.GetProperty(PR_STMP_ADDRESS) & vbNewLine
            End If
        Next

        If strMsg <> "" Then
            finalMsg = finalMsg & vbNewLine &
                        "This email is being sent to the following external email addresses:" & vbNewLine &
                        strMsg & "re-check the email addresses."
            ExtIdPresent = True
        End If
    End Function

    Function CheckForRestrictedAttachments() As Boolean
        CheckForRestrictedAttachments = False

        Dim myInspector As Outlook.Inspector = Application.ActiveInspector
        If Not TypeName(myInspector) = "Nothing" Then
            If TypeName(myInspector.CurrentItem) = "MailItem" Then
                Dim myItem As Outlook.MailItem = myInspector.CurrentItem
                Dim myAttachments As Outlook.Attachments = myItem.Attachments

                'Let us now check for the extensions of the files that have been attached
                For i = 1 To myAttachments.Count
                    'Find the position of the "." in the display name in the reverse order
                    Dim extensionPosition = InStrRev(myAttachments.Item(i).DisplayName, ".")
                    'Get the length of the attachment file
                    Dim attachmentFilenameLength = Len(myAttachments.Item(i).DisplayName)
                    'Obtain the extension type by reading to the right of "." in the display name of the attachment
                    Dim extensionType = Right(myAttachments.Item(i).DisplayName,
                                              attachmentFilenameLength - extensionPosition)
                    Select Case extensionType
                        Case "pdf", "docx", "doc", "xls", "xlsx"
                            finalMsg = finalMsg & vbNewLine &
                                        "Attachment may contain classified info, please review: " &
                                        myAttachments.Item(i).DisplayName
                            CheckForRestrictedAttachments = True
                            'OpenAttachment(myAttachments.Item(i).FileName)
                            If extensionType = "docx" Then
                                Dim attachmentFileName As String =
                                    System.IO.Path.GetFullPath(myAttachments.Item(i).FileName)
                                If ScanWordFile(attachmentFileName) Then
                                    finalMsg = finalMsg & vbNewLine &
                                                "This document " & myAttachments.Item(i).FileName &
                                                " contains restricted information and should be encrypted, if you like to send this."
                                    'myAttachments.Item(i).Delete()
                                End If
                            End If
                            If extensionType = "pdf" Then
                                Dim attachmentFileName As String =
                                    System.IO.Path.GetFullPath(myAttachments.Item(i).FileName)
                                If ScanFromPDF(attachmentFileName) Then
                                    finalMsg = finalMsg & vbNewLine &
                                                "This document " & myAttachments.Item(i).FileName &
                                                " contains restricted information and hence it will be removed."
                                    myAttachments.Item(i).Delete()
                                End If
                            End If

                        Case "java", "c", "h", "exe", "vb"
                            finalMsg = finalMsg & vbNewLine & "Restricted attachment present " &
                                        myAttachments.Item(i).DisplayName &
                                        ". This will be removed."
                            'For now the code only works if the attachment is the last one in the list
                            'Need some more analysis to remove the attachments in a more graceful manner
                            'Let's remove this attachment and decrease the counter by 1
                            myAttachments.Item(i).Delete()

                        Case "zip"
                            Dim sc As New Shell32.Shell()
                            Dim tempLoc As String = "C:\Users\Home\AppData\Local\Temp\" _
                                                    & myAttachments.Item(i).DisplayName _
                                                    & ".dir"

                            System.IO.Directory.CreateDirectory(tempLoc)
                            Dim extractDir As Shell32.Folder = sc.NameSpace(tempLoc)
                            Dim zipFile As Shell32.Folder = sc.NameSpace(System.IO.Path.GetFullPath(myAttachments.Item(i).FileName))
                            extractDir.CopyHere(zipFile.Items)

                    End Select
                Next i
            Else
                MsgBox("The Item doesn't seem to be MailItem")
            End If
        End If

    End Function

    Function OpenAttachment(ByVal FileName As String) As Boolean

        OpenAttachment = True
        If MsgBox("Do you want to review the contents of the attachment " & FileName, vbYesNo + vbQuestion) = vbYes Then
            Process.Start(FileName)
            'Need to add methods to stop the processes by referencing the process to an object and then call p.Close()
            'For now I'll rely on the user to close the opened attachment.
        End If

    End Function


    Private Function ScanWordFile(ByRef FileName As String) As Boolean
        ScanWordFile = False
        Dim text As New StringBuilder()
        Dim word As New Microsoft.Office.Interop.Word.Application()
        Dim miss As Object = System.Reflection.Missing.Value
        Dim path As Object = FileName
        Dim [readOnly] As Object = True
        Dim docs As Microsoft.Office.Interop.Word.Document
        If System.IO.File.Exists(path) Then
            docs = word.Documents.Open(path, miss, [readOnly], miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss)
            For i As Integer = 0 To docs.Paragraphs.Count - 1
                text.Append(" " & vbCr & vbLf & " " + docs.Paragraphs(i + 1).Range.Text.ToString())
            Next
            If IsCardNumberPresent(text.ToString()) Then
                ScanWordFile = True
            End If
        End If



        'MsgBox(text)

        'Return text.ToString()
    End Function

    Private Function ScanFromPDF(ByRef FileName As String) As Boolean
        ScanFromPDF = False
        Dim text As New StringBuilder()
        Using reader As New PdfReader(FileName)
            For i As Integer = 1 To reader.NumberOfPages
                text.Append(PdfTextExtractor.GetTextFromPage(reader, i))
            Next
            If IsCardNumberPresent(text.ToString()) Then
                ScanFromPDF = True
            End If
        End Using

    End Function




End Class


