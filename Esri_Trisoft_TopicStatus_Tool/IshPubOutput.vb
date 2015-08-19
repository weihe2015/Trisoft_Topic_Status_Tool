Imports System
Imports System.Xml
Imports System.IO
Imports System.Xml.XmlReader
Imports System.Text

Imports System.Collections
Imports Microsoft.VisualBasic.ControlChars
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports System.Convert
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel

Public Class IshPubOutput
    Inherits mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "IshPubOutput"
#End Region
#Region "Constructors"
    Public Sub New(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)
        oISHAPIObjs = New IshObjs(Username, Password, ServerURL)
        oISHAPIObjs.ISHAppObj.Login("InfoShareAuthor", Username, Password, Context)
    End Sub
#End Region

#Region "Methods"

    Public Function GetStatus(ByVal PubGUID As String, ByVal PubName As String, ByVal PubVer As String,
                                   ByVal Resolution As String, ByVal PubLang As String, ByVal PubType As String, ByVal outFolder As String,
                                   ByVal logFolder As String, ByRef excelRow As Integer, ByRef oSheet As Excel.Worksheet) As String

        Try
            Dim objWriter As New System.IO.StreamWriter(logFolder, True)
            objWriter.WriteLine("Start to get topic status for PubGUID: " & PubGUID)
            objWriter.Close()
            objWriter = Nothing
            Dim topics As HashSet(Of String) = GetTopicIDs(PubGUID, PubVer, PubLang, PubType)
            For Each topic As String In topics
                Dim perTopics As String() = topic.Split("#")
                Dim topicGUID As String = ""
                Dim topicVersion As String = ""
                Dim topicType As String = ""
                For k As Integer = 0 To perTopics.Length() - 1
                    If k = 0 Then
                        topicGUID = perTopics(k)
                    ElseIf k = 1 Then
                        topicVersion = perTopics(k)
                    ElseIf k = 2 Then
                        topicType = perTopics(k)
                    End If
                Next
                If topicGUID <> "" And topicVersion <> "" And topicType <> "" Then
                    oSheet.Cells(excelRow, 1).Value = PubGUID
                    oSheet.Cells(excelRow, 2).Value = PubName
                    oSheet.Cells(excelRow, 3).Value = PubVer
                    oSheet.Cells(excelRow, 4).Value = Resolution
                    oSheet.Cells(excelRow, 5).Value = topicGUID
                    oSheet.Cells(excelRow, 6).Value = topicType

                    oSheet.Cells(excelRow, 8).Value = topicVersion

                    Dim topicReso As String = ""
                    If topicType = "ISHIllustration" Then
                        topicReso = "Web"
                    End If
                    Dim topicEN As String = GetENCurrentState(topicGUID, topicVersion, topicReso)
                    If topicEN <> "" Then
                        Dim topicENs As String() = topicEN.Split("#")
                        Dim topicName As String = topicENs(0)
                        oSheet.Cells(excelRow, 7).Value = topicName
                        Dim topicAuthor As String = topicENs(1)
                        oSheet.Cells(excelRow, 12).Value = topicAuthor
                        Dim topicENStatus As String = topicENs(2)
                        Dim topicENL10N As String = topicENs(3)
                        oSheet.Cells(excelRow, 13).Value = topicENL10N
                        Dim topicENReleasedDate As String = "-"
                        Dim topicResult As String = ""

                        Dim topicStatus As String = "-"
                        Dim InTranslationDate As String = "-"
                        Dim TranslatedDate As String = "-"
                        Dim topicComment As String = "-"
                        Dim topicLanguage As String = PubLang
                        If topicENStatus = "Released" Then
                            topicENReleasedDate = topicENs(4)
                            topicResult = GetObjectInfo(topicGUID, topicVersion, PubLang, topicReso)
                            Dim topicResults As String() = topicResult.Split("#")
                            topicStatus = topicResults(0)
                            If topicStatus = "-" Then
                                topicStatus = topicENStatus
                                topicLanguage = "-"
                            End If
                            InTranslationDate = topicResults(1)
                            TranslatedDate = topicResults(2)
                            topicComment = topicResults(3)
                        Else
                            topicStatus = topicENStatus
                            topicLanguage = "-"
                        End If
                        oSheet.Cells(excelRow, 10).Value = topicLanguage
                        oSheet.Cells(excelRow, 9).Value = topicStatus
                        oSheet.Cells(excelRow, 11).Value = topicENReleasedDate
                        oSheet.Cells(excelRow, 14).Value = TranslatedDate
                        oSheet.Cells(excelRow, 15).Value = InTranslationDate
                        oSheet.Cells(excelRow, 16).Value = topicComment
                    Else
                        Dim objReporter As New System.IO.StreamWriter(logFolder, True)
                        objReporter.WriteLine(topicGUID & " : Problems with getting English topic status")
                        objReporter.Close()
                        objReporter = Nothing
                    End If

                End If
                excelRow = excelRow + 1
            Next


        Catch ex As Exception
            Dim objWriter As New System.IO.StreamWriter(logFolder, True)
            objWriter.WriteLine(PubGUID & " Problems with getting topic status")
            objWriter.Close()
            objWriter = Nothing
        End Try
        Return ""
    End Function

#End Region
End Class
