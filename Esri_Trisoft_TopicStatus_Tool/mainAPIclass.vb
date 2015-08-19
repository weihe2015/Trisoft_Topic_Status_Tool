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


Public Class mainAPIclass
#Region "Private Members"
    Private ReadOnly strModuleName As String = "CustomCMSFuncs"
#End Region
#Region "Constructors"
    Sub New()
        'Can be overridden by subs. Should report an error if ever called.
        'Do nothing.
    End Sub

#End Region
#Region "Properties"
    Public Context As String = ""
    Public oISHAPIObjs As IshObjs
    Public oCommonFuncs As New clsCommonFuncs
#End Region

#Region "Topic ID Funcs"
    Public Function GetTopicIDs(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal PubType As String) As HashSet(Of String)
        Dim TopicGUID As HashSet(Of String) = New HashSet(Of String)()
        'Get the pub's baseline ID from the pub object
        Dim outObjectList As String = ""
        Dim GUIDs(0) As String
        GUIDs(0) = GUID
        Dim Languages(0) As String
        Languages(0) = Language
        Dim Resolutions() As String
        Resolutions = {"High", "Low", " "}
        oISHAPIObjs.ISHMetaObj.GetLOVValues(Context, "DRESOLUTION", Resolutions)

        'Get the existing publication content at the specified version.
        oISHAPIObjs.ISHPubOutObj25.RetrieveVersionMetadata(Context, _
                        GUIDs, _
                        Version, _
                        oCommonFuncs.BuildMinPubMetadata.ToString, _
                        outObjectList)

        Dim VerDoc As New XmlDocument
        VerDoc.LoadXml(outObjectList)
        If VerDoc Is Nothing Or VerDoc.HasChildNodes = False Then
            'modErrorHandler.Errors.PrintMessage(3, "Unable to find publication for specified GUID: " + GUID, strModuleName + "-GetBaselineObjects")
            Return Nothing
        End If
        'Get the Baseline ID from the publication:
        Dim baselineID As String = ""
        Dim baselinename As String
        Dim ishfields As XmlNode = VerDoc.SelectSingleNode("//ishfields")
        baselinename = ishfields.SelectSingleNode("ishfield[@name='FISHBASELINE']").InnerText
        'Pull the baseline info
        oISHAPIObjs.ISHBaselineObj.GetBaselineId(Context, baselinename, baselineID)
        oISHAPIObjs.ISHBaselineObj.GetReport(Context, baselineID, Nothing, Languages, Languages, Languages, Resolutions, outObjectList)
        'Load the resulting baseline string as an xml document
        Dim baselineDoc As New XmlDocument
        baselineDoc.LoadXml(outObjectList)
        'for each object referenced, store the various info in an object and then store them in the hashSet.
        For Each baselineObject As XmlNode In baselineDoc.SelectNodes("/baseline/objects/object")
            Dim ishtype As String = baselineObject.Attributes.GetNamedItem("type").Value
            If ishtype = "ISHModule" Or ishtype = "ISHLibrary" Or ishtype = "ISHIllustration" Or ishtype = "ISHMasterDoc" Then
                Dim refGuid As String = baselineObject.Attributes.GetNamedItem("ref").Value
                Dim refver As String = baselineObject.Attributes.GetNamedItem("versionnumber").Value
                Dim refresolution As String = ""
                If PubType = "Both Image and Topic" Then
                    If ishtype = "ISHModule" Or ishtype = "ISHLibrary" Or ishtype = "ISHIllustration" Or ishtype = "ISHMasterDoc" Then
                        TopicGUID.Add(refGuid & "#" & refver & "#" & ishtype)
                    End If
                ElseIf PubType = "Image Only" Then
                    If ishtype = "ISHIllustration" Or ishtype = "ISHMasterDoc" Then
                        TopicGUID.Add(refGuid & "#" & refver & "#" & ishtype)
                    End If
                Else
                    If ishtype = "ISHModule" Or ishtype = "ISHLibrary" Or ishtype = "ISHMasterDoc" Then
                        TopicGUID.Add(refGuid & "#" & refver & "#" & ishtype)
                    End If
                End If

            End If
        Next
        Return TopicGUID
    End Function

#End Region

#Region "Doc Functions"

    Public Function GetObjectInfo(ByVal GUID As String, ByVal Version As String, ByVal Language As String, ByVal Resolution As String) As String
        Dim requestedmeta As StringBuilder = oCommonFuncs.BuildRequestedMetadata(Resolution)
        Dim XMLString As String = ""
        Dim Result As String = ""
        Dim MyDoc As New XmlDocument

        Try
            'oISHAPIObjs.ISHReportObj.GetReferencedDocObj(Context, GUID, Version, Language, False, requestedmeta.ToString, XMLString)
            oISHAPIObjs.ISHDocObj25.GetMetaData(Context, GUID, Version, Language, Resolution, requestedmeta.ToString, XMLString)
            MyDoc.LoadXml(XMLString)
            Dim TopicStatus As String = MyDoc.SelectSingleNode("//ishfield[@name=""FSTATUS""]").InnerText
            Dim Comment As String = MyDoc.SelectSingleNode("//ishfield[@name=""FCOMMENTS""]").InnerText
            Dim InTransDate As String = ""
            Dim TransDate As String = ""
            If TopicStatus = "Translated" Then
                Dim TransDates As String() = MyDoc.SelectSingleNode("//ishfield[@name=""MODIFIED-ON""]").InnerText.Split(" ")
                TransDate = TransDates(0)
                Dim InTransDates As String() = MyDoc.SelectSingleNode("//ishfield[@name=""CREATED-ON""]").InnerText.Split(" ")
                InTransDate = InTransDates(0)
            ElseIf TopicStatus = "In translation" Then
                Dim InTransDates As String() = MyDoc.SelectSingleNode("//ishfield[@name=""MODIFIED-ON""]").InnerText.Split(" ")
                InTransDate = InTransDates(0)
            End If
            Result = TopicStatus & "#" & InTransDate & "#" & TransDate & "#" & Comment
            Return Result
        Catch ex As Exception
            'MsgBox(ex.Message)
            'modErrorHandler.Errors.PrintMessage(3, "Failed to retrieve XML from CMS server: " + ex.Message, strModuleName)
            'Does not have the object in this language
            Result = "-" & "#" & "-" & "#" & "-" & "#" & "-"
            Return Result
        End Try
    End Function


    Public Function GetENCurrentState(ByVal GUID As String, ByVal Version As String, ByVal Resolution As String) As String
        Dim requestmeta As StringBuilder = oCommonFuncs.BuildMinEnMetadata(Resolution)
        Dim state As String = ""
        Dim L10N As String = ""
        Dim ENL10N As String = ""
        Dim OutXML As String = ""
        Dim title As String = ""
        Dim author As String = ""
        Dim result As String = ""
        Try
            '' Declare variable for the Application service
            oISHAPIObjs.ISHDocObj.GetMetaData(Context, GUID, Version, "en", Resolution, requestmeta.ToString, OutXML)
            Dim docInfo As New XmlDocument
            docInfo.LoadXml(OutXML)
            state = docInfo.SelectSingleNode("//ishfield[@name=""FSTATUS""]").InnerText
            L10N = docInfo.SelectSingleNode("//ishfield[@name=""FNOTRANSLATIONMGMT""]").InnerText
            title = docInfo.SelectSingleNode("//ishfield[@name=""FTITLE""]").InnerText
            author = docInfo.SelectSingleNode("//ishfield[@name=""FAUTHOR""]").InnerText
            If L10N = "No" Then
                ENL10N = "Yes"
            ElseIf L10N = "Yes" Then
                ENL10N = "No"
            End If
            Dim ENDates As String() = docInfo.SelectSingleNode("//ishfield[@name=""MODIFIED-ON""]").InnerText.Split(" ")
            Dim ENDate As String = ENDates(0)
            result = title & "#" & author & "#" & state & "#" & ENL10N & "#" & ENDate
            Return result
        Catch ex As Exception
            'MsgBox("Error getting current state for " + GUID + "" + Version + "" + Language + ". Message: " & ex.Message.ToString & strModuleName & "-GetCurrentState")
            'modErrorHandler.Errors.PrintMessage(2, "Error getting current state for " + GUID + "" + Version + "" + Language + ". Message: " + ex.Message.ToString, strModuleName + "-GetCurrentState")
            Return result
        End Try
    End Function
#End Region

End Class
