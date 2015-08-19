Imports System.IO
Imports System.Xml
Imports System.Text

Public Class clsCommonFuncs
    Public Function BuildMinPubMetadata() As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FISHISRELEASED"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("<ishfield name=""FISHMASTERREF"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FISHBASELINE"" level=""version""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function

    Public Function BuildRequestedMetadata(ByVal Resolution As String) As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        If Resolution <> "" Then
            requestedmeta.Append("<ishfield name=""FILLUSTRATOR"" level=""lng""/>")
            requestedmeta.Append("<ishfield name=""FRESOLUTION"" level=""lng""/>")
        End If

        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""MODIFIED-ON"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""CREATED-ON"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FCOMMENTS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FSTATUS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""DOC-LANGUAGE"" level=""lng""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta

        'requestedmeta.Append("<ishfield name=""FAUTHOR"" level=""lng""/>")
        'requestedmeta.Append("<ishfield name=""FNOTRANSLATIONMGMT"" level=""logical""/>")
        'requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")

        'useless:
        'requestedmeta.Append("<ishfield name=""FISHREVCOUNTER"" level=""version""/>")
        'FISHRELEASELABEL
        'requestedmeta.Append("<ishfield name=""FISHRELEASELABEL"" level=""version""/>")
        'requestedmeta.Append("<ishfield name=""FDESCRIPTION"" level=""logical""/>")
        'FDESCRIPTION
    End Function

    Public Function BuildMinEnMetadata(ByVal Resolution As String) As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""FNOTRANSLATIONMGMT"" level=""logical""/>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("<ishfield name=""FAUTHOR"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FSTATUS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""MODIFIED-ON"" level=""lng""/>")
        'requestedmeta.Append("<ishfield name=""FRESOLUTION"" level=""lng""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function

    Public Function BuildFullPubMetadata() As StringBuilder
        Dim requestedmeta As New StringBuilder
        requestedmeta.Append("<ishfields>")
        requestedmeta.Append("<ishfield name=""VERSION"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""FISHPUBSTATUS"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FISHEVENTID"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FISHISRELEASED"" level=""version""/>")
        requestedmeta.Append("<ishfield name=""DOC-LANGUAGE"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FISHOUTPUTFORMATREF"" level=""lng""/>")
        requestedmeta.Append("<ishfield name=""FTITLE"" level=""logical""/>")
        requestedmeta.Append("</ishfields>")
        Return requestedmeta
    End Function

End Class
