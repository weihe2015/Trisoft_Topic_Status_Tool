Public Class IshObjs
    Public ISHAppObj As New Application20.Application20
    Public ISHDocObj As New DocumentObj20.DocumentObj20
    Public ISHDocObj25 As New DocumentObj25.DocumentObj25
    Public ISHPubObj As New Publication20.Publication20
    Public ISHPubOutObj25 As New PublicationOutput25.PublicationOutput25
    Public ISHBaselineObj As New Baseline25.BaseLine25
    Public ISHMetaObj As New MetaDataAssist20.MetaDataAssist20

    Sub New(ByVal Username As String, ByVal Password As String, ByVal ServerURL As String)
        ISHAppObj.Url = ServerURL + "/infosharews/Application20.asmx"
        ISHDocObj.Url = ServerURL + "/infosharews/DocumentObj20.asmx"
        ISHDocObj25.Url = ServerURL + "/infosharews/DocumentObj25.asmx"
        ISHPubObj.Url = ServerURL + "/infosharews/Publication20.asmx"
        ISHPubOutObj25.Url = ServerURL + "/infosharews/PublicationOutput25.asmx"
        ISHBaselineObj.Url = ServerURL + "/infosharews/Baseline25.asmx"
        ISHMetaObj.Url = ServerURL + "/infosharews/MetaDataAssist20.asmx"

    End Sub

End Class
