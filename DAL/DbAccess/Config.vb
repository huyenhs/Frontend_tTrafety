Imports System.Configuration
Imports System.Xml
Imports System.Web
Imports System.Data.SqlClient

Public Class Config
#Region "CONSTRUCTORS"
    Shared Sub New()
    End Sub
#End Region

#Region "GET CONFIG METHODS"
    ''' <summary>
    ''' Lấy về ConnectionString theo tên đã định nghĩa trong Web.Config
    ''' </summary>
    Public Function GetConnectionString(ByVal name As String) As String
        Return ConfigurationManager.ConnectionStrings(name).ConnectionString
    End Function

    ''' <summary>
    ''' Lấy về AppSetting theo tên đã định nghĩa trong Web.Config
    ''' </summary>
    Public Shared Function GetAppSetting(ByVal setting As String) As String
        Return ConfigurationManager.AppSettings(setting)
    End Function

    Public Shared ReadOnly Property WebsiteAppPath() As String
        Get
            Return ConfigurationManager.AppSettings("WebsiteAppPath")
        End Get
    End Property

    Public Shared ReadOnly Property AdminAppPath() As String
        Get
            Return ConfigurationManager.AppSettings("AdminAppPath")
        End Get
    End Property

    Public Shared Function GetAdminAppPath(ByVal setting As String) As String
        Return ConfigurationManager.AppSettings(setting)
    End Function

    Public Shared Function GetWebsiteAppPath(ByVal setting As String) As String
        Return ConfigurationManager.AppSettings(setting)
    End Function

    Public Shared Function GetSection(ByVal section As String) As Object
        Return ConfigurationManager.GetSection(section)
    End Function
#End Region
#Region "database"
    Public Function ServerProperty(ByVal name As String) As SqlConnectionStringBuilder
        Dim SSB As New SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings(name).ConnectionString)
        Return SSB
    End Function

    Public ReadOnly Property ServerName(ByVal SSB As SqlConnectionStringBuilder) As String
        Get
            Return SSB.DataSource
        End Get
    End Property

    Public ReadOnly Property DBName(ByVal SSB As SqlConnectionStringBuilder) As String
        Get
            Return SSB.InitialCatalog
        End Get
    End Property
#End Region
#Region "SET CONFIG METHODS"
    Public Shared Function AddAppSetting(ByVal xmlDoc As XmlDocument, ByVal Key As String, ByVal Value As String) As XmlDocument
        Dim root As XmlNode = xmlDoc.SelectSingleNode("//appSettings")
        If root IsNot Nothing Then
            Dim child As XmlNode = root.SelectSingleNode("//add[@key='" & Key & "']")
            If child IsNot Nothing Then
                DirectCast(child, XmlElement).SetAttribute("value", Value)
                Return xmlDoc
            End If

            Dim element As XmlElement = xmlDoc.CreateElement("add")
            element.SetAttribute("key", Key)
            element.SetAttribute("value", Value)
            root.AppendChild(element)
        End If
        Return xmlDoc
    End Function

    Public Shared Function LoadConfig() As XmlDocument
        Dim xDoc As New XmlDocument()
        xDoc.Load(HttpContext.Current.Server.MapPath("web.config"))
        Return xDoc
    End Function

    Public Shared Function LoadConfig(ByVal filePath As String) As XmlDocument
        Dim xDoc As New XmlDocument()
        xDoc.Load(HttpContext.Current.Server.MapPath(filePath))
        Return xDoc
    End Function

    Public Shared Sub SaveConfig(ByVal xmlDoc As XmlDocument)
        Using xWriter As New XmlTextWriter(HttpContext.Current.Server.MapPath("web.config"), Nothing)
            xWriter.Formatting = Formatting.Indented
            xmlDoc.WriteTo(xWriter)
            xWriter.Flush()
            xWriter.Close()
        End Using
    End Sub






#End Region
End Class
