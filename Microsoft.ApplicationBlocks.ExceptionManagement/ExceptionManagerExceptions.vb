
   ' Summary description for ExceptionManagerExceptions.
    <Serializable()> Public Class CustomPublisherException
      Inherits BaseApplicationException

#Region "Constructors"

      ' Constructor with no params.
      Public Sub New()
         MyBase.New()
      End Sub 'New

      ' Constructor allowing the Message property to be set.
      ' Parameters:
      ' -message - String setting the message of the exception. 
      Public Sub New(ByVal message As String)
         MyBase.New(message)
      End Sub 'New

      ' Constructor allowing the Message and InnerException property to be set.
      ' Parameters:
      ' -message - String setting the message of the exception. 
      ' -inner - Sets a reference to the InnerException. 
      Public Sub New(ByVal message As String, ByVal inner As Exception)
         MyBase.New(message, inner)
      End Sub 'New

      ' Constructor allowing the message, assembly name, type name, and publisher format to be set.
      ' Parameters:
      ' -message - String setting the message of the exception. 
      ' -assemblyName - String setting the assembly name of the exception. 
      ' -typeName - String setting the type name of the exception. 
      ' -publisherFormat - String setting the publisher format of the exception. 
      Public Sub New(ByVal message As String, ByVal assemblyName As String, ByVal typeName As String, ByVal publisherFormat As PublisherFormat)
        MyBase.New(message)
        Me.m_assemblyName = assemblyName
        Me.m_typeName = typeName
        Me.m_publisherFormat = publisherFormat
      End Sub 'New

      ' Constructor allowing the Message and InnerException property to be set.
      ' Parameters:
      ' -message - String setting the message of the exception. 
      ' -assemblyName - String setting the assembly name of the exception. 
      ' -typeName - String setting the type name of the exception. 
      ' -publisherFormat - String setting the publisher format of the exception. 
      ' -inner - Sets a reference to the InnerException. 
      Public Sub New(ByVal message As String, ByVal assemblyName As String, ByVal typeName As String, ByVal publisherFormat As PublisherFormat, ByVal inner As Exception)
         MyBase.New(message, inner)
         Me.m_assemblyName = assemblyName
         Me.m_typeName = typeName
         Me.m_publisherFormat = publisherFormat
      End Sub 'New

      ' Constructor used for deserialization of the exception class.
      ' Parameters:
      ' -info - Represents the SerializationInfo of the exception. 
      ' -context - Represents the context information of the exception. 
      Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
         MyBase.New(info, context)
         m_assemblyName = info.GetString("assemblyName")
         m_typeName = info.GetString("typeName")
         m_publisherFormat = CType(info.GetValue("publisherFormat", GetType(PublisherFormat)), PublisherFormat)
      End Sub 'New

#End Region

      ' Member variable declarations
    Private m_assemblyName As String
    Private m_typeName As String
    Private m_publisherFormat As PublisherFormat

      ' Override the GetObjectData method to serialize custom values.
      ' Parameters:
      ' -info - Represents the SerializationInfo of the exception. 
      ' -context - Represents the context information of the exception. 
      <SecurityPermission(SecurityAction.Demand, SerializationFormatter:=True)> _
      Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        info.AddValue("assemblyName", m_assemblyName, GetType(String))
        info.AddValue("typeName", m_typeName, GetType(String))
        info.AddValue("publisherFormat", m_publisherFormat, GetType(PublisherFormat))
         MyBase.GetObjectData(info, context)
      End Sub 'GetObjectData

#Region "Public Properties"
      ' The exception format configured for the publisher that threw an exception.
      Public ReadOnly Property PublisherFormat() As PublisherFormat
         Get
            Return m_publisherFormat
         End Get
      End Property

      ' The Assembly name of the publisher that threw an exception.
      Public ReadOnly Property PublisherAssemblyName() As String
         Get
            Return m_assemblyName
         End Get
      End Property

      ' The Type name of the publisher that threw an exception.
      Public ReadOnly Property PublisherTypeName() As String
         Get
            Return m_typeName
         End Get
      End Property

#End Region
   End Class 'CustomPublisherException 
