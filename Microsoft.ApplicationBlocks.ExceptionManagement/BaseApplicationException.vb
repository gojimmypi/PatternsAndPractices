
   ' Base Application Exception Class. You can use this as the base exception object from
   ' which to derive your applications exception hierarchy.
    <Serializable()> Public Class BaseApplicationException
    Inherits ApplicationException

#Region "Constructors"
      ' Constructor with no params.
      Public Sub New()
        MyBase.New()
        InitializeEnvironmentInformation()
      End Sub 'New

      ' Constructor allowing the Message property to be set.
      ' Parameters:
      ' -message - String setting the message of the exception. 
      Public Sub New(ByVal message As String)
         MyBase.New(message)
         InitializeEnvironmentInformation()
      End Sub 'New

      ' Constructor allowing the Message and InnerException property to be set.
      ' Parameters:
      ' -message - String setting the message of the exception. 
      ' -inner - Sets a reference to the InnerException. 
      Public Sub New(ByVal message As String, ByVal inner As Exception)
         MyBase.New(message, inner)
         InitializeEnvironmentInformation()
      End Sub 'New

      ' Constructor used for deserialization of the exception class.
      ' Parameters:
      ' -info - Represents the SerializationInfo of the exception. 
      ' -context - Represents the context information of the exception. 
      Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
         MyBase.New(info, context)
         m_machineName = info.GetString("machineName")
         m_createdDateTime = info.GetDateTime("createdDateTime")
         m_appDomainName = info.GetString("appDomainName")
         m_threadIdentity = info.GetString("threadIdentity")
         m_windowsIdentity = info.GetString("windowsIdentity")
         m_additionalInformation = CType(info.GetValue("additionalInformation", GetType(NameValueCollection)), NameValueCollection)
      End Sub 'New

#End Region

#Region "Declare Member Variables"

      ' Member variable declarations
      Private m_machineName As String
      Private m_createdDateTime As DateTime = DateTime.Now
      Private m_appDomainName As String
      Private m_threadIdentity As String
      Private m_windowsIdentity As String

      Private Shared m_resourceManager As New ResourceManager(GetType(ExceptionManager).Namespace + ".ExceptionManagerText", [Assembly].GetAssembly(GetType(ExceptionManager)))
      ' Collection provided to store any extra information associated with the exception.
      Private m_additionalInformation As New NameValueCollection()

#End Region

      ' Override the GetObjectData method to serialize custom values.
      ' Parameters:
      ' -info - Represents the SerializationInfo of the exception. 
      ' -context - Represents the context information of the exception. 
      <SecurityPermission(SecurityAction.Demand, SerializationFormatter:=True)> Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
         info.AddValue("machineName", m_machineName, GetType(String))
         info.AddValue("createdDateTime", m_createdDateTime)
         info.AddValue("appDomainName", m_appDomainName, GetType(String))
         info.AddValue("threadIdentity", m_threadIdentity, GetType(String))
         info.AddValue("windowsIdentity", m_windowsIdentity, GetType(String))
         info.AddValue("additionalInformation", m_additionalInformation, GetType(NameValueCollection))
         MyBase.GetObjectData(info, context)
      End Sub 'GetObjectData

#Region "Public Properties"
      ' Machine name where the exception occurred.
      Public ReadOnly Property MachineName() As String
         Get
            Return m_machineName
         End Get
      End Property

      ' Date and Time the exception was created.
      Public ReadOnly Property CreatedDateTime() As DateTime
         Get
            Return m_createdDateTime
         End Get
      End Property

      ' AppDomain name where the exception occurred.
      Public ReadOnly Property AppDomainName() As String
         Get
            Return m_appDomainName
         End Get
      End Property

      ' Identity of the executing thread on which the exception was created.
      Public ReadOnly Property ThreadIdentityName() As String
         Get
            Return m_threadIdentity
         End Get
      End Property

      ' Windows identity under which the code was running.
      Public ReadOnly Property WindowsIdentityName() As String
         Get
            Return m_windowsIdentity
         End Get
      End Property

      ' Collection allowing additional information to be added to the exception.
      Public ReadOnly Property AdditionalInformation() As NameValueCollection
         Get
            Return m_additionalInformation
         End Get
      End Property

#End Region

  '  	/// <summary>
        '/// Initialization function that gathers environment information safely.
        '/// </summary>
    Private Sub InitializeEnvironmentInformation()
        Try
            m_machineName = Environment.MachineName
        Catch e As SecurityException
            m_machineName = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED")
        Catch
            m_machineName = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION")
         End Try

        Try
            m_threadIdentity = Thread.CurrentPrincipal.Identity.Name
        Catch e As SecurityException
            m_threadIdentity = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED")
        Catch
            m_threadIdentity = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION")
        End Try

        Try
            m_windowsIdentity = WindowsIdentity.GetCurrent().Name
        Catch e As SecurityException
            m_windowsIdentity = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED")
        Catch
            m_windowsIdentity = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION")
        End Try

        Try
            m_appDomainName = AppDomain.CurrentDomain.FriendlyName
        Catch e As SecurityException
            m_appDomainName = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED")
        Catch
            m_appDomainName = m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION")
        End Try
    End Sub


   End Class 'BaseApplicationException 