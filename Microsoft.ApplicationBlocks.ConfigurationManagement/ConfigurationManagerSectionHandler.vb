' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' ConfigurationManagerSectionHandler.vb
'
' Section handler used to read the CMAB configuration for the sectrions
' placed on the configuration file.
'
' For more information see the Configuration Management Application Block Implementation Overview. 
' 
' ===============================================================================
' Copyright (C) 2000-2001 Microsoft Corporation
' All rights reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
' FITNESS FOR A PARTICULAR PURPOSE.
' ==============================================================================

Imports System
Imports [SC] = System.Configuration
Imports System.Collections.Specialized
Imports System.IO
Imports System.Xml
Imports System.Xml.Schema

#Region "Class Definitions"

' <summary>
' Class that defines the settings within the configuration management settings in the config file.
' </summary>
Friend Class ConfigurationManagementSettings

#Region "Member fields"

    Private _configSections As New HybridDictionary(5, False)
    Private Shared _singletonSelf As ConfigurationManagementSettings = Nothing

#End Region

#Region "Constructors"

    Shared Sub New()
    End Sub 'New

    Public Sub New()
    End Sub 'New 
#End Region

#Region "Singleton"

    Public Shared ReadOnly Property Instance() As ConfigurationManagementSettings
        Get
            '  check if singleton exists yet; 
            If _singletonSelf Is Nothing Then
                '  NO, load it
                _singletonSelf = CType(SC.ConfigurationManager.GetSection("applicationConfigurationManagement"), _
                                        ConfigurationManagementSettings)
            End If

            '  return it
            Return _singletonSelf
        End Get
    End Property

#End Region

#Region "Public properties"

    ' <summary>
    ' An ArrayList containing all of the ConfigSectionSettings listed in the config file.
    ' </summary>
    Public ReadOnly Property Sections() As HybridDictionary
        Get
            Return _configSections
        End Get
    End Property

    ' <summary>
    ' The section used when the user adds the DefaultSectionName attribute to the
    ' applicationConfigurationManagement configuration section.
    ' </summary>
    Public Property DefaultSectionName() As String
        Get
            Return _defaultSectionName
        End Get
        Set(ByVal Value As String)
            _defaultSectionName = Value
        End Set
    End Property
    Private _defaultSectionName As String


    Default Public ReadOnly Property Item(ByVal key As String) As ConfigSectionSettings
        Get
            Return CType(_configSections(key), ConfigSectionSettings)
        End Get
    End Property
#End Region

#Region "Public methods"


    ' <summary>
    ' Adds a ConfigSectionSettings to the arraylist of sections.
    ' </summary>
    Public Sub AddConfigurationSection(ByVal configSection As ConfigSectionSettings)
        _configSections(configSection.Name) = configSection
    End Sub 'AddConfigurationSection


#End Region
End Class 'ConfigurationManagementSettings

' <summary>
' Class that defines the cache settings within the configuration management settings in the config file.
' </summary>
Friend Class ConfigCacheSettings
#Region "Member fields"
    Private _isEnabled As Boolean = False
    Private _refresh As String
#End Region

#Region "Public properties"

    ' <summary>
    ' This property specifies if the cache should be enabled or not.
    ' </summary>
    Public Property IsEnabled() As Boolean
        Get
            Return _isEnabled
        End Get
        Set(ByVal Value As Boolean)
            _isEnabled = Value
        End Set
    End Property

    ' <summary>
    ' Absolute time format for refresh of the config data cache. 
    ' </summary>
    Public Property Refresh() As String
        Get
            Return _refresh
        End Get
        Set(ByVal Value As String)
            _refresh = Value
        End Set
    End Property
#End Region
End Class 'ConfigCacheSettings

' <summary>
' Class that defines the provider settings within the configuration management settings in the config file.
' </summary>
Friend Class ConfigProviderSettings
#Region "Member fields"
    Private _typeName As String
    Private _assemblyName As String
    Private _otherAttributes As New ListDictionary
#End Region

#Region "Public properties"
    ' <summary>
    ' The assembly name of the configuration provider component that will be used to invoke the object.
    ' </summary>
    Public Property AssemblyName() As String
        Get
            Return _assemblyName
        End Get
        Set(ByVal Value As String)
            _assemblyName = Value
        End Set
    End Property

    ' <summary>
    ' The type name of the configuration provider component that will be used to invoke the object.
    ' </summary>
    Public Property TypeName() As String
        Get
            Return _typeName
        End Get
        Set(ByVal Value As String)
            _typeName = Value
        End Set
    End Property

    ' <summary>
    ' An collection of any other attributes included within the provider tag in the config file. 
    ' </summary>
    Public ReadOnly Property OtherAttributes() As ListDictionary
        Get
            Return _otherAttributes
        End Get
    End Property
#End Region

#Region "Public members"

    ' <summary>
    ' Allows name/value pairs to be added to the Other Attributes collection.
    ' </summary>
    Public Sub AddOtherAttributes(ByVal name As String, ByVal value As String)
        _otherAttributes.Add(name, value)
    End Sub 'AddOtherAttributes
#End Region
End Class 'ConfigProviderSettings

' <summary>
' Class that defines the protection provider settings within the configuration management settings in the config file.
' </summary>

Friend Class DataProtectionProviderSettings
#Region "Declare variables"
    Private _typeName As String
    Private _assemblyName As String
    Private _otherAttributes As New ListDictionary
#End Region

#Region "Public properties"
    ' <summary>
    ' The assembly name of the protection configuration provider component that will be used to invoke the object.
    ' </summary>
    Public Property AssemblyName() As String
        Get
            Return _assemblyName
        End Get
        Set(ByVal Value As String)
            _assemblyName = Value
        End Set
    End Property

    ' <summary>
    ' The type name of the protection provider component that will be used to invoke the object.
    ' </summary>
    Public Property TypeName() As String
        Get
            Return _typeName
        End Get
        Set(ByVal Value As String)
            _typeName = Value
        End Set
    End Property

    ' <summary>
    ' An collection of any other attributes included within the provider tag in the config file. 
    ' </summary>
    Public ReadOnly Property OtherAttributes() As ListDictionary
        Get
            Return _otherAttributes
        End Get
    End Property
#End Region

#Region "Public methods"

    ' <summary>
    ' Allows name/value pairs to be added to the Other Attributes collection.
    ' </summary>
    Public Sub AddOtherAttributes(ByVal name As String, ByVal value As String)
        _otherAttributes.Add(name, value)
    End Sub 'AddOtherAttributes
#End Region
End Class 'DataProtectionProviderSettings

' <summary>
' Class that defines the section settings within the configuration management settings in the config file.
' </summary>

Friend Class ConfigSectionSettings
#Region "Declare variables"
    Private _cache As ConfigCacheSettings
    Private _provider As ConfigProviderSettings
    Private _protection As DataProtectionProviderSettings
    Private _name As String
#End Region

#Region "Public properties"

    ' <summary>
    ' A ConfigCacheSettings configurated in the config file.
    ' </summary>

    Public Property Cache() As ConfigCacheSettings
        Get
            Return _cache
        End Get
        Set(ByVal Value As ConfigCacheSettings)
            _cache = Value
        End Set
    End Property

    ' <summary>
    ' A ConfigProviderSettings configurated in the config file.
    ' </summary>

    Public Property Provider() As ConfigProviderSettings
        Get
            Return _provider
        End Get
        Set(ByVal Value As ConfigProviderSettings)
            _provider = Value
        End Set
    End Property

    ' <summary>
    ' A ProtectionProviderSettings configurated in the config file.
    ' </summary>

    Public Property DataProtection() As DataProtectionProviderSettings
        Get
            Return _protection
        End Get
        Set(ByVal Value As DataProtectionProviderSettings)
            _protection = Value
        End Set
    End Property

    ' <summary>
    ' This property specifies the section name
    ' </summary>

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Value
        End Set
    End Property

#End Region
End Class 'ConfigSectionSettings

#End Region

#Region "ConfigurationManagerSectionHandler"

' <summary>
' The Configuration Section Handler for the "configurationManagement" section of the config file. 
' </summary>
Friend Class ConfigurationManagerSectionHandler
    Implements SC.IConfigurationSectionHandler

#Region "Members"
    Private _isValidDocument As Boolean = True
    Private _schemaErrors As String = ""
#End Region

#Region "Constructors"

    ' <summary>
    ' The constructor for the ConfigurationManagerSectionHandler to initialize the resource file.
    ' </summary>
    Public Sub New()
    End Sub 'New
#End Region

#Region "Implementation of IConfigurationSectionHandler"

    ' <summary>
    ' Builds the ConfigurationManagementSettings, ConfigurationProviderSettings and 
    ' ConfigurationItemsSettings structures based on the configuration file.
    ' </summary>
    ' <param name="parent">Composed from the configuration settings in a corresponding parent configuration section.</param>
    ' <param name="configContext">Provides access to the virtual path for which the configuration 
    ' section handler computes configuration values. Normally this parameter is reserved and is null.</param>
    ' <param name="section">The XML node that contains the configuration information to be handled. 
    ' Section provides direct access to the XML contents of the configuration section.</param>
    ' <returns>The ConfigurationManagementSettings struct built from the configuration settings.</returns>
    Public Function Create(ByVal parent As Object, ByVal configContext As Object, ByVal section As XmlNode) As Object Implements SC.IConfigurationSectionHandler.Create
        Try
            Dim configSettings As New ConfigurationManagementSettings

            ' Exit if there is no configuration settings.
            If section Is Nothing Then
                Return configSettings
            End If

            Dim xsdFile As Stream = Resource.ResourceManager.GetStream( _
                            "Microsoft.ApplicationBlocks.ConfigurationManagement.ConfigSchema.xsd")
            Try
                Dim sr As New StreamReader(xsdFile)
                Try
                    ' Set the validation settings on the XmlReaderSettings object.
                    Dim settings As XmlReaderSettings = New XmlReaderSettings()
                    settings.ValidationType = ValidationType.Schema

                    AddHandler settings.ValidationEventHandler, AddressOf ValidationCallBack
                    settings.Schemas.Add(XmlSchema.Read(New XmlTextReader(sr), Nothing))
                    settings.ValidationType = ValidationType.Schema

                    'Validate the document using a schema
                    Dim vreader As XmlReader = XmlReader.Create(New XmlTextReader(New StringReader(section.OuterXml)), settings)
                    While vreader.Read()
                        ' nothing, just read to validate
                    End While
                    If Not _isValidDocument Then
                        Throw New SC.ConfigurationErrorsException( _
                                Resource.ResourceManager("Res_ExceptionDocumentNotValidated", _schemaErrors))
                    End If
                Finally
                    CType(sr, IDisposable).Dispose()
                End Try
            Finally
                CType(xsdFile, IDisposable).Dispose()
            End Try
            Dim attr As XmlAttribute = section.Attributes("defaultSection")
            If Not (attr Is Nothing) Then
                configSettings.DefaultSectionName = attr.Value
            End If

            '#region "Loop through the section components and load them into the ConfigurationManagementSettings"

            Dim sectionSettings As ConfigSectionSettings = Nothing
            Dim configChildNode As XmlNode
            For Each configChildNode In section.ChildNodes
                If configChildNode.Name = "configSection" Then
                    ProcessConfigSection(configChildNode, sectionSettings, configSettings)
                End If
            Next configChildNode
            '#end region
            ' Return the ConfigurationManagementSettings loaded with the values from the config file.
            Return configSettings

        Catch exc As Exception
            Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionLoadingConfiguration"), _
                                            exc, section)
        End Try
    End Function 'Create


#End Region

#Region "Private methods"

    ' <summary>
    ' Process the "configSection" xml node section
    ' </summary>
    ' <param name="configChildNode"></param>
    ' <param name="sectionSettings"></param>
    ' <param name="configSettings"></param>
    Sub ProcessConfigSection(ByVal configChildNode As XmlNode, ByRef sectionSettings As ConfigSectionSettings, _
                        ByVal configSettings As ConfigurationManagementSettings)
        ' Initialize a new ConfigSectionSettings.
        sectionSettings = New ConfigSectionSettings

        ' Get a collection of all the attributes.
        Dim nodeAttributes As XmlAttributeCollection = configChildNode.Attributes

        '#region "Remove the known attributes and load the struct values"

        ' Remove the name attribute from the node and set its value in ConfigSectionSettings.
        Dim currentAttribute As XmlNode = nodeAttributes.RemoveNamedItem("name")
        If Not (currentAttribute Is Nothing) Then
            sectionSettings.Name = currentAttribute.Value
        End If
        ' Loop through the section components and load them into the ConfigurationManagementSettings.
        Dim cacheSettings As ConfigCacheSettings = Nothing
        Dim providerSettings As ConfigProviderSettings = Nothing
        Dim protectionSettings As DataProtectionProviderSettings = Nothing
        Dim sectionChildNode As XmlNode
        For Each sectionChildNode In configChildNode.ChildNodes
            Select Case sectionChildNode.Name
                Case "configCache"
                    ProcessConfigCacheSection(sectionChildNode, cacheSettings, sectionSettings)
                Case "configProvider"
                    ProcessConfigProviderSection(sectionChildNode, providerSettings, sectionSettings)
                Case "protectionProvider"
                    ProcessProtectionProvider(sectionChildNode, protectionSettings, sectionSettings)
                Case Else
            End Select
        Next sectionChildNode
        '#end region

        ' Add the ConfigurationSectionSettings to the sections collection.
        configSettings.AddConfigurationSection(sectionSettings)
    End Sub 'ProcessConfigSection


    ' <summary>
    ' Process the "configCache" xml node section
    ' </summary>
    ' <param name="sectionChildNode"></param>
    ' <param name="cacheSettings"></param>
    ' <param name="sectionSettings"></param>
    Sub ProcessConfigCacheSection(ByVal sectionChildNode As XmlNode, ByRef cacheSettings As ConfigCacheSettings, _
                            ByVal sectionSettings As ConfigSectionSettings)
        ' Initialize a new ConfigCacheSettings.
        cacheSettings = New ConfigCacheSettings

        ' Get a collection of all the attributes.
        Dim nodeAttributes As XmlAttributeCollection = sectionChildNode.Attributes

        ' Remove the enabled attribute from the node and set its value in ConfigCacheSettings.
        Dim currentAttribute As XmlNode = nodeAttributes.RemoveNamedItem("enabled")
        If Not (currentAttribute Is Nothing) AndAlso _
                    currentAttribute.Value.ToUpper(System.Globalization.CultureInfo.CurrentUICulture) = "TRUE" Then
            cacheSettings.IsEnabled = True
        End If
        ' Remove the refresh attribute from the node and set its value in ConfigCacheSettings.
        currentAttribute = nodeAttributes.RemoveNamedItem("refresh")
        If Not (currentAttribute Is Nothing) Then
            cacheSettings.Refresh = currentAttribute.Value
        End If
        ' Set the ConfigurationCacheSettings to the section cache.
        sectionSettings.Cache = cacheSettings
    End Sub 'ProcessConfigCacheSection


    ' <summary>
    ' Process the "configProvider" xml node section
    ' </summary>
    ' <param name="sectionChildNode"></param>
    ' <param name="providerSettings"></param>
    ' <param name="sectionSettings"></param>
    Sub ProcessConfigProviderSection(ByVal sectionChildNode As XmlNode, _
                ByRef providerSettings As ConfigProviderSettings, ByVal sectionSettings As ConfigSectionSettings)
        ' Initialize a new ConfigProviderSettings.
        providerSettings = New ConfigProviderSettings

        ' Get a collection of all the attributes.
        Dim nodeAttributes As XmlAttributeCollection = sectionChildNode.Attributes

        '#region "Remove the provider known attributes and load the struct values"

        ' Remove the assembly attribute from the node and set its value in ConfigProviderSettings.
        Dim currentAttribute As XmlNode = nodeAttributes.RemoveNamedItem("assembly")
        If Not (currentAttribute Is Nothing) Then
            providerSettings.AssemblyName = currentAttribute.Value.Trim()
        End If
        ' Remove the type attribute from the node and set its value in ConfigProviderSettings.
        currentAttribute = nodeAttributes.RemoveNamedItem("type")
        If Not (currentAttribute Is Nothing) Then
            providerSettings.TypeName = currentAttribute.Value
        End If '
        '#end region

        '#region "Loop through any other attributes and load them into OtherAttributes"

        ' Loop through any other attributes and load them into OtherAttributes.
        Dim i As Integer
        For i = 0 To nodeAttributes.Count - 1
            providerSettings.AddOtherAttributes(nodeAttributes.Item(i).Name, nodeAttributes.Item(i).Value)
        Next i

        '#end region

        ' Set the ConfigurationProviderSettings to the section provider.
        sectionSettings.Provider = providerSettings
    End Sub 'ProcessConfigProviderSection


    ' <summary>
    ' Process the "protectionProvider" xml node section
    ' </summary>
    ' <param name="sectionChildNode"></param>
    ' <param name="protectionSettings"></param>
    ' <param name="sectionSettings"></param>
    Sub ProcessProtectionProvider(ByVal sectionChildNode As XmlNode, _
                    ByRef protectionSettings As DataProtectionProviderSettings, _
                    ByVal sectionSettings As ConfigSectionSettings)

        ' Initialize a new DataProtectionProviderSettings.
        protectionSettings = New DataProtectionProviderSettings

        ' Get a collection of all the attributes.
        Dim nodeAttributes As XmlAttributeCollection = sectionChildNode.Attributes

        '#region "Remove the provider known attributes and load the struct values"

        ' Remove the assembly attribute from the node and set its value in DataProtectionProviderSettings.
        Dim currentAttribute As XmlNode = nodeAttributes.RemoveNamedItem("assembly")
        If Not (currentAttribute Is Nothing) Then
            protectionSettings.AssemblyName = currentAttribute.Value
        End If
        ' Remove the type attribute from the node and set its value in DataProtectionProviderSettings.
        currentAttribute = nodeAttributes.RemoveNamedItem("type")
        If Not (currentAttribute Is Nothing) Then
            protectionSettings.TypeName = currentAttribute.Value
        End If '
        '#end region

        '#region "Loop through any other attributes and load them into OtherAttributes"

        ' Loop through any other attributes and load them into OtherAttributes.
        Dim i As Integer
        For i = 0 To nodeAttributes.Count - 1
            protectionSettings.AddOtherAttributes(nodeAttributes.Item(i).Name, nodeAttributes.Item(i).Value)
        Next i
        '#end region

        ' Set the DataProtectionProviderSettings to the section provider.
        sectionSettings.DataProtection = protectionSettings
    End Sub 'ProcessProtectionProvider


    Private Sub ValidationCallBack(ByVal sender As Object, ByVal args As ValidationEventArgs)
        _isValidDocument = False
        _schemaErrors += args.Message + Environment.NewLine
        ' TODO - write to eventlog here
    End Sub 'ValidationCallBack

#End Region
End Class 'ConfigurationManagerSectionHandler
#End Region
