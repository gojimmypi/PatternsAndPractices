' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' ConfigurationManager.vb
'
' Public class and entry point for the CMAB.
' 
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
Imports [sc] = System.Configuration
Imports System.Collections
Imports System.Xml
 
Imports Microsoft.ApplicationBlocks.ConfigurationManagement.Storage
Imports Microsoft.ApplicationBlocks.ConfigurationManagement.DataProtection

#Region "Configuration Manager class"

' <summary>
' The Configuration Manager class manages the configuration information based on settings in the configuration file.
' </summary>
Public NotInheritable Class ConfigurationManager

#Region "Declare variables"
    Private Const DEFAULTSECTION_NAME As String = "Default"
    Private Shared _defaultSectionName As String = Nothing
    Private Shared _isInitialized As Boolean = False
    Private Shared _initException As Exception = Nothing
#End Region

#Region "Constructors"

    ' Do not allow this class to be instantiated
    Shared Sub New()
        Try
            InitAllProviders()
            _isInitialized = True
        Catch ex As Exception
            _initException = ex
        End Try
    End Sub 'New

    Private Sub New()
    End Sub 'New 
#End Region

#Region "Public methods"

#Region "Initialize"

    ' <summary>
    ' Initializes the configuration management support
    ' </summary>
    Public Shared Sub Initialize()
        ' Check the static initialization result
        If (Not _isInitialized) Then Throw _initException
    End Sub 'Initialize
#End Region

#Region "Read"

    ' <summary>
    ' Returns the section defined as the DefaultSection on the configuration
    ' file or the first section defined.
    ' </summary>
    ' <returns></returns>
    Public Overloads Shared Function Read() As Hashtable

        If (Not _isInitialized) Then Throw _initException
        Dim section As Object = Read(_defaultSectionName)
        If (section Is Nothing) Then
            Return Nothing
        End If
        If Not TypeOf section Is Hashtable Then
            Throw New sc.ConfigurationErrorsException( _
                    Resource.ResourceManager("RES_ExceptionConfigurationManagerDefaultSectionIsNotHashtable"))
        End If
        Return CType(section, Hashtable)
    End Function 'Read

    ' <summary>
    ' Returns a single specified value
    ' </summary>
    Public Overloads Shared Function Read(ByVal sectionName As String) As Object
        Dim cacheSettings As ConfigurationManagementCache = Nothing
        Dim cachedValue As CacheValue = Nothing
        Dim configReader As IConfigurationStorageReader = Nothing
        Dim configSectionNode As XmlNode = Nothing
        Dim customSectionHandler As sc.IConfigurationSectionHandler = Nothing
        Dim icshValue As Object = Nothing

        If (Not _isInitialized) Then Throw _initException

        'Validate the section name
        If sectionName Is Nothing OrElse sectionName.Length = 0 Then
            Throw New ArgumentNullException("sectionName", _
                        Resource.ResourceManager("RES_ExceptionConfigurationManagerInvalidSectionName"))
        End If

        If sectionName.IndexOf("/") <> -1 Then
            Throw New NotSupportedException(Resource.ResourceManager("RES_ExceptionSectionGroupsNotSupported"))
        End If

        If Not IsValidSection(sectionName) Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionSectionInvalid", sectionName))
        End If
        ' get a cacheSettings object from cachefactory, so we know what the cache settings are
        cacheSettings = CacheFactory.Create(sectionName)

        ' get cache wrapper object CacheValue from cachemanager--
        ' IF cache not null (i.e. config says "don't cache"...as in case of very lively data)
        If Not cacheSettings Is Nothing Then
            cachedValue = CType(cacheSettings(sectionName), CacheValue)
        End If

        ' If the value was cached return the value
        If Not (cachedValue Is Nothing) Then
            '  return its contents, cast to "object" (since we don't know what it is)
            Return CType(cachedValue.Value, Object)
        End If

        ' Create an instance of the storage reader for the given section
        configReader = StorageReaderFactory.Create(sectionName)

        'No provider for the requested section
        If configReader Is Nothing Then
            Throw New Exception(Resource.ResourceManager("RES_ExceptionStorageProviderException", sectionName))
        End If

        '  here we're actually using the storageReader to read an XmlNode.
        '  then feed that node to its associated IConfigurationSectionHandler implementation, 
        '  which decodes the node and returns an object of some type...here represented by icshValue.
        configSectionNode = Nothing
        Try
            configSectionNode = configReader.Read()
        Catch providerException As Exception
            Throw New sc.ConfigurationErrorsException( _
                    Resource.ResourceManager("RES_ExceptionStorageProviderException", sectionName), _
                    providerException)
        End Try

        'The storage provider returns null
        'OK to return null here, they may just not have written the section yet.
        If configSectionNode Is Nothing Then
            Return Nothing
        End If
        '  use the Factory to get our custom ICSH instance
        customSectionHandler = ConfigSectionHandlerFactory.Create(sectionName)

        '  using the ICSH _itself_ as a Factory, create our icshValue from it...in other words,
        '  the ICSH is taking this XmlNode and morphing it to an object instance.
        icshValue = customSectionHandler.Create(Nothing, Nothing, configSectionNode)

        'Update the cache.
        If Not (cacheSettings Is Nothing) Then
            ' The plain value is stored into the cache.
            cacheSettings(cacheSettings.SectionName) = icshValue
        End If

        Return icshValue
    End Function 'Read

#End Region

#Region "Write"


    ' <summary>
    ' Writes the default section (used with the Setting property), to the 
    ' storage provider.
    ' </summary>
    ' <remarks>This method uses the same instance returned by the Setting 
    ' property. If the Settings class is modified this method must be 
    ' called, otherwise the changes are not saved.</remarks>
    Public Overloads Shared Sub Write(ByVal value As Hashtable)
        If (Not _isInitialized) Then Throw _initException
        If value Is Nothing Then Throw New ArgumentNullException("value")
        Write(_defaultSectionName, value)
    End Sub 'Write


    ' <summary>
    ' Writes a single value using the specified section
    ' </summary>
    Public Overloads Shared Sub Write(ByVal sectionName As String, ByVal configValue As Object)
        Dim configStorageWriter As IConfigurationStorageWriter = Nothing
        Dim configStorageReader As IConfigurationStorageReader = Nothing
        Dim configSectionHandler As sc.IConfigurationSectionHandler = Nothing
        Dim configSectionHandlerWriter As IConfigurationSectionHandlerWriter = Nothing
        Dim xmlNode As xmlNode = Nothing
        Dim sectionCache As ConfigurationManagementCache = Nothing

        If (Not _isInitialized) Then Throw _initException
        '  Validate the section name
        If sectionName Is Nothing OrElse sectionName.Length = 0 Then
            Throw New ArgumentNullException("sectionName", _
                        Resource.ResourceManager("RES_ExceptionConfigurationManagerInvalidSectionName"))
        End If

        If sectionName.IndexOf("/") <> -1 Then
            Throw New NotSupportedException(Resource.ResourceManager("RES_ExceptionSectionGroupsNotSupported"))
        End If

        If Not IsValidSection(sectionName) Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionSectionInvalid", sectionName))
        End If
        '  get the storage READER from factory...then we will query its type to see if it can WRITE too
        configStorageReader = StorageReaderFactory.Create(sectionName)

        '  If the section handler is not ICSHW an exception is thrown
        If Not TypeOf configStorageReader Is IConfigurationStorageWriter Then
            Throw New sc.ConfigurationErrorsException( _
                    Resource.ResourceManager("RES_ExceptionHandlerNotWritable", sectionName))
        End If

        '  put Writer cast of the Reader into the configStorageWriter local object...yes, 
        '  we could just cast the original back onto itself
        '  but it's clearer here to have two distinct names
        configStorageWriter = CType(configStorageReader, IConfigurationStorageWriter)

        ' get the section handler
        configSectionHandler = ConfigSectionHandlerFactory.Create(sectionName)

        '  If the section handler is not ICSHW an exception is thrown
        If Not TypeOf configSectionHandler Is IConfigurationSectionHandlerWriter Then
            Throw New sc.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionHandlerNotWritable", sectionName))
        End If

        '  cast the ICSH instance to our custom ICSHWriter type
        configSectionHandlerWriter = CType(configSectionHandler, IConfigurationSectionHandlerWriter)

        '  Just as ICSH takes XML passed to it and creates an object essentially like XML serialization, 
        '  the ICSHWriter we define does the reverse--turns an object into XML.
        '  The IConfigurationSectionHandlerWriter implementation of course is free to use just that--the built-in 
        '.NET xml serialization, and the Quickstarts show this.
        xmlNode = configSectionHandlerWriter.Serialize(configValue)

        ' writes the configValue using the storage provider
        Try
            '  we have the node. Now WRITE this node to the storage location using an instance of
            '  IConfigurationStorageWriter
            configStorageWriter.Write(xmlNode)
        Catch storageProviderException As Exception
            Throw New sc.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionStorageProviderException", sectionName), _
                        storageProviderException)
        End Try

        sectionCache = CacheFactory.Create(sectionName)

        'Update the cache. 
        If Not (sectionCache Is Nothing) Then
            sectionCache(sectionCache.SectionName) = configValue
        End If
    End Sub 'Write

#End Region

#Region "IsReadOnly"

    ' <summary>
    ' Determines whether the section is readonly or not
    ' </summary>
    Public Shared Function IsReadOnly(ByVal sectionName As String) As Boolean

        If (Not _isInitialized) Then Throw _initException

        'Validate the section name
        If sectionName Is Nothing OrElse sectionName.Length = 0 Then
            Throw New ArgumentNullException("sectionName", _
                        Resource.ResourceManager("RES_ExceptionConfigurationManagerInvalidSectionName"))
        End If

        If Not StorageReaderFactory.ContainsKey(sectionName) Then
            'The section doesn't exist
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionSectionInvalid", sectionName))
        End If

        ' create an instance of provider for specified section name.
        Dim configProvider As IConfigurationStorageReader = StorageReaderFactory.Create(sectionName)

        ' create an instance of the section handler
        Dim configSection As sc.IConfigurationSectionHandler = ConfigSectionHandlerFactory.Create(sectionName)

        ' ask the specified storageProvider if it's IConfigurationStorageWriter
        If (TypeOf configProvider Is IConfigurationStorageWriter) And _
                    TypeOf configSection Is IConfigurationSectionHandlerWriter Then
            Return False
        Else
            Return True
        End If
    End Function 'IsReadOnly
#End Region

#End Region

#Region "Private/Internal Methods & Event Handlers"

    ' <summary>
    ' Method added to initialize all the providers.
    ' </summary>
    ' 
    Private Shared Function InitAllProviders() As Boolean
        Dim configMgmtSet As ConfigurationManagementSettings = Nothing
        Dim sectionSettings As ConfigSectionSettings = Nothing

        '  get the configuration settings object that wraps all our config info
        configMgmtSet = ConfigurationManagementSettings.Instance

        Try
            _defaultSectionName = configMgmtSet.DefaultSectionName

            '  have to deal with DictionaryEntry here
            Dim de As DictionaryEntry
            For Each de In configMgmtSet.Sections
                sectionSettings = CType(de.Value, ConfigSectionSettings)

                ' Set the default section
                If (_defaultSectionName Is Nothing OrElse _defaultSectionName.Trim().Length = 0) Then
                    _defaultSectionName = sectionSettings.Name
                End If

                ' use Factory class to make a storageReader; 
                ' NOTE that we only demand it be a Reader at this point, that is sufficient to initialize it.
                ' IF at a later point we wish to use a Writer,
                ' _we will query the object to see if it can be a Writer_
                StorageReaderFactory.Create(sectionSettings.Name)

                ' call into cachefactory so that (if valid) a cache object is created for this section name...
                ' remember the configMgmtCache objects are just encapsulations of the cache settings for
                ' a given config section
                CacheFactory.Create(sectionSettings.Name)
            Next de
            Return True
        Catch e As Exception
            'On any error loading the providers the cache list is cleaned.
            StorageReaderFactory.ClearCache()
            CacheFactory.ClearCache()
            DataProtectionFactory.ClearCache()

            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionProviderInit", _
                            sectionSettings.Name, sectionSettings.Provider.TypeName, _
                            sectionSettings.Provider.AssemblyName), e)
        End Try
    End Function 'InitAllProviders

    '"Creating Storage Provider and Cache for section '{0}' where storage type = '{1}' and assembly = '{2}' "

    ' <summary>
    ' ConfigChanges Event Handler
    ' </summary>
    Friend Shared Sub OnConfigChanges(ByVal storageProvider As IConfigurationStorageReader, _
                            ByVal sectionName As String)
        CacheFactory.ClearCache(sectionName)
    End Sub 'OnConfigChanges

    Friend Shared Function IsValidSection(ByVal sectionName As String) As Boolean
        Dim configMgmtSet As ConfigurationManagementSettings = Nothing

        '  get the configuration settings object that wraps all our config info
        configMgmtSet = ConfigurationManagementSettings.Instance

        If configMgmtSet Is Nothing Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionLoadingConfiguration"))
        End If

        Return Not (configMgmtSet.Sections(sectionName) Is Nothing)
    End Function 'IsValidSection


#End Region

#Region "Singleton"

    ' <summary>
    ' The connection manager singleton instance.
    ' </summary>

    Public Shared ReadOnly Property Items() As ConfigurationManager
        Get
            If (Not _isInitialized) Then Throw _initException
            '  check if singleton exists yet, if not make it
            If _singleton Is Nothing Then
                _singleton = New ConfigurationManager
            End If
            Return _singleton
        End Get
    End Property
    Private Shared _singleton As ConfigurationManager

#End Region

#Region "Item-Instance "

    ' <summary>
    ' Indexer used to get he hashtable instance when the default section
    ' returns a hashtable.
    ' </summary>

    Default Public Property Item(ByVal key As String) As Object
        Get
            Dim section As Hashtable = ConfigurationManager.Read()
            If section Is Nothing Then Return Nothing
            Return section(key)
        End Get

        Set(ByVal Value As Object)
            Dim htSection As Hashtable = ConfigurationManager.Read()
            htSection(key) = Value
            ConfigurationManager.Write(htSection)
        End Set
    End Property

#End Region
End Class 'ConfigurationManager
#End Region
