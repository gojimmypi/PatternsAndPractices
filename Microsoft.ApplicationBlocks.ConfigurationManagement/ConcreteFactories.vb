' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' ConcreteFactories.vb
'
' Factory pattern implementation for the visitors used on the block.
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
Imports System.Collections.Specialized
Imports [sc] = System.Configuration
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Xml

#Region "ConfigSectionHandlerFactory"

Friend NotInheritable Class ConfigSectionHandlerFactory
#Region "Declarations"

    Private Shared _icshCache As HybridDictionary
    Private Shared _xmlAppConfigDoc As XmlDocument

#End Region

#Region "Constructors"

    Shared Sub New()
        _icshCache = New HybridDictionary(5, True)
    End Sub 'New

    Private Sub New()
    End Sub 'New 

#End Region

#Region "Private Helper Methods"

    Private Shared Function FindSectionNode(ByVal xmlDoc As XmlDocument, ByVal sectionName As String) As XmlNode
        Dim currentNode As XmlNode = xmlDoc.SelectSingleNode( _
                            String.Format("/configuration/configSections/section[@name='{0}']", sectionName))
        Return currentNode
    End Function 'FindSectionNode

    Private Shared Sub PutInCache(ByVal icshInstance As sc.IConfigurationSectionHandler, ByVal sectionName As String)
        SyncLock _icshCache.SyncRoot
            _icshCache(sectionName) = icshInstance
        End SyncLock
    End Sub 'PutInCache

#End Region

#Region "Create Overloads"

    Public Shared Function Create(ByVal sectionName As String) As sc.IConfigurationSectionHandler

        If (_xmlAppConfigDoc Is Nothing) Then
            Monitor.Enter(GetType(ConfigSectionHandlerFactory))
            Try
                'Load the Application.config file
                If (_xmlAppConfigDoc Is Nothing) Then
                    _xmlAppConfigDoc = New XmlDocument
                    _xmlAppConfigDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)
                End If
            Finally
                Monitor.Exit(GetType(ConfigSectionHandlerFactory))
            End Try
        End If

        Dim icshInstance As sc.IConfigurationSectionHandler = Nothing

        '  try to find the ICSH in the cache
        icshInstance = CType(_icshCache(sectionName), sc.IConfigurationSectionHandler)

        If Not (icshInstance Is Nothing) Then
            Return icshInstance
        End If
        '  look in Application.config file
        Dim xmlSectionNode As XmlNode = FindSectionNode(_xmlAppConfigDoc, sectionName)

        '  if it is still not found we can not proceed further. Throw an error message.
        '  Without a type definition for the section handler we must stop executing.
        If xmlSectionNode Is Nothing Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionSectionNotFound", sectionName))
        End If

        '  get the fully-qualified type name
        Dim fullTypeName As String = xmlSectionNode.Attributes("type").Value
        If fullTypeName Is Nothing Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionTypeNotSpecified", sectionName))
        End If
        '  create instance using sister utility class
        icshInstance = CType(GenericFactory.Create(fullTypeName), sc.IConfigurationSectionHandler)  '

        '  the configuration section handler shouldn´t be null
        If icshInstance Is Nothing Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidSectionHandler", _
                                                                    fullTypeName))
        End If

        '  remember to cache it
        PutInCache(icshInstance, sectionName)

        Return icshInstance
    End Function 'Create


#End Region
End Class 'ConfigSectionHandlerFactory 


#End Region

#Region "StorageReaderFactory"

Friend NotInheritable Class StorageReaderFactory

#Region "Declarations"

    Private Shared _storageCache As HybridDictionary

#End Region

#Region "Constructors"

    Shared Sub New()
        _storageCache = New HybridDictionary(5, True)
    End Sub 'New

    Private Sub New()
    End Sub 'New 

#End Region

#Region "Private Helper Methods"

    Private Shared Function GetByCreating(ByVal sectionName As String) As IConfigurationStorageReader
        Dim configMgmtSet As ConfigurationManagementSettings = Nothing
        Dim sectionSettings As ConfigSectionSettings = Nothing
        Dim storageReader As IConfigurationStorageReader = Nothing

        '  get the configuration settings object that wraps all our config info
        configMgmtSet = ConfigurationManagementSettings.Instance

        '  get the requested config section by name
        sectionSettings = configMgmtSet(sectionName)

        'Create a new provider Instance
        If Not (sectionSettings.Provider Is Nothing) Then
            ' call the generic factory
            storageReader = CType(GenericFactory.Create(sectionSettings.Provider.AssemblyName, _
                                                sectionSettings.Provider.TypeName), _
                                                IConfigurationStorageReader)
        End If

        If storageReader Is Nothing Then
            Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionProviderInvalidType", _
                                        sectionSettings.Provider.AssemblyName + "," + _
                                        sectionSettings.Provider.TypeName))
        End If

        '  INITIALIZE the Reader
        InitializeStorage(sectionName, storageReader)

        ' Registers to the event on the provider
        AddHandler storageReader.ConfigChanges, AddressOf ConfigurationManager.OnConfigChanges

        'Cache the provider
        SyncLock _storageCache.SyncRoot
            ' cache Storage Provider
            _storageCache(sectionSettings.Name) = storageReader
        End SyncLock

        '  return the storage reader instance
        Return storageReader
    End Function 'GetByCreating

    Private Shared Sub InitializeStorage(ByVal sectionName As String, ByVal reader As IConfigurationStorageReader)
        '  get the config section we need:
        Dim configMngmtSet As ConfigurationManagementSettings = ConfigurationManagementSettings.Instance
        Dim sectionSettings As ConfigSectionSettings = CType(configMngmtSet.Sections(sectionName), _
                                                        ConfigSectionSettings)

        '  get the DataProtection object for this section (if there is one)
        Dim dataProtection As IDataProtection = DataProtectionFactory.Create(sectionName)

        If Not (sectionSettings.Provider Is Nothing) Then
            reader.Init(sectionName, sectionSettings.Provider.OtherAttributes, dataProtection)
        Else
            reader.Init(sectionName, New ListDictionary, dataProtection)
        End If
    End Sub 'InitializeStorage

#End Region

#Region "Create Overloads"

    Public Shared Function Create(ByVal sectionName As String) As IConfigurationStorageReader
        Dim storageReader As IConfigurationStorageReader = Nothing

        '  first try to get storageReader from cache
        storageReader = CType(_storageCache(sectionName), IConfigurationStorageReader) '

        'if storageReader is not null return it, else create, init, cache and return it.
        If Not storageReader Is Nothing Then
            Return storageReader
        Else
            Try
                '  create using helper and return
                Return GetByCreating(sectionName)
            Catch e As Exception
                Throw New sc.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionAppMisConfigured", _
                                                sectionName), e)
            End Try
        End If
    End Function 'Create

#End Region

#Region "Other Public Methods"

    ' <summary>
    ' pass-through to underlying collection object's ContainsKey method
    ' we never want to give direct access to collection so we must forward this call to protect 
    ' the private member's anonymity...and type in case we reoptimize storage later
    ' </summary>
    ' <param name="key">key whose existence we want to query</param>
    ' <returns>boolean true if key exists</returns>
    Public Shared Function ContainsKey(ByVal key As String) As Boolean
        Return _storageCache.Contains(key)
    End Function 'ContainsKey

    ' <summary>
    ' Clears the internal cache of all IConfigurationStorageReader objects
    ' </summary>
    Public Shared Sub ClearCache()
        SyncLock _storageCache.SyncRoot
            _storageCache.Clear()
        End SyncLock
    End Sub 'ClearCache

#End Region
End Class 'StorageReaderFactory 

#End Region

#Region "DataProtectionFactory"

Friend NotInheritable Class DataProtectionFactory

#Region "Declarations"

    Private Shared _protectionCache As HybridDictionary = Nothing

#End Region

#Region "Constructors"

    Shared Sub New()
        _protectionCache = New HybridDictionary(5, True)
    End Sub 'New

    Private Sub New()
    End Sub 'New 

#End Region

#Region "Private Helper Methods"

    Private Shared Function GetByCreating(ByVal sectionName As String) As IDataProtection
        Dim idp As IDataProtection = Nothing
        Dim exceptionDetail As String = ""
        Dim sectionSettings As ConfigSectionSettings = Nothing
        Dim configMgmtSet As ConfigurationManagementSettings = Nothing

        '  get ref to main configmgmtsettings instance (the singleton)
        configMgmtSet = ConfigurationManagementSettings.Instance

        '  Get the requested config section by name
        sectionSettings = CType(configMgmtSet.Sections(sectionName), ConfigSectionSettings)

        ' define detailed exception info here so we can use it at both possible throw points below
        exceptionDetail = Resource.ResourceManager("RES_ExceptionDetailType", _
                                        sectionSettings.Provider.TypeName, _
                                        sectionSettings.Provider.AssemblyName)

        '  make sure that having found correct Section, the DataProtection entry exists
        If Not sectionSettings.DataProtection Is Nothing Then
            Try
                ' Instatiate the IDP implementation using Generic factory
                idp = CType(GenericFactory.Create(sectionSettings.DataProtection.AssemblyName, _
                                        sectionSettings.DataProtection.TypeName), _
                                        IDataProtection)   '

                '  if the IDP is null, throw an exception. It's not a good thing, 
                '  we have a misconfiguration and need to stop executing.
                If idp Is Nothing Then
                    Throw New sc.ConfigurationErrorsException( _
                                        Resource.ResourceManager("RES_ExceptionProtectionProviderInvalidType", _
                                        exceptionDetail))
                End If

                '  NOW initialize the IDataProtection instance...for instance, 
                '  it might need to know where its key is stored
                idp.Init(sectionSettings.DataProtection.OtherAttributes)
            Catch e As Exception
                Throw New sc.ConfigurationErrorsException( _
                            Resource.ResourceManager("RES_ExceptionDataProtectionProviderInit", _
                                sectionSettings.DataProtection.AssemblyName + "-" + _
                                sectionSettings.DataProtection.TypeName, exceptionDetail), e)
            End Try
        End If


        '  is IDP null?  even if it is,  put in cache...that way, we will just cache a null value
        '  and avoid looking for one we know doesn't exist each time
        SyncLock _protectionCache.SyncRoot
            _protectionCache(sectionName) = idp
        End SyncLock

        '  RETURN the IDataProtection instance (or null if this section isn't configured to have one)
        Return idp
    End Function 'GetByCreating 

#End Region

#Region "Create Overloads"

    ' <summary>
    ' returns protection provider for default section if one exists 
    ' otherwise throws exception because dataprotection doesn't exist in default section, or default not defined
    ' </summary>
    ' <returns></returns>
    Public Overloads Shared Function Create() As IDataProtection
        Dim idp As IDataProtection = Nothing

        '  attempt to get from cache
        idp = CType(_protectionCache(ConfigurationManagementSettings.Instance.DefaultSectionName), IDataProtection)   '

        If Not idp Is Nothing Then
            Return idp
        Else
            '  get the default config section name, and pass to other overload
            Return GetByCreating(ConfigurationManagementSettings.Instance.DefaultSectionName)
        End If
    End Function 'Create

    ' <summary>
    ' Creates or loads from internal cache an IDataProtection instance, type-specific to the SectionName passed in
    ' </summary>
    ' <param name="sectionName"></param>
    ' <returns></returns>
    Public Overloads Shared Function Create(ByVal sectionName As String) As IDataProtection
        Dim idp As IDataProtection = Nothing

        '  attempt to get from cache
        idp = CType(_protectionCache(sectionName), IDataProtection)  '

        If Not idp Is Nothing Then
            Return idp
        Else
            Return GetByCreating(sectionName)
        End If
    End Function 'Create 

#End Region

#Region "Clear"

    ' <summary>
    ' Clears the internal cache of all DataProtectionFactory
    ' </summary>
    Public Shared Sub ClearCache()
        SyncLock _protectionCache.SyncRoot
            _protectionCache.Clear()
        End SyncLock
    End Sub 'ClearCache
#End Region
End Class 'DataProtectionFactory

#End Region

#Region "CacheFactory"

Friend NotInheritable Class CacheFactory

#Region "Declarations"

    Private Shared _cacheObjectCache As HybridDictionary

#End Region

#Region "Constructors"

    Shared Sub New()
        _cacheObjectCache = New HybridDictionary(5, True)
    End Sub 'New

    Private Sub New()
    End Sub 'New 
#End Region

#Region "Create Overloads"

    ' <summary>
    ' creates or retrieves from internal cache a ConfigurationManagementCache object, which is just an
    ' encapsulation of cache settings in the configuration file--such as frequency of refresh, 
    ' type of cache, location of cache etc.
    ' </summary>
    ' <param name="sectionName">the name of the config section from which we wish to read cache parameters</param>
    ' <returns>ConfigurationManagementCache object which encapsulates cache settings.</returns>
    Public Shared Function Create(ByVal sectionName As String) As ConfigurationManagementCache
        Dim cacheSettings As ConfigurationManagementCache = Nothing
        Dim configMgmtSet As ConfigurationManagementSettings = Nothing
        Dim sectionSettings As ConfigSectionSettings = Nothing

        '  look in our internal dictionary for the cacheSettings object
        cacheSettings = CType(_cacheObjectCache(sectionName), ConfigurationManagementCache)  '

        '  if it's not null return it otherwise create, cache, and return it;
        If Not cacheSettings Is Nothing Then
            '  found, return it
            Return cacheSettings
        Else
            '  need a fresh one
            '  get config settings object for this sectionname
            '  get the configuration settings object that wraps all our config info
            configMgmtSet = ConfigurationManagementSettings.Instance

            '  get the requested config section by name
            sectionSettings = configMgmtSet(sectionName)

            '  NOW actually check if this section asks for a cache at all; if it does not, just return null
            If (Not sectionSettings Is Nothing) AndAlso (Not sectionSettings.Cache Is Nothing) Then
                If False <> sectionSettings.Cache.IsEnabled Then
                    '  new the cacheSettings object
                    cacheSettings = New ConfigurationManagementCache(sectionSettings.Name, sectionSettings.Cache)

                    '  add it to internal cache
                    SyncLock _cacheObjectCache.SyncRoot
                        _cacheObjectCache(sectionSettings.Name) = cacheSettings
                    End SyncLock
                End If

                '  The sectionSettings DOES NOT HAVE CACHE, or it is NOT ENABLED--so just return null
            Else
                cacheSettings = Nothing
            End If

            '  return it
            Return cacheSettings
        End If
    End Function 'Create

#End Region

#Region "Clear"

    ' <summary>
    ' Clears the internal cache of ConfigurationManagementCache objects.  
    ' </summary>
    Public Overloads Shared Sub ClearCache()
        SyncLock _cacheObjectCache.SyncRoot
            _cacheObjectCache.Clear()
        End SyncLock
    End Sub 'ClearCache

    '  clears a particular cacheSettings entry based on section name
    Public Overloads Shared Sub ClearCache(ByVal sectionName As String)
        '  get the particular ConfigurationManagementCache referred to by key
        Dim cacheSettings As ConfigurationManagementCache = Create(sectionName)

        '  if it's not null, clear it
        If Not cacheSettings Is Nothing Then
            cacheSettings.Clear()
        End If
    End Sub 'ClearCache

#End Region
End Class 'CacheFactory 

#End Region
