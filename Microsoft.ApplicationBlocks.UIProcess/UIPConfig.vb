'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' UIPConfig.vb
'
' This file contains the implementations of configuration classes
'
' For more information see the User Interface Process Application Block Implementation Overview. 
' 
'===============================================================================
' Copyright (C) 2000-2001 Microsoft Corporation
' All rights reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
' FITNESS FOR A PARTICULAR PURPOSE.
'==============================================================================

Imports System
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Diagnostics
Imports System.Xml
Imports System.IO
Imports System.Xml.Schema


#Region "Enum Definitions"
'Specifies expiration modes supported by the state cache in the user process manager
Public Enum CacheExpirationMode
   Absolute = 1
   Sliding = 2
   None = 3
End Enum
#End Region

#Region "Configuration setting classes"
#Region "ObjectTypeSettings class"
        
'Base class for all providers settings within the UIP configuration settings in the config file.
Friend Class ObjectTypeSettings
    #Region "Declares Variables"
    Private Const AttributeName As String = "name"
    Private Const AttributeType As String = "type"
    Private Const Comma As String = ","
       
    Private _name As String
    Private _type As String
    Private _assembly As String
    #End Region
   
    #Region "Constructors"
         
    'Default constructor
    'Parameters: 
    Public Sub New()
    End Sub
          
    'Initializes a new instance of the <see cref="ObjectTypeSettings"/> class with a config string.
    'String expected to be of form:
    '
    '		"Microsoft.ApplicationBlocks.UIProcess.FooBar, 
    '		Microsoft.ApplicationBlocks.UIProcess, 
    '		Version=1.0.0.0, Culture=neutral, PublicKeyToken=55cfe03845fe5f4d"
    'Parameters: 
    '-fullType: The configuration type string of above format.
    Public Sub New(fullType As String, name As String)
        '  fix up type/asm strings
        SplitType(fullType)
          
        _name = name
    End Sub
          
    Public Sub New(fullType As String)
        MyClass.New(fullType, "")
    End Sub
       
    'Creates an instance of the ObjectTypeSettings class using the specified configNode
    Public Sub New(configNode As XmlNode)
        Dim currentAttribute As XmlNode
          
        'Gets the typed object attributes
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeType)
        Dim fullType As String
        If currentAttribute.Value.Trim().Length > 0 Then
            fullType = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeType, configNode.Name))
        End If 
        '  fix up type/asm strings
        SplitType(fullType)
          
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeName)
        If currentAttribute.Value.Trim().Length > 0 Then
            _name = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeName, configNode.Name))
        End If 
    End Sub
    #End Region
   
    'Takes incoming full type string, defined as:
    '"Microsoft.ApplicationBlocks.UIProcess.WinFormViewManager,   Microsoft.ApplicationBlocks.UIProcess, 
    '/			Version=1.0.0.4, Culture=neutral, PublicKeyToken=d69d63db1380c14d"
    ' And splits the type into two strings, the typeName and assemblyName.  Those are passed by as OUT params
    ' This routine also cleans up any extra whitespace, and throws an exception if the full type string
    ' does not have five comma-delimited parts....it expect the true full name complete with version and publickeytoken
    'Parameters: 
    '-fullType: 
    '-typeName: 
    '-assemblyName: 
    Private Sub SplitType(fullType As String)
        Dim parts As String() = fullType.Split(Comma.ToCharArray())
          
        If parts.Length = 1 Then
            _type = fullType
        ElseIf parts.Length = 5 Then
            '  set the object type name
            Me._type = parts(0).Trim()
            '  set the object assembly name
            Me._assembly = String.Concat(parts(1).Trim() + Comma, parts(2).Trim() + Comma, parts(3).Trim() + Comma, parts(4).Trim())
        Else
            Throw New ArgumentException(Resource.ResourceManager("RES_ExceptionBadTypeArgumentInFactory"), "fullType")
        End If
    End Sub

    #Region "Properties"
    'Gets the object name
    Public ReadOnly Property Name() As String
        Get
            Return _name
        End Get
    End Property 
       
    'Gets the object full qualified type name
    Public ReadOnly Property Type() As String
        Get
            Return _type
        End Get
    End Property 
       
    'Gets the fully qualified assembly name of the object.
    Public ReadOnly Property [Assembly]() As String
        Get
            Return _assembly
        End Get
    End Property 
       
    #End Region
End Class
#End Region

#Region "StatePersistenceProviderSettings class"
'Class that defines the state persistence provider settings within the UIP configuration settings in the config file.
Friend Class StatePersistenceProviderSettings
    Inherits ObjectTypeSettings
    #Region "Declares Variables"
    Private _attributes As NameValueCollection
    #End Region
       
    #Region "Constructor"
    'Default constructor
    'Parameters: 
    Public Sub New()
        _attributes = New NameValueCollection()
    End Sub 'New
           
    'Creates an instance of the StatePersistenceProviderSettings class using the specified configNode
    Public Sub New(configNode As XmlNode)
        MyBase.New(configNode)
        _attributes = New NameValueCollection()
        Dim currentAttribute As XmlAttribute
        For Each currentAttribute In  configNode.Attributes
            _attributes.Add(currentAttribute.Name, currentAttribute.Value)
        Next currentAttribute
    End Sub
    #End Region
       
    #Region "Properties"
       
    'Gets the state persistence attributes
    Public ReadOnly Property AdditionalAttributes() As NameValueCollection
        Get
            Return _attributes
        End Get
    End Property
    #End Region
End Class
#End Region

#Region "ViewSettings class"
'Class that defines the view settings within the UIP configuration settings in the config file.
Friend Class ViewSettings
    Inherits ObjectTypeSettings
    
    #Region "Declares Variables"
    Private Const AttributeController As String = "controller"
    Private Const AttributeStayOpen As String = "stayOpen"
    Private Const AttributeOpenModal As String = "openModal"
    Private _controller As String
    Private _isStayOpen As Boolean = False
    Private _isOpenModal As Boolean = false
    #End Region
   
    #Region "Constructor"
    'Default constructor
    Public Sub New()
    End Sub
           
    'Creates an instance of the ViewSettings class using the specified configNode
    Public Sub New(configNode As XmlNode)
        MyBase.New(configNode)
        Dim currentAttribute As XmlNode
          
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeController)
        If currentAttribute.Value.Trim().Length > 0 Then
            _controller = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeController, configNode.Name))
        End If 
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeStayOpen)
        If Not (currentAttribute Is Nothing) AndAlso currentAttribute.Value.Trim().Length > 0 Then
            _isStayOpen = XmlConvert.ToBoolean(currentAttribute.Value)
        End If
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeOpenModal)
		If Not ( currentAttribute Is Nothing) AndAlso currentAttribute.Value.Trim().Length > 0 Then
            _isOpenModal = XmlConvert.ToBoolean(currentAttribute.Value)
        End If
    End Sub
    #End Region
   
    #Region "Properties"
    'Gets the controller name related to this view
    Public ReadOnly Property Controller() As String
        Get
            Return _controller
        End Get
    End Property 
       
    'Specifies if the windows should stay open when the navigate
    'method is invoked
    Public ReadOnly Property IsStayOpen() As Boolean
        Get
            Return _isStayOpen
        End Get
    End Property

    'Gets a value indicating whether this view is displayed modally
	Public ReadOnly Property IsOpenModal() As Boolean 
	    Get
            Return _isOpenModal
		End Get
    End Property
    #End Region
End Class
#End Region

#Region "NavigationGraphSettings class"
'Class that defines the navigation graph settings within the UIP configuration settings in the config file.
Friend Class NavigationGraphSettings
    #Region "Declares variables"
    Private Const AttributeIViewManager As String = "iViewManager"
    Private Const AttributeName As String = "name"
    Private Const AttributeState As String = "state"
    Private Const AttributeStatePersist As String = "statePersist"
    Private Const AttributeStartView As String = "startView"
    Private Const AttributeExpirationMode As String = "cacheExpirationMode"
    Private Const AttributeExpirationInterval As String = "cacheExpirationInterval"
    Private Const NodeXPath As String = "node"
       
    Private _name As String
    Private _state As String
    Private _statePersist As String
    Private _iViewManager As String
    Private _startView As String
    Private _expirationMode As CacheExpirationMode = CacheExpirationMode.None
    Private _expirationInterval As TimeSpan = TimeSpan.MinValue
    Private _views As Hashtable
    #End Region
       
    #Region "Constructor"
    'Default constructor
    Public Sub New()
        _views = New Hashtable()
    End Sub
       
    'Creates an instance of the NavigationGraphSettings class using the specified configNode
    Public Sub New(configNode As XmlNode)
        Dim currentAttribute As XmlNode
          
        'Read iViewManager attribute
        '  ****  added 03.26.2003 mstuart to accomodate multiple ViewManager types across app
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeIViewManager)
        If currentAttribute.Value.Trim().Length > 0 Then
            _iViewManager = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeIViewManager, configNode.Name))
        End If 
        'Read name attribute
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeName)
        If currentAttribute.Value.Trim().Length > 0 Then
            _name = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeName, configNode.Name))
        End If 
        'Read state attribute
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeState)
        If currentAttribute.Value.Trim().Length > 0 Then
            _state = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeState, configNode.Name))
        End If 
        'Read statePersist attribute
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeStatePersist)
        If currentAttribute.Value.Trim().Length > 0 Then
            _statePersist = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeStatePersist, configNode.Name))
        End If 
        'Read startView attribute
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeStartView)
        If currentAttribute.Value.Trim().Length > 0 Then
            _startView = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeStartView, configNode.Name))
        End If 
        'Read cache expiration attributes
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeExpirationMode)
        If Not (currentAttribute Is Nothing) Then
            _expirationMode = CType([Enum].Parse(GetType(CacheExpirationMode), currentAttribute.Value, True), CacheExpirationMode)
             
            currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeExpirationInterval)
            Try
                Select Case _expirationMode
                Case CacheExpirationMode.Sliding
                    _expirationInterval = New TimeSpan(0, 0, 0, 0, Integer.Parse(currentAttribute.Value, System.Globalization.CultureInfo.CurrentCulture))
                Case CacheExpirationMode.Absolute
                    _expirationInterval = TimeSpan.Parse(currentAttribute.Value)
                    If _expirationInterval.Days > 0 Then
                            Throw New ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidAbsoluteInterval"))
                    End If
                End Select
            Catch e As Exception
                Throw New ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidCacheExpirationInterval"), e)
            End Try
        End If
          
        _views = New Hashtable()
        Dim currentNode As XmlNode
        For Each currentNode In  configNode.SelectNodes(NodeXPath)
            Dim node As New NodeSettings(currentNode)
            _views.Add(node.View, node)
        Next currentNode
    End Sub
    #End Region
       
    #Region "Properties"
    'Gets the specified node settings
    Default Public ReadOnly Property Item(view As String) As NodeSettings
        Get
            Return CType(_views(view), NodeSettings)
        End Get
    End Property 
       
    'Gets the IViewManager name
    Public ReadOnly Property ViewManager() As String
        Get
            Return _iViewManager
        End Get
    End Property 
       
    'Gets the navigation graph name
    Public ReadOnly Property Name() As String
        Get
            Return _name
        End Get
    End Property 
       
    'Gets the state object type used by this navigation graph
    Public ReadOnly Property State() As String
        Get
            Return _state
        End Get
    End Property 
       
    'Gets the state persist provider used by this navigation graph
    Public ReadOnly Property StatePersist() As String
        Get
            Return _statePersist
        End Get
    End Property 
       
    'Gets the first node configurated in the navigation graph
    Public ReadOnly Property FirstView() As NodeSettings
        Get
            Return CType(_views(_startView), NodeSettings)
        End Get
    End Property 
       
    'Gets the state cache expiration mode
    Public ReadOnly Property CacheExpirationMode() As CacheExpirationMode
        Get
            Return _expirationMode
        End Get
    End Property 
       
    'Gets the state cache expiration interval
    Public ReadOnly Property CacheExpirationInterval() As TimeSpan
        Get
            Return _expirationInterval
        End Get
    End Property
    #End Region
End Class
#End Region

#Region "NavigateToSettings class"
'Class that defines the navigateTo settings within the UIP configuration settings in the config file.
Friend Class NavigateToSettings
    #Region "Declares variables"
    Private Const AttributeNavigateValue As String = "navigateValue"
    Private Const AttributeView As String = "view"
    Private _navigateValue As String
    Private _view As String
    #End Region
   
    #Region "Constructor"
    Public Sub New()
    End Sub
       
    'Creates an instance of the NavigationToSettings class using the specified configNode
    Public Sub New(configNode As XmlNode)
        Dim currentAttribute As XmlNode = configNode.Attributes.RemoveNamedItem(AttributeNavigateValue)
        If currentAttribute.Value.Trim().Length > 0 Then
            _navigateValue = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeNavigateValue, configNode.Name))
        End If 
        currentAttribute = configNode.Attributes.RemoveNamedItem(AttributeView)
        If currentAttribute.Value.Trim().Length > 0 Then
            _view = currentAttribute.Value
        Else
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionInvalidXmlAttributeValue", AttributeView, configNode.Name))
        End If
    End Sub
    #End Region
   
    #Region "Properties"
    'Gets the navigation value
    Public ReadOnly Property NavigateValue() As String
        Get
            Return _navigateValue
        End Get
    End Property 
       
    'Gets/Sets the view name
    Public ReadOnly Property View() As String
        Get
            Return _view
        End Get
    End Property
    #End Region
End Class
#End Region

#Region "NodeSettings class"
'Class that defines the node settings within the UIP configuration settings in the config file.
Friend Class NodeSettings
    #Region "Declares variables"
    Private Const AttributeView As String = "view"
    Private Const NodeNavigateTo As String = "navigateTo"
    Private _view As String
    Private _navigateToCollection As HybridDictionary
    #End Region
       
    #Region "Constructor"
    Public Sub New()
        _navigateToCollection = New HybridDictionary()
    End Sub
           
    'Creates an instance of the NodeSettings class using the specified configNode
    Public Sub New(configNode As XmlNode)
        Dim currentAttribute As XmlNode = configNode.Attributes.RemoveNamedItem(AttributeView)
        If currentAttribute.Value.Trim().Length > 0 Then
            _view = currentAttribute.Value
        End If 
        _navigateToCollection = New HybridDictionary()
        Dim currentNode As XmlNode
        For Each currentNode In  configNode.SelectNodes(NodeNavigateTo)
            Dim navigateTo As New NavigateToSettings(currentNode)
            _navigateToCollection.Add(navigateTo.NavigateValue, navigateTo)
        Next currentNode
    End Sub 'New
    #End Region
       
    #Region "Properties"
    'Gets the specifed navigateTo settings.
    Default Public ReadOnly Property Item(navigateValue As String) As NavigateToSettings
        Get
            Return CType(_navigateToCollection(navigateValue), NavigateToSettings)
        End Get
    End Property 
       
    'Gets the view name
    Public ReadOnly Property View() As String
        Get
            Return _view
        End Get
    End Property
    #End Region
End Class
#End Region

#Region "UIPConfigSettings class"
'This class contains all UIP configuration from config file
'The UIPConfigSettings hierarchy looks like :
'  UIPConfigSettings
'    --- ObjectTypeSettings collection
'    --- ViewSettings collection
'    --- NavigationGraphSettings
'          --- NodeSettings collection
'                --- NavigateToSettings collection 
Friend Class UIPConfigSettings
    #Region "Declares variables"
    Private Const AttributeEnableStateCache As String = "enableStateCache"
    Private Const NodeObjectTypesXPath As String = "objectTypes"
    Private Const NodeIViewManagerXPath As String = "iViewManager"
    Private Const NodeStateXPath As String = "state"
    Private Const NodeControllerXpath As String = "controller"
    Private Const NodeViewXPath As String = "views/view"
    Private Const NodePersistProviderXPath As String = "statePersistenceProvider"
    Private Const NodeNavigationGraphXPath As String = "navigationGraph"
       
    Private _isStateCacheEnabled As Boolean = True
    Private _iViewManagerCollection As HybridDictionary
    Private _stateCollection As HybridDictionary
    Private _controllerCollection As HybridDictionary
    Private _statePersistenceCollection As HybridDictionary
    Private _viewByNameCollection As Hashtable
    Private _navigationGraphCollection As HybridDictionary
    #End Region
       
    #Region "Constructors"
    Public Sub New()
        'Init the hashtables
        _iViewManagerCollection = New HybridDictionary()
        _stateCollection = New HybridDictionary()
        _controllerCollection = New HybridDictionary()
        _statePersistenceCollection = New HybridDictionary()
        _viewByNameCollection = New Hashtable()
        _navigationGraphCollection = New HybridDictionary()
        _isStateCacheEnabled = True
    End Sub
       
    'Creates an instance of the UIPConfigSettings class using the specified configNode
    Public Sub New(configNode As XmlNode) 
        
        Me.New()

        'Get the enableStateCache attribute
        Dim currentAttribute As XmlNode = configNode.Attributes.RemoveNamedItem(AttributeEnableStateCache)
        If Not (currentAttribute Is Nothing) Then
            _isStateCacheEnabled = Convert.ToBoolean(currentAttribute.Value)
        End If 
        'Get the configured IViewManager object types
        Dim typedObject As ObjectTypeSettings
        Dim objectTypeNode As XmlNode
        For Each objectTypeNode In  configNode.SelectSingleNode(NodeObjectTypesXPath).ChildNodes
            Select Case objectTypeNode.LocalName
                Case NodeIViewManagerXPath
                    typedObject = New ObjectTypeSettings(objectTypeNode)
                    _iViewManagerCollection.Add(typedObject.Name, typedObject)
                Case NodeStateXPath
                    typedObject = New ObjectTypeSettings(objectTypeNode)
                    _stateCollection.Add(typedObject.Name, typedObject)
                Case NodeControllerXpath
                    typedObject = New ObjectTypeSettings(objectTypeNode)
                    _controllerCollection.Add(typedObject.Name, typedObject)
                Case NodePersistProviderXPath
                    typedObject = New StatePersistenceProviderSettings(objectTypeNode)
                    _statePersistenceCollection.Add(typedObject.Name, typedObject)
            End Select
        Next objectTypeNode
          
        'Get the configured views
        Dim viewNode As XmlNode
        For Each viewNode In  configNode.SelectNodes(NodeViewXPath)
            typedObject = New ViewSettings(viewNode)
            If Not _viewByNameCollection.ContainsKey(typedObject.Name) Then
                _viewByNameCollection.Add(typedObject.Name, typedObject)
            Else
                Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionViewSettingAlreadyConfigured", typedObject.Name))
            End If
        Next viewNode 
        'Get the configured navigation graphs
        Dim navigationGraph As NavigationGraphSettings
        Dim navigationGraphNode As XmlNode
        For Each navigationGraphNode In  configNode.SelectNodes(NodeNavigationGraphXPath)
            navigationGraph = New NavigationGraphSettings(navigationGraphNode)
            _navigationGraphCollection.Add(navigationGraph.Name, navigationGraph)
        Next navigationGraphNode
    End Sub
    #End Region
       
    #Region "Get Methods"
    Private Function GetNavigationGraphSettings(navigationGraphName As String) As NavigationGraphSettings
        Dim navigationGraph As NavigationGraphSettings = CType(_navigationGraphCollection(navigationGraphName), NavigationGraphSettings)
        If navigationGraph Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionNavigationGraphNotFound", navigationGraphName))
        End If 
        Return navigationGraph
    End Function
    
    'Returns an ObjectTypeSettings wrapper around Type information about IViewManager found in config file
    'Parameters: 
    '-navigationGraphName: name of the Navigation Graph
    'Returns: The IViewManager settings configured for the specified navigation graph 
    Public Overridable Function GetIViewManagerSettings(navigationGraphName As String) As ObjectTypeSettings
        Dim navigationGraph As NavigationGraphSettings = GetNavigationGraphSettings(navigationGraphName)
        Return CType(_iViewManagerCollection(navigationGraph.ViewManager), ObjectTypeSettings)
    End Function
           
    'Specifies if the state cache is enabled or not
    Public Overridable ReadOnly Property IsStateCacheEnabled() As Boolean
        Get
            Return _isStateCacheEnabled
        End Get
    End Property
        
    '/Looks up a view name based on view name
    Public Overridable Function GetViewSettingsFromName(viewName As String) As ViewSettings
        Return CType(_viewByNameCollection(viewName), ViewSettings)
    End Function
       
    'Looks up next view based on incoming graph, view, and nav value
    'Parameters: 
    Public Overridable Function GetNextViewSettings(navigationGraphName As String, currentViewName As String, navigateValue As String) As ViewSettings
        'Retrieve a navgraph class based on nav name
        Dim navigationGraph As NavigationGraphSettings = GetNavigationGraphSettings(navigationGraphName)
          
        '  get the current view node settings
        Dim node As NodeSettings = navigationGraph(currentViewName)
        If node Is Nothing Then
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionCouldNotGetNextViewType", navigationGraphName, currentViewName, navigateValue))
        End If 
        
        '  get the next view name from the navigateTo node
        Dim navigateTo As NavigateToSettings = node(navigateValue)
        If navigateTo Is Nothing Then
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionCouldNotGetNextViewType", navigationGraphName, currentViewName, navigateValue))
        End If 

        '  get the view settings using the view name 
        Return GetViewSettingsFromName(navigateTo.View)
    End Function
       
    'This just looks for the first View name in given navgraph and returns it; 
    Public Overridable Function GetFirstViewSettings(navigationGraphName As String) As ViewSettings
        'Retrieve a navgraph class based on nav name
        Dim navigationGraph As NavigationGraphSettings = GetNavigationGraphSettings(navigationGraphName)
          
        '  return first view settings
        If navigationGraph.FirstView Is Nothing Then
            Return Nothing
        Else
            Return CType(_viewByNameCollection(navigationGraph.FirstView.View), ViewSettings)
        End If
    End Function
       
    'Looks up state persistence provider from graph name
    Public Overridable Function GetPersistenceProviderSettings(navigationGraphName As String) As StatePersistenceProviderSettings
        'Retrieve a navgraph class based on nav name
        Dim navigationGraph As NavigationGraphSettings = GetNavigationGraphSettings(navigationGraphName)
          
        Return CType(_statePersistenceCollection(navigationGraph.StatePersist), StatePersistenceProviderSettings)
    End Function 
           
    'Gets the settings of the state object used by the specified navigation graph
    Public Overridable Function GetStateSettings(navigationGraphName As String) As ObjectTypeSettings
        'Retrieve a navgraph class based on nav name
        Dim navigationGraph As NavigationGraphSettings = GetNavigationGraphSettings(navigationGraphName)
          
        Return CType(_stateCollection(navigationGraph.State), ObjectTypeSettings)
    End Function
       
    'Gets the settings of the controller object used by the specified view
    Public Overridable Function GetControllerSettings(viewName As String) As ObjectTypeSettings
        Dim viewSettings As ViewSettings = GetViewSettingsFromName(viewName)
        If viewSettings Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionViewConfigNotFound", viewName))
        End If 
        Return CType(_controllerCollection(viewSettings.Controller), ObjectTypeSettings)
    End Function
          
    'Gets the cache settings used by the specified navigation graph
    Public Overridable Sub GetCacheConfiguration(navigationGraphName As String, ByRef mode As CacheExpirationMode, ByRef interval As TimeSpan)
        'Retrieve a navgraph class based on nav name
        Dim navigationGraph As NavigationGraphSettings = GetNavigationGraphSettings(navigationGraphName)
          
        mode = navigationGraph.CacheExpirationMode
        interval = navigationGraph.CacheExpirationInterval
    End Sub
    #End Region
End Class
#End Region
#End Region

#Region "Configuration handler"
'The Configuration Section Handler for the "uipConfiguration" section of the config file. 
Friend Class UIPConfigHandler
    Implements IConfigurationSectionHandler
    Private _isValidDocument As Boolean = True
    Private _schemaErrors As String
              
    Public Sub New()
    End Sub
        
    Public Function Create(parent As Object, input As Object, section As XmlNode) As Object Implements IConfigurationSectionHandler.Create
        ValidateSchema(section)
        Dim config As New UIPConfigSettings(section)
        Return config
    End Function
              
    Private Sub ValidateSchema(section As XmlNode)
        
        Dim validatingReader As XmlReader = Nothing
        Dim xsdFile As Stream = Nothing
        Dim streamReader As StreamReader = Nothing
        Try
            ' Set the validation settings on the XmlReaderSettings object.
            Dim settings As XmlReaderSettings = New XmlReaderSettings()
            settings.ValidationType = ValidationType.Schema
            settings.Schemas.Add(XmlSchema.Read(New XmlTextReader(streamReader), Nothing))
            AddHandler settings.ValidationEventHandler, AddressOf ValidationCallBack

            'Validate the document using a schema            
            validatingReader = XmlReader.Create(New XmlTextReader(New StringReader(section.OuterXml)), settings)
        
            xsdFile = Resource.ResourceManager.GetStream("UIPConfigSchema.xsd")
            streamReader = New StreamReader(xsdFile)

            ' Validate the document
            While validatingReader.Read()
                ' nothing, just read to validate
            End While

            If Not _isValidDocument Then
                Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionDocumentNotValidated", _schemaErrors))
            End If
        Finally
            If Not validatingReader Is Nothing Then validatingReader.Close()
            If Not streamReader Is Nothing Then streamReader.Close()
            If Not xsdFile Is Nothing Then xsdFile.Close()
        End Try
    End Sub
        
    Private Sub ValidationCallBack(sender As Object, args As ValidationEventArgs)
        _isValidDocument = False
        _schemaErrors += args.Message + Environment.NewLine
    End Sub
End Class
#End Region

#Region "UIPConfiguration class"
'Helper class to obtain UIP configuration from config file
Friend Class UIPConfiguration
    #Region "Constant members"
    Private Const UipConfigSection As String = "uipConfiguration"
    #End Region
       
    Private Shared CurrentConfig As UIPConfigSettings = Nothing
          
    'Gets the UIP configuration
    Public Shared ReadOnly Property Config() As UIPConfigSettings
        Get
            If CurrentConfig Is Nothing Then
                Try
                    CurrentConfig = CType(ConfigurationManager.GetSection(UipConfigSection), UIPConfigSettings)
                Catch e As Exception
                    Throw New UIPException(Resource.ResourceManager("RES_ExceptionLoadUIPConfig"), e)
                End Try
                
                If CurrentConfig Is Nothing Then
                    Throw New ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionUIPConfigNotFound"))
                End If
            End If
            Return CurrentConfig
        End Get
    End Property
End Class
#End Region
