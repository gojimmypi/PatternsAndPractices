Imports System.Configuration

#Region "Configuration Class Definitions"
#Region "Enum Definitions"
   ' Enum containing the mode options for the exceptionManagement tag.
   Public Enum ExceptionManagementMode
      ' The ExceptionManager should not process exceptions.
      [Off]
      ' The ExceptionManager should process exceptions. This is the default.
      [On]
   End Enum 'ExceptionManagementMode

   ' Enum containing the mode options for the publisher tag.
   Public Enum PublisherMode
      ' The ExceptionManager should not call the publisher.
      [Off]
      ' The ExceptionManager should call the publisher. This is the default.
      [On]
   End Enum 'PublisherMode

   ' Enum containing the format options for the publisher tag.
   Public Enum PublisherFormat
      ' The ExceptionManager should call the IExceptionPublisher interface of the publisher. 
      ' This is the default.
      Exception
      ' The ExceptionManager should call the IExceptionXmlPublisher interface of the publisher. 
      Xml
   End Enum 'PublisherFormat

#End Region

#Region "Class Definitions"
   ' Class that defines the exception management settings in the config file.
   Public Class ExceptionManagementSettings
      Private m_mode As ExceptionManagementMode = ExceptionManagementMode.On
      Private m_publishers As New ArrayList()

      ' Specifies the whether the exceptionManagement settings are "on" or "off".
      Public Property Mode() As ExceptionManagementMode
         Get
            Return m_mode
         End Get
         Set(ByVal Value As ExceptionManagementMode)
           m_mode = Value
         End Set
      End Property

      ' An ArrayList containing all of the PublisherSettings listed in the config file.
      Public ReadOnly Property Publishers() As ArrayList
         Get
            Return m_publishers
         End Get
      End Property

      ' Adds a PublisherSettings to the arraylist of publishers.
      Public Sub AddPublisher(ByVal publisher As PublisherSettings)
         Publishers.Add(publisher)
      End Sub 'AddPublisher
   End Class 'ExceptionManagementSettings

   ' Class that defines the publisher settings within the exception management settings in the config file.
   Public Class PublisherSettings
      Private m_mode As PublisherMode = PublisherMode.On
      Private m_exceptionFormat As PublisherFormat = PublisherFormat.Exception
      Private m_assemblyName As String
      Private m_typeName As String
      Private m_includeTypes As TypeFilter
      Private m_excludeTypes As TypeFilter
      Private m_otherAttributes As New NameValueCollection()

      ' Specifies the whether the exceptionManagement settings are "on" or "off".
      Public Property Mode() As PublisherMode
         Get
            Return m_mode
         End Get
         Set(ByVal Value As PublisherMode)
            m_mode = Value
         End Set
      End Property

      ' Specifies the whether the publisher supports the IExceptionXmlPublisher interface (value is set to "xml")
      ' or the publisher supports the IExceptionPublisher interface (value is either left off or set to "exception").
      Public Property ExceptionFormat() As PublisherFormat
         Get
            Return m_exceptionFormat
         End Get
         Set(ByVal Value As PublisherFormat)
            m_exceptionFormat = Value
         End Set
      End Property

      ' The assembly name of the publisher component that will be used to invoke the object.
      Public Property AssemblyName() As String
         Get
            Return m_assemblyName
         End Get
         Set(ByVal Value As String)
            m_assemblyName = Value
         End Set
      End Property

      ' The type name of the publisher component that will be used to invoke the object.
      Public Property TypeName() As String
         Get
            Return m_typeName
         End Get
         Set(ByVal Value As String)
            m_typeName = Value
         End Set
      End Property

      ' A semicolon delimited list of all exception types that the publisher will be invoked for.  
      ' A "*" can be used to specify all types and is the default value if this is left off.
      Public Property IncludeTypes() As TypeFilter
         Get
            Return m_includeTypes
         End Get
         Set(ByVal Value As TypeFilter)
            m_includeTypes = Value
         End Set
      End Property

      ' A semicolon delimited list of all exception types that the publisher will not be invoked for. 
      ' A "*" can be used to specify all types. The default is to exclude no types.
      Public Property ExcludeTypes() As TypeFilter
         Get
            Return m_excludeTypes
         End Get
         Set(ByVal Value As TypeFilter)
            m_excludeTypes = Value
         End Set
      End Property

      ' Determines whether the exception type is to be filtered out based on the includes and exclude
      ' types specified.
      ' Parameters:
      ' -exceptionType - The Type of the exception to check for filtering. 
      ' Returns: True is the exception type is to be filtered out, false if it is not filtered out. 
      Public Function IsExceptionFiltered(ByVal exceptionType As Type) As Boolean
         ' If no types are excluded then the exception type is not filtered.
         If m_excludeTypes Is Nothing Then
            Return False
         End If

        ' If the Type is in the Exclude Filter
        If (MatchesFilter(exceptionType, ExcludeTypes)) Then
            ' If the Type is in the Include Filter
            If (MatchesFilter(exceptionType, IncludeTypes)) Then
                'The Type is not filtered out because it was explicitly Included.
                Return False
            'If the Type is not in the Include Filter
            Else
                'The Type is filtered because it was Excluded and did not match the Include Filter.
                Return True
            End If

        'Otherwise it is not Filtered.
        Else
            ' The Type is not filtered out because it did not match the Exclude Filter.
            Return False
        End If

      End Function 'IsExceptionFiltered

        'Determines if a type is contained the supplied filter. 
        Private Function MatchesFilter(ByVal TypeToCompare As Type, ByVal Filter As TypeFilter) As Boolean
            Dim m_typeInfo As TypeInfo
            Dim i As Short

            'If no filter is provided type does not match the filter.
            If Filter Is Nothing Then Return False

            'If all types are accepted in the filter (using the "*") return true.
            If Filter.AcceptAllTypes Then Return True

        For i = 0 To CShort(Filter.Types.Count - 1)
            m_typeInfo = CType(Filter.Types(i), TypeInfo)

            'If the Type matches this type in the Filter, then return true.
            If m_typeInfo.ClassType.Equals(TypeToCompare.GetType) Then Return True

            'If the filter type includes SubClasses of itself (it had a "+" before the type in the
            'configuration file) AND the Type is a SubClass of the filter type, then return true.
            If m_typeInfo.IncludeSubClasses = True AndAlso m_typeInfo.ClassType.IsAssignableFrom(TypeToCompare) Then Return True
        Next

        'If no matches are found return false
        Return False

    End Function

    ' A collection of any other attributes included within the publisher tag in the config file. 
    Public ReadOnly Property OtherAttributes() As NameValueCollection
        Get
            Return m_otherAttributes
        End Get
    End Property

    ' Allows name/value pairs to be added to the Other Attributes collection.
    Public Sub AddOtherAttributes(ByVal name As String, ByVal value As String)
        OtherAttributes.Add(name, value)
    End Sub 'AddOtherAttributes
End Class 'PublisherSettings

    'TypeFilter class stores contents of the Include and Exclude filters provided in the
    'configuration file
    Public Class TypeFilter
        Private m_acceptAllTypes As Boolean = False
        Private m_types As ArrayList = New ArrayList()

        'Indicates if all types should be accepted for a filter
        Public Property AcceptAllTypes() As Boolean
        Get
            Return m_acceptAllTypes
        End Get
        Set(ByVal Value As Boolean)
            m_acceptAllTypes = Value
        End Set
        End Property

        'Collection of types for the filter
        Public ReadOnly Property Types() As ArrayList
        Get
            Return m_types
        End Get
        End Property

    End Class


   'TypeInfo class contains information about each type within a TypeFilter
   Public Class TypeInfo
        Private m_classType As Type
        Private m_includeSubClasses As Boolean = False

        'Indicates if subclasses are to be included with the type specified in the Include and Exclude filters
        Public Property IncludeSubClasses() As Boolean
        Get
            Return m_includeSubClasses
        End Get
        Set(ByVal Value As Boolean)
            m_includeSubClasses = Value
        End Set
        End Property

        'The Type class representing the type specified in the Include and Exclude filters
        Public Property ClassType() As Type
        Get
            Return m_classType
        End Get
        Set(ByVal Value As Type)
            m_classType = Value
        End Set
        End Property

   End Class

#End Region
#End Region

#Region "ExceptionManagerSectionHandler"

   ' The Configuration Section Handler for the "exceptionManagement" section of the config file. 
   Public Class ExceptionManagerSectionHandler
    Implements IConfigurationSectionHandler

#Region "Constructors"
      ' The constructor for the ExceptionManagerSectionHandler to initialize the resource file.
      Public Sub New()
         ' Load Resource File for localized text.
         m_resourceManager = New ResourceManager(Me.GetType().Namespace + ".ExceptionManagerText", Me.GetType().Assembly)
      End Sub 'New

#End Region

#Region "Declare Variables"
      ' Member variables.
      Private Shared EXCEPTION_TYPE_DELIMITER As Char = Convert.ToChar(";") '
      Private EXCEPTIONMANAGEMENT_MODE As String = "mode"
      Private PUBLISHER_NODENAME As String = "publisher"
      Private PUBLISHER_MODE As String = "mode"
      Private PUBLISHER_ASSEMBLY As String = "assembly"
      Private PUBLISHER_TYPE As String = "type"
      Private PUBLISHER_EXCEPTIONFORMAT As String = "exceptionFormat"
      Private PUBLISHER_INCLUDETYPES As String = "include"
      Private PUBLISHER_EXCLUDETYPES As String = "exclude"
      Private m_resourceManager As ResourceManager

#End Region

      ' Builds the ExceptionManagementSettings and PublisherSettings structures based on the configuration file.
      ' Parameters:
      ' -parent - Composed from the configuration settings in a corresponding parent configuration section. 
      ' -configContext - Provides access to the virtual path for which the configuration section handler computes configuration values. Normally this parameter is reserved and is null. 
      ' -section - The XML node that contains the configuration information to be handled. section provides direct access to the XML contents of the configuration section. 
      ' Returns: The ExceptionManagementSettings struct built from the configuration settings. 
      Public Overridable Function Create(ByVal parent As Object, ByVal configContext As Object, ByVal section As XmlNode) As Object Implements IConfigurationSectionHandler.Create
         'Dim settings As New ExceptionManagementSettings()
         'Dim currentAttribute As XmlNode
         'Dim nodeAttributes As XmlAttributeCollection = section.Attributes
         Dim publisherSettings As publisherSettings
         Dim node As XmlNode
         Dim i As Integer
         Dim j As Integer

         Try

            Dim settings As New ExceptionManagementSettings()

            ' Exit if there are no configuration settings.
            If section Is Nothing Then
               Return settings
            End If

            Dim currentAttribute As XmlNode
            Dim nodeAttributes As XmlAttributeCollection = section.Attributes

            ' Get the mode attribute.
            currentAttribute = nodeAttributes.RemoveNamedItem(EXCEPTIONMANAGEMENT_MODE)
            If Not (currentAttribute Is Nothing) AndAlso currentAttribute.Value.ToUpper(CultureInfo.InvariantCulture) = "OFF" Then
                settings.Mode = ExceptionManagementMode.Off
            End If

            ' Loop through the publisher components and load them into the ExceptionManagementSettings.

            For Each node In section.ChildNodes
               If node.Name = PUBLISHER_NODENAME Then
                  ' Initialize a new PublisherSettings.
                  publisherSettings = New publisherSettings()

                  ' Get a collection of all the attributes.
                  nodeAttributes = node.Attributes

                  ' Remove the mode attribute from the node and set its value in PublisherSettings.
                  currentAttribute = nodeAttributes.RemoveNamedItem(PUBLISHER_MODE)
                  If Not (currentAttribute Is Nothing) AndAlso currentAttribute.Value.ToUpper(CultureInfo.InvariantCulture) = "OFF" Then
                     publisherSettings.Mode = PublisherMode.Off
                  End If

                  ' Remove the assembly attribute from the node and set its value in PublisherSettings.
                  currentAttribute = nodeAttributes.RemoveNamedItem(PUBLISHER_ASSEMBLY)
                  If Not (currentAttribute Is Nothing) Then
                     publisherSettings.AssemblyName = currentAttribute.Value
                  End If

                  ' Remove the type attribute from the node and set its value in PublisherSettings.
                  currentAttribute = nodeAttributes.RemoveNamedItem(PUBLISHER_TYPE)
                  If Not (currentAttribute Is Nothing) Then
                     publisherSettings.TypeName = currentAttribute.Value
                  End If

                  ' Remove the exceptionFormat attribute from the node and set its value in PublisherSettings.
                  currentAttribute = nodeAttributes.RemoveNamedItem(PUBLISHER_EXCEPTIONFORMAT)
                  If Not (currentAttribute Is Nothing) AndAlso currentAttribute.Value.ToUpper(CultureInfo.InvariantCulture) = "XML" Then
                     publisherSettings.ExceptionFormat = PublisherFormat.Xml
                  End If

                  ' Remove the include attribute from the node and set its value in PublisherSettings.
                  currentAttribute = nodeAttributes.RemoveNamedItem(PUBLISHER_INCLUDETYPES)
                  If Not (currentAttribute Is Nothing) Then
                     publisherSettings.IncludeTypes = LoadTypeFilter(currentAttribute.Value.Split(EXCEPTION_TYPE_DELIMITER))
                  End If

                  ' Remove the exclude attribute from the node and set its value in PublisherSettings.
                  currentAttribute = nodeAttributes.RemoveNamedItem(PUBLISHER_EXCLUDETYPES)
                  If Not (currentAttribute Is Nothing) Then
                     publisherSettings.ExcludeTypes = LoadTypeFilter(currentAttribute.Value.Split(EXCEPTION_TYPE_DELIMITER))
                  End If

                  ' Loop through any other attributes and load them into OtherAttributes.
                  j = nodeAttributes.Count - 1
                  For i = 0 To j '
                     publisherSettings.AddOtherAttributes(nodeAttributes.Item(i).Name, nodeAttributes.Item(i).Value)
                  Next i

                  ' Add the PublisherSettings to the publishers collection.
                  settings.Publishers.Add(publisherSettings)
               End If
            Next node

            ' Remove extra allocated space of the ArrayList of Publishers. 
            settings.Publishers.TrimToSize()

            ' Return the ExceptionManagementSettings loaded with the values from the config file.
            Return settings

         Catch exc As Exception
            Throw New ConfigurationErrorsException(m_resourceManager.GetString("RES_ExceptionLoadingConfiguration"), exc, section)

         End Try

      End Function 'Create

        ' Creates TypeFilter with type information from the string array of type names.
        ' Parameters:
        ' -rawFilter - String array containing names of types to be included in the filter.
        ' Returns: TypeFilter object containing type information.
        Private Function LoadTypeFilter(ByVal rawFilter As String()) As TypeFilter
            ' Initialize filter
            Dim m_typeFilter As TypeFilter = New TypeFilter()

            ' Verify information was passed in
            If Not rawFilter Is Nothing Then
                Dim m_exceptionTypeInfo As TypeInfo
                Dim i As Short

                'Loop through the string array
            For i = 0 To CShort(rawFilter.GetLength(0) - 1)
                ' If the wildcard character "*" exists set the TypeFilter to accept all types.
                If rawFilter(i) = "*" Then
                    m_typeFilter.AcceptAllTypes = True

                Else
                    Try
                        If rawFilter(i).Length > 0 Then
                            'Create the TypeInfo class
                            m_exceptionTypeInfo = New TypeInfo()

                            'If the string starts with a "+"
                            If rawFilter(i).Trim().StartsWith("+") Then
                                'Set the TypeInfo class to include subclasses
                                m_exceptionTypeInfo.IncludeSubClasses = True

                                'Get the Type class from the filter privided.
                                m_exceptionTypeInfo.ClassType = Type.GetType(rawFilter(i).Trim().TrimStart(Convert.ToChar("+")), True)

                            Else
                                ' Set the TypeInfo class not to include subclasses
                                m_exceptionTypeInfo.IncludeSubClasses = False
                                ' Get the Type class from the filter privided.
                                m_exceptionTypeInfo.ClassType = Type.GetType(rawFilter(i).Trim(), True)

                            End If

                            ' Add the TypeInfo class to the TypeFilter
                            m_typeFilter.Types.Add(m_exceptionTypeInfo)

                        End If

                    Catch e As TypeLoadException
                        ' If the Type could not be created throw a configuration exception.
                        ExceptionManager.PublishInternalException(New ConfigurationErrorsException(m_resourceManager.GetString("RES_EXCEPTION_LOADING_CONFIGURATION"), e), Nothing)
                    End Try
                End If
            Next
            End If

            Return m_typeFilter

        End Function 'LoadTypeFilter
   End Class 'ExceptionManagerSectionHandler
#End Region