Imports [SC] = System.Configuration

#Region "ExceptionManager Class"
   ' The Exception Manager class manages the publishing of exception information based on settings in the configuration file.
   Public NotInheritable Class ExceptionManager

      ' Private constructor to restrict an instance of this class from being created.
      Private Sub New()
      End Sub 'New

      ' Member variable declarations
      Private Shared EXCEPTIONMANAGER_NAME As String = GetType(ExceptionManager).Name
      Private Shared EXCEPTIONMANAGEMENT_CONFIG_SECTION As String = "exceptionManagement"

      ' Resource Manager for localized text.
      Private Shared m_resourceManager As New ResourceManager(GetType(ExceptionManager).Namespace + ".ExceptionManagerText", [Assembly].GetAssembly(GetType(ExceptionManager)))

      ' Static method to publish the exception information.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      Public Overloads Shared Sub Publish(ByVal exception As Exception)
         ExceptionManager.Publish(exception, Nothing)
      End Sub 'Publish

      ' Static method to publish the exception information and any additional information.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
      Public Overloads Shared Sub Publish(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection)

         Dim config As ExceptionManagementSettings
         Dim Publisher As PublisherSettings

         Try

            ' Create the Additional Information collection if it does not exist.
            If additionalInfo Is Nothing Then additionalInfo = New NameValueCollection()

                Try
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".MachineName", Environment.MachineName)
                Catch e As SecurityException
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".MachineName", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED"))
                Catch
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".MachineName", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION"))
                End Try

                Try
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".TimeStamp", DateTime.Now.ToString())
                Catch e As SecurityException
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".TimeStamp", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED"))
                Catch
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".TimeStamp", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION"))
                End Try

                Try
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".FullName", Reflection.Assembly.GetExecutingAssembly().FullName)

                Catch e As SecurityException
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".FullName", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED"))
                Catch
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".FullName", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION"))
                End Try

                Try
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".AppDomainName", AppDomain.CurrentDomain.FriendlyName)
                Catch e As SecurityException
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".AppDomainName", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED"))
                Catch
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".AppDomainName", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION"))
                End Try

                Try
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".ThreadIdentity", Thread.CurrentPrincipal.Identity.Name)
                Catch e As SecurityException
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".ThreadIdentity", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED"))
                Catch
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".ThreadIdentity", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION"))
                End Try

                Try
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".WindowsIdentity", WindowsIdentity.GetCurrent().Name)
                Catch e As SecurityException
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".WindowsIdentity", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_PERMISSION_DENIED"))
                Catch
                    additionalInfo.Add(EXCEPTIONMANAGER_NAME + ".WindowsIdentity", m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_INFOACCESS_EXCEPTION"))
                End Try


            ' Check for any settings in config file.
            If SC.ConfigurationManager.GetSection(EXCEPTIONMANAGEMENT_CONFIG_SECTION) Is Nothing Then

                ' Publish the exception and additional information to the default publisher if no settings are present.
                PublishToDefaultPublisher(exception, additionalInfo)
            Else
                ' Get settings from config file
                config = CType(SC.ConfigurationManager.GetSection(EXCEPTIONMANAGEMENT_CONFIG_SECTION), ExceptionManagementSettings)

                ' If the mode is not equal to "off" call the Publishers, otherwise do nothing.
                If config.Mode = ExceptionManagementMode.On Then
                    ' If no publishers are specified, use the default publisher.
                    If config.Publishers Is Nothing OrElse config.Publishers.Count = 0 Then
                        ' Publish the exception and additional information to the default publisher if no mode is specified.
                        PublishToDefaultPublisher(exception, additionalInfo)
                    Else

                        ' Loop through the publisher components specified in the config file.
                        For Each Publisher In config.Publishers  '

                            ' Call the Publisher component specified in the config file.
                            Try
                                ' Verify the publishers mode is not set to "OFF".
                                ' This publisher will be called even if the mode is not specified.  
                                ' The mode must explicitly be set to OFF to not be called.
                                If Publisher.Mode = PublisherMode.On Then
                                    If exception Is Nothing OrElse Not Publisher.IsExceptionFiltered(exception.GetType()) Then
                                        ' Publish the exception and any additional information
                                        PublishToCustomPublisher(exception, additionalInfo, Publisher)
                                    End If
                                End If
                                ' Catches any failure to call a custom publisher.
                            Catch e As Exception
                                ' Publish the exception thrown within the ExceptionManager.
                                PublishInternalException(e, Nothing)

                                ' Publish the original exception and additional information to the default publisher.
                                PublishToDefaultPublisher(exception, additionalInfo)
                            End Try
                        Next Publisher ' End Catch block.
                    End If
                End If ' End foreach loop through publishers.
            End If '

         Catch e As exception '
            ' Publish the exception thrown when trying to call the custom publisher.
            PublishInternalException(e, Nothing)

            ' Publish the original exception and additional information to the default publisher.
            PublishToDefaultPublisher(exception, additionalInfo)
         End Try
      End Sub 'Publish


      ' Private static helper method to publish the exception information to a custom publisher.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
      ' -publisher - The PublisherSettings that contains the values of the publishers configuration. 
      Private Shared Sub PublishToCustomPublisher(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection, ByVal publisher As PublisherSettings)

         Dim XMLPublisher As IExceptionXmlPublisher
         Dim m_Publisher As IExceptionPublisher

         Try
            ' Check if the exception format is "xml".
            If publisher.ExceptionFormat = PublisherFormat.Xml Then
               ' If it is load the IExceptionXmlPublisher interface on the custom publisher.
               ' Instantiate the class
                XMLPublisher = CType(Activate(publisher.AssemblyName, publisher.TypeName), IExceptionXmlPublisher)

               ' Publish the exception and any additional information
               XMLPublisher.Publish(SerializeToXML(exception, additionalInfo), publisher.OtherAttributes)
            ' Otherwise load the IExceptionPublisher interface on the custom publisher.
            Else

               ' Instantiate the class
                m_Publisher = CType(Activate(publisher.AssemblyName, publisher.TypeName), IExceptionPublisher)

               ' Publish the exception and any additional information
               m_Publisher.Publish(exception, additionalInfo, publisher.OtherAttributes)
            End If
         Catch e As exception
            Dim publisherException As CustomPublisherException = New CustomPublisherException(m_resourceManager.GetString("RES_CUSTOM_PUBLISHER_FAILURE_MESSAGE"), publisher.AssemblyName, publisher.TypeName, publisher.ExceptionFormat, e)
            publisherException.AdditionalInformation.Add(publisher.OtherAttributes)

            Throw publisherException
         End Try
      End Sub 'PublishToCustomPublisher

      ' Private static helper method to publish the exception information to the default publisher.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
      Private Shared Sub PublishToDefaultPublisher(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection)
         ' Get the Default Publisher
         Dim Publisher As New DefaultPublisher()

         ' Publish the exception and any additional information
         Publisher.Publish(exception, additionalInfo, Nothing)
      End Sub 'PublishToDefaultPublisher

      ' Private static helper method to publish the exception information to the default publisher.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
      Protected Friend Shared Sub PublishInternalException(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection)
         ' Get the Default Publisher
         Dim Publisher As New DefaultPublisher("Application", m_resourceManager.GetString("RES_EXCEPTIONMANAGER_INTERNAL_EXCEPTIONS"))

         ' Publish the exception and any additional information
         Publisher.Publish(exception, additionalInfo, Nothing)
      End Sub 'PublishInternalException

      ' Private helper function to assist in run-time activations. Returns
      ' an object from the specified assembly and type.
      ' Parameters:
      ' -assembly - Name of the assembly file (w/out extension) 
      ' -typeName - Name of the type to create 
      ' Returns: Instance of the type specified in the input parameters. 
      Private Shared Function Activate(ByVal [assembly] As String, ByVal typeName As String) As Object
         Return AppDomain.CurrentDomain.CreateInstanceAndUnwrap([assembly], typeName)
      End Function 'Activate

      ' Public static helper method to serialize the exception information into XML.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
    Public Shared Function SerializeToXML(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection) As XmlDocument

        Try
            ' Variables representing the XmlElement names.
            Dim xmlNodeName_ROOT As String = m_resourceManager.GetString("RES_XML_ROOT")
            Dim xmlNodeName_ADDITIONAL_INFORMATION As String = m_resourceManager.GetString("RES_XML_ADDITIONAL_INFORMATION")
            Dim xmlNodeName_EXCEPTION As String = m_resourceManager.GetString("RES_XML_EXCEPTION")
            Dim xmlNodeName_STACK_TRACE As String = m_resourceManager.GetString("RES_XML_STACK_TRACE")

            ' Create a new XmlDocument.
            Dim xmlDoc As New XmlDocument()

            ' Create the root node.
            Dim m_root As XmlElement = xmlDoc.CreateElement(xmlNodeName_ROOT)
            xmlDoc.AppendChild(m_root)

            ' Variables to hold values while looping through the exception chain.
            Dim element As XmlElement
            Dim exceptionAddInfoElement As XmlElement
            Dim stackTraceElement As XmlElement
            Dim stackTraceText As XmlText
            Dim attribute As XmlAttribute

            Dim i As String
            Dim currentException As exception
            ' Temp variable to hold InnerException object during the loop.
            Dim parentElement As XmlElement = Nothing ' Temp variable to hold the parent exception node during the loop.
            Dim aryPublicProperties As PropertyInfo()
            Dim currentAdditionalInfo As NameValueCollection
            Dim p As PropertyInfo

            ' Check if the collection has values.
            If Not (additionalInfo Is Nothing) AndAlso additionalInfo.Count > 0 Then

                ' Create the element for the collection.
                element = xmlDoc.CreateElement(xmlNodeName_ADDITIONAL_INFORMATION)

                ' Loop through the collection and add the values as attributes on the element.

                For Each i In additionalInfo
                    attribute = xmlDoc.CreateAttribute(i.Replace(" ", "_"))
                    attribute.Value = additionalInfo.Get(i)
                    element.Attributes.Append(attribute)
                Next i

                ' Add the element to the root.
                m_root.AppendChild(element)
            End If

   


            If exception Is Nothing Then
                ' Create an empty exception element.
                element = xmlDoc.CreateElement(xmlNodeName_EXCEPTION)

                ' Append to the root node.
                m_root.AppendChild(element)
            Else
                currentException = exception 'Temp variable to hold InnerException object during the loop.
                'Loop through each exception class in the chain of exception objects and record its information
                Do
                    ' Create the exception element.
                    element = xmlDoc.CreateElement(xmlNodeName_EXCEPTION)

                    ' Add the exceptionType as an attribute.
                    attribute = xmlDoc.CreateAttribute("ExceptionType")
                    attribute.Value = currentException.GetType().FullName
                    element.Attributes.Append(attribute)

                    ' Loop through the public properties of the exception object and record their value.
                    aryPublicProperties = currentException.GetType().GetProperties() '

                    For Each p In aryPublicProperties
                        ' Do not log information for the InnerException or StackTrace. This information is 
                        ' captured later in the process.
                        If p.Name <> "InnerException" And p.Name <> "StackTrace" Then
                            ' Only record properties whose value is not null.
                            If Not (p.GetValue(currentException, Nothing) Is Nothing) Then
                                ' Check if the property is AdditionalInformation and the exception type is a BaseApplicationException.
                                If p.Name = "AdditionalInformation" And TypeOf currentException Is BaseApplicationException Then
                                    ' Verify the collection is not null.
                                    If Not (p.GetValue(currentException, Nothing) Is Nothing) Then
                                        ' Cast the collection into a local variable.
                                        currentAdditionalInfo = CType(p.GetValue(currentException, Nothing), NameValueCollection)

                                        ' Verify the collection has values.
                                        If currentAdditionalInfo.Count > 0 Then
                                            ' Create element.
                                            exceptionAddInfoElement = xmlDoc.CreateElement(xmlNodeName_ADDITIONAL_INFORMATION)

                                            ' Loop through the collection and add values as attributes.
                                            For Each i In currentAdditionalInfo
                                                attribute = xmlDoc.CreateAttribute(i.Replace(" ", "_"))
                                                attribute.Value = currentAdditionalInfo.Get(i)
                                                exceptionAddInfoElement.Attributes.Append(attribute)
                                            Next i

                                            element.AppendChild(exceptionAddInfoElement)
                                        End If
                                    End If
                                    ' Otherwise just add the ToString() value of the property as an attribute.
                                Else
                                    attribute = xmlDoc.CreateAttribute(p.Name)
                                    attribute.Value = p.GetValue(currentException, Nothing).ToString()
                                    element.Attributes.Append(attribute)
                                End If
                            End If
                        End If
                    Next p

                    ' Record the StackTrace within a separate element.
                    If Not (currentException.StackTrace Is Nothing) Then

                        ' Create Stack Trace Element.
                        stackTraceElement = xmlDoc.CreateElement(xmlNodeName_STACK_TRACE)

                        stackTraceText = xmlDoc.CreateTextNode(currentException.StackTrace.ToString())

                        stackTraceElement.AppendChild(stackTraceText)

                        element.AppendChild(stackTraceElement)
                    End If

                    ' Check if this is the first exception in the chain.
                    If parentElement Is Nothing Then
                        ' Append to the root node.
                        m_root.AppendChild(element)
                    Else
                        ' Append to the parent exception object in the exception chain.
                        parentElement.AppendChild(element)
                    End If

                    ' Reset the temp variables.
                    parentElement = element
                    currentException = currentException.InnerException
                Loop While Not (currentException Is Nothing)
            End If ' Continue looping until we reach the end of the exception chain.

            ' Return the XmlDocument.
            Return xmlDoc

        Catch e As exception
            Throw New SerializationException(m_resourceManager.GetString("RES_EXCEPTIONMANAGEMENT_XMLSERIALIZATION_EXCEPTION"), e)
        End Try

    End Function 'SerializeToXML
End Class 'ExceptionManager

#End Region

#Region "DefaultPublisher Class"

   ' Component used as the default publishing component if one is not specified in the config file.
   Public NotInheritable Class DefaultPublisher
      Implements IExceptionPublisher

      ' Default Constructor.
      Public Sub New()
      End Sub 'New 

      ' Constructor allowing the log name and application names to be set.
      ' Parameters:
      ' -logName - The name of the log for the DefaultPublisher to use. 
      ' -applicationName - The name of the application.  This is used as the Source name in the event log. 
      Public Sub New(ByVal logName As String, ByVal applicationName As String)
         Me.logName = logName
         Me.applicationName = applicationName
      End Sub 'New

      Private Shared m_resourceManager As New ResourceManager(GetType(ExceptionManager).Namespace + ".ExceptionManagerText", [Assembly].GetAssembly(GetType(ExceptionManager)))

      ' Member variable declarations
      Private logName As String = "Application"
      Private applicationName As String = m_resourceManager.GetString("RES_EXCEPTIONMANAGER_PUBLISHED_EXCEPTIONS")
      Private TEXT_SEPARATOR As String = "*********************************************"
	
      ' Method used to publish exception information and additional information.
      ' Parameters:
      ' -exception - The exception object whose information should be published. 
      ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
      ' -configSettings - A collection of any additional attributes provided in the config settings for the custom publisher.
      Public Sub Publish(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection, ByVal configSettings As NameValueCollection) Implements IExceptionPublisher.Publish

         ' Create StringBuilder to maintain publishing information.
         Dim strInfo As New StringBuilder()
         Dim i As String
         Dim currentException As exception
         Dim intExceptionCount As Integer = 1 ' Count variable to track the number of exceptions in the chain.
         Dim aryPublicProperties As PropertyInfo()
         Dim currentAdditionalInfo As NameValueCollection
         Dim p As PropertyInfo
         Dim j As Integer
         Dim k As Integer

         ' Load Config values if they are provided.
         If Not (configSettings Is Nothing) Then
            If Not (configSettings("applicationName") Is Nothing) AndAlso configSettings("applicationName").Length > 0 Then
               applicationName = configSettings("applicationName")
            End If
            If Not (configSettings("logName") Is Nothing) AndAlso configSettings("logName").Length > 0 Then
               logName = configSettings("logName")
            End If
         End If

         ' Verify that the Source exists before gathering exception information.
         VerifyValidSource()

         ' Record the contents of the AdditionalInfo collection.
         If Not (additionalInfo Is Nothing) Then

            ' Record General information.
            strInfo.AppendFormat("{0}General Information {0}{1}{0}Additional Info:", Environment.NewLine, TEXT_SEPARATOR)

            For Each i In additionalInfo
               strInfo.AppendFormat("{0}{1}: {2}", Environment.NewLine, i, additionalInfo.Get(i))
            Next i
         End If

         If exception Is Nothing Then
            strInfo.AppendFormat("{0}{0}No Exception object has been provided..{0}", Environment.NewLine)
         Else
            ' Loop through each exception class in the chain of exception objects.

            ' Temp variable to hold InnerException object during the loop.
            currentException = exception '

            Do
               ' Write title information for the exception object.
               strInfo.AppendFormat("{0}{0}{1}) Exception Information{0}{2}", Environment.NewLine, intExceptionCount.ToString(), TEXT_SEPARATOR)
               strInfo.AppendFormat("{0}Exception Type: {1}", Environment.NewLine, currentException.GetType().FullName)

               ' Loop through the public properties of the exception object and record their value.
               aryPublicProperties = currentException.GetType().GetProperties()  '

               For Each p In aryPublicProperties
                  ' Do not log information for the InnerException or StackTrace. This information is 
                  ' captured later in the process.
                  If p.Name <> "InnerException" And p.Name <> "StackTrace" Then
                     If p.GetValue(currentException, Nothing) Is Nothing Then
                        strInfo.AppendFormat("{0}{1}: NULL", Environment.NewLine, p.Name)
                     Else
                        ' Loop through the collection of AdditionalInformation if the exception type is a BaseApplicationException.
                        If p.Name = "AdditionalInformation" And TypeOf currentException Is BaseApplicationException Then
                           ' Verify the collection is not null.
                           If Not (p.GetValue(currentException, Nothing) Is Nothing) Then
                              ' Cast the collection into a local variable.
                              currentAdditionalInfo = CType(p.GetValue(currentException, Nothing), NameValueCollection)

                              ' Check if the collection contains values.
                              If currentAdditionalInfo.Count > 0 Then
                                 strInfo.AppendFormat("{0}AdditionalInformation:", Environment.NewLine)

                                 ' Loop through the collection adding the information to the string builder.
                                 k = currentAdditionalInfo.Count - 1
                                 For j = 0 To k
                                    strInfo.AppendFormat("{0}{1}: {2}", Environment.NewLine, currentAdditionalInfo.GetKey(j), currentAdditionalInfo(j))
                                 Next
                              End If
                           End If
                        ' Otherwise just write the ToString() value of the property.
                        Else
                           strInfo.AppendFormat("{0}{1}: {2}", Environment.NewLine, p.Name, p.GetValue(currentException, Nothing))
                        End If
                     End If
                  End If
               Next p

               ' Record the StackTrace with separate label.
               If Not (currentException.StackTrace Is Nothing) Then '
                  strInfo.AppendFormat("{0}{0}StackTrace Information{0}{1}", Environment.NewLine, TEXT_SEPARATOR)
                  strInfo.AppendFormat("{0}{1}", Environment.NewLine, currentException.StackTrace)
               End If

               ' Reset the temp exception object and iterate the counter.
               currentException = currentException.InnerException
               intExceptionCount += 1
            Loop While Not (currentException Is Nothing)
         End If '

         ' Write the entry to the event log.   
         WriteToLog(strInfo.ToString(), EventLogEntryType.Error)
      End Sub 'Publish

      ' Helper function to write an entry to the Event Log.
      ' Parameters:
      ' -entry - The entry to enter into the Event Log. 
      ' -type - The EventLogEntryType to be used when the entry is logged to the Event Log. 
      Private Sub WriteToLog(ByVal entry As String, ByVal type As EventLogEntryType)
            Try
                ' Write the entry to the Event Log.
                EventLog.WriteEntry(applicationName, entry, type)
            Catch e As SecurityException
                Throw New SecurityException(String.Format(m_resourceManager.GetString("RES_DEFAULTPUBLISHER_EVENTLOG_DENIED"), applicationName), e)
            End Try
      End Sub 'WriteToLog

        Private Sub VerifyValidSource()
            Try
                If Not EventLog.SourceExists(applicationName) Then
                    EventLog.CreateEventSource(applicationName, logName)
                End If
            Catch e As SecurityException
                Throw New SecurityException(String.Format(m_resourceManager.GetString("RES_DEFAULTPUBLISHER_EVENTLOG_DENIED"), applicationName), e)
            End Try
       End Sub

   End Class 'DefaultPublisher

#End Region