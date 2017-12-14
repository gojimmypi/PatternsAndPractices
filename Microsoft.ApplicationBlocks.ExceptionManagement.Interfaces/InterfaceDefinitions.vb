
#Region "Publishing Interfaces"
' Interface to publish exception information.  All exception information is passed as the chain of exception objects.
Public Interface IExceptionPublisher
    ' Method used to publish exception information and additional information.
    ' Parameters:
    ' -exception - The exception object whose information should be published. 
    ' -additionalInfo - A collection of additional data that should be published along with the exception information. 
    ' -configSettings - A collection of name/value attributes specified in the config settings. 
    Sub Publish(ByVal exception As Exception, ByVal additionalInfo As NameValueCollection, ByVal configSettings As NameValueCollection)
End Interface 'IPublishException

' Interface to publish exception information.  All exception information is passed as XML.
Public Interface IExceptionXmlPublisher
    ' Method used to publish exception information and any additional information in XML.
    ' Parameters:
    ' -exceptionInfo - An XML Document containing the all exception information. 
    ' -configSettings - A collection of name/value attributes specified in the config settings. 
    Sub Publish(ByVal exceptionInfo As XmlDocument, ByVal configSettings As NameValueCollection)
End Interface 'IPublishXMLException

#End Region