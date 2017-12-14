' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' XmlHashtableSectionHandler.vb
'
' A sample section handler that converts a hashtable so it can be used on
' any application that uses a Hashtable to handle the configuration.
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
Imports System.Collections
Imports [SC] = System.Configuration
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Threading

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

' <summary>
' A sample section handler which stores a hashtable on the confguration
' storage.
' </summary>
Friend Class XmlHashtableSectionHandler
    Implements MABCM.IConfigurationSectionHandlerWriter
    Private Shared _xmlSerializer As XmlSerializer

    Shared Sub New()
        _xmlSerializer = New XmlSerializer(GetType(XmlSerializableHashtable))
    End Sub 'New

#Region "Implementation of IConfigurationSectionHandler"

    Function Create(ByVal parent As Object, ByVal configContext As Object, ByVal section As XmlNode) _
            As Object Implements SC.IConfigurationSectionHandler.Create

        If section.ChildNodes.Count = 0 Then
            Return New Hashtable
        End If
        Monitor.Enter(Me)
        Try
            Dim xmlHt As XmlSerializableHashtable = _
                    CType(_xmlSerializer.Deserialize(New XmlNodeReader(section)), XmlSerializableHashtable)
            Return xmlHt.InnerHashtable
        Catch ex As Exception
            Throw New SC.ConfigurationErrorsException( _
                Resource.ResourceManager("RES_ExceptionCantDeserializeHashtable"), ex)
        Finally
            Monitor.Exit(Me)
        End Try
    End Function 'IConfigurationSectionHandler.Create

#End Region

#Region "Implementation of IConfigurationSectionHandlerWriter"

    Function Serialize(ByVal value As Object) As XmlNode Implements IConfigurationSectionHandlerWriter.Serialize
        If Not TypeOf value Is Hashtable Then
            Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidConfigurationInstance"))
        End If

        '  get the stringwriter instance
        Dim sw As New StringWriter

        '  serialize a new XmlSerializableHashtable...
        '  use its constructor to pass in the actual (non-serializable) hashtable we've been given
        _xmlSerializer.Serialize(sw, New XmlSerializableHashtable(CType(value, Hashtable)))

        '  put the xml-serialized text into an xml doc
        Dim doc As New XmlDocument
        doc.LoadXml(sw.ToString())

        '  return it
        Return doc.DocumentElement
    End Function 'IConfigurationSectionHandlerWriter.Serialize

#End Region

End Class 'XmlHashtableSectionHandler
