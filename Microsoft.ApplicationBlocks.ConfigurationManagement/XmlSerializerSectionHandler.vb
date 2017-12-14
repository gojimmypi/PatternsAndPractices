' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' XmlSerializerSectionHandler.vb
'
' A sample section handler that uses xmlserializer to store any xml serializable
' class on the configuration.
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
' 
' </summary>
Friend Class XmlSerializerSectionHandler
    Implements MABCM.IConfigurationSectionHandlerWriter

    Private _xmlSerializerCache As New Hashtable

    Public Sub New()
    End Sub 'New 

#Region "Implementation of IConfigurationSectionHandler"

    Function Create(ByVal parent As Object, ByVal configContext As Object, ByVal section As XmlNode) _
            As Object Implements SC.IConfigurationSectionHandler.Create

        Dim xmlSerializerSection As XmlNode = section
        Dim typeName As String = xmlSerializerSection.Attributes("type").Value
        Dim classType As Type = Type.GetType(typeName, True)

        Dim xs As XmlSerializer = CType(_xmlSerializerCache(classType), XmlSerializer) '
        If xs Is Nothing Then
            xs = New XmlSerializer(classType)
            _xmlSerializerCache(classType) = xs
        End If
        Monitor.Enter(xs)
        Try
            Return xs.Deserialize(New XmlNodeReader(xmlSerializerSection.ChildNodes(0)))
        Finally
            Monitor.Exit(xs)
        End Try
    End Function 'IConfigurationSectionHandler.Create

#End Region

#Region "Implementation of IConfigurationSectionHandlerWriter"

    Function Serialize(ByVal value As Object) As XmlNode Implements MABCM.IConfigurationSectionHandlerWriter.Serialize
        Dim xs As XmlSerializer = CType(_xmlSerializerCache(value.GetType()), XmlSerializer) '
        If xs Is Nothing Then
            xs = New XmlSerializer(value.GetType())
            _xmlSerializerCache(value.GetType()) = xs
        End If

        Dim sw As New StringWriter(System.Globalization.CultureInfo.CurrentUICulture)
        Dim xmlTw As New XmlTextWriter(sw)
        xmlTw.WriteStartElement("XmlSerializerSection")
        xmlTw.WriteAttributeString("type", value.GetType().FullName + ", " + value.GetType().Assembly.FullName)
        Monitor.Enter(xs)
        Try
            xs.Serialize(xmlTw, value)
        Finally
            Monitor.Exit(xs)
        End Try

        xmlTw.WriteEndElement()
        xmlTw.Flush()

        Dim doc As New XmlDocument
        doc.LoadXml(sw.ToString())
        Return doc.DocumentElement
    End Function 'IConfigurationSectionHandlerWriter.Serialize
#End Region

End Class 'XmlSerializerSectionHandler
