' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' InterfaceDefinitions.vb
'
' This file contains the definition for the interfaces used on the application block
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
Imports [SC] = System.Configuration
Imports System.Collections.Specialized
Imports System.Runtime.InteropServices
Imports System.Xml


Namespace Microsoft.ApplicationBlocks.ConfigurationManagement

#Region "Delegates"

    ' <summary>
    ' Used the specify the configuration have been changed on the storage
    ' </summary>
    <Serializable()> _
    Public Delegate Sub ConfigurationChanged(ByVal storageProvider As IConfigurationStorageReader, _
                                ByVal sectionName As String)

#End Region

#Region "Provider Interfaces"

    ' <summary>
    ' Allows end users to implement their own configuration management storage.
    ' All storage providers must implement this interface
    ' </summary>
    <ComVisible(False)> _
    Public Interface IConfigurationStorageReader
        ' <summary>
        ' Inits the config provider 
        ' </summary>
        ' <param name="sectionName"></param>
        ' <param name="configStorageParameters">Configuration parameters</param>
        ' <param name="dataProtection">Data protection interface.</param>param>
        Sub Init(ByVal sectionName As String, ByVal configStorageParameters As ListDictionary, _
                    ByVal dataProtection As IDataProtection)

        ' <summary>
        ' Returns an XML representation of the configuration data
        ' </summary>
        ' <returns></returns>
        Function Read() As XmlNode

        ' <summary>
        ' Event to indicate a change in the configuration storage
        ' </summary>
        Event ConfigChanges As ConfigurationChanged
        '

        ' <summary>
        ' Whether the provider has been initialized
        ' </summary>
        ReadOnly Property Initialized() As Boolean

    End Interface 'IConfigurationStorageReader 
    ' <summary>
    ' Implemented by configuration providers to allow for writeable storage of configuration 
    ' information
    ' </summary>
    <ComVisible(False)> _
    Public Interface IConfigurationStorageWriter
        Inherits IConfigurationStorageReader

        ' <summary>
        '  This method writes the xml-serialized object to the underlying storage 
        ' </summary>
        Sub Write(ByVal value As XmlNode)
    End Interface

    ' <summary>
    ' Implemented by custom section handlers in order to allow a writeable implementation
    ' </summary>
    <ComVisible(False)> _
    Public Interface IConfigurationSectionHandlerWriter
        Inherits SC.IConfigurationSectionHandler

        ' <summary>
        ' This method converts the public fields and read/write properties of an object into XML.
        ' </summary>
        Function Serialize(ByVal value As Object) As XmlNode
    End Interface 'IConfigurationSectionHandlerWriter
#End Region

#Region "DataProtection Interfaces"
    ' <summary>
    ' Implemented by data protection providers to allow for encrypt information
    ' </summary>
    <ComVisible(False)> _
    Public Interface IDataProtection
        Inherits IDisposable

        ' <summary>
        ' Inits the data protection provider 
        ' </summary>
        ' <param name="dataProtectionParameters">Data protection parameters</param>
        Sub Init(ByVal dataProtectionParameters As ListDictionary)

        ' <summary>
        ' Encrypts a raw of bytes that represents a plain text
        ' </summary>
        ' <param name="plainText">plain text</param>
        ' <returns>a cipher value</returns>
        Function Encrypt(ByVal plainText() As Byte) As Byte()

        ' <summary>
        ' Decrypts a cipher value
        ' </summary>
        ' <param name="cipherText">cipher text</param>
        ' <returns>a raw of bytes that represents a plain text</returns>
        Function Decrypt(ByVal cipherText() As Byte) As Byte()

        ' <summary>
        ' Computes a hash
        ' </summary>
        ' <param name="plainText">plain text</param>
        ' <returns>hash data</returns>
        Function ComputeHash(ByVal plainText() As Byte) As Byte()
    End Interface 'IDataProtection

#End Region

End Namespace 'Microsoft.ApplicationBlocks.ConfigurationManagement