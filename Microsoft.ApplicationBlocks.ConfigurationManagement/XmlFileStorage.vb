' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' XmlFileStorage.vb
'
' This file contains a read-write storage implementation that uses an 
' xml file to save the configuration.
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
Imports System.Text
Imports System.Xml

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

Namespace Storage

    ' <summary>
    ' NOTE only need to implement IConfigurationStorageWriter interface, 
    ' that interface aggregates Reader AND Writer functionality.
    ' If we wished only a Reader, we would implement the lesser IConfigurationStorageReader interface.
    ' </summary>
    Friend Class XmlFileStorage
        Implements IConfigurationStorageWriter
#Region "Declare Variables"
        Private _applicationDocumentPath As String = Nothing
        Private _machineDocumentPath As String = Nothing
        Private _sectionName As String = Nothing
        Private _isSigned As Boolean = False
        Private _isEncrypted As Boolean = False
        Private _isRefreshOnChange As Boolean = False
        Private _isInitOk As Boolean = False
        Private _fileWatcherApp As FileSystemWatcher = Nothing
        Private _dataProtection As IDataProtection
#End Region

#Region "Constructor"

        Public Sub New()
        End Sub 'New

#End Region

#Region "IConfigurationStorageReader implementation"

        ' <summary>
        ' Inits the provider
        ' </summary>
        Public Sub Init(ByVal sectionName As String, ByVal configStorageParameters As ListDictionary, _
                    ByVal dataProtection As IDataProtection) Implements IConfigurationStorageReader.Init
            'Inits the provider properties
            _sectionName = sectionName

            _applicationDocumentPath = CType(configStorageParameters("path"), String) '
            If ApplicationDocumentPath Is Nothing Then
                _applicationDocumentPath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
            Else
                _applicationDocumentPath = Path.GetFullPath(_applicationDocumentPath)
            End If

            Dim signedString As String = CType(configStorageParameters("signed"), String) '
            If Not (signedString Is Nothing) AndAlso signedString.Length <> 0 Then
                _isSigned = Boolean.Parse(signedString)
            End If
            Dim encryptedString As String = CType(configStorageParameters("encrypted"), String) '
            If Not (encryptedString Is Nothing) AndAlso encryptedString.Length <> 0 Then
                _isEncrypted = Boolean.Parse(encryptedString)
            End If
            Dim refreshOnChangeString As String = CType(configStorageParameters("refreshOnChange"), String) '
            If Not (refreshOnChangeString Is Nothing) AndAlso refreshOnChangeString.Length <> 0 Then
                _isRefreshOnChange = Boolean.Parse(refreshOnChangeString)
            End If
            Me._dataProtection = dataProtection

            If ((_isSigned OrElse _isEncrypted) AndAlso _dataProtection Is Nothing) Then
                Throw New SC.ConfigurationErrorsException( _
                    Resource.ResourceManager("RES_ExceptionInvalidDataProtectionConfiguration", _sectionName))
            End If

            '  here we set up a file-watch that, if refreshOnChange is enabled, will fire an event when config 
            '  changes and cause cache to flush
            If _isRefreshOnChange Then
                If Path.IsPathRooted(_applicationDocumentPath) Then
                    _fileWatcherApp = New FileSystemWatcher(Path.GetDirectoryName(_applicationDocumentPath), Path.GetFileName(_applicationDocumentPath))
                Else
                    _fileWatcherApp = New FileSystemWatcher(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, Path.GetFileName(_applicationDocumentPath))
                End If

                _fileWatcherApp.EnableRaisingEvents = True
                AddHandler _fileWatcherApp.Changed, AddressOf OnChanged
            End If
            _isInitOk = True
        End Sub 'Init

        Public ReadOnly Property Initialized() As Boolean Implements IConfigurationStorageReader.Initialized
            Get
                Return _isInitOk
            End Get
        End Property

        ' <summary>
        ' Return a section node
        ' </summary>
        Public Function Read() As XmlNode Implements IConfigurationStorageReader.Read

            Dim xmlApplicationDocument As XmlDocument = New XmlDocument
            Try
                If File.Exists(_applicationDocumentPath) Then
                    LoadXmlFile(xmlApplicationDocument, _applicationDocumentPath)
                Else
                    Throw New SC.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionConfigurationFileNotFound", xmlApplicationDocument, _sectionName))
                End If
            Catch e As XmlException
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidConfigurationDocument"))
            End Try

            'Select the item nodes
            Dim sectionNode As XmlNode = xmlApplicationDocument.SelectSingleNode(("/configuration/" + SectionName))
            If sectionNode Is Nothing Then
                Return Nothing
            End If

            'Clone the XmlNode to prevent concurrency problems
            sectionNode = sectionNode.CloneNode(True)

            If _isSigned OrElse _isEncrypted Then
                Dim encryptedNode As XmlNode = sectionNode.SelectSingleNode("encryptedData")
                Dim sectionData As String = ""
                If Not (encryptedNode Is Nothing) Then
                    sectionData = encryptedNode.InnerXml
                End If
                Dim signatureNode As XmlNode = sectionNode.SelectSingleNode("signature")
                Dim xmlSignature As String = ""
                If Not (signatureNode Is Nothing) Then
                    xmlSignature = signatureNode.InnerXml
                End If
                If _isSigned Then
                    'Compute the hash
                    Dim hash As Byte() = _dataProtection.ComputeHash(Encoding.UTF8.GetBytes(sectionData))
                    Dim newHashString As String = Convert.ToBase64String(hash)

                    'Compare the hashes
                    If newHashString <> xmlSignature Then
                        Throw New SC.ConfigurationErrorsException( _
                                Resource.ResourceManager("RES_ExceptionSignatureCheckFailed", SectionName))
                    End If
                Else
                    If Not (xmlSignature Is Nothing) AndAlso xmlSignature.Length <> 0 Then
                        Throw New SC.ConfigurationErrorsException( _
                                Resource.ResourceManager("RES_ExceptionInvalidSignatureConfig", _sectionName))
                    End If
                End If
                If _isEncrypted Then
                    Dim encryptedBytes As Byte() = Nothing
                    Dim decryptedBytes As Byte() = Nothing
                    Try
                        Try
                            encryptedBytes = Convert.FromBase64String(sectionData)
                            decryptedBytes = _dataProtection.Decrypt(encryptedBytes)
                            sectionData = Encoding.UTF8.GetString(decryptedBytes)
                        Catch ex As Exception
                            Throw New SC.ConfigurationErrorsException( _
                                    Resource.ResourceManager("RES_ExceptionInvalidEncryptedString"), ex)
                        End Try
                    Finally
                        If Not (encryptedBytes Is Nothing) Then
                            Array.Clear(encryptedBytes, 0, encryptedBytes.Length)
                        End If
                        If Not (decryptedBytes Is Nothing) Then
                            Array.Clear(decryptedBytes, 0, decryptedBytes.Length)
                        End If
                    End Try
                End If

                Dim xmlDoc As New XmlDocument
                Dim newNode As XmlNode = xmlDoc.CreateElement(SectionName)
                newNode.InnerXml = sectionData
                xmlDoc.AppendChild(newNode)
                Return xmlDoc.FirstChild.FirstChild
            Else
                Return sectionNode.FirstChild
            End If
        End Function 'Read

        ' <summary>
        ' Event to indicate a change in the configuration storage
        ' </summary>
        Public Event ConfigChanges As ConfigurationChanged Implements IConfigurationStorageReader.ConfigChanges

#End Region

#Region "IConfigurationStorageWriter implementation"
        ' <summary>
        ' Writes a section on the XML document
        ' </summary>
        Public Sub Write(ByVal value As XmlNode) Implements IConfigurationStorageWriter.Write

            Dim xmlApplicationDocument As XmlDocument = New XmlDocument
            SyncLock Me.GetType()
                If File.Exists(_applicationDocumentPath) Then
                    Try
                        LoadXmlFile(xmlApplicationDocument, _applicationDocumentPath)
                    Catch e As XmlException
                        Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidConfigurationDocument"))
                    End Try
                Else
                    Dim configNode As XmlNode = xmlApplicationDocument.CreateElement("configuration")
                    xmlApplicationDocument.AppendChild(configNode)
                End If
            End SyncLock

            'Select the item nodes
            Dim sectionNode As XmlNode = xmlApplicationDocument.SelectSingleNode(("/configuration/" + SectionName))
            If Not sectionNode Is Nothing Then
                'Remove the node contents
                sectionNode.RemoveAll()
            Else
                ' If the node does not exist, then it's created
                sectionNode = xmlApplicationDocument.CreateElement(SectionName)
                Dim configurationNode As XmlNode = xmlApplicationDocument.SelectSingleNode("/configuration")
                configurationNode.AppendChild(sectionNode)
            End If

            Dim cloneNode As XmlNode
            If IsSigned OrElse IsEncrypted Then
                Dim paramValue As String = value.OuterXml
                Dim paramSignature As String = ""
                If IsEncrypted Then
                    Dim encryptedBytes As Byte() = Nothing
                    Dim decryptedBytes As Byte() = Nothing
                    Try
                        decryptedBytes = Encoding.UTF8.GetBytes(paramValue)
                        encryptedBytes = DataProtection.Encrypt(decryptedBytes)
                        paramValue = Convert.ToBase64String(encryptedBytes)
                    Finally
                        If Not (encryptedBytes Is Nothing) Then
                            Array.Clear(encryptedBytes, 0, encryptedBytes.Length)
                        End If
                        If Not (decryptedBytes Is Nothing) Then
                            Array.Clear(decryptedBytes, 0, decryptedBytes.Length)
                        End If
                    End Try
                End If
                If IsSigned Then
                    'Compute the hash
                    Dim hash As Byte() = _dataProtection.ComputeHash(Encoding.UTF8.GetBytes(paramValue))
                    paramSignature = Convert.ToBase64String(hash)
                End If

                Dim signatureNode As XmlNode = sectionNode.OwnerDocument.CreateElement("signature")
                signatureNode.InnerText = paramSignature
                sectionNode.AppendChild(signatureNode)
                Dim encryptedNode As XmlNode = sectionNode.OwnerDocument.CreateElement("encryptedData")
                encryptedNode.InnerXml = paramValue
                sectionNode.AppendChild(encryptedNode)
            Else
                cloneNode = xmlApplicationDocument.ImportNode(value, True)

                'Appends the node to the document
                sectionNode.AppendChild(cloneNode)
            End If

            'Save the document
            SyncLock Me.GetType()
                If Not _fileWatcherApp Is Nothing Then
                    _fileWatcherApp.EnableRaisingEvents = False
                End If

                Dim fs As FileStream = Nothing
                Try
                    fs = New FileStream(_applicationDocumentPath, FileMode.Create)
                    xmlApplicationDocument.Save(fs)
                    fs.Flush()
                Finally
                    If Not (fs Is Nothing) Then
                        CType(fs, IDisposable).Dispose()
                    End If
                End Try

                If Not _fileWatcherApp Is Nothing Then
                    _fileWatcherApp.EnableRaisingEvents = True
                End If
            End SyncLock
        End Sub 'Write 
#End Region

#Region "Protected properties"
        Protected ReadOnly Property ApplicationDocumentPath() As String
            Get
                Return _applicationDocumentPath
            End Get
        End Property

        Protected ReadOnly Property IsSigned() As Boolean
            Get
                Return _isSigned
            End Get
        End Property

        Protected ReadOnly Property IsEncrypted() As Boolean
            Get
                Return _isEncrypted
            End Get
        End Property

        Public ReadOnly Property IsRefreshOnChange() As Boolean
            Get
                Return _isRefreshOnChange
            End Get
        End Property

        Protected ReadOnly Property IsInitOK() As Boolean
            Get
                Return _isInitOk
            End Get
        End Property

        Protected ReadOnly Property DataProtection() As IDataProtection
            Get
                Return _dataProtection
            End Get
        End Property

        Protected ReadOnly Property SectionName() As String
            Get
                Return _sectionName
            End Get
        End Property
#End Region

#Region "Protected methods"

        ' <summary>
        ' FileSystemWatcher Event Handler
        ' </summary>
        Protected Overridable Sub OnChanged(ByVal sender As Object, ByVal e As FileSystemEventArgs)
            'Notify file changes to the configuration manager
            RaiseEvent ConfigChanges(Me, SectionName)
        End Sub 'OnChanged

        ' <summary>
        ' Loads the Xml file on the document
        ' </summary>
        ' <param name="doc">An Xml document instance</param>
        ' <param name="fileName">The file name</param>
        Sub LoadXmlFile(ByVal doc As XmlDocument, ByVal fileName As String)
            Dim fs As FileStream = Nothing
            Try
                fs = New FileStream(fileName, FileMode.Open, FileAccess.Read)
                doc.Load(fs)
            Finally
                If Not (fs Is Nothing) Then
                    CType(fs, IDisposable).Dispose()
                End If
            End Try
        End Sub 'LoadXmlFile
#End Region
    End Class 'XmlFileStorage
End Namespace 'Storage
