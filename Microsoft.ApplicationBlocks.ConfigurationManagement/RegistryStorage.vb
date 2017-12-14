' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' RegistryStorage.vb
'
' This file contains a read-write storage implementation that uses the windows
' registry to save the configuration.
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
Imports [SC] = System.Configuration
Imports System.Text
Imports System.Xml
Imports Microsoft.Win32

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

Namespace Storage

#Region "RegistryStorage class"

    ' <summary>
    ' Sample storage provider used to get data from the windows's registry
    ' </summary>
    ' <remarks>This privider uses the following attributes on the XML file:
    ' <list type="">
    '		<item><b>registrySubKey</b>. Used to specify the subkey for a Windows's registry value</item>
    '		<item><b>encrypted</b>. Used to specify if the section must be encrypted </item>
    '		<item><b>signed</b>. Used to specify if the section must be signed </item>
    ' </list>
    ' </remarks>

    Friend Class RegistryStorage
        Implements IConfigurationStorageWriter
#Region "Declare Variables"
        Private _registrySubkey As RegistryKey = Nothing
        Private _registryRoot As RegistryHive = RegistryHive.CurrentUser
        Private _isSigned As Boolean = False
        Private _isEncrypted As Boolean = False
        Private _isInitOk As Boolean = False
        Private _dataProtection As IDataProtection
        Private _sectionName As String = Nothing
#End Region

#Region "Default constructor"

        Public Sub New()
        End Sub 'New
#End Region

#Region "IConfigurationStorageReader implementation"

        Sub Init(ByVal sectionName As String, ByVal configStorageParameters As ListDictionary, _
                    ByVal dataProtection As IDataProtection) Implements IConfigurationStorageWriter.Init

            _sectionName = sectionName

            'Inits the provider properties
            Dim registryRootString As String = CType(configStorageParameters("registryRoot"), String) '
            If Not (registryRootString Is Nothing) AndAlso registryRootString.Length <> 0 Then
                _registryRoot = CType([Enum].Parse(GetType(RegistryHive), registryRootString, True), RegistryHive)
            End If
            Dim registrySubKeyString As String = CType(configStorageParameters("registrySubKey"), String) '
            If registrySubKeyString Is Nothing OrElse registrySubKeyString.Length = 0 Then
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidRegistrySubKey", _
                                        registrySubKeyString))
            End If

            Dim registryPath As String = ""
            Select Case (RegistryRoot)
                Case RegistryHive.CurrentUser
                    registryPath = Registry.CurrentUser.Name
                Case RegistryHive.LocalMachine
                    registryPath = Registry.LocalMachine.Name
                Case RegistryHive.Users
                    registryPath = Registry.Users.Name
            End Select

            _registrySubkey = GetRegistrySubKey(RegistryRoot, registrySubKeyString)
            If RegistrySubkey Is Nothing Then
                Throw New SC.ConfigurationErrorsException( _
                    Resource.ResourceManager("RES_ExceptionInvalidRegistrySubKey", registrySubKeyString))
            End If
            Dim signedString As String = CType(configStorageParameters("signed"), String) '
            If Not (signedString Is Nothing) AndAlso signedString.Length <> 0 Then
                _isSigned = Boolean.Parse(signedString)
            End If
            Dim encryptedString As String = CType(configStorageParameters("encrypted"), String) '
            If Not (encryptedString Is Nothing) AndAlso encryptedString.Length <> 0 Then
                _isEncrypted = Boolean.Parse(encryptedString)
            End If
            Me._dataProtection = dataProtection

            If (_isSigned OrElse _isEncrypted) AndAlso (_dataProtection Is Nothing) Then
                Throw New SC.ConfigurationErrorsException( _
                    Resource.ResourceManager("RES_ExceptionInvalidDataProtectionConfiguration", _sectionName))
            End If
            _isInitOk = True
        End Sub 'IConfigurationStorageReader.Init

        ReadOnly Property Initialized() As Boolean Implements IConfigurationStorageWriter.Initialized
            Get
                Return _isInitOk
            End Get
        End Property

        Function Read() As XmlNode Implements IConfigurationStorageWriter.Read
            Dim xmlSection, xmlSignature As String
            Try
                Dim sectionKey As RegistryKey = Nothing
                Try
                    sectionKey = RegistrySubkey.OpenSubKey(SectionName, False)
                    If sectionKey Is Nothing Then
                        Return Nothing
                    End If
                    xmlSection = CStr(sectionKey.GetValue("value"))
                    xmlSignature = CStr(sectionKey.GetValue("signature"))
                Finally
                    If Not (sectionKey Is Nothing) Then
                        CType(sectionKey, IDisposable).Dispose()
                    End If
                End Try
            Catch e As Exception
                Throw New SC.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionCantGetSectionData", SectionName), e)
            End Try
            If xmlSection Is Nothing Then
                Throw New SC.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionCantGetSectionData", SectionName))
            End If
            Dim xmlDoc As XmlDocument = Nothing
            If _isSigned Then
                'Compute the hash
                Dim hash As Byte() = _dataProtection.ComputeHash(Encoding.UTF8.GetBytes(xmlSection))
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
                        encryptedBytes = Convert.FromBase64String(xmlSection)
                        decryptedBytes = _dataProtection.Decrypt(encryptedBytes)
                        xmlSection = Encoding.UTF8.GetString(decryptedBytes)
                    Catch ex As Exception
                        Throw New SC.ConfigurationErrorsException( _
                                Resource.ResourceManager("RES_ExceptionInvalidEncryptedString"), ex)
                    End Try
                Finally
                    If Not (encryptedBytes Is Nothing) Then Array.Clear(encryptedBytes, 0, encryptedBytes.Length)
                    If Not (decryptedBytes Is Nothing) Then Array.Clear(decryptedBytes, 0, decryptedBytes.Length)
                End Try
            End If
            xmlDoc = New XmlDocument
            xmlDoc.LoadXml(xmlSection)
            Return xmlDoc.DocumentElement
        End Function 'IConfigurationStorageReader.Read

        ' <summary>
        ' Event used internally to know when the storage detect some changes 
        ' on the config
        ' </summary>
        Public Event ConfigChanges As ConfigurationChanged Implements IConfigurationStorageWriter.ConfigChanges

#End Region

#Region "IConfigurationStorageWriter implementation"


        ' <summary>
        ' Writes a section on the XML document
        ' </summary>
        Sub Write(ByVal value As XmlNode) Implements IConfigurationStorageWriter.Write
            If SectionName Is Nothing OrElse SectionName.Length = 0 Then
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionCantUseNullKeys"))
            End If

            Dim paramSignature As String = ""
            Dim paramValue As String = value.OuterXml

            If paramValue.Length > 500000 Then
                Throw New SC.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionRegistryValueLimit", paramValue.Length))
            End If
            Try
                If _isEncrypted Then
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
                If _isSigned Then
                    'Compute the hash
                    Dim hash As Byte() = _dataProtection.ComputeHash(Encoding.UTF8.GetBytes(paramValue))
                    paramSignature = Convert.ToBase64String(hash)
                End If

                Dim sectionKey As RegistryKey = Nothing
                Try
                    sectionKey = RegistrySubkey.OpenSubKey(SectionName, True)
                    If sectionKey Is Nothing Then
                        sectionKey = RegistrySubkey.CreateSubKey(SectionName)
                    End If
                    sectionKey.SetValue("value", paramValue)
                    sectionKey.SetValue("signature", paramSignature)
                Finally
                    If Not (sectionKey Is Nothing) Then
                        CType(sectionKey, IDisposable).Dispose()
                    End If
                End Try
            Catch e As Exception
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionCantWriteRegistry", _
                            RegistrySubkey.ToString(), SectionName, e.ToString()), e)
            End Try
        End Sub 'IConfigurationStorageWriter.Write

#End Region

#Region "Private methods"


        '  only HKCU, HKLM, and HKU are valid.  Others are too much of a security risk, even with lock-down 
        '  DON'T store transient config data in sensitive registry hives
        '  the hives accessed by ConfigMan should be locked down by ACL to restrict activity to the least necessary
        Private Function GetRegistrySubKey(ByVal root As RegistryHive, ByVal subKey As String) As RegistryKey
            Dim subKeyObject As RegistryKey = Nothing
            Select Case root
                Case RegistryHive.CurrentUser
                    subKeyObject = Registry.CurrentUser.OpenSubKey(subKey, True)
                Case RegistryHive.LocalMachine
                    subKeyObject = Registry.LocalMachine.OpenSubKey(subKey, True)
                Case RegistryHive.Users
                    subKeyObject = Registry.Users.OpenSubKey(subKey, True)
                Case Else
                    '  if they ask for a disallowed Hive, throw here...limit to HKCU, HKU, HKLM
                    Dim errString As String = String.Format( _
                            Resource.ResourceManager("RES_ExceptionDisallowedRegistryKey"), _
                            [Enum].GetName(GetType(RegistryHive), root))
                    Throw New Exception(errString)
            End Select

            Return subKeyObject
        End Function 'GetRegistrySubKey


#End Region

#Region "protected properties"

        Protected ReadOnly Property RegistrySubkey() As RegistryKey
            Get
                Return _registrySubkey
            End Get
        End Property

        Protected ReadOnly Property RegistryRoot() As RegistryHive
            Get
                Return _registryRoot
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
    End Class 'RegistryStorage
#End Region

End Namespace 'Storage
