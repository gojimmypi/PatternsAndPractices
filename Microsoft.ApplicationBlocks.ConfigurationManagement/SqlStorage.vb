' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' SqlStorage.vb
'
' This file contains a read-write storage implementation that uses microsoft
' SqlServer 2000 to save the configuration.
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
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Xml

Imports Microsoft.ApplicationBlocks.Data
Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

Namespace Storage
    ' <summary>
    ' Sample SqlServer storage provider used to get data from a database
    ' </summary>
    ' <remarks>This privider uses the following attributes on the XML file:
    ' <list type="">
    '		<item><b>connectionString</b>. Used to specify the connection string for a SqlServer database</item>
    '		<item><b>getConfigSP</b>. The stored procedure used to get the database configuration. 
    '               This stored procedure must return an XML. This stored procedure must have the 
    '               following parameters: <c>@sectionName varchar(50)</c></item>
    '		<item><b>setConfigSP</b>. A stored procedure used to save the configuration on the database. 
    '               This stored prodedure must have the following parameters: 
    '               <c>@param_section varchar(50), @param_name varchar(50), @param_value varchar(255)</c>
    '       </item>
    ' </list>
    ' </remarks>

    Friend Class SqlStorage
        Implements IConfigurationStorageWriter
#Region "Declare Variables"

        Private _connectionString As String = Nothing
        Private _getConfigSP As String = Nothing
        Private _setConfigSP As String = Nothing
        Private _sectionName As String = Nothing
        Private _isSigned As Boolean = False
        Private _isEncrypted As Boolean = False
        Private _isInitOk As Boolean = False
        Private _dataProtection As IDataProtection = Nothing

#End Region

#Region "Constructor"

        Public Sub New()
        End Sub 'New


#End Region

#Region "IConfigurationStorageReader implementation"


        Public Sub Init(ByVal sectionName As String, ByVal configStorageParameters As ListDictionary, _
                ByVal dataProtection As IDataProtection) Implements IConfigurationStorageReader.Init
            'Inits the provider properties
            _sectionName = sectionName

            'Use the registry path attribute first
            Dim regKey As String = CType(configStorageParameters("connectionStringRegKeyPath"), String) '
            If Not (regKey Is Nothing) AndAlso regKey.Length <> 0 Then
                _connectionString = DataProtectionHelper.GetRegistryDefaultValue( _
                            regKey, "connectionString", "connectionStringRegKeyPath")
            End If
            'If the connection string was not in the regustry, use the 'connectionString' attribute
            If _connectionString Is Nothing OrElse _connectionString.Length = 0 Then
                _connectionString = CType(configStorageParameters("connectionString"), String) '
                If _connectionString Is Nothing OrElse _connectionString.Length = 0 Then
                    Throw New SC.ConfigurationErrorsException( _
                            Resource.ResourceManager("RES_ExceptionInvalidConnectionString", _connectionString))
                End If
            End If
            _getConfigSP = CType(configStorageParameters("getConfigSP"), String) '
            If _getConfigSP Is Nothing Then
                _getConfigSP = "cmab_get_config"
            End If
            _setConfigSP = CType(configStorageParameters("setConfigSP"), String) '
            If _setConfigSP Is Nothing Then
                _setConfigSP = "cmab_set_config"
            End If
            Dim signedString As String = CType(configStorageParameters("signed"), String) '
            If Not (signedString Is Nothing) AndAlso signedString.Length <> 0 Then
                _isSigned = Boolean.Parse(signedString)
            End If
            Dim encryptedString As String = CType(configStorageParameters("encrypted"), String) '
            If Not (encryptedString Is Nothing) AndAlso encryptedString.Length <> 0 Then
                _isEncrypted = Boolean.Parse(encryptedString)
            End If
            Me._dataProtection = DataProtection

            If (_isSigned OrElse _isEncrypted) AndAlso (_dataProtection Is Nothing) Then
                Throw New SC.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionInvalidDataProtectionConfiguration", _sectionName))
            End If
            Try
                Dim sqlConnection As sqlConnection = Nothing
                Try
                    sqlConnection = New sqlConnection(_connectionString)
                    '  attempt to open the database...catch the exception early here rather than elsewhere 
                    '  when we're trying to get data
                    sqlConnection.Open()
                    _isInitOk = True
                Finally
                    If Not (sqlConnection Is Nothing) Then
                        CType(sqlConnection, IDisposable).Dispose()
                    End If
                End Try
            Catch e As Exception
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionCantOpenConnection", _
                                        _connectionString), e)
            End Try
        End Sub 'Init


        Public ReadOnly Property Initialized() As Boolean Implements IConfigurationStorageReader.Initialized
            Get
                Return _isInitOk
            End Get
        End Property


        Public Function Read() As XmlNode Implements IConfigurationStorageReader.Read
            Dim xmlSection, xmlSignature As String
            Try
                Dim reader As SqlDataReader = Nothing
                Try
                    reader = SqlHelper.ExecuteReader(_connectionString, _getConfigSP, SectionName)
                    If Not reader.Read() Then
                        Return Nothing
                    End If
                    xmlSection = CType(IIf(reader.IsDBNull(0), Nothing, reader.GetString(0)), String)
                    xmlSignature = CType(IIf(reader.IsDBNull(1), Nothing, reader.GetString(1)), String)
                Finally
                    If Not (reader Is Nothing) Then
                        CType(reader, IDisposable).Dispose()
                    End If
                End Try
            Catch e As Exception
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionCantGetSectionData", _
                        SectionName), e)
            End Try
            If xmlSection Is Nothing Then
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionCantGetSectionData", _
                        SectionName))
            End If
            Dim xmlDoc As XmlDocument = Nothing
            If _isSigned Then
                'Compute the hash
                Dim hash As Byte() = _dataProtection.ComputeHash(Encoding.UTF8.GetBytes(xmlSection))
                Dim newHashString As String = Convert.ToBase64String(hash)

                'Compare the hashes
                If newHashString <> xmlSignature Then
                    Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionSignatureCheckFailed", _
                                                SectionName))
                End If
            Else
                If Not (xmlSignature Is Nothing) AndAlso xmlSignature.Length <> 0 Then
                    Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionInvalidSignatureConfig", _
                                                _sectionName))
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
        ' Not used.
        ' </summary>
        Public Event ConfigChanges As ConfigurationChanged Implements IConfigurationStorageReader.ConfigChanges

#End Region

#Region "IConfigurationStorageWriter implementation"

        ' <summary>
        ' Writes a section on the XML document
        ' </summary>
        Public Sub Write(ByVal value As XmlNode) Implements IConfigurationStorageWriter.Write
            Dim paramSignature As String = ""
            Dim paramValue As String = value.OuterXml

            If IsEncrypted Then
                Dim encryptedBytes As Byte() = Nothing
                Dim decryptedBytes As Byte() = Nothing
                Try
                    decryptedBytes = Encoding.UTF8.GetBytes(paramValue)
                    encryptedBytes = DataProtection.Encrypt(decryptedBytes)
                    paramValue = Convert.ToBase64String(encryptedBytes)
                Finally
                    If Not (encryptedBytes Is Nothing) Then Array.Clear(encryptedBytes, 0, encryptedBytes.Length)
                    If Not (decryptedBytes Is Nothing) Then Array.Clear(decryptedBytes, 0, decryptedBytes.Length)
                End Try
            End If
            If IsSigned Then
                'Keep the hashed value
                Dim hash As Byte() = _dataProtection.ComputeHash(Encoding.UTF8.GetBytes(paramValue))
                paramSignature = Convert.ToBase64String(hash)
            End If
            Dim rows As Integer
            Try
                Dim sectionValueParameter As New SqlParameter("@section_value", SqlDbType.NText)
                sectionValueParameter.Value = paramValue

                rows = SqlHelper.ExecuteNonQuery(ConnectionString, CommandType.StoredProcedure, SetConfigSP, _
                                    New SqlParameter("@section_name", SectionName), sectionValueParameter, _
                                    New SqlParameter("@section_signature", paramSignature))
                If rows <> 1 Then
                    Throw New SC.ConfigurationErrorsException( _
                        Resource.ResourceManager("RES_ExceptionStoredProcedureUpdatedNoRecords", _
                                SetConfigSP, SectionName, paramValue, rows))
                End If
            Catch e As Exception
                Throw New SC.ConfigurationErrorsException(Resource.ResourceManager("RES_ExceptionCantExecuteStoredProcedure", _
                                    SetConfigSP, SectionName, paramValue, e))
            End Try
        End Sub 'Write


#End Region

#Region "Protected properties"

        Protected ReadOnly Property ConnectionString() As String
            Get
                Return _connectionString
            End Get
        End Property

        Protected ReadOnly Property GetConfigSP() As String
            Get
                Return _getConfigSP
            End Get
        End Property

        Protected ReadOnly Property SetConfigSP() As String
            Get
                Return _setConfigSP
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
    End Class 'SqlStorage
End Namespace 'Storage