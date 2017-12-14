'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' SecureSqlServerStatePersistence.cs
'
' This file contains the implementations of the SecureSqlServerStatePersistence and CryptHelper classes
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
Imports System.IO
Imports System.Data
Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Security.Permissions
Imports System.Security.Cryptography
Imports System.Diagnostics
Imports Microsoft.Win32

Imports Microsoft.ApplicationBlocks.Data

#Region "SecureSqlServerPersistState Provider"
'IStatePersistence implementation.
'This provider encrypts the serialized data using a symmetric algorithm.
'The algorithm key is obtained from the LOCAL_MACHINE hive in the windows registry.
Friend Class SecureSqlServerPersistState
    Implements IStatePersistence 
   
    #Region "CryptHelper class"
    'Helper class used to crypt intentions.
    Friend Class CryptHelper
        Private IV As Byte() =  {162, 239, 121, 27, 111, 214, 206, 34}
      
        'Encrypts a stream of bytes using a TripleDES symmetric algorithm
        'Parameters:
        '-plainValue: stream of bytes that going to be encrypted
        '-key: symmetric algorithm key
        'Returns:
        ' a encrypted strem of bytes
        Public Function Encrypt(plainValue() As Byte, key() As Byte) As Byte()
            Dim chiperValue() As Byte
         
            Dim algorithm As TripleDES = TripleDESCryptoServiceProvider.Create()
            Dim memStream As New MemoryStream()
         
            Dim cryptoStream As New CryptoStream(memStream, algorithm.CreateEncryptor(key, IV), CryptoStreamMode.Write)
         
            Try
                cryptoStream.Write(plainValue, 0, plainValue.Length)
                cryptoStream.Flush()
                cryptoStream.FlushFinalBlock()
            
                chiperValue = memStream.ToArray()
            Catch e As Exception
                Throw New UIPException(Resource.ResourceManager("RES_ExceptionSecureSqlProviderCantEncrypt"), e)
            Finally
                memStream.Close()
                cryptoStream.Close()
            End Try
         
            Return chiperValue
        End Function
           
        'Decrypts a encrypted stream of bytes using a TripleDES symmetric algorithm
        'Parameters: 
        '-cipherValue: encrypted stream of bytes that going to be decrypted
        '-key: symmetric algorithm key
        'Returns:
        ' a strem of bytes
        Public Function Decrypt(cipherValue() As Byte, key() As Byte) As Byte()
            Dim plainValue(cipherValue.Length - 1) As Byte
         
            Dim algorithm As TripleDES = TripleDESCryptoServiceProvider.Create()
            Dim memStream As New MemoryStream(cipherValue)
         
            Dim cryptoStream As New CryptoStream(memStream, algorithm.CreateDecryptor(key, IV), CryptoStreamMode.Read)
         
            Try
                cryptoStream.Read(plainValue, 0, plainValue.Length)
            Catch e As Exception
                Throw New UIPException(Resource.ResourceManager("RES_ExceptionSecureSqlProviderCantDecrypt"), e)
            Finally
                'Flush the stream buffer
                cryptoStream.Close()
            End Try
         
            Return plainValue
        End Function
    End Class
    #End Region
   
    #Region "Declares variables"
    Private Const ConfigRegistryValue As String = "symmetric key"
    Private Const ConfigDefaultRegistryValue As String = "SOFTWARE\Microsoft\UIP Application Block"
    Private Const ConfigConnectionString As String = "connectionString"
    Private Const ConfigRegistry As String = "registryPath"
    Private Const DbSelectState As String = "SelectState"
    Private Const DbParamStateGuid As String = "@StateGuid"
    Private Const DbInsertState As String = "InsertState"
    Private Const DbParamXmlState As String = "@XmlState"
    Private Const ReadSize As Integer = 1400
    Private connectionString As String = Nothing
    Private registryPath As String = Nothing
    Private _registryKey As RegistryKey
    #End Region
   
    #Region "IPersistState implementation"
    'The possible provider config attributes are:
    '   - connectionString: Specifies the database connection string
    '   - registryPath: Specifies the registry key path where is stored
    '                   the encryption symmetric key. 
    Public Sub Init(statePersistenceParameters As NameValueCollection) Implements IStatePersistence.Init
        connectionString = statePersistenceParameters(ConfigConnectionString)
        If connectionString Is Nothing Then
            Throw New ApplicationException(Resource.ResourceManager.FormatMessage("RES_ExceptionSQLStatePersistenceProviderInit", ConfigConnectionString))
        End If 
        registryPath = statePersistenceParameters(ConfigRegistry)
        If registryPath Is Nothing Then
            registryPath = ConfigDefaultRegistryValue
        End If 
        Try
            '  mstuart 03.30.2003:  here we're requesting permission to access the given registry key.
            '  if we can't, we throw right away.
            Dim permission As New RegistryPermission(RegistryPermissionAccess.Read, Registry.LocalMachine.Name + "\" + registryPath)
            permission.Demand()
        Catch e As System.Security.SecurityException
            Throw New UIPException(Resource.ResourceManager("RES_ExceptionSecureSqlProviderRegistryPermissions"), e)
        End Try
            
        _registryKey = Registry.LocalMachine.OpenSubKey(registryPath, False)
      
        If _registryKey Is Nothing Then
            Throw New UIPException(Resource.ResourceManager("RES_ExceptionSecureSqlProviderSymmetricKey"))
        End If
    End Sub
    
    Private Function Encrypt(plainValue() As Byte) As Byte()
        Dim chiperValue() As Byte
      
        'Get encryption key
        Dim base64Key As String = _registryKey.GetValue(ConfigRegistryValue).ToString()
        Dim key As Byte() = Convert.FromBase64String(base64Key)
      
        Dim cryptHelper As New CryptHelper()
        chiperValue = cryptHelper.Encrypt(plainValue, key)
      
        'Clean encryption key
        base64Key = Nothing
        key = New Byte(-1) {}
      
        Return chiperValue
    End Function
   
    Private Function Decrypt(cipherValue() As Byte) As Byte()
        Dim plainValue(cipherValue.Length - 1) As Byte
      
        'Get encryption key
        Dim base64Key As String = _registryKey.GetValue(ConfigRegistryValue).ToString()
        Dim key As Byte() = Convert.FromBase64String(base64Key)
      
        Dim cryptHelper As New CryptHelper()
        plainValue = cryptHelper.Decrypt(cipherValue, key)
      
        'Clean encryption key
        base64Key = Nothing
        key = New Byte(-1) {}
      
        Return plainValue
    End Function
   
    'Saves the state object into a SQL Server database
    'The provider encrypts the serialized state before stores it in the data base
    'Parameters: 
    '-state: a valid state object
    <SqlClientPermission(System.Security.Permissions.SecurityAction.Demand)>  _
    Public Sub Save(state As State) Implements IStatePersistence.Save
        Dim formatter As New BinaryFormatter()
        Dim memoryStream As New MemoryStream()
        formatter.Serialize(memoryStream, state)
      
        Dim serializedObject As Byte() = memoryStream.GetBuffer()
        Dim cipherObject As Byte() = Encrypt(serializedObject)
      
        Try
            Dim binState As New SqlParameter(DbParamXmlState, System.Data.SqlDbType.Image)
            binState.Value = cipherObject
         
            SqlHelper.ExecuteNonQuery(connectionString, CommandType.StoredProcedure, DbInsertState, New SqlParameter() {New SqlParameter(DbParamStateGuid, state.TaskId), binState})
        Catch ex As Exception
            Throw New ApplicationException(Resource.ResourceManager("RES_ExceptionSQLStatePersistenceProviderDehydrate"), ex)
        Finally
            memoryStream.Close()
        End Try
    End Sub
   
    'Loads a existing state object from a SQL Server database
    'The provider decrypts the serialized state before restores it
    'Parameters: 
    '-taskId: the task identifier
    'Returns: 
    'a valid state object
    <SqlClientPermission(System.Security.Permissions.SecurityAction.Demand)>  _
    Public Function Load(taskId As Guid) As State Implements IStatePersistence.Load
        Dim requestedState As State = Nothing
        Dim reader As SqlDataReader = Nothing
        Dim memoryStream As MemoryStream = Nothing
        Try
            reader = SqlHelper.ExecuteReader(connectionString, CommandType.StoredProcedure, DbSelectState, New SqlParameter(DbParamStateGuid, taskId))
         
            If Not reader.Read() Then
                reader.Close()
                Return Nothing
            End If
         
            'Get size of image data  pass null as the byte array parameter
            Dim byteTotal As Long = reader.GetBytes(0, 0, Nothing, 0, 0)
         
            ' Allocate byte array to hold image data
            Dim cipherObject(CInt(byteTotal)-1) As Byte
            Dim index As Integer = 0
            Dim bytesRead As Long = 0
            While bytesRead < byteTotal
                ' read the object binary data 
                bytesRead += reader.GetBytes(0, index, cipherObject, index, ReadSize)
                index += ReadSize
            End While
         
            'Decrypt the cipher object
            Dim serializedObject As Byte() = Decrypt(cipherObject)
         
            'Deserialize the object
            memoryStream = New MemoryStream(serializedObject)
            Dim formatter As New BinaryFormatter()
            requestedState = CType(formatter.Deserialize(memoryStream), State)
        Catch ex As Exception
            Throw New ApplicationException(Resource.ResourceManager("RES_ExceptionSQLStatePersistenceProviderRehydrate"), ex)
        Finally
            If Not (reader Is Nothing) Then
                reader.Close()
            End If
            If Not (memoryStream Is Nothing) Then
                memoryStream.Close()
            End If
        End Try 
        Return requestedState
   End Function
   #End Region
End Class
#End Region
