'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' SqlServerStatePersistence.vb
'
' This file contains the implementations of the SqlServerStatePersistence class
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
Imports System.Configuration
Imports System.Runtime.Serialization.Formatters.Binary
Imports Microsoft.ApplicationBlocks.Data

#Region "SqlPersistState Provider"
'IStatePersistence SQL Server implementation.
'This provider uses SQL Server storage to dehydrate/rehydrate state objects
Friend Class SqlServerPersistState
    Implements IStatePersistence
   
    #Region "Declares variables"
    Private Const ConfigConnectionString As String = "connectionString"
    Private Const DbSelectState As String = "SelectState"
    Private Const DbParamStateGuid As String = "@StateGuid"
    Private Const DbInsertState As String = "InsertState"
    Private Const DbParamXmlState As String = "@XmlState"
    Private Const ReadSize As Integer = 1400
    Private connectionString As String = Nothing
    #End Region
   
    #Region "Constructors"
       
    Public Sub New()
    End Sub
    #End Region
   
    #Region "IPersistState implementation"
    'The possible provider config attributes are:
    '   - connectionString: Specifies the database connection string
    Public Sub Init(statePersistenceParameters As NameValueCollection) Implements IStatePersistence.Init 
        connectionString = statePersistenceParameters(ConfigConnectionString)
        If connectionString Is Nothing Then
            Throw New ApplicationException(Resource.ResourceManager.FormatMessage("RES_ExceptionSQLStatePersistenceProviderInit", ConfigConnectionString))
        End If
    End Sub
       
    'Saves the state object into a SQL Server database
    'Parameters: 
    '-state: a valid state object
    <SqlClientPermission(System.Security.Permissions.SecurityAction.Demand)>  _
    Public Sub Save(state As State) Implements IStatePersistence.Save  
        Dim formatter As New BinaryFormatter()
        Dim memoryStream As New MemoryStream()
        formatter.Serialize(memoryStream, state)
          
        Dim serializedObject As Byte() = memoryStream.GetBuffer()
          
        Try
            Dim binState As New SqlParameter(DbParamXmlState, System.Data.SqlDbType.Image)
            binState.Value = serializedObject
             
            SqlHelper.ExecuteNonQuery(connectionString, CommandType.StoredProcedure, DbInsertState, New SqlParameter() {New SqlParameter(DbParamStateGuid, state.TaskId), binState})
        Catch ex As Exception
            Throw New ApplicationException(Resource.ResourceManager("RES_ExceptionSQLStatePersistenceProviderDehydrate"), ex)
        Finally
            memoryStream.Close()
        End Try
    End Sub
      
    'Loads a existing state object from a SQL Server database
    'Parameters: 
    '-taskId: the task identifier
    'Returns: a valid state object
    <SqlClientPermission(System.Security.Permissions.SecurityAction.Demand)>  _
    Public Function Load(taskGuid As Guid) As State Implements IStatePersistence.Load
        Dim requestedState As State = Nothing
        Dim reader As SqlDataReader = Nothing
        Dim memoryStream As MemoryStream = Nothing
        Try
            reader = SqlHelper.ExecuteReader(connectionString, CommandType.StoredProcedure, DbSelectState, New SqlParameter(DbParamStateGuid, taskGuid))
             
            If Not reader.Read() Then
                reader.Close()
                Return Nothing
            End If
             
            'Get size of image data  pass null as the byte array parameter
            Dim byteTotal As Long = reader.GetBytes(0, 0, Nothing, 0, 0)
             
            ' Allocate byte array to hold image data
            Dim serializedObject(CInt(byteTotal)) As Byte
            Dim index As Integer = 0
            Dim bytesRead As Long = 0
            While bytesRead < byteTotal
                ' read the object binary data
                bytesRead += reader.GetBytes(0, index, serializedObject, index, ReadSize)
                index += ReadSize
            End While
             
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