' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' CacheStorage.vb
'
' Cache storage support. This file defines the interface and a sample implementation.
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
Imports System.Collections

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

#Region "ICacheStorage"

' <summary>
' This interface must be implemented by cache storage providers
' </summary>
Friend Interface ICacheStorage

    ' <summary>
    ' Adds an item to the cache
    ' </summary>
    ' <param name="key">item key</param>
    ' <param name="item">item value</param>
    Sub Add(ByVal key As String, ByVal item As CacheValue)

    ' <summary>
    ' Gets a item from the cache
    ' </summary>
    ' <param name="key">item key</param>
    ' <returns>item value</returns>
    Function [Get](ByVal key As String) As CacheValue

    ' <summary>
    ' Determines whether the cache contains the specific key
    ' </summary>
    ' <param name="key">item key</param>
    ' <returns>true if the item exist, otherwise false</returns>
    Function ContainsKey(ByVal key As String) As Boolean

    ' <summary>
    ' Removes all elements from the cache
    ' </summary>
    Sub Clear()
End Interface 'ICacheStorage

#End Region

#Region "Cache Storage classes"

' <summary>
' The cache value used to hold the cache type
' </summary>
Friend Class CacheValue

    Public Sub New(ByVal value As Object, ByVal itemAge As DateTime)
        _value = value
        _itemAge = itemAge
    End Sub 'New


    Public ReadOnly Property Value() As Object
        Get
            Return _value
        End Get
    End Property
    Private _value As Object


    Public ReadOnly Property ItemAge() As DateTime
        Get
            Return _itemAge
        End Get
    End Property
    Private _itemAge As DateTime
End Class 'CacheValue

#End Region

#Region "MemoryCacheStorage"

' <summary>
' This class implements a cache in memory
' </summary>
Friend Class MemoryCacheStorage
    Implements ICacheStorage

#Region "Declare Variables"
    Private items As New Hashtable
#End Region

    ' <summary>
    ' Adds an item to the cache
    ' </summary>
    Public Sub Add(ByVal key As String, ByVal item As CacheValue) Implements ICacheStorage.Add
        SyncLock items.SyncRoot
            items(key) = item
        End SyncLock
    End Sub 'Add


    ' <summary>
    ' Gets a item from the cache
    ' </summary>
    Public Function [Get](ByVal key As String) As CacheValue Implements ICacheStorage.Get
        Return CType(items(key), CacheValue)
    End Function 'Get


    ' <summary>
    ' Determines whether the cache contains the specific key
    ' </summary>
    Public Function ContainsKey(ByVal key As String) As Boolean Implements ICacheStorage.ContainsKey
        Return items.ContainsKey(key)
    End Function 'ContainsKey


    ' <summary>
    ' Removes all elements from the cache
    ' </summary>
    Public Sub Clear() Implements ICacheStorage.Clear
        SyncLock items.SyncRoot
            items.Clear()
        End SyncLock
    End Sub 'Clear
End Class 'MemoryCacheStorage '

#End Region
