' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' ConfigurationManagementCache.vb
'
' Cache management for the application block.
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

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement.DataProtection

#Region "ConfigurationManagementCache Class"

' <summary>
' This class provides cache services to the Configuration Manager.
' </summary>
Friend Class ConfigurationManagementCache

#Region "Declare Variables"
    Private _sectionName As String
    Private _storage As ICacheStorage
    Private _refresh As String
#End Region

#Region "Constructors"

    ' <summary>
    ' Constructor allowing the cache configuration to be set.
    ' </summary>
    ' <param name="sectionName">The section to be cached</param>
    ' <param name="cacheSettings">The settings for the cache</param>
    Public Sub New(ByVal sectionName As String, ByVal cacheSettings As ConfigCacheSettings)
        _refresh = cacheSettings.Refresh
        _storage = ActivateCacheStorage()

        Me._sectionName = sectionName
    End Sub 'New
#End Region

#Region "Indexer"

    ' <summary>
    ' Class indexer
    ' </summary>
    Default Public Property Item(ByVal key As String) As Object
        Get
            Return [Get](key)
        End Get
        Set(ByVal Value As Object)
            Add(key, Value)
        End Set
    End Property

#End Region

#Region "Properties"

    ' <summary>
    ' Absolute time format for refresh of the config data cache. 
    ' </summary>
    Public ReadOnly Property Refresh() As String
        Get
            Return _refresh
        End Get
    End Property

    ' <summary>
    ' This property specifies the section name associated with this cache
    ' </summary>

    Public ReadOnly Property SectionName() As String
        Get
            Return _sectionName
        End Get
    End Property

#End Region

    ' <summary>
    ' Puts an item in the cache
    ' </summary>
    Private Sub Add(ByVal key As String, ByVal item As Object)
        If key Is Nothing Then
            Throw New ArgumentNullException("key", Resource.ResourceManager("RES_ExceptionInvalidKeyValue"))
        End If
        If item Is Nothing Then
            Throw New ArgumentNullException("item", Resource.ResourceManager("RES_ExceptionInvalidCacheElement"))
        End If
        _storage.Add(key, New CacheValue(item, DateTime.Now))
    End Sub 'Add

    ' <summary>
    ' Gets the item with the specific key
    ' </summary>
    Private Function [Get](ByVal key As String) As CacheValue
        Dim returnValue As CacheValue = Nothing

        returnValue = _storage.Get(key)

        If returnValue Is Nothing OrElse _
                ExtendedFormatHelper.IsExtendedExpired(_refresh, returnValue.ItemAge, DateTime.Now) = True Then
            ' item has expired from cache
            ' use WeakReference for cache items to respond better to memory pressure.  
            ' The cache is freed on memory pressure.
            Return Nothing
        Else
            Return returnValue
        End If
    End Function 'Get

    ' <summary>
    ' Determines whether the cache contains the specific key
    ' </summary>
    Public Function ContainsKey(ByVal key As String) As Boolean
        Return _storage.ContainsKey(key)
    End Function 'ContainsKey

    ' <summary>
    ' Removes all elements from the cache
    ' </summary>
    Public Sub Clear()
        If Not (_storage Is Nothing) Then
            _storage.Clear()
        End If
    End Sub 'Clear

    ' <summary>
    ' Private helper function to assist in cache storage activations. Returns
    ' an CacheStorage from the specified location type.
    ' </summary>
    ' <returns>Instance of a specific CacheStorage implementation.</returns>
    Private Function ActivateCacheStorage() As ICacheStorage
        Return New MemoryCacheStorage
    End Function 'ActivateCacheStorage
End Class 'ConfigurationManagementCache 
#End Region

