'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' MemoryStatePersistence.vb
'
' This file contains the implementations of the MemoryStatePersistence class
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
Imports System.Collections.Specialized
Imports Microsoft.ApplicationBlocks.UIProcess

' This is a simple Memory-based state persistence for Winform applications.  DO NOT use it server-side, the locking 
' will bottleneck busy web apps.
Friend Class MemoryStatePersistence
    Implements IStatePersistence
    Private _stateReferences As New HybridDictionary

#Region "IStatePersistence Members"

    Public Sub Init(ByVal statePersistenceParameters As System.Collections.Specialized.NameValueCollection) Implements IStatePersistence.Init
    End Sub 'Init

    Public Function Load(ByVal taskId As Guid) As State Implements IStatePersistence.Load
        Return CType(_stateReferences(taskId), State)
    End Function 'Load

    Public Sub Save(ByVal state As State) Implements IStatePersistence.Save
        '  lock on syncroot to prevent collisions
        SyncLock _stateReferences.SyncRoot
            _stateReferences(state.TaskId) = state
        End SyncLock
    End Sub 'Save

#End Region
End Class 'MemoryStatePersistence
