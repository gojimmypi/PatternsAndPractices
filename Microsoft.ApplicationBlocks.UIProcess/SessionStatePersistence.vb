'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' SessionStatePersistence.vb
'
' This file contains the implementations of the SessionStatePersistence class
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
Imports System.Web

Imports Microsoft.ApplicationBlocks.UIProcess

'This is a simple Session-based state persistence for Webform applications.  
'It pushes State directly to Session variables keyed to the Task GUID.
'It is useful for applications where State might be ephemeral and not worth saving to SQL; it
'is also useful for web-farm applications where multiple front-end servers need to "see" the same state.
'In this case, ASP Session would be used either as a State Server OR in SQL Session mode, and this 
'persistence provider would piggyback on ASP Session.
'By using this, it is possible to easily migrate applications to non-ASP by simply replacing the 
'State Persistence Provider (among the other normally necessary changes)
Friend Class SessionStatePersistence
    Implements IStatePersistence


    Public Sub New()
    End Sub 'New 

#Region "IStatePersistence Members"

    Public Sub Init(ByVal statePersistenceParameters As System.Collections.Specialized.NameValueCollection) Implements IStatePersistence.Init
    End Sub 'Init

    Public Function Load(ByVal taskId As Guid) As State Implements IStatePersistence.Load
        '  pull State object directly out of Session
        Return CType(HttpContext.Current.Session(taskId.ToString()), State)
    End Function 'Load

    Public Sub Save(ByVal inState As State) Implements IStatePersistence.Save
        '  put State object directly into Session
        HttpContext.Current.Session(inState.TaskId.ToString()) = inState
    End Sub 'Save

#End Region
End Class 'SessionStatePersistence
