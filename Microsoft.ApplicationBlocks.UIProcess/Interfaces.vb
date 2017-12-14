'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' Interfaces.vb
'
' This file contains the definitions of the ITask, IView, IViewManager and IStatePersistence
' interfaces
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

#Region "ITask Interface"

'Defines a Task Object wich can be passed to UIPManager.  Used by applications that wish to get the TaskID for 
'their internal use.  Example:  Logon of shopping cart...Redirects to another page, but needs to know TaskID to correlate 
'with Logon.
Public Interface ITask
    'Get the taskid for the logon
    Function [Get]() As Guid
   
    'Create a task id for the logon
    'Parameters:
    '-taskId: task identifier
    Sub Create(taskId As Guid)
End Interface
#End Region

#Region "IView Interface"
'Represents an view used in web and win applications
Public Interface IView
    'Gets the current view controller
    ReadOnly Property Controller() As ControllerBase
   
    'Gets the view name
    ReadOnly Property ViewName() As String
   
    'Gets the current view navigation graph. 
    'This view is actually shown in this navigation graph
    ReadOnly Property NavigationGraph() As String
   
    'Gets the task id related to this view
    ReadOnly Property TaskId() As Guid
End Interface
#End Region

#Region "IViewManager Interface"
'Represents a view manager. 
'Each type of application has associate an view manager, 
'therefore is easier to add more application types.
Public Interface IViewManager
    'Stores a property into the view manager. 
    'Each task has its own properties
    'The property storage is a view manager responsibility
    'Parameters:
    '-taskId: task identifier    
    '-name: property name
    '-value: property value
    Sub StoreProperty(taskId As Guid, name As String, value As Object)
    
    'Activates the specified view
    'Parameters:
    '-previousView: the view actually displayed
    '-taskId: a existing task id
    '-navigationGraph: a configured navigation graph name
    '-view: the view name to be displayed
    Sub ActivateView(previousView As String, taskId As Guid, navigationGraph As String, view As String)
   
    'Utility method that checks requests to ensure that requested view and current view match.
    'Parameters:
    '-view: IView requested
    '-stateViewName: View name saved into the state
    Function IsRequestCurrentView(view As IView, stateViewName As String) As Boolean
   
    'Gets the running tasks in the manager
    'Returns:
    'a array with the task identifiers
    Function GetCurrentTasks() As Guid()
End Interface
#End Region

#Region "State Persist Provider Interface"
'Interface defines how State and Task objects may be dehydrated/rehydrated by Manager object.
'Allows us to abstract storage so we can use SQL, file, binary, XML, whatever.
Public Interface IStatePersistence
    'Inits the provider
    '-statePersistenceParameters: provider settings
    Sub Init(statePersistenceParameters As NameValueCollection)
   
    'Serializes and saves the state on a specific storage
    'Parameters:
    '-state: a valid state object
    Sub Save(state As State)
   
    'Restores and deserializes the state from a specific storage 
    'If the task doesn´t exists then a null value is returned
    'Parameters:
    '-taskId: A task identifier. This identifier will be used to restore a saved state
    'Returns:
    'a valid state object
    Function Load(taskId As Guid) As State
End Interface
#End Region
