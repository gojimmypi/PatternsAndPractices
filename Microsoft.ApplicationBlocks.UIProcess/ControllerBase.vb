'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' ControllerBase.vb
'
' This file contains the implementations of the ControllerBase, StartTaskEventArgs,
' StartTaskEventHandler and TaskArgumentsHolder classes
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

'This class hold the arguments that can be passed from one nav graph to another when chaining use cases.
'The data include a "return pointer"--the originating graph and taskid--and an object wich can be used
'to encapsulate data that is passed from one nav to another.
Public Class TaskArgumentsHolder
    Private _originatingTaskID As Guid
    Private _originatingNavGraphName As String
    Private _taskArguments As Object
        
    'Constructor
	'-originatingTaskID: originating task id
	'-originatingNavGraphName: originating navigation graph name
	'-taskArguments: an object with generic data
    Public Sub New(originatingTaskID As Guid, originatingNavGraphName As String, taskArguments As Object)
        _originatingTaskID = originatingTaskID
        _originatingNavGraphName = originatingNavGraphName
        _taskArguments = taskArguments
    End Sub
        
    'Gets/Sets an object wich can be used
    'to encapsulate data that is passed from one nav to another
    Public Property TaskArguments() As Object
        Get
            Return _taskArguments
        End Get
        Set
            _taskArguments = value
        End Set
    End Property
        
    'Gets/Sets the originating task id
    Public Property OriginatingTaskID() As Guid
        Get
            Return _originatingTaskID
        End Get
        Set
            _originatingTaskID = value
        End Set
    End Property
        
    'Gets/Sets the originating navigation graph name
    Public Property OriginatingNavGraphName() As String
        Get
            Return _originatingNavGraphName
        End Get
        Set
            _originatingNavGraphName = value
        End Set
    End Property
End Class

#Region "StartTask Event Args and Delegate"
' Represents the method that will handle the StartTask event
Delegate Sub StartTaskEventHandler(sender As Object, e As StartTaskEventArgs)

' Provides data for the StartTask event
Friend Class StartTaskEventArgs
    Inherits EventArgs
    Implements ITask
    Private _nextNavGraphName As String = ""
    Private _taskArguments As TaskArgumentsHolder = Nothing
    Private _nextTaskID As Guid = Guid.Empty
    
    Public Sub New(nextNavigationGraphName As String)
        MyClass.New(nextNavigationGraphName, Nothing, Guid.Empty)
    End Sub
        
    'Constructor that assumes we are chaining nav graphs, and need all information relevant to have a "return pointer" on a "stack"
    'Parameters:
    '-nextNavigationGraphName: next graph
    '-taskArguments: task arguments
    '-nextTaskID: If we know we are returning to a nav graph we've been before, and wish to enter at a known point
    Public Sub New(nextNavigationGraphName As String, taskArguments As TaskArgumentsHolder, nextTaskID As Guid)
        _nextNavGraphName = nextNavigationGraphName
        _taskArguments = taskArguments
        _nextTaskID = nextTaskID
    End Sub
    
    Public Property NextNavigationGraph() As String
        Get
            Return _nextNavGraphName
        End Get
        Set
            _nextNavGraphName = value
        End Set
    End Property 
    
    Public Property TaskArguments() As TaskArgumentsHolder
        Get
            Return _taskArguments
        End Get
        Set
            _taskArguments = value
        End Set
    End Property 
        
    Public Property NextTaskID() As Guid
        Get
            Return _nextTaskID
        End Get
        Set
            _nextTaskID = value
        End Set
    End Property 
    
    #Region "ITask Members"
    Public Function [Get]() As Guid Implements ITask.Get 
        Return _nextTaskID
    End Function 

    Public Sub Create(taskId As Guid) Implements ITask.Create
    End Sub   
    
    #End Region
End Class
#End Region

#Region "Controller abstract class definition"
'This class coordinates the user process.
'Represents the controller in a Model-View-Controller pattern.
MustInherit Public Class ControllerBase
   #Region "Declares variables"
   
   Private _state As State
   
   #End Region
   
   #Region "Constructors"
   
   Private Sub New()
   End Sub

   'Constructor
   'Parameters:
   '-controllerState: State object, encapsulating all state for this interaction graph
   Protected Sub New(controllerState As State)
      Me._state = controllerState
   End Sub
   #End Region
   
   #Region "Events"
   'Occurs before the controller navigates to the next view
   Friend Event BeforeNavigate As EventHandler
   
    'Occurs when a task has started
   Friend Event StartTask As StartTaskEventHandler
   #End Region
   
   'Gets the controller state
   Public ReadOnly Property State() As State
      Get
         Return _state
      End Get
   End Property
 
   'Simply allows for a default "next" button that causes Manager to navigate to first next page.
   Protected Overridable Sub Navigate()
      RaiseEvent BeforeNavigate(Me, EventArgs.Empty)
   End Sub
   
   Protected Overridable Sub OnEndTask()
      State.Clear()
   End Sub
   
   'Navigates to the next navigation graph.
   Overloads Protected Sub OnStartTask(nextNavigationGraphName As String)
        RaiseEvent StartTask(Me, New StartTaskEventArgs(nextNavigationGraphName))
   End Sub
     
   'Navigates to the next navigation graph; passes the Object property bag.
   Overloads Protected Sub OnStartTask(nextNavigationGraphName As String, taskArguments As TaskArgumentsHolder, nextTaskID As Guid)
        RaiseEvent StartTask(Me, New StartTaskEventArgs(nextNavigationGraphName, taskArguments, nextTaskID))
   End Sub
      
   'This method is called by the UIPManager when a new task starts
   'Parameters:
   '-taskArguments: A holder for originating navgraph and taskid, and an object for other "stuff" that
   'will be used by the controller to get state information from the previous nav graph</param>
   Public Overridable Sub EnterTask(taskArguments As TaskArgumentsHolder)
   End Sub
End Class
#End Region
