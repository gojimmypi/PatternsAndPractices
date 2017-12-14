'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' UIPManager.cs
'
' This file contains the implementations of the UIPManager class
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
Imports System.Windows.Forms
Imports System.Diagnostics

#Region "UIPManager class"
'Manager dispenses Controllers to Views, senses when Controllers have finished, spawns new Views, 
'coordinates Task
NotInheritable Public Class UIPManager
    #Region "Variable Declarations"
    Private Const CommaSeparator As String = ","
    Private Const ParentFormKey As String = "ParentForm"
    #End Region

    #Region "Constructors "

    Shared Sub New()
    End Sub

    #End Region

    #Region "InitializeController & StartTask "
    'Creates and intitializes a controller for the specified view
    'Parameters: 
    '-view: a valid view object
    'Returns: a controller object
    Public Shared Function InitializeController(view As IView) As ControllerBase
        Dim ctrl As ControllerBase = Nothing
        Dim viewManager As IViewManager = Nothing
        Dim state As State = Nothing
        
        ' create a new controller object for the specified view
        ctrl = ControllerFactory.Create(view)
        
        AddHandler ctrl.BeforeNavigate, AddressOf Navigate
        
        'add EndTaskEvent sink so Controller can fire event to trigger next Navigation Graph
		AddHandler ctrl.StartTask, AddressOf StartTask 
        
        '  create view manager
        viewManager = ViewManagerFactory.Create(view.NavigationGraph)
        
        '  grab reference to the Controller's State object so we can ask it stuff
        state = ctrl.State
        
        '  query ViewManager...if it cares about out-of-order View accesses, such as "Back" button on browser or 
        '  out-of-bounds extra session opened "ctrl-N", then it can code defensively against that...and perhaps force-open correct view
        If Not viewManager.IsRequestCurrentView(view, state.CurrentView) Then
            viewManager.ActivateView(Nothing, state.TaskId, state.NavigationGraph, state.CurrentView)
        End If 
        '  finally, return requested Controller instance
        Return ctrl
    End Function 'InitializeController

    'Overload for StartTask that sinks EndTask event fired by Controller--permits us 
	'to chain NavGraphs w/out ANY nav code in Views, Controller does All
	Overloads Private Shared Sub StartTask(sender As Object, e As StartTaskEventArgs)
	    StartTask( Nothing, e.NextNavigationGraph, CType(e, ITask), e.TaskArguments )
	End Sub

    'Starts a new task
    'Parameters: 
    '-navigationGraphName: navigation graph that will be used for the task
    Overloads Public Shared Sub StartTask(navigationGraphName As String)
        StartTask(Nothing, navigationGraphName, Nothing, Nothing)
    End Sub

    'Starts a new task
    'This overload can be used only from MDI Applications
    'Parameters: 
    '-parentForm: MDI parent form
    '-navigationGraphName: navigation graph that will be used for the task
    Overloads Public Shared Sub StartTask(parentForm As Form, navigationGraphName As String)
        StartTask(parentForm, navigationGraphName, Nothing, Nothing)
    End Sub

    'Starts a new task
    'Parameters: 
    '-navigationGraphName: navigation graph that will be used for the task
    '-task: A task object used to get the task id
    Overloads Public Shared Sub StartTask(navigationGraphName As String, task As ITask)
        StartTask(Nothing, navigationGraphName, task, Nothing)
    End Sub

    'Starts a new task
    'This overload can be used only from MDI Applications
    'Parameters: 
    '-parentForm: MDI parent form
    '-navigationGraphName: navigation graph that will be used for the task
    '-task: A task object used to get the task id
    '-taskArguments:
    Overloads Public Shared Sub StartTask(parentForm As Form, navigationGraphName As String, task As ITask, taskArguments As TaskArgumentsHolder)
        Dim taskId As Guid = Guid.Empty
		Dim	state As State = Nothing
		Dim viewManager As IViewManager = Nothing
		Dim previousView As String = Nothing
            
		'  THREE-WAY DECISION:
		'  1)  no ITask object--assume new Task, create new TaskID, create new State
		'  2)  ITask object, but no valid TaskID--assume new Task, create new State
		'  3)  ITask object, valid TaskID--assume known Task, retrieve known State
		'
		'  CASE (1):  
		'  if incoming ITask is null, we KNOW this is a new Task.  Act appropriately--create new Task/State
		If task Is Nothing Then 
            state = StateFactory.Create( navigationGraphName )
			taskId = state.TaskId
		Else
			'  ask the incoming ITask if it has a TaskID already.  
			'  this would mean the application has already hooked up a Task ID to whatever representation it uses...
			'  for example, correlating TaskID to windows logon
				
			taskId = task.Get()

			If taskId.Equals(Guid.Empty) Then
			
                'CASE (2)  
				'  OK, the application has not pre-set a task.  Therefore tell application what the new TaskId is, 
				'  and internally the client app will use it...for example, by creating an entry 
				'  in its DB lookup table to correlate logon with Task.

				'  set up the new State object, since we know now that we're on a new Task so we need new State
				state = StateFactory.Create( navigationGraphName )
				
                '  new State now contains new TaskID, get it here
				taskId = state.TaskId
				
                '  tell the taskObject (wich is our sink back into client application) about the new TaskID
				task.Create(taskId)
            Else
				
				'  CASE (3)
				'  in this case, ITask was not null, AND it returned a valid TaskID...
				'  SO this is a known Task...and there is a State for it.  Retrieve that State.
				state = StateFactory.Load( navigationGraphName, taskId ) 
			End If
        End If

		'  now, the new State may or may not already know what its CurrentView is.  If it does not, it must;
		'  here we check that, and if it does not know its first View we must tell it to be the first view of the new navgraph:
		If state.CurrentView.Length = 0 Then
		    '  get first view name/type
			Dim viewSettings As ViewSettings = UIPConfiguration.Config.GetFirstViewSettings(navigationGraphName)
			If viewSettings Is Nothing Then
			    Throw New UIPException( Resource.ResourceManager.FormatMessage( "RES_ExceptionStartViewConfigNotFound", navigationGraphName ) )
            End If        			      
			
            '  populate and persist the new State object
			state.CurrentView = viewSettings.Name
			state.NavigationGraph = navigationGraphName
			state.NavigateValue = ""
    				
			'  persist
			state.Save()
		End If

		'get the controller for the NEXT NavGraph
		'  NOTE:  we use the nav graph name passed to us; the ORIGINATING navgraph says "start task on next nav, here it is 'navB'";
		'  also, the ORIGINATING navgraph says "I have a return TaskID for the new navgraph; we've been there before"
		'  OR it might say "I don't have an originating task ID for the 'new' nav graph, start it fresh..."
		'  either way, controller factory semantics take care of creating appropriate State object for us.
		Dim controller As ControllerBase = ControllerFactory.Create( navigationGraphName, taskId )

		'  now initialize this next controller, let it do what it needs to with the task arguments
		controller.EnterTask( taskArguments )
		
        '  finally save the state attached to this controller, the controller EnterTask() may have put important state
		'  from the previous nav graph into this NEXT controller's state
		controller.State.Save()

		'  grab reference to this state; REMEMBER, the state object was saved before, then we created controller
		'  to allow controller to put new stuff in state (stuff from previous navgraph); so now it may be different
		state = controller.State
		
    	'  OK now we have latest copy of state.
		'  create view manager
		viewManager = ViewManagerFactory.Create( navigationGraphName )
            
		Try
		    If Not parentForm Is Nothing Then
			    viewManager.StoreProperty( state.TaskId, ParentFormKey, parentForm )
			End If
		Catch e As  Exception
		    Throw New UIPException( Resource.ResourceManager("RES_ExceptionCantSetViewProperty"), e )
		End Try

		Try
		    viewManager.ActivateView( previousView, state.TaskId, state.NavigationGraph, state.CurrentView )
        Catch e As System.Threading.ThreadAbortException
		Catch e As Exception
		    Throw New UIPException( Resource.ResourceManager.FormatMessage( "RES_ExceptionCantActivateView", state.CurrentView ), e )
		End Try
    End Sub
    #End Region

    #Region "Navigation Code"

    'Handler for the navigate method
    'Parameters:
    '-source: the controller that raises the event
    '-e:event argument
    Private Shared Sub Navigate([source] As Object, e As EventArgs)
        Dim controller As ControllerBase = CType([source], ControllerBase)
        Dim state As State = controller.State
        Dim previousView As String = state.CurrentView
        
        '  create view manager
        Dim viewManager As IViewManager = ViewManagerFactory.Create(state.NavigationGraph)
        
        '  CurrentView reflects THIS page; we are about to go to next page, so 
        '  get that next page's name now and put it in currentview
        Dim nextView As ViewSettings = UIPConfiguration.Config.GetNextViewSettings(state.NavigationGraph, state.CurrentView, state.NavigateValue)
        
        If nextView Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionCouldNotGetNextViewType", state.NavigationGraph, state.CurrentView, state.NavigateValue))
        End If 
        
        state.CurrentView = nextView.Name
        state.NavigateValue = ""
        state.Save()
        
        Try
            viewManager.ActivateView(previousView, state.TaskId, state.NavigationGraph, state.CurrentView)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantActivateView", nextView.Name), ex)
        End Try
    End Sub
    #End Region
End Class
#End Region
