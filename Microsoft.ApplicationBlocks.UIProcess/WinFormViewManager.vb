'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' WinFormViewManager.vb
'
' This file contains the implementations of the WinFormViewManager class
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
Imports System.Collections
Imports System.Windows.Forms

#Region "WinFormViewManager class definition"
'Provides methods to manipulate winform views 
Friend Class WinFormViewManager
    Implements IViewManager
    #Region "Declares variables"
    Private Const CommaSeparator As String = ","
    Private Const ParentFormKey As String = "ParentForm"

    'Stores active forms 
    Private Shared ActiveForms As New Hashtable()

    'Stores active views
    Private Shared ActiveViews As New Hashtable()

    'Stores active views
    Private Shared Properties As New Hashtable()
    #End Region

    #Region "Constructor"
    Public Sub New()
    End Sub
    #End Region

    #Region "IViewManager Members"
    Private Function GetProperty(taskId As Guid, name As String) As Object 
        Dim taskProperties As Hashtable = CType(Properties(taskId), Hashtable)
        If Not (taskProperties Is Nothing) Then
            Return taskProperties(name)
        Else
            Return Nothing
        End If
    End Function

    'Stores a property into the view manager. 
    'Each task has its own properties
    'The property storage is a view manager responsibility
    'Parameters: 
    '-taskId: task identifier
    '-name: property name
    '-value: property value
    Public Sub StoreProperty(taskId As Guid, name As String, value As Object) Implements IViewManager.StoreProperty
        If Properties(taskId) Is Nothing Then
            Properties(taskId) = New Hashtable()
        End If 
        CType(Properties(taskId), Hashtable)(name) = value
    End Sub

    'Activates a specific view
    'Parameters: 
    '-previousView: the view actually displayed
    '-taskId: a existing task id
    '-navigationGraph: a configured navigation graph name
    '-view: the view name to be displayed
    Public Sub ActivateView(previousView As String, taskId As Guid, navigationGraph As String, view As String) Implements IViewManager.ActivateView
        If ActiveForms(taskId) Is Nothing Then
            ActiveForms(taskId) = New Hashtable()
            ActiveViews(taskId) = New Hashtable()
        End If
        
        Dim taskActiveForms As Hashtable = CType(ActiveForms(taskId), Hashtable)
        Dim taskActiveViews As Hashtable = CType(ActiveViews(taskId), Hashtable)
        
        Dim winFormView As WinFormView = CType(taskActiveForms(view), WinFormView)
        If Not (winFormView Is Nothing) Then
            'Use the existing instance
            winFormView.Activate()
        Else
            
            'Create a new instance
            Dim viewSettings As ViewSettings = UIPConfiguration.Config.GetViewSettingsFromName(view)
            If viewSettings Is Nothing Then
                Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionViewConfigNotFound", view))
            End If 
            winFormView = CType(GenericFactory.Create(viewSettings), WinFormView)
            winFormView.InternalTaskId = taskId
            winFormView.InternalNavigationGraph = navigationGraph
            winFormView.InternalViewName = view
            
            taskActiveForms(view) = winFormView
            taskActiveViews(winFormView) = view
            
            AddHandler winFormView.Activated, AddressOf Form_Activated
            AddHandler winFormView.Closed, AddressOf Form_Closed
            
            'Get the parent form
            Dim parentForm As Form = CType(GetProperty(taskId, ParentFormKey), Form)
            If Not (parentForm Is Nothing) AndAlso parentForm.IsMdiContainer = True Then
                winFormView.MdiParent = parentForm
            End If 
            
            If viewSettings.IsOpenModal Then
			    If Not previousView Is Nothing Then
				    ' Get the current form
					Dim previousViewSettings As ViewSettings = UIPConfiguration.Config.GetViewSettingsFromName(previousView)
					If previousViewSettings Is Nothing Then
					    Throw New UIPException( Resource.ResourceManager.FormatMessage( "RES_ExceptionViewConfigNotFound", previousView ) )
                    End If                
					Dim previousForm As WinFormView = CType(taskActiveForms(previousView), WinFormView)
					
					winFormView.ShowDialog(CType(previousForm, IWin32Window) )
				Else
					'the previous view is unknown, so the first view of the navgraph is modal, 
					' as a last resort we try to get the parentForm from our properties.
					winFormView.ShowDialog(CType(parentForm, IWin32Window ))
				End If
			Else
				winFormView.Show()
            End If
        End If
        
        If Not previousView Is Nothing AndAlso previousView.Length <> 0 Then
            'Get the current form
            Dim previousViewSettings As ViewSettings = UIPConfiguration.Config.GetViewSettingsFromName(previousView)
            If previousViewSettings Is Nothing Then
                throw new UIPException( Resource.ResourceManager.FormatMessage( "RES_ExceptionViewConfigNotFound", previousView ) )
            End If    
            
            Dim previousForm As WinFormView = CType(taskActiveForms(previousView), WinFormView)
            If Not previousForm Is Nothing AndAlso Not previousViewSettings.IsStayOpen Then
                'The current window must be closed
                previousForm.Close()
            End If
        End If
    End Sub

    'Utility method that checks web requests to ensure that requested page and current view match.
    'If user bookmarks page D, then proceeds to page F, then returns to bookmark--
    'State when loaded will have F as CurrentView.  
    'Any submissions on page D will fail, because navigation graph may not have appropriate view-navigateResult pairs.
    'THEREFORE, code defensively against this.  Check current page, check referrer, check state object's currentview
    'Parameters: 
    '-view: the next view
    '-stateView: the view saved in the state
    Public Function IsRequestCurrentView(view As IView, stateViewName As String) As Boolean Implements IViewManager.IsRequestCurrentView
        ' Not implemented. In winform applications isnt necesary 
        Return True
    End Function

    'Gets the running tasks in the manager
    'Returns: a array with the task identifiers
    Public Function GetCurrentTasks() As Guid() Implements IViewManager.GetCurrentTasks
        Dim tasks As New ArrayList()
        Dim key As Guid
        For Each key In  ActiveViews.Keys
            tasks.Add(key)
        Next key
        
        Return CType(tasks.ToArray(GetType(Guid)), Guid())
    End Function
    #End Region

    #Region "WinForm Event Handlers"
    'Updates the current view
    'Parameters: 
    Private Sub Form_Activated([source] As Object, e As EventArgs)
        Dim winFormView As WinFormView = CType([source], WinFormView)
        Dim state As State = winFormView.Controller.State
        
        'Get the views related to the current task
        Dim taskActiveViews As Hashtable = CType(ActiveViews(state.TaskId), Hashtable)
        
        'Get the view related to the form that fires this event
        Dim currentView As String = CStr(taskActiveViews([source]))
        
        'Update the state current view
        If Not (currentView Is Nothing) Then
            state.CurrentView = currentView
        End If
    End Sub

    'Removes the closed form from the collection of active forms
    Private Sub Form_Closed([source] As Object, e As EventArgs)
        Dim winFormView As IView = CType([source], IView)
        
        'Get the views related to the current task
        Dim taskActiveViews As Hashtable = CType(ActiveViews(winFormView.TaskId), Hashtable)
        Dim taskActiveForms As Hashtable = CType(ActiveForms(winFormView.TaskId), Hashtable)
        
        'Get the view related to the form that fires this event
        Dim currentView As String = CStr(taskActiveViews([source]))
        
        'Remove the view and its form
        If Not (currentView Is Nothing) Then
            taskActiveForms.Remove(currentView)
            taskActiveViews.Remove([source])
        End If
    End Sub
    #End Region
End Class
#End Region
