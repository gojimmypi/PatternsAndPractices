'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' WebFormViewManager.vb
'
' This file contains the implementations of the SessionMoniker and WebFormViewManager classes
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
Imports System.Collections
Imports System.Globalization

#Region "SessionMoniker"
'Internal class used to store session information
Friend Class SessionMoniker
    #Region "Variable Declarations"
    Private Const ColonSeparator As String = ":"
    Private Const SessionTaskMonikerKey As String = "TaskMoniker"

    Private _navigationGraph As String
    Private _currentview As String
    Private _taskId As Guid
    #End Region

    #Region "Constructors"
    Public Sub New(navigationGraphName As String, currentViewName As String, taskId As Guid)
        _navigationGraph = navigationGraphName
        _currentview = currentViewName
        _taskId = taskId
    End Sub

    Public Sub New(navigationGraphName As String, currentViewName As String, taskId As String)
        _navigationGraph = navigationGraphName
        _currentview = currentViewName
        _taskId = New Guid(taskId)
    End Sub

    Public Sub New(taskMoniker As String)
        '  split string on colons
        Dim armoniker As String() = taskMoniker.Split(ColonSeparator.ToCharArray())
        
        '  check if it has expected 3 items
        If 2 = armoniker.GetUpperBound(0) Then
            _navigationGraph = armoniker(0)
            _currentview = armoniker(1)
            _taskId = New Guid(armoniker(2))
        Else
            Throw New ArgumentException(Resource.ResourceManager("RES_ExceptionIncorrectNumberOfItemsInTaskMonikerString"))
        End If
    End Sub
    #End Region

    #Region "Static helper methods"
    'Tests a moniker retrieved from context to see if it matches the pattern "NavGraphName:CurrentViewName:TaskGuid"
    Private Shared Function IsMonikerValid(moniker As String) As Boolean
        '  check if it's null or zero length first
        If Nothing = moniker OrElse 0 = moniker.Length Then
            Return False
        End If 
        '  potential security risk, this input comes from user
        Dim armon As String() = moniker.Split(ColonSeparator.ToCharArray())
        
        '  check for correct # elements
        If Not 2 = armon.GetUpperBound(0) Then
            Return False
        End If 
        '  put into string args
        Dim navgraph As String = armon(0)
        Dim view As String = armon(1)
        Dim task As String = armon(2)
        
        '  check lengths
        If navgraph.Length < 255 AndAlso view.Length < 255 AndAlso task.Length = 36 Then
            Return True
        Else
            Return False
        End If
    End Function

    'Gets the session moniker related to the specified task
    Public Shared Function GetFromSession(taskId As Guid) As SessionMoniker
        Dim stringMoniker As String = CStr(HttpContext.Current.Session((SessionTaskMonikerKey + taskId.ToString())))
        If IsMonikerValid(stringMoniker) Then
            Return New SessionMoniker(stringMoniker)
        Else
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantGetSessionMoniker", taskId.ToString()))
        End If
    End Function

    'Gets all session monikers actually stored in the user session
    Public Shared Function GetAllFromSession() As SessionMoniker()
        Dim monikers As New ArrayList()
        Dim sm As SessionMoniker
        Dim key As String
        For Each key In  HttpContext.Current.Session.Keys
            If key.StartsWith(SessionTaskMonikerKey) Then
            sm = CType(HttpContext.Current.Session(key), SessionMoniker)
            monikers.Add(sm)
            End If
        Next key
        
        Return CType(monikers.ToArray(GetType(SessionMoniker)), SessionMoniker())
    End Function
    #End Region

    #Region "Public methods"
    'Stores the session moniker in the user session
    Public Sub StoreInSession()
        HttpContext.Current.Session((SessionTaskMonikerKey + Me.TaskId.ToString())) = Me.ToString()
    End Sub
    #End Region

    #Region "Public Properties Get/Set"
    Public Property NavGraphName() As String
        Get
            Return _navigationGraph
        End Get
        Set
            _navigationGraph = value
        End Set
    End Property

    Public Property CurrentViewName() As String
        Get
            Return _currentview
        End Get
        Set
            _currentview = value
        End Set
    End Property

    Public Property TaskId() As Guid
        Get
            Return _taskId
        End Get
        Set
            _taskId = value
        End Set
    End Property
    #End Region

    #Region "ToString Override"

    Public Overrides Function ToString() As String
        Return Me._navigationGraph + ":" + Me._currentview + ":" + Me._taskId.ToString("", CultureInfo.CurrentCulture)
    End Function
    #End Region
End Class 
#End Region

#Region "WebFormViewManager class definition"
'Provides methods to manipulate webform views 
Friend Class WebFormViewManager
    Implements IViewManager

    Public Sub New()
    End Sub

    #Region "IViewManager Members"
    'Activates a specific view
    'Parameters: 
    '-previousView: the view actually displayed
    '-taskId: a existing task id
    '-navigationGraph: a configured navigation graph name
    '-view: the view name to be displayed
    Public Sub ActivateView(previousView As String, taskId As Guid, navGraph As String, view As String) Implements IViewManager.ActivateView
        
        '  create a session moniker
        Dim sessionMoniker As New SessionMoniker(navGraph, view, taskId)
        
        ' Stores the Moniker into the Session
        sessionMoniker.StoreInSession()
        
        Dim viewSettings As ViewSettings = UIPConfiguration.Config.GetViewSettingsFromName(view)
        If viewSettings Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionViewConfigNotFound", view))
        End If 
        Dim queryString As String = WebFormView.CurrentTaskKey + "=" + taskId.ToString()
        
        '  ThreadAbortException par for course on Redirect...trap here and squelch
        '  used "false" param to allow thread to continue execution after Redirect.
        '  this means cleanup work etc. can continue and we're not throwing a TAE every time we redirect that aspnet has to swallow up.
        Try
            If previousView Is Nothing Then
                HttpContext.Current.Response.Redirect(HttpContext.Current.Request.ApplicationPath + "/" + viewSettings.Type + "?" + queryString, True)
            Else
                HttpContext.Current.Response.Redirect(HttpContext.Current.Request.ApplicationPath + "/" + viewSettings.Type + "?" + queryString, False)
            End If
        Catch e as System.Threading.ThreadAbortException 
        End Try
    End Sub

    'Stores a property into the view manager. 
    'Each task has its own properties
    'The property storage is a view manager responsibility
    'Parameters: 
    '-taskId: task identifier    
    '-name: property name
    '-value: property value
    Public Sub StoreProperty(taskId As Guid, name As String, value As Object) Implements IViewManager.StoreProperty 
    End Sub

    'Utility method that checks web requests to ensure that requested page and current view match.
    'If user bookmarks page D, then proceeds to page F, then returns to bookmark
    'State when loaded will have F as CurrentView.  
    'Any submissions on page D will fail, because navigation graph may not have appropriate view-navigateResult pairs.
    'THEREFORE, code defensively against this.  Check current page, check referrer, check state object's currentview
    'Parameters: 
    '-view: the next view
    '-stateViewName: the view saved in the state
    Public Function IsRequestCurrentView(view As IView, stateViewName As String) As Boolean Implements IViewManager.IsRequestCurrentView 
        '  get state currentview; must all match
        Dim viewSettings As ViewSettings = UIPConfiguration.Config.GetViewSettingsFromName(stateViewName)
        If viewSettings Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionViewConfigNotFound", stateViewName))
        End If 
        Dim stateViewType As String = viewSettings.Type
        
        Dim page As System.Web.UI.Page = CType(view, System.Web.UI.Page)
        Dim viewType As String = page.Request.CurrentExecutionFilePath.Replace(page.Request.ApplicationPath + "/", "")
        viewType = viewType.ToLower(System.Globalization.CultureInfo.CurrentUICulture)
        
        If stateViewType.ToLower(System.Globalization.CultureInfo.CurrentUICulture).Equals(viewType) Then
            Return True
        Else
            Return False
        End If
    End Function

    'Gets the running tasks in the manager
    'Returns: a array with the task identifiers
    Public Function GetCurrentTasks() As Guid() Implements IViewManager.GetCurrentTasks 
        Dim monikers As SessionMoniker() = SessionMoniker.GetAllFromSession()
        Dim tasks(monikers.Length) As Guid
        Dim index As Integer
        For index = 0 To monikers.Length - 1
            tasks(index) = monikers(index).TaskId
        Next index
        
        Return tasks
    End Function
    #End Region
End Class
#End Region
