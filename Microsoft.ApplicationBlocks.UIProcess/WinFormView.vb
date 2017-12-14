'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' WinFormView.vb
'
' This file contains the implementations of the WinFormView class
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

#Region "WinFormView class definition"
'Represents a view used in windows applications
Public Class WinFormView
    Inherits Form
    Implements IView
    #Region "Declares variables"
    Private _controller As ControllerBase
    Private _taskId As Guid
    Private _navigationGraph As String
    Private _viewName As String
    #End Region

    Public Sub New()
        AddHandler Me.Load, AddressOf WinFormViewOnLoad
    End Sub

    #Region "IView implementation"
    'Gets the task id related to this view
    ReadOnly Property TaskId() As Guid Implements IView.TaskId
        Get
            Return _taskId
        End Get 
    End Property 
    
    'Gets the current view navigation graph. 
    'This view is actually shown in this navigation graph
    ReadOnly Property NavigationGraph() As String Implements IView.NavigationGraph
        Get
            Return _navigationGraph
        End Get 
    End Property 
    
    'Gets the view name
    ReadOnly Property ViewName() As String Implements IView.ViewName
        Get
            Return _viewName
        End Get 
    End Property 
    
    'Gets the view controller
    Public ReadOnly Property Controller() As ControllerBase Implements IView.Controller 
        Get
            Return _controller
        End Get
    End Property
    #End Region

    #Region "Internal properties"
    'Gets the current task identifier
    Friend WriteOnly Property InternalTaskId() As Guid
        Set
            _taskId = value
        End Set
    End Property 

    'Gets the current navigation graph
    Friend WriteOnly Property InternalNavigationGraph() As String
        Set
            _navigationGraph = value
        End Set
    End Property 

    'Gets the view name
    Friend WriteOnly Property InternalViewName() As String
        Set
            _viewName = value
        End Set
    End Property
    #End Region

    Public Sub WinFormViewOnLoad([source] As Object, e As EventArgs)
        '  because all WinForms in UIP apps inherit from this, 
        '  we have to be conscious of design-time problems.
        '  The full UIP can't be invoked when designing the form, so short-circuit here to avoid design time exception
        If True = Me.DesignMode Then
            _controller = Nothing
        Else
            _controller = UIPManager.InitializeController(Me)
        End If
    End Sub
End Class
#End Region
