'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' WebFormView.vb
'
' This file contains the implementations of the WebFormView class
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
Imports System.Web.UI

#Region "WebFormView class definition"
'Represents a view used in web applications
Public Class WebFormView
    Inherits Page
    Implements IView

    #Region "Declares variables"
    Private _controller As ControllerBase
    Private _sessionMoniker As SessionMoniker

    'QueryString key used to get the current task id
    Public Const CurrentTaskKey As String = "CurrentTask"
    #End Region

    #Region "Constructor"
    Public Sub New()
        AddHandler Me.Load, AddressOf WebFormView_Load
    End Sub
    #End Region

    #Region "IView implementation "
    'Gets the view controller
    Public ReadOnly Property Controller() As ControllerBase Implements IView.Controller
        Get
            Return _controller
        End Get
    End Property 

    'Gets the current view navigation graph. 
    'This view is actually shown in this navigation graph
    ReadOnly Property NavigationGraph() As String Implements IView.NavigationGraph
        Get
            Return _sessionMoniker.NavGraphName
        End Get
    End Property 

    'Gets the task id related to this view
    ReadOnly Property TaskId() As Guid Implements IView.TaskId
        Get
            Return _sessionMoniker.TaskId
        End Get
    End Property 

    'Gets the view name
    ReadOnly Property ViewName() As String Implements IView.ViewName
        Get
            Return _sessionMoniker.CurrentViewName
        End Get
    End Property 

    #End Region

    Private Sub WebFormView_Load(sender As Object, e As System.EventArgs)
        _sessionMoniker = GetSessionMoniker()
        _controller = UIPManager.InitializeController(Me)
    End Sub

    Private Function GetSessionMoniker() As SessionMoniker
        Dim sessionMoniker As SessionMoniker = SessionMoniker.GetFromSession(New Guid(Request.QueryString(CurrentTaskKey)))
        Return sessionMoniker
    End Function 
End Class
#End Region
