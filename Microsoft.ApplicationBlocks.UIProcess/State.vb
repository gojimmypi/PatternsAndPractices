'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' State.vb
'
' This file contains the implementations of the State class
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
Imports System.Runtime.Serialization
Imports System.Security.Permissions

#Region "CacheEntry class definition"

'This class represents a entry in the state cache
Friend Class CacheEntry
    Private _itemAge As DateTime = DateTime.MinValue
    Private _value As Object
   
    'Constructor
    'Parameters: 
    '-value: item value
    '-itemAge: item age. Specifies when expires the item
    Public Sub New(value As Object, itemAge As DateTime)
        _itemAge = itemAge
        _value = value
    End Sub
      
    'Constructor
    'The item never expires
    'Parameters: 
    '-value: item value
    Public Sub New(value As Object)
        _value = value
    End Sub
   
    'Specifies if the item has expired
    Public ReadOnly Property IsExpired() As Boolean
        Get
            If _itemAge.Equals(DateTime.MinValue) Then 'The item never expires
                Return False
            Else
                Return DateTime.Now > _itemAge
            End If
        End Get
    End Property 
   
    'Gets the item value
    Public ReadOnly Property Value() As Object
        Get
            Return _value
        End Get
    End Property
End Class
#End Region

#Region "StateChangedEventArgs class definition"
'Provides data for StateChanged event
Public Class StateChangedEventArgs
   Inherits EventArgs
   
   
   'Constructor
   'Parameters: 
   '-key: state item key
   Public Sub New(key As String)
      _key = key
   End Sub 'New
   
   Private _key As String
   
   'Gets the changed state item key
   Public ReadOnly Property Key() As String
      Get
         Return _key
      End Get
   End Property
End Class
#End Region

#Region "State class definition"
'This class mantains user process state. Represents the model in a 
'Model-View-Controller pattern
'This class must be serializable. If a derived class requires a 
'complex serialization mecanism, then it must implement the ISerializable interface.
'NOTE also that derived classes must call base GetObjectData and serialization constructor appropriately to ensure
'full serialize/deserialize.
<Serializable()>  _
Public Class State
    Inherits DictionaryBase
    Implements ISerializable

    #Region "Declare variables"
    Private _taskId As Guid = Guid.Empty
    Private _currentView As String = ""
    Private _navigationGraph As String = ""
    Private _navigateValue As String = ""

    <NonSerialized()>  _
    Private _stateVisitor As IStatePersistence = Nothing
    Private Const CommaSeparator As String = ","
    '  string names used to put local property values into serialization stream.
    '  GUID appended to end to make collisions with InnerHashtable-items VERY unlikely...mstuart 03.24.2003
    Private ReadOnly _tagGuid1 As New Guid( new byte() {&H5F, &H4D, &H69, &H63, &H68, &H61, &H65, &H6C, &H20, &H53, &H74, &H75, &H61, &H72, &H74, &H5F } )

    Private Const NameCurrentView As String = "_currentView_{FF9B8CB4-E13B-44a7-B3C6-B385D8EB8167}"
    Private Const NameNavigationGraph As String = "_navigationGraph_{FF9B8CB4-E13B-44a7-B3C6-B385D8EB8167}"
    Private Const NameNavigationValue As String = "_navigateValue_{FF9B8CB4-E13B-44a7-B3C6-B385D8EB8167}"
    Private Const NameTaskId As String = "_taskId_{FF9B8CB4-E13B-44a7-B3C6-B385D8EB8167}"
    #End Region
   
    #Region "Constructors"
    'Constructor
    'Parameters: 
    '-statePersistenceProvider: A valid State persistence provider
    Public Sub New(statePersistenceProvider As IStatePersistence)
        Me.Accept(statePersistenceProvider)
    End Sub
       
    'Constructor
    Public Sub New()
        MyClass.New(Guid.Empty, Nothing, Nothing, Nothing, Nothing)
    End Sub
       
    'Constructor
    'Parameters: 
    '-taskId: A task identifier
    Public Sub New(taskId As Guid)
        MyClass.New(taskId, Nothing, Nothing, Nothing, Nothing)
    End Sub
       
    'Constructor
    'Parameters: 
    '-taskId: A task identifier
    '-navGraph: A valid navigation graph name
    Public Sub New(taskId As Guid, navGraph As String)
        MyClass.New(taskId, navGraph, Nothing, Nothing, Nothing)
    End Sub
    
    'Constructor
    'Parameters: 
    '-taskId: A task identifier
    '-navGraph: A valid navigation graph name
    '-currentView: The current view in the navigation graph 
    Public Sub New(taskId As Guid, navGraph As String, currentView As String)
        MyClass.New(taskId, navGraph, currentView, Nothing, Nothing)
    End Sub
    
    'Constructor
    'Parameters: 
    '-taskId: A task identifier
    '-navGraph: A valid navigation graph name
    '-currentView: The current view in the navigation graph 
    '-navigateValue: Used by the controller to determine wich is the next view
    '-statePersistenceProvider: A valid State persistence provider
    Public Sub New(taskId As Guid, navigationGraph As String, currentView As String, navigateValue As String, statePersistence As IStatePersistence)
        Me._taskId = taskId
        Me._navigationGraph = navigationGraph
        Me._currentView = currentView
        Me._navigateValue = navigateValue
        Me.Accept(statePersistence)
    End Sub
      
    'ISerializable-required constructor.  
    'Parameters: 
    '-si: 
    '-context: 
    <SecurityPermissionAttribute(SecurityAction.Demand, SerializationFormatter := True), SecurityPermissionAttribute(SecurityAction.LinkDemand, Flags := SecurityPermissionFlag.SerializationFormatter)>  _
    Protected Sub New(si As SerializationInfo, context As StreamingContext) 
        Dim name As String = ""
		Dim	tag As String = _tagGuid1.ToString()
		Dim tagIndex As Integer = 0

		' iterate over contents of serialization info; 
		' put each key-value pair back into our inner hashtable
		Dim se As SerializationEntry
        For each se in si 
		    name = se.Name
			tagIndex = name.IndexOf( tag )
			
			' check that the name ends in our tag guid; we tag all hashtable items with the tag guid to 
			' allow distinguishing regular items, and items added by derived classes, from our actual hashtable items
			If tagIndex > 0 Then
				me.InnerHashtable.Add( name.Substring(0, tagIndex), se.Value)
			End If
        Next

		'  deserialize the rest of our properties
		me._currentView = si.GetString(NameCurrentView)
		me._navigationGraph = si.GetString(NameNavigationGraph)
		me._navigateValue = si.GetString(NameNavigationValue)
		me._taskId = CType(si.GetValue(NameTaskId, GetType(System.Guid)), Guid)
    End Sub
    #End Region
   
    #Region "ISerializable Members"
    'Required "GetObjectData" of ISerializable.  Packages class info into a SerializationInfo.
    'Parameters: 
    '-info: 
    '-context: 
    <SecurityPermissionAttribute(SecurityAction.Demand, SerializationFormatter := True), SecurityPermissionAttribute(SecurityAction.LinkDemand, Flags := SecurityPermissionFlag.SerializationFormatter)>  _
    Public Overridable Sub GetObjectData(info As SerializationInfo, context As StreamingContext) Implements ISerializable.GetObjectData 
        '  iterate over inner hashtable and add key-value pairs to serialization info
		Dim ht As Hashtable = me.InnerHashtable
	    Dim key As Object 	
    	For each key in ht.Keys 
		    '  assumption is that the inner hash table contains no duplicate keys (of course);
			'  further that keys won't collide with properties added later, wich are keyed with "name_GUID"
			'  ALSO:  to be sure we don't try to add/retrieve derived classes serialization streams to our hashtable, 
			'  TAG HASH ENTRIES with a uniquefier (guid)
			info.AddValue( key.ToString() & _tagGuid1.ToString(), ht(key) )
		Next
		
    	'  Individually add the other properties of interest
		'  The StateChanged event is ignored, because we can't serialize our listeners too.
		info.AddValue(NameCurrentView, me._currentView)
		info.AddValue(NameNavigationGraph, me._navigationGraph)
		info.AddValue(NameNavigationValue, me._navigateValue)
		info.AddValue(NameTaskId, me._taskId, GetType(System.Guid))
    End Sub
    #End Region
   
    #Region "StateChangedEvent and Delegate Definitions"
       
       
    'This event is raised when the state has changed. So, the views can refresh themselves to stay in-sync
    Delegate Sub StateChangedEventHandler(sender As Object, e As StateChangedEventArgs)
       
    Public Event StateChanged As StateChangedEventHandler
       
    #End Region
       
    'Visitor pattern for StatePersistence
    'Parameters: 
    '-statePersistence: A valid state persistence provider object
    Public Sub Accept(statePersistence As IStatePersistence)
        Me._stateVisitor = statePersistence
    End Sub
          
    'Stores the state into a storage using the persistence provider related to this state
    Public Sub Save()
        If _stateVisitor Is Nothing Then
            Throw New UIPException(Resource.ResourceManager("RES_ExceptionStateNotInitialized"))
        End If

        _stateVisitor.Save(Me)
    End Sub

    'Gets/Sets the navigation value. This value determines 
    'wich is the next view in the navigation graph.
    Public Property NavigateValue() As String
        Get
            Return _navigateValue
        End Get
        Set(ByVal Value As String)
            _navigateValue = Value
        End Set
    End Property

    'Gets/Sets the state navigation graph
    Public Property NavigationGraph() As String
        Get
            Return _navigationGraph
        End Get
        Set(ByVal Value As String)
            _navigationGraph = Value
        End Set
    End Property

    'Gets/Sets the current view in the navigation graph
    Public Property CurrentView() As String
        Get
            Return _currentView
        End Get
        Set(ByVal Value As String)
            _currentView = Value
        End Set
    End Property

    'Gets/Sets the current task id
    Public Property TaskId() As Guid
        Get
            Return _taskId
        End Get
        Set(ByVal Value As Guid)
            _taskId = Value
        End Set
    End Property

#Region "DictionaryBase members"

    'Indexer. Gets the item with the specified key
    Default Public Overridable Property Item(ByVal key As String) As Object
        Get
            Return Me.InnerHashtable(key)
        End Get
        Set(ByVal Value As Object)
            Me.InnerHashtable(key) = Value
            RaiseEvent StateChanged(Me, New StateChangedEventArgs(key))
        End Set
    End Property

    'Gets an object that can be used to synchonize access to the state
    Public Overridable ReadOnly Property SyncRoot() As Object
        Get
            Return Me.InnerHashtable.SyncRoot
        End Get
    End Property

    'Adds an element with the specified key to state
    Public Overridable Sub Add(ByVal key As String, ByVal value As Object)
        Me.InnerHashtable.Add(key, value)
        RaiseEvent StateChanged(Me, New StateChangedEventArgs(key))
    End Sub

    'Removes the element with the specified key from the state
    Public Overridable Sub Remove(ByVal key As String)
        Me.InnerHashtable.Remove(key)
        RaiseEvent StateChanged(Me, New StateChangedEventArgs(key))
    End Sub

    'Determines whether the state contains a specific key
    'Parameters: 
    '-key: 
    'Returns: 
    Public Overridable Function Contains(ByVal key As String) As Boolean
        Return Me.InnerHashtable.Contains(key)
    End Function

    'Copies the state elements to a array	
    'Parameters: 
    '-array: the array to copy to.  must be capable of accepting objects of type "DictionaryEntry"
    '-index: the zero-based array at wich to begin copying the State contents to the array
    Public Overloads Sub CopyTo(ByVal array() As DictionaryEntry, ByVal index As Integer)
        If array Is Nothing Then
            Throw New ArgumentNullException("array", Resource.ResourceManager("RES_ExceptionNullArrayInCopyToArray"))
        End If

        If 1 <> array.Rank Then
            Throw New ArgumentException(Resource.ResourceManager("RES_ExceptionInvalidArrayDimensionsInCopyToArray"), "array")
        End If

        If 0 > index OrElse array.GetUpperBound(0) < index OrElse array.GetUpperBound(0) < index + Me.InnerHashtable.Count Then
            Throw New ArgumentOutOfRangeException("index", index, Resource.ResourceManager("RES_ExceptionOutOfBoundsIndexInCopyToArray"))
        End If

        Try
            '  Attempt to copy to array
            Dim de As DictionaryEntry
            For Each de In Me.InnerHashtable
                index += 1
                array.SetValue(de, index)
            Next de
        Catch e As Exception
            Throw New InvalidCastException(Resource.ResourceManager("RES_ExceptionInvalidCastInCopyToArray"), e)
        End Try
    End Sub
#End Region
End Class
#End Region
