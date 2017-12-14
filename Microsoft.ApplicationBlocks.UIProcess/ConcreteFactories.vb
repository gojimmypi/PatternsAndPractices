'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' ConcreteFactories.vb
'
' This file contains the implementations of the StatePersistenceFactory, ViewManagerFactory,
' StateFactory and ControllerFactory classes
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
Imports System.Configuration
Imports System.Reflection
Imports System.Diagnostics

#Region "StatePersistenceFactory" 
'This class acts as a Factory for StatePersistence providers
NotInheritable Friend Class StatePersistenceFactory
    #Region "Declarations"
    Private Shared StatePersistenceCache As HybridDictionary
    #End Region
   
    #Region "Constructors"
    'Static constructor
    Shared Sub New()
        StatePersistenceCache = New HybridDictionary(5, True)
    End Sub
   
    Private Sub New()
    End Sub
    #End Region
   
    'Returns an instance of IStatePersistence according to type derived from nav graph
    'this is an optimization to avoid having to look up type info from config object each time, just pass in nav graph name.
    'Parameters:
    '-navigationGraph: nav graph
    'Returns:
    'instance of ISP, of specified type.  Gets from internal Cache if possible.
    Public Shared Function Create(navigationGraph As String) As IStatePersistence
        Dim providerSettings As StatePersistenceProviderSettings = UIPConfiguration.Config.GetPersistenceProviderSettings(navigationGraph)
        If providerSettings Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionStatePersistenceProviderConfigNotFound", navigationGraph))
        End If 
        
        Dim statePersistenceKey As String = providerSettings.Type + "," + providerSettings.Assembly
        Dim spp As IStatePersistence = CType(StatePersistenceCache(statePersistenceKey), IStatePersistence)
        If spp Is Nothing Then
            Try
                '  now create instance based on that type info
                spp = CType(GenericFactory.Create(providerSettings), IStatePersistence)
            
                '  pass in parameters to spp init method.  this is where spp's find data they need such as
                '  connection strings, etc.
                spp.Init(providerSettings.AdditionalAttributes)
            Catch e As Exception
                Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantCreateStatePersistenceProvider", providerSettings.Type), e)
            End Try
         
            '  lock collection
            SyncLock StatePersistenceCache.SyncRoot
                StatePersistenceCache(statePersistenceKey) = spp
            End SyncLock
        End If 
        'return it
        Return spp
    End Function
End Class
#End Region

#Region "ViewManagerFactory" 
'This class acts as a Factory for IViewManager objects
NotInheritable Friend Class ViewManagerFactory
    #Region "Declarations"
    Private Shared ViewManagerCache As HybridDictionary
    #End Region
   
    #Region "Constructors"

    'Static constructor
    Shared Sub New()
        ViewManagerCache = New HybridDictionary(5, True)
    End Sub
   
    Private Sub New()
    End Sub
   
    #End Region
   
    'Creates an IViewManager of a type specific to the named NavigationGraph; 
    'if it can, it simply returns a reference from an internal cache since these are (presumed) to be stateless
    'Parameters:
    '-navigationGraph: name of a nav graph
    'Returns:
    'instance of IViewManager. Gets from internal cache if possible
    Public Shared Function Create(navigationGraph As String) As IViewManager
        Dim ivmSettings As ObjectTypeSettings = Nothing
      
        'check if we have an instance of requested ivm in cache
        Dim ivm As IViewManager = CType(ViewManagerCache(navigationGraph), IViewManager)
      
        'not found in cache--create, store in cache, and return
        If ivm Is Nothing Then
            'get the type info from config
            'Get the view manager according to the configured application type
            ivmSettings = UIPConfiguration.Config.GetIViewManagerSettings(navigationGraph)
         
            If ivmSettings Is Nothing Then
                Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionIViewManagerNotFound", navigationGraph))
            End If
            Try
                '  activate an instance
                ivm = CType(GenericFactory.Create(ivmSettings), IViewManager)
            Catch e As Exception
                Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantCreateIViewManager", ivmSettings.Type), e)
            End Try
         
            'lock collection
            SyncLock ViewManagerCache.SyncRoot
                ViewManagerCache(navigationGraph) = ivm '  add this IViewManager to the cache
            End SyncLock
        End If 
        
        'return it
        Return ivm
    End Function 
End Class
#End Region

#Region "StateFactory" 
'This class acts as a Factory for State objects
NotInheritable Friend Class StateFactory
    #Region "Declarations"
    Private Shared StateCache As ListDictionary
    #End Region
   
    #Region "Constructors"
     
    'Static constructor
    Shared Sub New()
        StateCache = New ListDictionary()
    End Sub
   
    Private Sub New()
    End Sub
   
    #End Region
      
    'Creates a new cache entry
    'Parameters:
    '-mode: cache expiration mode
    '-interval: cache expiration interval
    '-cacheValue: cache value
    'Returns: a valid CacheEntry object
    Private Shared Function CreateCacheEntry(mode As CacheExpirationMode, interval As TimeSpan, cacheValue As Object) As CacheEntry
        Dim now As DateTime = DateTime.Now
        Select Case mode
            Case CacheExpirationMode.Absolute
                Dim absoluteDate As DateTime = New DateTime(now.Year, now.Month, now.Day).Add(interval)
                If absoluteDate > now Then
                    Return New CacheEntry(cacheValue, absoluteDate)
                Else
                    Return New CacheEntry(cacheValue, absoluteDate.AddDays(1))
                End If
            Case CacheExpirationMode.Sliding
                Return New CacheEntry(cacheValue, DateTime.Now.Add(interval))
            Case Else
                Return New CacheEntry(cacheValue)
        End Select
    End Function
   
    'Lookups a state object from the cache for the specified navigation graph and task id
    'Parameters:
    '-taskId: the guid associated to the task
    'Returns:
    'The state object
    Private Shared Function LoadFromCache(navigationGraph As String, taskId As Guid) As State
        Dim state As State = Nothing
      
        'attempt to retrieve from cache
        Dim cacheEntry As CacheEntry = CType(StateCache(taskId), CacheEntry)
      
        If Not cacheEntry Is Nothing Then
            'Check if the entry has expired
            If Not cacheEntry.IsExpired Then
                Dim weakReference As WeakReference = CType(cacheEntry.Value, WeakReference)
                If weakReference.IsAlive Then
                    state = CType(weakReference.Target, State)
                End If
            End If
        End If 
        
        'return it
        Return state
    End Function
   
    'Loads State object based on navgraph and taskID.  Internally, attempts to get object
    'from cache first, then creates SPP and uses it to load explicitly.
    'Parameters:
    '-navigationGraph: a navigation graph
    '-taskId: the task id
    'Returns:
    ' a valid state object
    Overloads Public Shared Function Load(navigationGraph As String, taskId As Guid) As State
        Dim state As State = Nothing
      
        If UIPConfiguration.Config.IsStateCacheEnabled Then
            state = LoadFromCache(navigationGraph, taskId)
         
            If state Is Nothing Then
                'State is not there in the cache, so a new state will be created here
                Dim spp As IStatePersistence = StatePersistenceFactory.Create(navigationGraph)
                state = Load(spp, taskId)
            
                'Get expiration configuration
                Dim mode As CacheExpirationMode
                Dim interval As TimeSpan
                UIPConfiguration.Config.GetCacheConfiguration(navigationGraph, mode, interval)
            
                '  move LOCK so that DePersist happens BEFORE lock acquired, otherwise we're holding
                '  a lock during a potentially lengthy database operation and deserialization.
                '  We only need the lock to have exclusive rights in cache hashtable.
                SyncLock StateCache.SyncRoot
                    StateCache(taskId) = CreateCacheEntry(mode, interval, New WeakReference(state, False))
                End SyncLock
            End If
        Else
            'The cache is disabled, so a new one is created
            Dim spp As IStatePersistence = StatePersistenceFactory.Create(navigationGraph)
            state = Load(spp, taskId)
        End If
      
        Return state
    End Function
      
    'Explicitly loads State from provided IStatePersistence instance, retrieving based on TaskID
    'Note that this overload does not attempt to fetch from cache
    'Parameters:
    '-statePersistenceProvider
    '-taskId
    'Returns:
    Overloads Private Shared Function Load(statePersistenceProvider As IStatePersistence, taskId As Guid) As State
        Dim state As State = statePersistenceProvider.Load(taskId)
      
        If Not state Is Nothing Then
            state.Accept(statePersistenceProvider)
        Else
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionTaskNotFound", taskId))
        End If 
        
        Return state
    End Function
      
    'Creates a State object, loading it from persistence provider
    'Parameters:
    '-navigationGraph: name of navigation graph
    'Returns:
    'State instance of type specified in config file
    Overloads Public Shared Function Create(navigationGraphName As String) As State
        
        ' Create a State persistence provider to be used by the state object
        Dim spp As IStatePersistence = StatePersistenceFactory.Create(navigationGraphName)
        
        Return Create(spp, navigationGraphName)
    End Function
    
    'Creates a State object, loading it from persistence provider
    'Parameters:
    '-navigationGraph: name of navigation graph
    '-spp: state persistence provider
    'Returns:
    'State instance of type specified in config file
    Overloads Private Shared Function Create(spp As IStatePersistence, navigationGraphName As String) As State
        Dim stateType As String = ""
        Dim state As State = Nothing
        Dim typeSettings As ObjectTypeSettings = Nothing
        Dim taskId As Guid = Guid.Empty
      
        typeSettings = UIPConfiguration.Config.GetStateSettings(navigationGraphName)
        If typeSettings Is Nothing Then
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionStateConfigNotFound", navigationGraphName))
        End If 
        
        ' Set the arguments used by the State object constructor
        Dim args As Object() = {spp}
            
        Try
            'pass to Base class' reflection code
            'DON'T look for this State in Cache, "CREATE" semantics in this class
            'demand that we create it freshly...
            'UNLIKE other Factories, State is stateful and we don't recycle in Create;
            'instead if the consuming class wishes a Cached entry, they might get it
            'from Load() methods instead...
            state = CType(GenericFactory.Create(typeSettings, args), State)
        Catch e As Exception
            Throw New ConfigurationErrorsException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantCreateState", stateType), e)
        End Try
      
        'creates a new Task id
        taskId = Guid.NewGuid()
      
        'store the task id into the state object 
        state.TaskId = taskId
      
        'Check if the cache is enabled
        If UIPConfiguration.Config.IsStateCacheEnabled Then
            'Get expiration configuration
            Dim mode As CacheExpirationMode
            Dim interval As TimeSpan
            UIPConfiguration.Config.GetCacheConfiguration(navigationGraphName, mode, interval)
         
            ' Create a new StateCacheEntry object using the existing state object and cache configuration
            Dim entry As CacheEntry = CreateCacheEntry(mode, interval, New WeakReference(state, True))
         
            'PUT IT IN Cache, we manage State Cache in State Factory...
            'as with all other concrete Factory implementations...
            SyncLock StateCache.SyncRoot
                StateCache(taskId) = entry
                Debug.Assert(entry Is StateCache(taskId) , "Cache object DID NOT contain StateCacheEntry just added to it.", "")
            End SyncLock
        End If
      
        'return it
        Return state
   End Function
End Class
#End Region

#Region "ControllerFactory"
'This class acts as a Factory for controller objects
NotInheritable Friend Class ControllerFactory
    #Region "Constructors"
    'Static constructor
    Shared Sub New()
    End Sub
      
    Private Sub New()
    End Sub
    #End Region
  
    'Returns a Controller, appropriate to a particular View.
	'Parameters:
	'-view: an instance of IView
	'Returns: a valid controller
	Public Shared Overloads Function Create( ByVal view As IView ) As ControllerBase
	    Return Create( view.NavigationGraph, view.TaskId )
	End Function


	'Returns a controller appropriate to a particular nav Graph and Task id.
	'Figures out a Controller type based on State's CurrentView
	'Parameters:
	'-navigationGraph: the name of the navgraph
	'-taskID: the task id guid
	'Returns: an instance of object typed ControllerBase wich will be a derivation specific to this view
	Public Shared Overloads Function Create( ByVal navigationGraph As String, ByVal taskID As Guid ) As ControllerBase
		Dim controller As ControllerBase = Nothing
		Dim typeSettings As ObjectTypeSettings = Nothing
		Dim	args() As object = Nothing
		Dim state As State = Nothing

		Try
		    'Get the task state
			If taskID.Equals( Guid.Empty ) Then 
			    state = StateFactory.Create( navigationGraph )
			Else
			    state = StateFactory.Load( navigationGraph, taskID )
            End If
			
            ' get the settings of the controller related to the current view
			typeSettings = UIPConfiguration.Config.GetControllerSettings(state.CurrentView)
			If typeSettings Is Nothing Then
			    Throw New UIPException( Resource.ResourceManager.FormatMessage( "RES_ExceptionControllerNotFound", navigationGraph ) )
            End If

			' set the arguments to be used by the controller object constructor
			args = New Object() {state}

			'  create a new controller object
			controller = CType(GenericFactory.Create( typeSettings, args ), ControllerBase )

			Return controller
		Catch e As Exception 
            Throw New UIPException( Resource.ResourceManager.FormatMessage( "RES_ExceptionCantInitializeController", navigationGraph ), e )
		End Try
	End Function
End Class
#End Region
