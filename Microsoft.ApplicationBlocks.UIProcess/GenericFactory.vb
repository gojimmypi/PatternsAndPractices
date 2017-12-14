'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' GenericFactory.vb
'
' This file contains the implementations of the GenericFactory class
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
Imports System.Configuration
Imports System.Reflection

'Acts as the basic implementation for the multiple Factory classes used in UIProcess.
'We need to create instances based on config info for State, SPP's, ViewManagers...
'Have Factories for those, and since there's much common code for doing Reflection-based activation 
'keep that code in one central place.
NotInheritable Friend Class GenericFactory
    #Region "Constructors"
         
        'Static constructor
        Shared Sub New()
        End Sub
            
        Private Sub New()
        End Sub
      
    #End Region
      
    #Region "Create Overloads"
      
    'Create an object using full name type contained in typeSettings
    'Parameters:
    '-typeSettings: A typeSetting object with the needed type information to create a class instance
    'Returns:
    'An instance of the specified type
    Overloads Public Shared Function Create(typeSettings As ObjectTypeSettings) As Object
        Return Create(typeSettings, Nothing)
    End Function
      
      
    'Create an object using full name type contained in typeSettings
    'Parameters:
    '-typeSettings: A typeSetting object with the needed type information to create a class instance
    '-args: constructor arguments
    'Returns:
    'an instance of the specified type
    Overloads Public Shared Function Create(typeSettings As ObjectTypeSettings, args() As Object) As Object
        Dim assemblyInstance As [Assembly] = Nothing
        Dim typeInstance As Type = Nothing
         
        Try
            '  use full assembly name to get assembly instance
            assemblyInstance = [Assembly].Load(typeSettings.Assembly)
        Catch e As Exception
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantLoadAssembly", typeSettings.Assembly), e)
        End Try  
         
        '  use type name to get type from assembly
        typeInstance = assemblyInstance.GetType(typeSettings.Type, True, False)
         
        Try
            If Not (args Is Nothing) Then
                Return Activator.CreateInstance(typeInstance, args)
            Else
                Return Activator.CreateInstance(typeInstance)
            End If
        Catch e As Exception
            Throw New UIPException(Resource.ResourceManager.FormatMessage("RES_ExceptionCantCreateInstanceUsingActivate", typeInstance), e)
        End Try
    End Function
      
    #End Region
End Class 
