'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' Resource.vb
'
' This file contains the implementations of the Resource class
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
Imports System.Resources
Imports System.Reflection
Imports System.IO

'Helper class used to manage application Resources
NotInheritable Friend Class Resource
    #Region "Static part"
    Private Const ResourceFileName As String = ".UIPText"
    
    Private Shared InternalResource As New Resource
    
    'Gets a resource manager for the assembly resource file
    Public Shared ReadOnly Property ResourceManager() As Resource
        Get
            Return InternalResource
        End Get
    End Property
    #End Region
      
    #Region "Instance part "
    Private rm As ResourceManager = Nothing
      
    'Constructor
    Public Sub New()
        rm = New ResourceManager(Me.GetType().Namespace + ResourceFileName, [Assembly].GetExecutingAssembly())
    End Sub
    
    'Gets the message with the specified key from the assembly resource file
    Default Public ReadOnly Property Item(key As String) As String
        Get
            Return rm.GetString(key, System.Globalization.CultureInfo.CurrentCulture)
        End Get
    End Property
    
    'Gets a resource stream with the messages used by the UIP classes
    'Parameters:
    '-name: resource key
    'Returns:
    'a resource stream
    Public Function GetStream(name As String) As Stream
        Return [Assembly].GetExecutingAssembly().GetManifestResourceStream((Me.GetType().Namespace + "." + name))
    End Function
    
    'Formats a message stored in the UIP assembly resource file.
    'Parameters:
    '-key: resource key
    '-format: format arguments
    'Returns:
    'a formated string
    Public Function FormatMessage(key As String, ParamArray format() As Object) As String
        Return String.Format(System.Globalization.CultureInfo.CurrentUICulture, Me(key), format)
    End Function
    #End Region
End Class
