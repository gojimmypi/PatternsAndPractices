' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' Resource.vb
'
' Wrapper to make the use of resources easy on the code.
'
' For more information see the Configuration Management Application Block Implementation Overview. 
' 
' ===============================================================================
' Copyright (C) 2000-2001 Microsoft Corporation
' All rights reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
' FITNESS FOR A PARTICULAR PURPOSE.
' ==============================================================================

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Resources
Imports System.Reflection

' <summary>
   ' Helper class used to manage application resources
   ' </summary>
NotInheritable Class Resource
#Region "Static part"

    Private Shared ReadOnly _resource As New resource

    '  static constructor private by nature.  Initialize our read-only member _resourceManager here, 
    '  there will only ever be one copy.
    Shared Sub New()

    End Sub 'New


    '  return the singleton instance of Resource

    Public Shared ReadOnly Property ResourceManager() As Resource
        Get
            Return _resource
        End Get
    End Property


#End Region

#Region "Instance part"

    '  this is the ACTUAL resource manager, for which this class is just a convenience wrapper
    Private _resourceManager As ResourceManager = Nothing


    '  make constructor private so no one can directly create an instance of Resource, only use the Static Property ResourceManager
    Private Sub New()
        _resourceManager = New ResourceManager(Me.GetType().Namespace + ".ConfigurationManagerText", _
                                        [Assembly].GetExecutingAssembly())
    End Sub 'New


    '  a convenience Indexer that access the internal resource manager

    Default Public ReadOnly Property Item(ByVal key As String) As String
        Get
            Return _resourceManager.GetString(key, CultureInfo.CurrentCulture)
        End Get
    End Property



    Default Public ReadOnly Property Item(ByVal key As String, ByVal ParamArray par() As Object) As String
        Get
            Return String.Format(CultureInfo.CurrentUICulture, _
                    _resourceManager.GetString(key, CultureInfo.CurrentCulture), par)
        End Get
    End Property


    Public Function GetStream(ByVal name As String) As Stream
        Return [Assembly].GetExecutingAssembly().GetManifestResourceStream(name)
    End Function 'GetStream
#End Region
End Class 'Resource
