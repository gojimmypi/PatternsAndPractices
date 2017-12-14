'===============================================================================
' Microsoft User Interface Process Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp
'
' UIPException.vb
'
' This file contains the implementations of the UIPException class
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
Imports System.Runtime.Serialization

'The UIP main exception
<Serializable()>  _
Public Class UIPException
    Inherits ApplicationException
    
    'Constructor
    Public Sub New()
    End Sub
    
    
    'Constructor
    'Parameters: 
    '-msg: Exception message
    Public Sub New(msg As String)
        MyBase.New(msg)
    End Sub
        
    'Constructor
    'Parameters: 
    '-msg: Exception message
    '-inner: Inner exception
    Public Sub New(msg As String, inner As Exception)
        MyBase.New(msg, inner)
    End Sub
    
    'Deserialization constructor.
    Protected Sub New(info As SerializationInfo, context As StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class
