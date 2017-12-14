'===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' AssemblyInfo.vb
'
' This file contains the the definitions of assembly level attributes.
'
' For more information see the Configuration Management Application Block Implementation Overview. 
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
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

<Assembly: FileIOPermission(SecurityAction.RequestMinimum)> 
<Assembly: SqlClientPermission(SecurityAction.RequestMinimum)> 
<Assembly: SecurityPermission(SecurityAction.RequestMinimum, _
 Flags:=SecurityPermissionFlag.UnmanagedCode Or _
    SecurityPermissionFlag.SerializationFormatter Or _
    SecurityPermissionFlag.ControlThread)> 
<Assembly: RegistryPermission(SecurityAction.RequestMinimum)> 
<Assembly: ReflectionPermission(SecurityAction.RequestMinimum, _
 Flags:=ReflectionPermissionFlag.MemberAccess)> 

<Assembly: AssemblyTitle("Microsoft.ApplicationBlocks.ConfigurationManagement")>

<Assembly: AssemblyCompany("")>

<Assembly: AssemblyVersion("2.0.1.3172")> 

<assembly: AssemblyDelaySign(False)>

<Assembly: ComVisible(False)> 

<assembly: CLSCompliant(True)>

<Assembly: AssemblyFileVersionAttribute("2.0.1.3172")>
<Assembly: AssemblyCopyrightAttribute("")>