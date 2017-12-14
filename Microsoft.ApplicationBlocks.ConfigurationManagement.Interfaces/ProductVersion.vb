' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' ProductVersion.vb
'
' The product version information so all the assemblies have the same version
' and product name.
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

Namespace Microsoft.ApplicationBlocks.ConfigurationManagement
   ' <summary>
   ' Used to set the same version on every project on the solution
   ' </summary>
   
   Public Class Product
      ' <summary>
      ' The product version
      ' </summary>
        Public Const Version As String = "1.0.0.0"
      
      ' <summary>
      ' The company name
      ' </summary>
      Public Const Company As String = "Microsoft Corp."
      
      ' <summary>
      ' The project name
      ' </summary>
      Public Const Name As String = "Configuration Management Application Block"
   End Class 'Product
End Namespace 'Microsoft.ApplicationBlocks.ConfigurationManagement