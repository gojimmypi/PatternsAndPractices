' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' GenericFactory.vb
'
' Factory pattern implementation, this file defines generic functionality
' for all the factories.
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
Imports System.Reflection

' <summary>
' Acts as the basic implementation for the multiple Factory classes used elsewhere.
' We need to create instances based on config info ...
' Have Factories for those, and since there's much common code for doing Reflection-based activation 
' keep that code in one central place.
' 
' </summary>
NotInheritable Class GenericFactory

#Region "Declarations"

    Private Const COMMA_DELIMITER As String = ","

#End Region

#Region "Constructors"


    Shared Sub New()
    End Sub 'New


    Private Sub New()
    End Sub 'New 

#End Region

#Region "Private Helper Methods"


    ' <summary>
    ' Takes incoming full type string, defined as:
		'   FULL TYPE NAME AS WRITTEN IN CONFIG IS: 
		'   "Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage,   Microsoft.ApplicationBlocks.ConfigurationManagement, 
		'  			Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    '  And splits the type into two strings, the typeName and assemblyName.  Those are passed by as OUT params
    '  This routine also cleans up any extra whitespace, and throws an exception if the full type string
    '  does not have five comma-delimited parts....it expect the true full name complete with version and publickeytoken
    ' </summary>
    ' <param name="fullType"></param>
    ' <param name="typeName"></param>
    ' <param name="assemblyName"></param>
    Private Shared Sub SplitType(ByVal fullType As String, ByRef typeName As String, ByRef assemblyName As String)
        Dim parts As String() = fullType.Split(COMMA_DELIMITER.ToCharArray())

        If 5 <> parts.Length Then
            Throw New ArgumentException( _
                            Resource.ResourceManager("RES_ExceptionBadTypeArgumentInFactory"), "fullType")
        Else
            '  package type name:
            typeName = parts(0).Trim()
            '  package fully-qualified assembly name separated by commas
            assemblyName = String.Concat(parts(1).Trim() + COMMA_DELIMITER, parts(2).Trim() + COMMA_DELIMITER, _
                                        parts(3).Trim() + COMMA_DELIMITER, parts(4).Trim())
            '  return
            Return
        End If
    End Sub 'SplitType


#End Region

#Region "Create Overloads"


    ' <summary>
    ' Returns an object instantiated by the Activator, using fully-qualified combined assembly-type  supplied.
    ' Assembly parameter example: 
		'   Assembly parameter example: "Microsoft.ApplicationBlocks.ConfigurationManagement, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
		'   Type parameter example: Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage"
		'   FULL TYPE NAME AS WRITTEN IN CONFIG IS: 
		'   "Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage,   Microsoft.ApplicationBlocks.ConfigurationManagement, 
		'  			Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    ' </summary>
    ' <param name="fullTypeName">the fully-qualified type name</param>
    ' <returns>instance of requested assembly/type typed as System.Object</returns>
    Public Overloads Shared Function Create(ByVal fullTypeName As String) As Object
        Dim assemblyName As String = ""
        Dim typeName As String = ""
        '  use helper to split
        SplitType(fullTypeName, typeName, assemblyName)
        '  just call main overload
        Return Create(assemblyName, typeName, Nothing)
    End Function 'Create



    ' <summary>
    ' Returns an object instantiated by the Activator, using fully-qualified asm/type supplied.		
    '   Assembly parameter example: "Microsoft.ApplicationBlocks.ConfigurationManagement, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
		'   Type parameter example: Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage"
		'   FULL TYPE NAME AS WRITTEN IN CONFIG IS: 
		'   "Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage,   Microsoft.ApplicationBlocks.ConfigurationManagement, 
		'  			Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    ' </summary>
    ' <param name="assemblyName">fully-qualified assembly name</param>
    ' <param name="typeName">the type name</param>
    ' <returns>instance of requested assembly/type typed as System.Object</returns>
    Public Overloads Shared Function Create(ByVal assemblyName As String, ByVal typeName As String) As Object
        Dim aName As String = ""
        Dim tName As String = ""

        '  use helper to split
        SplitType(typeName + "," + assemblyName, tName, aName)

        '  just call main overload
        Return Create(aName, tName, Nothing)
    End Function 'Create



    ' <summary>
    ' Returns an object instantiated by the Activator, using fully-qualified asm/type supplied.
		'   Assembly parameter example: "Microsoft.ApplicationBlocks.ConfigurationManagement, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
		'   Type parameter example: Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage"
		'   FULL TYPE NAME AS WRITTEN IN CONFIG IS: 
		'   "Microsoft.ApplicationBlocks.ConfigurationManagement.XmlFileStorage,   Microsoft.ApplicationBlocks.ConfigurationManagement, 
		'  			Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    ' </summary>
    ' <param name="assemblyName">fully-qualified assembly name</param>
    ' <param name="typeName">the type name</param>
    ' <param name="constructorArguments">constructor arguments for type to be created</param>
    ' <returns>instance of requested assembly/type typed as System.Object</returns>
    Public Overloads Shared Function Create(ByVal assemblyName As String, ByVal typeName As String, _
                        ByVal constructorArguments() As Object) As Object

        Dim assemblyInstance As [Assembly] = Nothing
        Dim typeInstance As Type = Nothing

        Try
            '  use full asm name to get assembly instance
            assemblyInstance = [Assembly].Load(assemblyName.Trim())
        Catch e As Exception
            Throw New TypeLoadException(Resource.ResourceManager("RES_ExceptionCantLoadAssembly", _
                                    assemblyName, typeName), e)
        End Try

        Try
            '  use type name to get type from asm; note we WANT case specificity 
            typeInstance = assemblyInstance.GetType(typeName.Trim(), True, False)

            '  now attempt to actually create an instance, passing constructor args if available
            If Not (constructorArguments Is Nothing) Then
                Return Activator.CreateInstance(typeInstance, constructorArguments)
            Else
                Return Activator.CreateInstance(typeInstance)
            End If
        Catch e As Exception
            Throw New TypeLoadException(Resource.ResourceManager("RES_ExceptionCantCreateInstanceUsingActivate", _
                                assemblyName, typeName), e)
        End Try
    End Function 'Create

#End Region
End Class 'GenericFactory
