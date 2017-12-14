' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' DataProtection.vb
'
' Data protection implementation that uses DPAPI for encryption and MACHASH for
' hashing.
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
Imports System.Runtime.InteropServices
Imports System.Collections.Specialized
Imports System.Security.Permissions
Imports System.Security.Cryptography
Imports Microsoft.Win32

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

Namespace DataProtection
    ' <summary>
    ' The key store used for the DPAPI encryption
    ' </summary>

    Public Enum Store
        ' <summary>
        ' Use the machine key to encrypt the data
        ' </summary>
        Machine = 1

        ' <summary>
        ' Use the user key to encrypt the data
        ' </summary>
        User
    End Enum 'Store

    ' <summary>
    ' The DPAPI wrapper
    ' </summary>

    Friend Class DPAPIDataProtection
        Implements IDataProtection
#Region "Constants"
        Private Const CRYPTPROTECT_UI_FORBIDDEN As Integer = &H1
        Private Const CRYPTPROTECT_LOCAL_MACHINE As Integer = &H4
        Private store As store
#End Region

#Region "P/Invoke structures"
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
        Friend Structure DATA_BLOB
            Public cbData As Integer
            Public pbData As IntPtr
        End Structure 'DATA_BLOB

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
        Friend Structure CRYPTPROTECT_PROMPTSTRUCT
            Public cbSize As Integer
            Public dwPromptFlags As Integer
            Public hwndApp As IntPtr
            Public szPrompt As String
        End Structure 'CRYPTPROTECT_PROMPTSTRUCT
#End Region

#Region "External methods"

        Declare Auto Function CryptProtectData Lib "Crypt32.dll" (ByRef pDataIn As DATA_BLOB, _
                            ByVal szDataDescr As String, ByRef pOptionalEntropy As DATA_BLOB, _
                            ByVal pvReserved As IntPtr, ByRef pPromptStruct As CRYPTPROTECT_PROMPTSTRUCT, _
                            ByVal dwFlags As Integer, ByRef pDataOut As DATA_BLOB) As Boolean

        Declare Auto Function CryptUnprotectData Lib "Crypt32.dll" (ByRef pDataIn As DATA_BLOB, _
                            ByVal szDataDescr As String, ByRef pOptionalEntropy As DATA_BLOB, _
                            ByVal pvReserved As IntPtr, ByRef pPromptStruct As CRYPTPROTECT_PROMPTSTRUCT, _
                            ByVal dwFlags As Integer, ByRef pDataOut As DATA_BLOB) As Boolean

        Declare Auto Function FormatMessage Lib "kernel32.dll" (ByVal dwFlags As Integer, _
                            ByRef lpSource As IntPtr, ByVal dwMessageId As Integer, _
                            ByVal dwLanguageId As Integer, ByRef lpBuffer As String, _
                            ByVal nSize As Integer, ByVal Arguments As IntPtr) As Integer
#End Region

#Region "Constructor"

        Public Sub New()
            Me.New( CType( 1, Store ) ) 
        End Sub

        Public Sub New(ByVal tempStore As store)
            store = tempStore
        End Sub 'New
#End Region

#Region "Declare members"
        Private _macKey As Byte() = Nothing
#End Region

#Region "IDataProtection implementation"


        Public Sub Init(ByVal initParams As ListDictionary) Implements IDataProtection.Init
            Dim sp As New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)
            sp.Assert()

            Dim keyStoreString As String = CType(initParams("keyStore"), String) '
            If Not (keyStoreString Is Nothing) AndAlso keyStoreString.Length <> 0 Then
                store = CType([Enum].Parse(GetType(store), keyStoreString, True), store)
            Else
                store = store.Machine
            End If
            Dim base64Key As String = Nothing
            Dim regKey As String = CType(initParams("hashKeyRegistryPath"), String) '
            If Not (regKey Is Nothing) AndAlso regKey.Length <> 0 Then
                base64Key = DataProtectionHelper.GetRegistryDefaultValue(regKey, "hashKey", "hashKeyRegistryPath")
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                base64Key = CType(initParams("hashKey"), String) '
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                Throw New Exception(Resource.ResourceManager("Res_ExceptionEmptyHashKey"))
            End If
            'Get the hashkey bytes from the base64 string
            _macKey = Convert.FromBase64String(base64Key)
        End Sub 'Init


        ' <summary>
        ' Encrypt the given data
        ' </summary>
        Public Overloads Function Encrypt(ByVal plainText() As Byte) As Byte() Implements IDataProtection.Encrypt
            Dim sp As New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)
            sp.Assert()

            Return Encrypt(plainText, Nothing)
        End Function 'Encrypt

        Public Overloads Function Decrypt(ByVal cipherText() As Byte) As Byte() Implements IDataProtection.Decrypt
            Dim sp As New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)
            sp.Assert()

            Return Decrypt(cipherText, Nothing)
        End Function 'Decrypt

        Function ComputeHash(ByVal hashedData() As Byte) As Byte() Implements IDataProtection.ComputeHash
            Dim sp As New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)
            sp.Assert()

            'Compute the hash
            Dim mac As New HMACSHA1(_macKey)
            Return mac.ComputeHash(hashedData, 0, hashedData.Length)
        End Function 'ComputeHash

#End Region

#Region "Private methods"
        Public Overloads Function Encrypt(ByVal plainText() As Byte, ByVal optionalEntropy() As Byte) As Byte()
            Dim plainTextBlob As New DATA_BLOB
            Dim cipherTextBlob As New DATA_BLOB
            Dim entropyBlob As New DATA_BLOB

            Dim prompt As New CRYPTPROTECT_PROMPTSTRUCT
            InitPromptstruct(prompt)

            Dim dwFlags As Integer
            Try
                Try
                    Dim bytesSize As Integer = plainText.Length
                    plainTextBlob.pbData = Marshal.AllocHGlobal(bytesSize)
                    If IntPtr.Zero.Equals(plainTextBlob.pbData) Then
                        Throw New Exception(Resource.ResourceManager("Res_UnableToAllocateBuffer"))
                    End If
                    plainTextBlob.cbData = bytesSize
                    Marshal.Copy(plainText, 0, plainTextBlob.pbData, bytesSize)
                Catch ex As Exception
                    Throw New Exception(Resource.ResourceManager("Res_ExceptionMarshallingData"), ex)
                End Try
                If store.Machine = store Then
                    'Using the machine store, should be providing entropy.
                    dwFlags = CRYPTPROTECT_LOCAL_MACHINE Or CRYPTPROTECT_UI_FORBIDDEN
                    'Check to see if the entropy is null
                    If optionalEntropy Is Nothing Then
                        'Allocate something
                        optionalEntropy = New Byte(-1) {}
                    End If
                    Try
                        Dim bytesSize As Integer = optionalEntropy.Length
                        entropyBlob.pbData = Marshal.AllocHGlobal(optionalEntropy.Length)
                        If IntPtr.Zero.Equals(entropyBlob.pbData) Then
                            Throw New Exception(Resource.ResourceManager("Res_ExceptionAllocatingEntropyBuffer"))
                        End If
                        Marshal.Copy(optionalEntropy, 0, entropyBlob.pbData, bytesSize)
                        entropyBlob.cbData = bytesSize
                    Catch ex As Exception
                        Throw New Exception(Resource.ResourceManager("Res_ExceptionEntropyMarshallingData"), ex)
                    End Try
                Else
                    'Using the user store
                    dwFlags = CRYPTPROTECT_UI_FORBIDDEN
                End If
                If Not CryptProtectData(plainTextBlob, "", entropyBlob, IntPtr.Zero, prompt, _
                                    dwFlags, cipherTextBlob) Then
                    Throw New Exception(Resource.ResourceManager("Res_ExceptionEncryptionFailed") + _
                                GetErrorMessage(Marshal.GetLastWin32Error()))
                End If
            Catch ex As Exception
                Throw New Exception(Resource.ResourceManager("Res_ExceptionEncryptionFailed") + ex.Message, ex)
            Finally
                If Not (plainText Is Nothing) Then
                    Array.Clear(plainText, 0, plainText.Length)
                End If
                If Not (optionalEntropy Is Nothing) Then
                    Array.Clear(optionalEntropy, 0, optionalEntropy.Length)
                End If
            End Try
            Dim cipherText(cipherTextBlob.cbData) As Byte
            Marshal.Copy(cipherTextBlob.pbData, cipherText, 0, cipherTextBlob.cbData)
            Return cipherText
        End Function 'Encrypt

        Public Overloads Function Decrypt(ByVal cipherText() As Byte, ByVal optionalEntropy() As Byte) As Byte()
            Dim plainTextBlob As New DATA_BLOB
            Dim cipherBlob As New DATA_BLOB
            Dim prompt As New CRYPTPROTECT_PROMPTSTRUCT
            InitPromptstruct(prompt)
            Try
                Try
                    Dim cipherTextSize As Integer = cipherText.Length
                    cipherBlob.pbData = Marshal.AllocHGlobal(cipherTextSize)
                    If IntPtr.Zero.Equals(cipherBlob.pbData) Then
                        Throw New Exception( _
                                Resource.ResourceManager("Res_ExceptionUnableToAllecateCipherTextBuffer"))
                    End If
                    cipherBlob.cbData = cipherTextSize
                    Marshal.Copy(cipherText, 0, cipherBlob.pbData, cipherBlob.cbData)
                Catch ex As Exception
                    Throw New Exception(Resource.ResourceManager("Res_ExceptionMarshallingData"), ex)
                End Try
                Dim entropyBlob As New DATA_BLOB
                Dim dwFlags As Integer
                If store.Machine = store Then
                    'Using the machine store, should be providing entropy.
                    dwFlags = CRYPTPROTECT_LOCAL_MACHINE Or CRYPTPROTECT_UI_FORBIDDEN
                    'Check to see if the entropy is null
                    If optionalEntropy Is Nothing Then
                        'Allocate something
                        optionalEntropy = New Byte(-1) {}
                    End If
                    Try
                        Dim bytesSize As Integer = optionalEntropy.Length
                        entropyBlob.pbData = Marshal.AllocHGlobal(bytesSize)
                        If IntPtr.Zero.Equals(entropyBlob.pbData) Then
                            Throw New Exception(Resource.ResourceManager("Res_ExceptionAllocatingEntropyBuffer"))
                        End If
                        entropyBlob.cbData = bytesSize
                        Marshal.Copy(optionalEntropy, 0, entropyBlob.pbData, bytesSize)
                    Catch ex As Exception
                        Throw New Exception(Resource.ResourceManager("Res_ExceptionEntropyMarshallingData"), ex)
                    End Try
                Else
                    'Using the user store
                    dwFlags = CRYPTPROTECT_UI_FORBIDDEN
                End If
                If Not CryptUnprotectData(cipherBlob, Nothing, entropyBlob, IntPtr.Zero, _
                                prompt, dwFlags, plainTextBlob) Then

                    Throw New Exception(Resource.ResourceManager("Res_ExceptionDecryptionFailed") + _
                                GetErrorMessage(Marshal.GetLastWin32Error()))
                End If
                'Free the blob and entropy.
                If Not IntPtr.Zero.Equals(cipherBlob.pbData) Then
                    Marshal.FreeHGlobal(cipherBlob.pbData)
                End If
                If Not IntPtr.Zero.Equals(entropyBlob.pbData) Then
                    Marshal.FreeHGlobal(entropyBlob.pbData)
                End If
            Catch ex As Exception
                Throw New Exception(Resource.ResourceManager("Res_ExceptionDecryptionFailed") + ex.Message, ex)
            End Try
            Dim plainText(plainTextBlob.cbData) As Byte
            Marshal.Copy(plainTextBlob.pbData, plainText, 0, plainTextBlob.cbData)
            Return plainText
        End Function 'Decrypt

        Private Sub InitPromptstruct(ByRef ps As CRYPTPROTECT_PROMPTSTRUCT)
            ps.cbSize = Marshal.SizeOf(GetType(CRYPTPROTECT_PROMPTSTRUCT))
            ps.dwPromptFlags = 0
            ps.hwndApp = IntPtr.Zero
            ps.szPrompt = Nothing
        End Sub 'InitPromptstruct


        Private Shared Function GetErrorMessage(ByVal errorCode As Integer) As String
            Dim FORMAT_MESSAGE_ALLOCATE_BUFFER As Integer = &H100
            Dim FORMAT_MESSAGE_IGNORE_INSERTS As Integer = &H200
            Dim FORMAT_MESSAGE_FROM_SYSTEM As Integer = &H1000
            Dim messageSize As Integer = 255
            Dim lpMsgBuf As String = ""
            Dim dwFlags As Integer = FORMAT_MESSAGE_ALLOCATE_BUFFER Or _
                                     FORMAT_MESSAGE_FROM_SYSTEM Or _
                                     FORMAT_MESSAGE_IGNORE_INSERTS
            Dim ptrlpSource As New IntPtr
            Dim prtArguments As New IntPtr
            Dim retVal As Integer = FormatMessage(dwFlags, ptrlpSource, errorCode, 0, lpMsgBuf, _
                                            messageSize, IntPtr.Zero)
            If 0 = retVal Then
                Throw New Exception(Resource.ResourceManager("Res_ExceptionFormattingMessage", errorCode))
            End If
            Return lpMsgBuf
        End Function 'GetErrorMessage
#End Region

#Region "IDisposable implementation"

        ' <summary>
        ' Close the unmanaged resources
        ' </summary>
        Sub Dispose() Implements IDisposable.Dispose
        End Sub 'IDisposable.Dispose
        'Not used, no unmanaged members

#End Region
    End Class 'DPAPIDataProtection
End Namespace 'Microsoft.ApplicationBlocks.ConfigurationManagement.DataProtection


    Friend Class DataProtectionHelper

    Public Shared Function GetRegistryDefaultValue(ByVal regKey As String, ByVal valueName As String, _
                                ByVal attributeName As String) As String
        Dim baseKey As RegistryKey = Nothing
        Try
            If regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith( _
                    Registry.LocalMachine.Name) Then

                baseKey = Registry.LocalMachine
            ElseIf regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith( _
                    Registry.CurrentUser.Name) Then

                baseKey = Registry.CurrentUser
            ElseIf regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith( _
                Registry.Users.Name) Then

                baseKey = Registry.Users
            Else
                Throw New Exception(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormat", _
                                regKey, attributeName))
            End If
            Dim idxFirstSlash As Integer = regKey.IndexOf("\") + 1
            Dim keyPath As String = regKey.Substring(idxFirstSlash, regKey.Length - idxFirstSlash)

            Dim valueKey As RegistryKey = Nothing
            Try
                valueKey = baseKey.OpenSubKey(keyPath, False)
                If valueKey Is Nothing Then
                    Throw New Exception(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormatCantFoundKey", _
                                regKey, attributeName))
                End If
                Return CType(valueKey.GetValue(valueName, ""), String) '
            Finally
                If Not (valueKey Is Nothing) Then
                    CType(valueKey, IDisposable).Dispose()
                End If
            End Try
        Finally
            If Not (baseKey Is Nothing) Then
                CType(baseKey, IDisposable).Dispose()
            End If
        End Try
    End Function 'GetRegistryDefaultValue
End Class 'DataProtectionHelper
