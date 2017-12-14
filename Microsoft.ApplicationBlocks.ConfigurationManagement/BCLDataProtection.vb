' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' BCLDataProtection.vb
'
' Data protection provider sample implementation that uses base class library
' support for data protection. Uses TripleDES for encryption and MACHASH for hashing.
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
Imports System.Collections.Specialized
Imports System.Security.Cryptography
Imports System.IO

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

Namespace DataProtection

    ' <summary>
    ' Implementation of a Data Provider using the base class library cryptography support
    ' </summary>
    Friend Class BCLDataProtection
        Implements IDataProtection
#Region "Declare members"
        Private _encryptionAlgorithm As SymmetricAlgorithm
        Private _macKey As Byte() = Nothing
#End Region

        ' <summary>
        ' Default constructor
        ' </summary>
        Public Sub New()
        End Sub 'New 

#Region "Implementation of IDataProtection"

        ' <summary>
        ' The initialization method used to get defaults from the configuration file
        ' </summary>
        ' <param name="dataProtectionParameters"></param>
        Public Sub Init(ByVal dataProtectionParameters As ListDictionary) Implements IDataProtection.Init
            'create an instance of the Triple-DES crypto 
            _encryptionAlgorithm = TripleDESCryptoServiceProvider.Create()

            ' Process the configuration parameters
            Dim base64Key As String = Nothing
            Dim regKey As String = CType(dataProtectionParameters("hashKeyRegistryPath"), String)

            If Not (regKey Is Nothing) AndAlso regKey.Length <> 0 Then
                base64Key = DataProtectionHelper.GetRegistryDefaultValue(regKey, "hashKey", "hashKeyRegistryPath")
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                base64Key = CType(dataProtectionParameters("hashKey"), String)
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                Throw New Exception(Resource.ResourceManager("Res_ExceptionEmptyHashKey"))
            End If
            'Get the key bytes from the base64 string
            _macKey = Convert.FromBase64String(base64Key)

            base64Key = Nothing
            regKey = CType(dataProtectionParameters("symmetricKeyRegistryPath"), String)

            If Not (regKey Is Nothing) AndAlso regKey.Length <> 0 Then
                base64Key = DataProtectionHelper.GetRegistryDefaultValue(regKey, _
                                                        "symmetricKey", _
                                                        "symmetricKeyRegistryPath")
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                base64Key = CType(dataProtectionParameters("symmetricKey"), String)
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                Throw New Exception(Resource.ResourceManager("Res_ExceptionEmptySymmetricKey"))
            End If
            ' Set the key
            _encryptionAlgorithm.Key = Convert.FromBase64String(base64Key)

            base64Key = Nothing
            regKey = CType(dataProtectionParameters("initializationVectorRegistryPath"), String)  '

            If Not (regKey Is Nothing) AndAlso regKey.Length <> 0 Then
                base64Key = DataProtectionHelper.GetRegistryDefaultValue(regKey, _
                                                            "initializationVector", _
                                                            "initializationVectorRegistryPath")
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                base64Key = CType(dataProtectionParameters("initializationVector"), String)
            End If
            If base64Key Is Nothing OrElse base64Key.Length = 0 Then
                Throw New Exception(Resource.ResourceManager("Res_ExceptionEmptyInitializationVectorKey"))
            End If
            'Set the IV
            _encryptionAlgorithm.IV = Convert.FromBase64String(base64Key)
        End Sub 'Init

        ' <summary>
        ' Encryption method
        ' </summary>
        Public Function Encrypt(ByVal plainText() As Byte) As Byte() Implements IDataProtection.Encrypt
            Dim cipherValue As Byte() = Nothing
            Dim plainValue As Byte() = plainText

            Dim memStream As New MemoryStream
            Dim cryptoStream As New cryptoStream(memStream, _
                                            _encryptionAlgorithm.CreateEncryptor(), _
                                            CryptoStreamMode.Write)
            Try
                ' Write the encrypted information
                cryptoStream.Write(plainValue, 0, plainValue.Length)
                cryptoStream.Flush()
                cryptoStream.FlushFinalBlock()

                ' Get the encrypted stream
                cipherValue = memStream.ToArray()
            Catch e As Exception
                Throw New ApplicationException(Resource.ResourceManager("Res_ExceptionCantEncrypt"), e)
            Finally
                ' Clear the arrays
                If Not (plainValue Is Nothing) Then
                    Array.Clear(plainValue, 0, plainValue.Length)
                End If
                memStream.Close()
                cryptoStream.Close()
            End Try

            Return cipherValue
        End Function 'Encrypt

        ' <summary>
        ' Decryption method
        ' </summary>
        Public Function Decrypt(ByVal cipherText() As Byte) As Byte() Implements IDataProtection.Decrypt
            Dim cipherValue As Byte() = cipherText
            Dim plainValue(cipherValue.Length) As Byte

            Dim memStream As New MemoryStream(cipherValue)
            Dim cryptoStream As New cryptoStream(memStream, _
                                        _encryptionAlgorithm.CreateDecryptor(), _
                                        CryptoStreamMode.Read)
            Try
                ' Decrypt the data
                cryptoStream.Read(plainValue, 0, plainValue.Length)
            Catch e As Exception
                Throw New ApplicationException(Resource.ResourceManager("Res_ExceptionCantDecrypt"), e)
            Finally
                ' Clear the arrays
                If Not (cipherValue Is Nothing) Then
                    Array.Clear(cipherValue, 0, cipherValue.Length)
                End If
                memStream.Close()
                'Flush the stream buffer
                cryptoStream.Close()
            End Try
            Return plainValue
        End Function 'Decrypt

        ' <summary>
        ' Compute a hash for an arbitrary string
        ' </summary>
        ' <param name="plainText">The plain text to hash</param>
        ' <returns>The hash to compute</returns>
        Public Function ComputeHash(ByVal plainText() As Byte) As Byte() Implements IDataProtection.ComputeHash
            'Compute the hash
            Dim mac As New HMACSHA1(_macKey)
            Dim hash As Byte() = mac.ComputeHash(plainText, 0, plainText.Length)

            Return hash
        End Function 'ComputeHash


#End Region

#Region "IDisposable implementation"

        ' <summary>
        ' Close the unmanaged resources
        ' </summary>
        Sub Dispose() Implements IDisposable.Dispose
            _encryptionAlgorithm.Clear()
        End Sub 'IDisposable.Dispose

#End Region
    End Class 'BCLDataProtection
End Namespace 'Microsoft.ApplicationBlocks.ConfigurationManagement.DataProtection