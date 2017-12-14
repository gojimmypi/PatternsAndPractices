' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' Form1.vb
'
' Sample utility that creates the registry files that are used to make
' deployment easy for the CMAB.
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
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Windows.Forms

Imports Microsoft.Win32


' <summary>
' Summary description for Form1.
' </summary>

Public Class keyMgmt
    Inherits System.Windows.Forms.Form
    Private label1 As System.Windows.Forms.Label
    Private label2 As System.Windows.Forms.Label
    Private label3 As System.Windows.Forms.Label
    Private WithEvents button1 As System.Windows.Forms.Button
    Private WithEvents button2 As System.Windows.Forms.Button
    Private toolTip1 As System.Windows.Forms.ToolTip
    Private registryKeyPath As System.Windows.Forms.TextBox
    Private hashKey As System.Windows.Forms.TextBox
    Private symmetricKey As System.Windows.Forms.TextBox
    Private regFilePath As System.Windows.Forms.TextBox
    Private label4 As System.Windows.Forms.Label
    Private WithEvents button5 As System.Windows.Forms.Button
    Private WithEvents generate As System.Windows.Forms.Button
    Private WithEvents testRegKey As System.Windows.Forms.Button
    Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Private WithEvents btnExit As System.Windows.Forms.Button
    Private components As System.ComponentModel.IContainer
    
    
    Public Sub New()
        '
        ' Required for Windows Form Designer support
        '
        InitializeComponent()
    End Sub 'New
    
    
    ' <summary>
    ' Clean up any resources being used.
    ' </summary>
    Protected Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub 'Dispose

#Region "Windows Form Designer generated code"

    ' <summary>
    ' Required method for Designer support - do not modify
    ' the contents of this method with the code editor.
    ' </summary>
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents initializationVector As System.Windows.Forms.TextBox
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.label1 = New System.Windows.Forms.Label
        Me.label2 = New System.Windows.Forms.Label
        Me.label3 = New System.Windows.Forms.Label
        Me.registryKeyPath = New System.Windows.Forms.TextBox
        Me.hashKey = New System.Windows.Forms.TextBox
        Me.symmetricKey = New System.Windows.Forms.TextBox
        Me.button1 = New System.Windows.Forms.Button
        Me.button2 = New System.Windows.Forms.Button
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.testRegKey = New System.Windows.Forms.Button
        Me.button5 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.generate = New System.Windows.Forms.Button
        Me.regFilePath = New System.Windows.Forms.TextBox
        Me.label4 = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.initializationVector = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(8, 8)
        Me.label1.Name = "label1"
        Me.label1.TabIndex = 0
        Me.label1.Text = "Registry key:"
        '
        'label2
        '
        Me.label2.Location = New System.Drawing.Point(8, 48)
        Me.label2.Name = "label2"
        Me.label2.TabIndex = 1
        Me.label2.Text = "Hash key:"
        '
        'label3
        '
        Me.label3.Location = New System.Drawing.Point(8, 128)
        Me.label3.Name = "label3"
        Me.label3.TabIndex = 2
        Me.label3.Text = "Symmetric key:"
        '
        'registryKeyPath
        '
        Me.registryKeyPath.Location = New System.Drawing.Point(120, 8)
        Me.registryKeyPath.Name = "registryKeyPath"
        Me.registryKeyPath.Size = New System.Drawing.Size(216, 20)
        Me.registryKeyPath.TabIndex = 3
        Me.registryKeyPath.Text = ""
        '
        'hashKey
        '
        Me.hashKey.Location = New System.Drawing.Point(120, 48)
        Me.hashKey.Name = "hashKey"
        Me.hashKey.ReadOnly = True
        Me.hashKey.Size = New System.Drawing.Size(216, 20)
        Me.hashKey.TabIndex = 4
        Me.hashKey.Text = ""
        '
        'symmetricKey
        '
        Me.symmetricKey.Location = New System.Drawing.Point(120, 128)
        Me.symmetricKey.Name = "symmetricKey"
        Me.symmetricKey.ReadOnly = True
        Me.symmetricKey.Size = New System.Drawing.Size(216, 20)
        Me.symmetricKey.TabIndex = 5
        Me.symmetricKey.Text = ""
        '
        'button1
        '
        Me.button1.Location = New System.Drawing.Point(352, 48)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(24, 23)
        Me.button1.TabIndex = 6
        Me.button1.Text = "..."
        Me.toolTip1.SetToolTip(Me.button1, "Generate")
        '
        'button2
        '
        Me.button2.Location = New System.Drawing.Point(352, 128)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(24, 23)
        Me.button2.TabIndex = 7
        Me.button2.Text = "..."
        Me.toolTip1.SetToolTip(Me.button2, "Generate")
        '
        'testRegKey
        '
        Me.testRegKey.Location = New System.Drawing.Point(352, 8)
        Me.testRegKey.Name = "testRegKey"
        Me.testRegKey.Size = New System.Drawing.Size(24, 23)
        Me.testRegKey.TabIndex = 8
        Me.testRegKey.Text = "..."
        Me.toolTip1.SetToolTip(Me.testRegKey, "Test reg key")
        '
        'button5
        '
        Me.button5.Location = New System.Drawing.Point(352, 168)
        Me.button5.Name = "button5"
        Me.button5.Size = New System.Drawing.Size(24, 23)
        Me.button5.TabIndex = 12
        Me.button5.Text = "..."
        Me.toolTip1.SetToolTip(Me.button5, "Select")
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(352, 88)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(24, 23)
        Me.Button3.TabIndex = 14
        Me.Button3.Text = "..."
        Me.toolTip1.SetToolTip(Me.Button3, "Generate")
        '
        'generate
        '
        Me.generate.Location = New System.Drawing.Point(120, 208)
        Me.generate.Name = "generate"
        Me.generate.Size = New System.Drawing.Size(120, 40)
        Me.generate.TabIndex = 9
        Me.generate.Text = "Generate .reg file"
        '
        'regFilePath
        '
        Me.regFilePath.Location = New System.Drawing.Point(120, 168)
        Me.regFilePath.Name = "regFilePath"
        Me.regFilePath.Size = New System.Drawing.Size(216, 20)
        Me.regFilePath.TabIndex = 10
        Me.regFilePath.Text = ""
        '
        'label4
        '
        Me.label4.Location = New System.Drawing.Point(8, 168)
        Me.label4.Name = "label4"
        Me.label4.TabIndex = 11
        Me.label4.Text = ".reg file:"
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(256, 208)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(120, 40)
        Me.btnExit.TabIndex = 13
        Me.btnExit.Text = "Exit"
        '
        'saveFileDialog1
        '
        Me.saveFileDialog1.FileName = "doc1"
        '
        'initializationVector
        '
        Me.initializationVector.Location = New System.Drawing.Point(120, 88)
        Me.initializationVector.Name = "initializationVector"
        Me.initializationVector.ReadOnly = True
        Me.initializationVector.Size = New System.Drawing.Size(216, 20)
        Me.initializationVector.TabIndex = 15
        Me.initializationVector.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "initializationVector:"
        '
        'keyMgmt
        '
        Me.AcceptButton = Me.generate
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(480, 286)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.initializationVector)
        Me.Controls.Add(Me.regFilePath)
        Me.Controls.Add(Me.symmetricKey)
        Me.Controls.Add(Me.hashKey)
        Me.Controls.Add(Me.registryKeyPath)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.button5)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.generate)
        Me.Controls.Add(Me.testRegKey)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.button1)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Name = "keyMgmt"
        Me.Text = "Key Management for CMAB"
        Me.ResumeLayout(False)

    End Sub 'InitializeComponent 
#End Region


    ' <summary>
    ' The main entry point for the application.
    ' </summary>
    <STAThread()> _
    Shared Sub Main()
        Application.Run(New keyMgmt)
    End Sub 'Main


    Private Sub testRegKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles testRegKey.Click
        GetRegistryDefaultValue(registryKeyPath.Text)
    End Sub 'testRegKey_Click


    Private Shared Sub GetRegistryDefaultValue(ByVal regKey As String)
        Dim baseKey As RegistryKey = Nothing
        Try
            If regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith("HKLM\") Then
                baseKey = Registry.LocalMachine
            ElseIf regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith("HKCU\") Then
                baseKey = Registry.CurrentUser
            ElseIf regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith("HKU\") Then
                baseKey = Registry.Users
            Else
                MessageBox.Show(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormat", regKey))
                Return
            End If

            Dim idxFirstSlash As Integer = regKey.IndexOf("\") + 1
            If idxFirstSlash = 0 Then
                MessageBox.Show(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormat", regKey))
                Return
            End If
            Dim keyPath As String = regKey.Substring(idxFirstSlash, regKey.Length - idxFirstSlash)
            Try
                Dim valueKey As RegistryKey = Nothing
                Try
                    valueKey = baseKey.OpenSubKey(keyPath, False)
                    Return
                Finally
                    If Not (valueKey Is Nothing) Then
                        CType(valueKey, IDisposable).Dispose()
                    End If
                End Try
            Catch e As Exception
                MessageBox.Show(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormatMessage", regKey, e.Message))
            End Try
        Finally
            If Not (baseKey Is Nothing) Then
                CType(baseKey, IDisposable).Dispose()
            End If
        End Try
    End Sub 'GetRegistryDefaultValue

    Private Function GetRandomKey(ByVal intLength As Integer) As String
        Dim key(intLength) As Byte
        Dim r As New Random
        r.NextBytes(key)
        Dim str As String = Convert.ToBase64String(key)
        Dim buffer As String = "="
        Dim i As Integer = 0
        Do While Convert.FromBase64String(str).Length <> intLength And (i < Len(str))
            ' not all byte arrays end up with a Base64Length the same as intLength!!!
            str = Convert.ToBase64String(key, 0, intLength - i)
            i += 1
        Loop

        Return str
    End Function

    Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
        hashKey.Text = GetRandomKey(8)
        Exit Sub
        'Dim key(8) As Byte
        'Dim r As New Random
        'r.NextBytes(key)
        'hashKey.Text = Convert.ToBase64String(key)
    End Sub 'button1_Click

    Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button2.Click
        ' http://www.gotdotnet.com/community/messageboard/Thread.aspx?id=216425
        ' The KeyConfig tool is creating a key that is 124 characters long (1024 bits) but the Triple Des Encryption 
        ' class only accepts keys that are 124 or 196 BITS long. I changed the code behind the button for the key so 
        ' that it only creates a 24 character (96 bit) key. 
        symmetricKey.Text = GetRandomKey(24)
        Exit Sub

        'Dim key(24) As Byte ' was Dim key(127) As Byte
        'Dim r As New Random
        'r.NextBytes(key)
        'symmetricKey.Text = Convert.ToBase64String(key)
    End Sub 'button2_Click

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        initializationVector.Text = GetRandomKey(8)
        Exit Sub

        'Dim key(8) As Byte
        'Dim r As New Random
        'r.NextBytes(key)
        'initializationVector.Text = Convert.ToBase64String(key)
    End Sub

    Private Sub button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button5.Click
        saveFileDialog1.InitialDirectory = [Assembly].GetEntryAssembly().Location
        saveFileDialog1.Filter = "Registry files (*.reg)|*.reg"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            regFilePath.Text = saveFileDialog1.FileName
        End If
    End Sub 'button5_Click

    Private Function GetRegistryRoot(ByVal regKey As String, ByVal displayMessage As Boolean) As String
        If regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith("HKLM\") Then
            Return "HKEY_LOCAL_MACHINE"
        ElseIf regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith("HKCU\") Then
            Return "HKEY_CURRENT_USER"
        ElseIf regKey.ToUpper(System.Globalization.CultureInfo.CurrentUICulture).StartsWith("HKU\") Then
            Return "HKEY_USERS"
        Else
            If (displayMessage) Then
                MessageBox.Show(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormat", regKey))
                Return Nothing
            Else
                Throw New Exception(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormat", regKey))
            End If
        End If
    End Function 'GetRegistryRoot


    Private Sub generate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles generate.Click
        If regFilePath.Text.Trim() = String.Empty Then
            MessageBox.Show(Resource.ResourceManager("Res_ExceptionInvalidFileName"))
            Exit Sub
        End If

        Dim targetFileName As String = ""
        Dim fi As FileInfo = New FileInfo(regFilePath.Text)
        If (fi.Extension.ToLower(System.Globalization.CultureInfo.CurrentUICulture) <> ".reg") Then
            targetFileName = Path.GetFileNameWithoutExtension(fi.Name) + ".reg"
        Else
            targetFileName = regFilePath.Text
        End If

        If (File.Exists(targetFileName) AndAlso _
         MessageBox.Show(Resource.ResourceManager("Res_WarningFileExists", targetFileName), "KeyConfig", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No) Then
            Return
        End If

        Dim sw As StreamWriter = Nothing
        Try
            sw = New StreamWriter(targetFileName)
            sw.WriteLine("Windows Registry Editor Version 5.00")
            sw.WriteLine()
            sw.Write("[")
            sw.Write(GetRegistryRoot(registryKeyPath.Text, False))
            sw.Write("\")
            Dim idxFirstSlash As Integer = registryKeyPath.Text.IndexOf("\") + 1
            If idxFirstSlash = 0 Then
                MessageBox.Show(Resource.ResourceManager("Res_ExceptionInvalidRegKeyFormat", registryKeyPath.Text))
                Return
            End If
            Dim keyPath As String = registryKeyPath.Text.Substring(idxFirstSlash, registryKeyPath.Text.Length - idxFirstSlash)
            sw.Write(keyPath)
            sw.WriteLine("]")
            sw.WriteLine()
            sw.Write("""hashKey""=""")
            sw.Write(hashKey.Text)
            sw.WriteLine("""")
            sw.Write("""symmetricKey""=""")
            sw.Write(symmetricKey.Text)
            sw.WriteLine("""")

            MessageBox.Show(Resource.ResourceManager("Res_Success"))
        Catch ex As Exception
            Dim message As String = ""
            Dim tempException As Exception = ex

            While Not (tempException Is Nothing)
                message += tempException.Message + Environment.NewLine + "----------" + Environment.NewLine
                tempException = tempException.InnerException
            End While
            MessageBox.Show(message)
            If Not (sw Is Nothing) Then
                CType(sw, IDisposable).Dispose()
            End If
            If (File.Exists(targetFileName)) Then File.Delete(targetFileName)
        Finally
            If Not (sw Is Nothing) Then
                CType(sw, IDisposable).Dispose()
            End If
        End Try
    End Sub 'generate_Click

    Private Sub exit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub 'exit_Click

End Class 'keyMgmt
