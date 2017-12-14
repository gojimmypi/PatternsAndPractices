Imports System.ComponentModel
Imports System.Configuration.Install

' Installer class used to create two event sources for the 
' Exception Management Application Block to function correctly.
<RunInstaller(True)> Public Class ExceptionManagerInstaller
    Inherits System.Configuration.Install.Installer

    Private exceptionManagerEventLogInstaller As System.Diagnostics.EventLogInstaller
    Private exceptionManagementEventLogInstaller As System.Diagnostics.EventLogInstaller
    Private Shared m_resourceManager As ResourceManager = New ResourceManager(GetType(ExceptionManager).Namespace + ".ExceptionManagerText", [Assembly].GetAssembly(GetType(ExceptionManager)))

    'Constructor with no params.
    Public Sub New()
        MyBase.New()

        'Initialize variables.
        InitializeComponent()

    End Sub

    'Installer overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' Initialization function to set internal variables.
    Private Sub InitializeComponent()

            Me.exceptionManagerEventLogInstaller = New System.Diagnostics.EventLogInstaller()
            Me.exceptionManagementEventLogInstaller = New System.Diagnostics.EventLogInstaller()

            ' exceptionManagerEventLogInstaller

            Me.exceptionManagerEventLogInstaller.Log = "Application"
            Me.exceptionManagerEventLogInstaller.Source = m_resourceManager.GetString("RES_EXCEPTIONMANAGER_INTERNAL_EXCEPTIONS")

            ' exceptionManagementEventLogInstaller

            Me.exceptionManagementEventLogInstaller.Log = "Application"
            Me.exceptionManagementEventLogInstaller.Source = m_resourceManager.GetString("RES_EXCEPTIONMANAGER_PUBLISHED_EXCEPTIONS")

            Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.exceptionManagerEventLogInstaller, Me.exceptionManagementEventLogInstaller})

    End Sub

End Class
