VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RubberduckRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const mstrRubberduckLibraryGuid As String = "{e07c841c-14b4-4890-83e9-8c80b06dd59d}"
Private Const mstrRubberduckSourceControlLibraryGuid As String = "{41382441-e6fc-473b-91eb-cdd2a515ac8d}"
Private Const mstrRubberduckExtensionProgId As String = "Rubberduck.Extension"

Private mobjLibraryRegistration As LibraryRegistration


' ================================================================================================
' CONSTRUCTOR / DESTRUCTOR
' ================================================================================================

Private Sub Class_Initialize()
    Set mobjLibraryRegistration = New LibraryRegistration
End Sub

Private Sub Class_Terminate()
    If Not mobjLibraryRegistration Is Nothing Then Set mobjLibraryRegistration = Nothing
End Sub


' ================================================================================================
' PUBLIC METHODS
' ================================================================================================

Public Sub RegisterRubberduck()
    RegisterRubberduckLibrary
    RegisterRubberduckSourceControlLibrary
    RegisterRubberduckAddin mstrRubberduckExtensionProgId, "Rubberduck", "Rubberduck"
End Sub

' ================================================================================================

Public Sub UnregisterRubberduck()
    UnregisterRubberduckLibrary
    UnregisterRubberduckSourceControlLibrary
    UnregisterRubberduckAddin mstrRubberduckExtensionProgId
End Sub


' ================================================================================================
' PRIVATE METHODS
' ================================================================================================

Private Static Sub RegisterRubberduckLibrary()

    With mobjLibraryRegistration
        .AssemblyCodebaseFile = "Rubberduck.dll"
        .AssemblyCodebasePath = Environ$("USERPROFILE") & "\AppData\Local\Rubberduck"
        .AssemblyName = "Rubberduck, Version=2.1.2.2766, Culture=neutral, PublicKeyToken=null"
        .AssemblyVersion = "2.1.2.2766"
        .RuntimeVersion = "v4.0.30319"
        .TypeLibraryFileName = "Rubberduck.tlb"
        .TypeLibraryFilePath = Environ$("USERPROFILE") & "\AppData\Local\Rubberduck"
        .TypeLibraryGuid = mstrRubberduckLibraryGuid
        .TypeLibraryName = "Rubberduck"
        .TypeLibraryVersion = "2.1"
    
        .RegisterTypeLibrary
    
        ' Classes
        .RegisterClass "Rubberduck.API.VBA.IdentifierReference", "{57F78E64-8ADF-4D81-A467-A0139B877D14}", "Rubberduck.IdentifierReference"
        .RegisterClass "Rubberduck.API.VBA.ParserState", "{28754D11-10CC-45FD-9F6A-525A65412B7A}", "Rubberduck.ParserState"
        .RegisterClass "Rubberduck.API.VBA.Declaration", "{67940D0B-081A-45BE-B0B9-CAEAFE034BC0}", "Rubberduck.Declaration"
        .RegisterClass "Rubberduck._Extension", "{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}", "Rubberduck.Extension"
        .RegisterClass "Rubberduck.UI._DockableWindowHost", "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}", "Rubberduck.UI.DockableWindowHost"
        .RegisterClass "Rubberduck.UnitTesting.AssertClass", "{69E194DA-43F0-3B33-B105-9B8188A6F040}", "Rubberduck.AssertClass"
        .RegisterClass "Rubberduck.UnitTesting.FakesProvider", "{69E194DD-43F0-3B33-B105-9B8188A6F040}", "Rubberduck.FakesProvider"
        .RegisterClass "Rubberduck.UnitTesting.PermissiveAssertClass", "{40F71F29-D63F-4481-8A7D-E04A4B054501}", "Rubberduck.PermissiveAssertClass"
        
        ' Enums
        .RegisterEnum "Rubberduck.API.VBA.Accessibility", "{272C7FBA-25BD-4785-93ED-18B5601F830D}" ' Guid generated
        .RegisterEnum "Rubberduck.API.VBA.DeclarationType", "{83567AE4-4F84-4F8A-B41D-A387FC25A464}" ' Guid generated
        
        ' Interfaces
        .RegisterInterface "IAssert", "{69E194DB-43F0-3B33-B105-9B8188A6F040}"
        .RegisterInterface "IDeclaration", "{2038ADD4-8376-42C4-9A79-5A5C8013EBD2}" ' Guid generated
        .RegisterInterface "IFake", "{69E194DF-43F0-3B33-B105-9B8188A6F040}"
        .RegisterInterface "IFakesProvider", "{69E194DE-43F0-3B33-B105-9B8188A6F040}"
        .RegisterInterface "IIdentifierReference", "{C0E43B26-E2DE-4274-829D-41D221F7B92A}" ' Guid generated
        .RegisterInterface "IParserState", "{A086CF8E-3FA3-4D03-A3D1-7BF8AAC799B7}" ' Guid generated
        .RegisterInterface "IParserStateEvents", "{3D8EAA28-8983-44D5-83AF-2EEC4C363079}"
        .RegisterInterface "IStub", "{69E194E1-43F0-3B33-B105-9B8188A6F040}"
        .RegisterInterface "IVerify", "{69E194E0-43F0-3B33-B105-9B8188A6F040}"
    
    End With
    
End Sub

' ================================================================================================

Private Static Sub RegisterRubberduckSourceControlLibrary()

    With mobjLibraryRegistration
    
        .AssemblyCodebaseFile = "Rubberduck.SourceControl.dll"
        .AssemblyCodebasePath = Environ$("USERPROFILE") & "\AppData\Local\Rubberduck"
        .AssemblyName = "Rubberduck.SourceControl, Version=2.0.0.0, Culture=neutral, PublicKeyToken=null"
        .AssemblyVersion = "2.0.0.0"
        .RuntimeVersion = "v4.0.30319"
        .TypeLibraryFileName = "Rubberduck.SourceControl.dll"
        .TypeLibraryFilePath = Environ$("USERPROFILE") & "\AppData\Local\Rubberduck"
        .TypeLibraryGuid = mstrRubberduckSourceControlLibraryGuid
        .TypeLibraryName = "Rubberduck Source Control 2.0"
        .TypeLibraryVersion = "2.0"
    
        ' Type Library
        .RegisterTypeLibrary
    
        ' Classes
        .RegisterClass "Rubberduck.SourceControl.Branch", "{6154532B-8880-40E9-B41E-2419C30B9F9A}", "Rubberduck.Branch"
        .RegisterClass "Rubberduck.SourceControl.FileStatusEntry", "{13AA3AF6-1397-4017-9E97-CBAD6A65FAFA}", "Rubberduck.FileStatus"
        .RegisterClass "Rubberduck.SourceControl.FileStatusEntries", "{E68A88BB-E15C-40D8-8D18-CAF7637312B5}", "Rubberduck.FileStatusEntries"
        .RegisterClass "Rubberduck.SourceControl.Interop.Branches", "{423A3B28-376B-4F96-A2E0-96E354965048}", "Rubberduck.Branches"
        .RegisterClass "Rubberduck.SourceControl.Interop.Credentials", "{AE54B926-49EB-4FB1-9F8A-AFE504A5A569}", "Rubberduck.Credentials"
        .RegisterClass "Rubberduck.SourceControl.Interop.GitProvider", "{0C22A01D-3255-4BB6-8D67-DCC40A548A32}", "Rubberduck.GitProvider"
        .RegisterClass "Rubberduck.SourceControl.Interop.SourceControlClassFactory", "{29FB0A0E-F113-458F-823B-1CD1B60D2CA7}", "Rubberduck.SourceControlClassFactory"
        .RegisterClass "Rubberduck.SourceControl.Repository", "{B2965961-7240-40CD-BE16-9425E2FB003C}", "Rubberduck.Repository"
        
        ' Enums
        .RegisterEnum "Rubberduck.SourceControl.FileStatus", "{4DDA743E-E3A7-440A-A030-92DF616B2C7B}"
    
        ' Interfaces
        .RegisterInterface "IBranch", "{B64F14D4-D083-4B41-BE99-4736C1D24B56}"
        .RegisterInterface "ICredentials", "{4EDFBFED-F6E7-4AFA-ADB4-B9FCAD21C256}"
        .RegisterInterface "IFileStatusEntry", "{577CB2D3-A84B-44FF-94EF-F4FC78363D74}"
        .RegisterInterface "IRepository", "{E8509738-3A06-4E8F-85FE-16F63F5A6DC3}"
        .RegisterInterface "_ISourceControlClassFactory", "{335DA0D8-625C-4CB9-90CD-C9A306B9B787}"
        .RegisterInterface "ISourceControlProvider", "{A44AF849-3C48-4303-A855-4B156958F3C4}"

    End With

End Sub

' ================================================================================================

Private Static Sub RegisterRubberduckAddin(ByVal strProgId As String, ByVal strDescription As String, ByVal strFriendlyName As String)
    mobjLibraryRegistration.RegisterVbeAddin strProgId, strDescription, strFriendlyName
End Sub

' ================================================================================================

Private Static Sub UnregisterRubberduckLibrary()
       
    With mobjLibraryRegistration
    
        .TypeLibraryGuid = mstrRubberduckLibraryGuid
    
        ' Type Library
        .UnregisterTypeLibrary
        
        ' Classes
        .UnregisterClass "{57F78E64-8ADF-4D81-A467-A0139B877D14}", "Rubberduck.IdentifierReference"
        .UnregisterClass "{28754D11-10CC-45FD-9F6A-525A65412B7A}", "Rubberduck.ParserState"
        .UnregisterClass "{67940D0B-081A-45BE-B0B9-CAEAFE034BC0}", "Rubberduck.Declaration"
        .UnregisterClass "{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}", "Rubberduck.Extension"
        .UnregisterClass "{9CF1392A-2DC9-48A6-AC0B-E601A9802608}", "Rubberduck.UI.DockableWindowHost"
        .UnregisterClass "{69E194DA-43F0-3B33-B105-9B8188A6F040}", "Rubberduck.AssertClass"
        .UnregisterClass "{69E194DD-43F0-3B33-B105-9B8188A6F040}", "Rubberduck.FakesProvider"
        .UnregisterClass "{40F71F29-D63F-4481-8A7D-E04A4B054501}", "Rubberduck.PermissiveAssertClass"
        
        ' Enums
        .UnregisterEnum "{272C7FBA-25BD-4785-93ED-18B5601F830D}" ' Rubberduck.API.VBA.Accessibility
        .UnregisterEnum "{83567AE4-4F84-4F8A-B41D-A387FC25A464}" ' Rubberduck.API.VBA.DeclarationType
        
        ' Interfaces
        .UnregisterInterface "{69E194DB-43F0-3B33-B105-9B8188A6F040}" ' IAssert
        .UnregisterInterface "{2038ADD4-8376-42C4-9A79-5A5C8013EBD2}" ' IDeclaration
        .UnregisterInterface "{69E194DF-43F0-3B33-B105-9B8188A6F040}" ' IFake
        .UnregisterInterface "{69E194DE-43F0-3B33-B105-9B8188A6F040}" ' IFakesProvider
        .UnregisterInterface "{C0E43B26-E2DE-4274-829D-41D221F7B92A}" ' IIdentifierReference
        .UnregisterInterface "{A086CF8E-3FA3-4D03-A3D1-7BF8AAC799B7}" ' IParserState
        .UnregisterInterface "{3D8EAA28-8983-44D5-83AF-2EEC4C363079}" ' IParserStateEvents
        .UnregisterInterface "{69E194E1-43F0-3B33-B105-9B8188A6F040}" ' IStub
        .UnregisterInterface "{69E194E0-43F0-3B33-B105-9B8188A6F040}" ' IVerify

    End With

End Sub

' ================================================================================================

Private Static Sub UnregisterRubberduckSourceControlLibrary()
    
    With mobjLibraryRegistration
        
        .TypeLibraryGuid = mstrRubberduckSourceControlLibraryGuid
        
        ' Type Library
        .UnregisterTypeLibrary
        
        ' Classes
        .UnregisterClass "{6154532B-8880-40E9-B41E-2419C30B9F9A}", "Rubberduck.Branch"
        .UnregisterClass "{13AA3AF6-1397-4017-9E97-CBAD6A65FAFA}", "Rubberduck.FileStatus"
        .UnregisterClass "{E68A88BB-E15C-40D8-8D18-CAF7637312B5}", "Rubberduck.FileStatusEntries"
        .UnregisterClass "{423A3B28-376B-4F96-A2E0-96E354965048}", "Rubberduck.Branches"
        .UnregisterClass "{AE54B926-49EB-4FB1-9F8A-AFE504A5A569}", "Rubberduck.Credentials"
        .UnregisterClass "{0C22A01D-3255-4BB6-8D67-DCC40A548A32}", "Rubberduck.GitProvider"
        .UnregisterClass "{29FB0A0E-F113-458F-823B-1CD1B60D2CA7}", "Rubberduck.SourceControlClassFactory"
        .UnregisterClass "{B2965961-7240-40CD-BE16-9425E2FB003C}", "Rubberduck.Repository"
        
        ' Enums
        .UnregisterEnum "{4DDA743E-E3A7-440A-A030-92DF616B2C7B}" ' Rubberduck.SourceControl.FileStatus
    
        ' Interfaces
        .UnregisterInterface "{B64F14D4-D083-4B41-BE99-4736C1D24B56}" ' IBranch
        .UnregisterInterface "{4EDFBFED-F6E7-4AFA-ADB4-B9FCAD21C256}" ' ICredentials
        .UnregisterInterface "{577CB2D3-A84B-44FF-94EF-F4FC78363D74}" ' IFileStatusEntry
        .UnregisterInterface "{E8509738-3A06-4E8F-85FE-16F63F5A6DC3}" ' IRepository
        .UnregisterInterface "{335DA0D8-625C-4CB9-90CD-C9A306B9B787}" ' _ISourceControlClassFactory
        .UnregisterInterface "{A44AF849-3C48-4303-A855-4B156958F3C4}" ' ISourceControlProvider

    End With

End Sub

' ================================================================================================

Private Static Sub UnregisterRubberduckAddin(ByVal strProgId As String)
    mobjLibraryRegistration.UnregisterVbeAddin strProgId
End Sub
