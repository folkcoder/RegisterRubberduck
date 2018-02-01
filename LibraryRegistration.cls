VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibraryRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const mstrSoftwareClassesKeyBase32Bit As String = "Software\Classes\WOW6432Node\"
Private Const mstrSoftwareClassesKeyBase64Bit As String = "Software\Classes\"

Private mobjRegistry As Registry

Private mstrAssemblyCodebaseFile As String
Private mstrAssemblyCodebasePath As String
Private mstrAssemblyName As String
Private mstrAssemblyVersion As String
Private mstrRuntimeVersion As String

Private mstrTypeLibraryFileName As String
Private mstrTypeLibraryFilePath As String
Private mstrTypeLibraryGuid As String
Private mstrTypeLibraryName As String
Private mstrTypeLibraryVersion As String


' ================================================================================================
' CONSTRUCTOR / DESTRUCTOR
' ================================================================================================

Private Sub Class_Initialize()
    Set mobjRegistry = New Registry
    mobjRegistry.Hive = HKEY_CURRENT_USER
    mobjRegistry.Provider = Bit64
End Sub

Private Sub Class_Terminate()
    If Not mobjRegistry Is Nothing Then Set mobjRegistry = Nothing
End Sub


' ================================================================================================
' PROPERTIES
' ================================================================================================

Public Property Get AssemblyCodebaseFile() As String
    AssemblyCodebaseFile = mstrAssemblyCodebaseFile
End Property
Public Property Let AssemblyCodebaseFile(ByVal Value As String)
    mstrAssemblyCodebaseFile = Value
End Property

Public Property Get AssemblyCodebasePath() As String
    AssemblyCodebasePath = mstrAssemblyCodebasePath
End Property
Public Property Let AssemblyCodebasePath(ByVal Value As String)
    mstrAssemblyCodebasePath = Value
End Property

Public Property Get AssemblyName() As String
    AssemblyName = mstrAssemblyName
End Property
Public Property Let AssemblyName(ByVal Value As String)
    mstrAssemblyName = Value
End Property

Public Property Get RuntimeVersion() As String
    RuntimeVersion = mstrRuntimeVersion
End Property
Public Property Let RuntimeVersion(ByVal Value As String)
    mstrRuntimeVersion = Value
End Property

Public Property Get AssemblyVersion() As String
    AssemblyVersion = mstrAssemblyVersion
End Property
Public Property Let AssemblyVersion(ByVal Value As String)
    mstrAssemblyVersion = Value
End Property

Public Property Get TypeLibraryFileName() As String
    TypeLibraryFileName = mstrTypeLibraryFileName
End Property
Public Property Let TypeLibraryFileName(ByVal Value As String)
    mstrTypeLibraryFileName = Value
End Property

Public Property Get TypeLibraryFilePath() As String
    TypeLibraryFilePath = mstrTypeLibraryFilePath
End Property
Public Property Let TypeLibraryFilePath(ByVal Value As String)
    mstrTypeLibraryFilePath = Value
End Property

Public Property Get TypeLibraryGuid() As String
    TypeLibraryGuid = mstrTypeLibraryGuid
End Property
Public Property Let TypeLibraryGuid(ByVal Value As String)
    mstrTypeLibraryGuid = Value
End Property

Public Property Get TypeLibraryName() As String
    TypeLibraryName = mstrTypeLibraryName
End Property
Public Property Let TypeLibraryName(ByVal Value As String)
    mstrTypeLibraryName = Value
End Property

Public Property Get TypeLibraryVersion() As String
    TypeLibraryVersion = mstrTypeLibraryVersion
End Property
Public Property Let TypeLibraryVersion(ByVal Value As String)
    mstrTypeLibraryVersion = Value
End Property



' ================================================================================================
' PUBLIC METHODS
' ================================================================================================

Public Static Sub RegisterClass( _
    ByVal strName As String, _
    ByVal strGuid As String, _
    ByVal strProgId As String _
    )

Const strComComponentGuid As String = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"

Dim strClassRegistrationKeyBase32Bit As String
Dim strClsidKey32Bit As String
Dim strImplementedCategoriesKey32Bit As String
Dim strInprocServer32Key32Bit As String
Dim strInprocServer32AssemblyVersionKey32Bit As String
Dim strClsidProgIdKey32Bit As String

Dim strClassRegistrationKeyBase64Bit As String
Dim strClsidKey64Bit As String
Dim strImplementedCategoriesKey64Bit As String
Dim strInprocServer32Key64Bit As String
Dim strInprocServer32AssemblyVersionKey64Bit As String
Dim strClsidProgIdKey64Bit As String

    ' SOFTWARE\CLASSES\ProgId
    strClassRegistrationKeyBase32Bit = mstrSoftwareClassesKeyBase32Bit & strProgId
    mobjRegistry.CreateRegistryKey strClassRegistrationKeyBase32Bit
    mobjRegistry.WriteRegistryString strClassRegistrationKeyBase32Bit, vbNullString, strName

    strClassRegistrationKeyBase64Bit = mstrSoftwareClassesKeyBase64Bit & strProgId
    mobjRegistry.CreateRegistryKey strClassRegistrationKeyBase64Bit
    mobjRegistry.WriteRegistryString strClassRegistrationKeyBase64Bit, vbNullString, strName
    

    ' SOFTWARE\CLASSES\ProgId\CLSID
    strClsidKey32Bit = strClassRegistrationKeyBase32Bit & "\CLSID"
    mobjRegistry.CreateRegistryKey strClsidKey32Bit
    mobjRegistry.WriteRegistryString strClsidKey32Bit, vbNullString, strGuid

    strClsidKey64Bit = strClassRegistrationKeyBase64Bit & "\CLSID"
    mobjRegistry.CreateRegistryKey strClsidKey64Bit
    mobjRegistry.WriteRegistryString strClsidKey64Bit, vbNullString, strGuid
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid
    strClsidKey32Bit = mstrSoftwareClassesKeyBase32Bit & "CLSID\" & strGuid
    mobjRegistry.CreateRegistryKey strClsidKey32Bit
    mobjRegistry.WriteRegistryString strClsidKey32Bit, vbNullString, strName
    
    strClsidKey64Bit = mstrSoftwareClassesKeyBase64Bit & "CLSID\" & strGuid
    mobjRegistry.CreateRegistryKey strClsidKey64Bit
    mobjRegistry.WriteRegistryString strClsidKey64Bit, vbNullString, strName
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\Implemented Categories
    strImplementedCategoriesKey32Bit = strClsidKey32Bit & "\Implemented Categories"
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey32Bit
    
    strImplementedCategoriesKey64Bit = strClsidKey64Bit & "\Implemented Categories"
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey64Bit
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\Implemented Categories\COM COMPONENT GUID
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey32Bit & "\" & strComComponentGuid
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey64Bit & "\" & strComComponentGuid
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\InprocServer32
    strInprocServer32Key32Bit = strClsidKey32Bit & "\InprocServer32"
    mobjRegistry.CreateRegistryKey strInprocServer32Key32Bit
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, vbNullString, "mscoree.dll"
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "RuntimeVersion", Me.RuntimeVersion
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "ThreadingModel", "Both"

    strInprocServer32Key64Bit = strClsidKey64Bit & "\InprocServer32"
    mobjRegistry.CreateRegistryKey strInprocServer32Key64Bit
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, vbNullString, "mscoree.dll"
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "RuntimeVersion", Me.RuntimeVersion
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "ThreadingModel", "Both"
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\InprocServer32\AssemblyVersion
    strInprocServer32AssemblyVersionKey32Bit = strInprocServer32Key32Bit & "\" & Me.AssemblyVersion
    mobjRegistry.CreateRegistryKey strInprocServer32AssemblyVersionKey32Bit
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "RuntimeVersion", Me.RuntimeVersion
    
    strInprocServer32AssemblyVersionKey64Bit = strInprocServer32Key64Bit & "\" & Me.AssemblyVersion
    mobjRegistry.CreateRegistryKey strInprocServer32AssemblyVersionKey64Bit
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "RuntimeVersion", Me.RuntimeVersion
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\ProgId
    strClsidProgIdKey32Bit = strClsidKey32Bit & "\ProgId"
    mobjRegistry.CreateRegistryKey strClsidProgIdKey32Bit
    mobjRegistry.WriteRegistryString strClsidProgIdKey32Bit, vbNullString, strProgId
    
    strClsidProgIdKey64Bit = strClsidKey64Bit & "\ProgId"
    mobjRegistry.CreateRegistryKey strClsidProgIdKey64Bit
    mobjRegistry.WriteRegistryString strClsidProgIdKey64Bit, vbNullString, strProgId
    
End Sub

' ================================================================================================

Public Static Sub RegisterEnum( _
    ByVal strName As String, _
    ByVal strGuid As String _
    )

Dim strRecordKey As String

    ' SOFTWARE\CLASSES\Record\EnumGuid\AssemblyVersion
    strRecordKey = mstrSoftwareClassesKeyBase64Bit & "Record\" & strGuid & "\" & Me.AssemblyVersion
    mobjRegistry.CreateRegistryKey strRecordKey
    mobjRegistry.WriteRegistryString strRecordKey, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strRecordKey, "Class", strName
    mobjRegistry.WriteRegistryString strRecordKey, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strRecordKey, "RuntimeVersion", Me.RuntimeVersion
    
End Sub

' ================================================================================================

Public Static Sub RegisterInterface( _
    ByVal strName As String, _
    ByVal strGuid As String _
    )

Dim strInterfaceKey32Bit As String
Dim strProxyStubClsid32Key32Bit As String
Dim strTypeLibraryKey32Bit As String

Dim strInterfaceKey64Bit As String
Dim strProxyStubClsid32Key64Bit As String
Dim strTypeLibraryKey64Bit As String

    ' SOFTWARE\CLASSES\Interface\InterfaceGuid
    strInterfaceKey32Bit = mstrSoftwareClassesKeyBase32Bit & "Interface\" & strGuid
    mobjRegistry.CreateRegistryKey strInterfaceKey32Bit
    mobjRegistry.WriteRegistryString strInterfaceKey32Bit, vbNullString, strName

    strInterfaceKey64Bit = mstrSoftwareClassesKeyBase64Bit & "Interface\" & strGuid
    mobjRegistry.CreateRegistryKey strInterfaceKey64Bit
    mobjRegistry.WriteRegistryString strInterfaceKey64Bit, vbNullString, strName


    ' SOFTWARE\CLASSES\Interface\InterfaceGuid\ProxyStubClsid32
    strProxyStubClsid32Key32Bit = strInterfaceKey32Bit & "\ProxyStubClsid32"
    mobjRegistry.CreateRegistryKey strProxyStubClsid32Key32Bit
    mobjRegistry.WriteRegistryString strProxyStubClsid32Key32Bit, vbNullString, Me.TypeLibraryGuid
    
    strProxyStubClsid32Key64Bit = strInterfaceKey64Bit & "\ProxyStubClsid32"
    mobjRegistry.CreateRegistryKey strProxyStubClsid32Key64Bit
    mobjRegistry.WriteRegistryString strProxyStubClsid32Key64Bit, vbNullString, Me.TypeLibraryGuid
    

    ' SOFTWARE\CLASSES\Interface\InterfaceGuid\TypeLib
    strTypeLibraryKey32Bit = strInterfaceKey32Bit & "\TypeLib"
    mobjRegistry.CreateRegistryKey strTypeLibraryKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibraryKey32Bit, vbNullString, strGuid
    mobjRegistry.WriteRegistryString strTypeLibraryKey32Bit, "Version", Me.TypeLibraryVersion
    
    strTypeLibraryKey64Bit = strInterfaceKey64Bit & "\TypeLib"
    mobjRegistry.CreateRegistryKey strTypeLibraryKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibraryKey64Bit, vbNullString, strGuid
    mobjRegistry.WriteRegistryString strTypeLibraryKey64Bit, "Version", Me.TypeLibraryVersion
    
End Sub

' ================================================================================================

Public Static Sub RegisterTypeLibrary()

Dim strTypeLibKey32Bit As String
Dim strTypeLibWin32Key32Bit As String
Dim strTypeLibFlagsKey32Bit As String
Dim strTypeLibHelpdirKey32Bit As String

Dim strTypeLibKey64Bit As String
Dim strTypeLibWin32Key64Bit As String
Dim strTypeLibFlagsKey64Bit As String
Dim strTypeLibHelpdirKey64Bit As String

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion
    strTypeLibKey32Bit = mstrSoftwareClassesKeyBase32Bit & "TypeLib\" & Me.TypeLibraryGuid & "\" & Me.TypeLibraryVersion
    mobjRegistry.CreateRegistryKey strTypeLibKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibKey32Bit, vbNullString, Me.TypeLibraryName
    
    strTypeLibKey64Bit = mstrSoftwareClassesKeyBase64Bit & "TypeLib\" & Me.TypeLibraryGuid & "\" & Me.TypeLibraryVersion
    mobjRegistry.CreateRegistryKey strTypeLibKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibKey64Bit, vbNullString, Me.TypeLibraryName
        

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion\0\win32
    strTypeLibWin32Key32Bit = strTypeLibKey32Bit & "\0\win32"
    mobjRegistry.CreateRegistryKey strTypeLibWin32Key32Bit
    mobjRegistry.WriteRegistryString strTypeLibWin32Key32Bit, vbNullString, Me.TypeLibraryFilePath & "\" & Me.TypeLibraryFileName

    strTypeLibWin32Key64Bit = strTypeLibKey64Bit & "\0\win32"
    mobjRegistry.CreateRegistryKey strTypeLibWin32Key64Bit
    mobjRegistry.WriteRegistryString strTypeLibWin32Key64Bit, vbNullString, Me.TypeLibraryFilePath & "\" & Me.TypeLibraryFileName
    

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion\FLAGS
    strTypeLibFlagsKey32Bit = strTypeLibKey32Bit & "\FLAGS"
    mobjRegistry.CreateRegistryKey strTypeLibFlagsKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibFlagsKey32Bit, vbNullString, "0"

    strTypeLibFlagsKey64Bit = strTypeLibKey64Bit & "\FLAGS"
    mobjRegistry.CreateRegistryKey strTypeLibFlagsKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibFlagsKey64Bit, vbNullString, "0"
    

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion\HELPDIR
    strTypeLibHelpdirKey32Bit = strTypeLibKey32Bit & "\HELPDIR"
    mobjRegistry.CreateRegistryKey strTypeLibHelpdirKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibHelpdirKey32Bit, vbNullString, Me.TypeLibraryFilePath
    
    strTypeLibHelpdirKey64Bit = strTypeLibKey64Bit & "\HELPDIR"
    mobjRegistry.CreateRegistryKey strTypeLibHelpdirKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibHelpdirKey64Bit, vbNullString, Me.TypeLibraryFilePath
    
End Sub

' ================================================================================================

Public Static Sub RegisterVbeAddin( _
    ByVal strProgId As String, _
    ByVal strDescription As String, _
    ByVal strFriendlyName As String _
    )
    
Const strVbeAddinKey32Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\"
Const strVbeAddinKey64Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns64\"
Const intLoadBehaviorAtStartup As Integer = 3

Dim strAddinClassKey32Bit As String
Dim strAddinClassKey64Bit As String

    ' SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\ProgId
    strAddinClassKey32Bit = strVbeAddinKey32Bit & strProgId
    mobjRegistry.CreateRegistryKey strAddinClassKey32Bit
    mobjRegistry.WriteRegistryString strAddinClassKey32Bit, "Description", strDescription
    mobjRegistry.WriteRegistryString strAddinClassKey32Bit, "FriendlyName", strFriendlyName
    mobjRegistry.WriteRegistryInteger strAddinClassKey32Bit, "LoadBehavior", intLoadBehaviorAtStartup
    
    strAddinClassKey64Bit = strVbeAddinKey64Bit & strProgId
    mobjRegistry.CreateRegistryKey strAddinClassKey64Bit
    mobjRegistry.WriteRegistryString strAddinClassKey64Bit, "Description", strDescription
    mobjRegistry.WriteRegistryString strAddinClassKey64Bit, "FriendlyName", strFriendlyName
    mobjRegistry.WriteRegistryInteger strAddinClassKey64Bit, "LoadBehavior", intLoadBehaviorAtStartup
    
End Sub

' ================================================================================================

Public Static Sub UnregisterClass( _
    ByVal strGuid As String, _
    ByVal strProgId As String _
    )

Dim strClassRegistrationKeyBase32Bit As String
Dim strClsidKey32Bit As String

Dim strClassRegistrationKeyBase64Bit As String
Dim strClsidKey64Bit As String

    ' SOFTWARE\CLASSES\ProgId
    strClassRegistrationKeyBase32Bit = mstrSoftwareClassesKeyBase32Bit & strProgId
    mobjRegistry.DeleteRegistryTree strClassRegistrationKeyBase32Bit
               
    strClassRegistrationKeyBase64Bit = mstrSoftwareClassesKeyBase64Bit & strProgId
    mobjRegistry.DeleteRegistryTree strClassRegistrationKeyBase64Bit
               
               
    ' SOFTWARE\CLASSES\CLSID\ClassGuid
    strClsidKey32Bit = mstrSoftwareClassesKeyBase32Bit & "CLSID\" & strGuid
    mobjRegistry.DeleteRegistryTree strClsidKey32Bit
        
    strClsidKey64Bit = mstrSoftwareClassesKeyBase64Bit & "CLSID\" & strGuid
    mobjRegistry.DeleteRegistryTree strClsidKey64Bit
        
End Sub

' ================================================================================================

Public Static Sub UnregisterEnum(ByVal strGuid As String)

Dim strRecordKey32Bit As String
Dim strRecordKey64Bit As String
   
    ' SOFTWARE\CLASSES\Record\EnumGuid
    strRecordKey32Bit = mstrSoftwareClassesKeyBase32Bit & "Record\" & strGuid
    mobjRegistry.DeleteRegistryTree strRecordKey32Bit

    strRecordKey64Bit = mstrSoftwareClassesKeyBase64Bit & "Record\" & strGuid
    mobjRegistry.DeleteRegistryTree strRecordKey64Bit

End Sub

' ================================================================================================

Public Static Sub UnregisterInterface(ByVal strGuid As String)

Dim strInterfaceKey32Bit As String
Dim strInterfaceKey64Bit As String

    ' SOFTWARE\CLASSES\Interface\InterfaceGuid
    strInterfaceKey32Bit = mstrSoftwareClassesKeyBase32Bit & "Interface\" & strGuid
    mobjRegistry.DeleteRegistryTree strInterfaceKey32Bit
    
    strInterfaceKey64Bit = mstrSoftwareClassesKeyBase64Bit & "Interface\" & strGuid
    mobjRegistry.DeleteRegistryTree strInterfaceKey64Bit
    
End Sub

' ================================================================================================

Public Static Sub UnregisterTypeLibrary()

Dim strTypeLibKey32Bit As String
Dim strTypeLibKey64Bit As String

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid
    strTypeLibKey32Bit = mstrSoftwareClassesKeyBase32Bit & "TypeLib\" & Me.TypeLibraryGuid
    mobjRegistry.DeleteRegistryTree strTypeLibKey32Bit
        
    strTypeLibKey64Bit = mstrSoftwareClassesKeyBase64Bit & "TypeLib\" & Me.TypeLibraryGuid
    mobjRegistry.DeleteRegistryTree strTypeLibKey64Bit
        
End Sub

' ================================================================================================

Public Static Sub UnregisterVbeAddin(ByVal strProgId As String)
    
Const strVbeAddinKey32Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\"
Const strVbeAddinKey64Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns64\"
   
Dim strAddinClassKey32Bit As String
Dim strAddinClassKey64Bit As String
   
    ' SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\ProgId
    strAddinClassKey32Bit = strVbeAddinKey32Bit & strProgId
    mobjRegistry.DeleteRegistryTree strAddinClassKey32Bit
    
    strAddinClassKey64Bit = strVbeAddinKey64Bit & strProgId
    mobjRegistry.DeleteRegistryTree strAddinClassKey64Bit
    
End Sub

