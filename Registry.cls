VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum RegistryHive
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum RegistryProvider
    Bit32 = 32
    Bit64 = 64
End Enum

Private mobjContext As Object
Private mobjRegistry As Object
Private mlngRegistryHive As RegistryHive


' ================================================================================================
' CONSTRUCTOR / DESTRUCTOR
' ================================================================================================

Private Sub Class_Initialize()

Dim objLocator As Object

    Set mobjContext = CreateObject("WbemScripting.SWbemNamedValueSet")
    mobjContext.Add "__ProviderArchitecture", RegistryProvider.Bit64
    
    Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set mobjRegistry = objLocator.ConnectServer(".", "root\default", vbNullString, vbNullString, , , , mobjContext).Get("StdRegProv")

End Sub

' ================================================================================================

Private Sub Class_Terminate()
    If Not mobjContext Is Nothing Then Set mobjContext = Nothing
    If Not mobjRegistry Is Nothing Then Set mobjRegistry = Nothing
End Sub


' ================================================================================================
' PROPERTIES
' ================================================================================================

Public Property Get Hive() As RegistryHive
    Hive = mlngRegistryHive
End Property
Public Property Let Hive(ByRef Value As RegistryHive)
    mlngRegistryHive = Value
End Property

Public Property Get Provider() As RegistryProvider
    Provider = mobjContext.Item("__ProviderArchitecture")
End Property
Public Property Let Provider(ByRef Value As RegistryProvider)
    mobjContext.Item("__ProviderArchitecture") = Value
End Property


' ================================================================================================
' PUBLIC METHODS
' ================================================================================================

Public Sub CreateRegistryKey(ByVal strKeyPath As String)

Dim objInParameters As Object
Dim objOutParameters As Object

    Set objInParameters = mobjRegistry.Methods_("CreateKey").InParameters
    objInParameters.hDefKey = mlngRegistryHive
    objInParameters.sSubKeyName = strKeyPath
    
    Set objOutParameters = mobjRegistry.ExecMethod_("CreateKey", objInParameters, , mobjContext)
    If Not objOutParameters.ReturnValue = 0 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.CreateRegistryKey", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If
   
End Sub

' ================================================================================================

Public Sub DeleteRegistryTree(ByVal strKeyPath As String)

    DeleteSubkeys strKeyPath

End Sub

' ================================================================================================

Public Sub WriteRegistryString( _
    ByVal strKeyPath As String, _
    ByVal strValueName As String, _
    ByVal strValue As String _
    )

Dim objInParameters As Object
Dim objOutParameters As Object
      
    Set objInParameters = mobjRegistry.Methods_("SetStringValue").InParameters
    objInParameters.hDefKey = mlngRegistryHive
    objInParameters.sSubKeyName = strKeyPath
    objInParameters.sValueName = strValueName
    objInParameters.sValue = strValue
    
    Set objOutParameters = mobjRegistry.ExecMethod_("SetStringValue", objInParameters, , mobjContext)
    If Not objOutParameters.ReturnValue = 0 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.WriteRegistryString", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If

End Sub

' ================================================================================================

Public Sub WriteRegistryInteger( _
    ByVal strKeyPath As String, _
    ByVal strValueName As String, _
    ByVal intValue As Integer _
    )

Dim objInParameters As Object
Dim objOutParameters As Object
   
    Set objInParameters = mobjRegistry.Methods_("SetDWORDValue").InParameters
    objInParameters.hDefKey = mlngRegistryHive
    objInParameters.sSubKeyName = strKeyPath
    objInParameters.sValueName = strValueName
    objInParameters.uValue = intValue
    
    Set objOutParameters = mobjRegistry.ExecMethod_("SetDWORDValue", objInParameters, , mobjContext)
    If Not objOutParameters.ReturnValue = 0 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.WriteRegistryString", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If

End Sub

' ================================================================================================
' PRIVATE METHODS
' ================================================================================================

Private Static Sub DeleteSubkeys(ByVal strKeyPath As String)
    
Dim vntSubkeys As Variant
Dim vntSubkey As Variant
    
    mobjRegistry.EnumKey mlngRegistryHive, strKeyPath, vntSubkeys
    If IsArray(vntSubkeys) Then
        For Each vntSubkey In vntSubkeys
            DeleteSubkeys strKeyPath & "\" & CStr(vntSubkey)
        Next vntSubkey
    End If
     
    DeleteRegistryKey strKeyPath
     
End Sub

' ================================================================================================

Private Static Sub DeleteRegistryKey(ByVal strKeyPath As String)
    
Dim objInParameters As Object
Dim objOutParameters As Object

    Set objInParameters = mobjRegistry.Methods_("DeleteKey").InParameters
    objInParameters.hDefKey = mlngRegistryHive
    objInParameters.sSubKeyName = strKeyPath
    
    Set objOutParameters = mobjRegistry.ExecMethod_("DeleteKey", objInParameters, , mobjContext)
    If Not objOutParameters.ReturnValue = 0 And Not objOutParameters.ReturnValue = 2 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.DeleteRegistryKey", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If

End Sub
