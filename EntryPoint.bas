Attribute VB_Name = "EntryPoint"
Option Compare Database
Option Explicit

' =============================================================================

Public Sub RegisterRubberduck()

Dim objRubberduckRegistration As RubberduckRegistration

    Set objRubberduckRegistration = New RubberduckRegistration
    objRubberduckRegistration.RegisterRubberduck

End Sub

' =============================================================================

Public Sub UnregisterRubberduck()

Dim objRubberduckRegistration As RubberduckRegistration
    
    Set objRubberduckRegistration = New RubberduckRegistration
    objRubberduckRegistration.UnregisterRubberduck

End Sub

' =============================================================================
