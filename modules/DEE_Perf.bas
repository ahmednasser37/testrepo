Attribute VB_Name = "DEE_Perf"
Option Explicit

' =============================================================================
' DEE_Perf.bas -- Performance Guards (P0.3)
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Reference-counted performance guard so nested EnterHeavyMode/ExitHeavyMode
' calls do not prematurely restore application state.
'
' USAGE:
'   Call DEE_Perf.EnterHeavyMode
'   ' ... heavy operation ...
'   Call DEE_Perf.ExitHeavyMode  ' always in finally/error handler
' =============================================================================

Private m_HeavyDepth          As Long
Private m_SavedCalc           As XlCalculation
Private m_SavedScreen         As Boolean
Private m_SavedEvents         As Boolean
Private m_SavedDisplayAlerts  As Boolean
Private m_StartTime           As Double

' ---------------------------------------------------------------------------
' EnterHeavyMode -- Disable screen/calc/events (ref-counted)
' ---------------------------------------------------------------------------
Public Sub EnterHeavyMode()
    If m_HeavyDepth = 0 Then
        m_SavedCalc          = Application.Calculation
        m_SavedScreen        = Application.ScreenUpdating
        m_SavedEvents        = Application.EnableEvents
        m_SavedDisplayAlerts = Application.DisplayAlerts
        m_StartTime          = Timer

        Application.Calculation   = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents  = False
        Application.DisplayAlerts = False
    End If
    m_HeavyDepth = m_HeavyDepth + 1
End Sub

' ---------------------------------------------------------------------------
' ExitHeavyMode -- Restore application state (ref-counted)
' ---------------------------------------------------------------------------
Public Sub ExitHeavyMode()
    If m_HeavyDepth > 0 Then m_HeavyDepth = m_HeavyDepth - 1
    If m_HeavyDepth = 0 Then
        Application.Calculation   = m_SavedCalc
        Application.ScreenUpdating = m_SavedScreen
        Application.EnableEvents  = m_SavedEvents
        Application.DisplayAlerts = m_SavedDisplayAlerts
        Application.StatusBar     = False
    End If
End Sub

' ---------------------------------------------------------------------------
' ForceReset -- Emergency recovery (e.g. after unhandled crash)
' ---------------------------------------------------------------------------
Public Sub ForceReset()
    m_HeavyDepth = 0
    Application.Calculation   = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents  = True
    Application.DisplayAlerts = True
    Application.StatusBar     = False
End Sub

' ---------------------------------------------------------------------------
' ElapsedSec -- Seconds since last EnterHeavyMode
' ---------------------------------------------------------------------------
Public Function ElapsedSec() As Double
    ElapsedSec = Timer - m_StartTime
End Function

' ---------------------------------------------------------------------------
' HeavyDepth -- Current nesting depth (for diagnostics)
' ---------------------------------------------------------------------------
Public Function HeavyDepth() As Long
    HeavyDepth = m_HeavyDepth
End Function
