Attribute VB_Name = "DEE_RibbonXML"
Option Explicit

' =============================================================================
' DEE_RibbonXML.bas -- Ribbon XML served via IRibbonExtensibility
' Protocol DEE v2 -- Project Controls Add-in
' =============================================================================
' Called by ThisWorkbook.GetCustomUI(). Returns the complete Ribbon XML.
' No ZIP injection required -- Excel calls this at load time.
' =============================================================================

Public Function GetRibbonXML() As String
    Dim x As String
    x = "<?xml version=""1.0"" encoding=""UTF-8""?>"
    x = x & "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    x = x & "<ribbon><tabs>"
    x = x & "<tab id=""tabDEE"" label=""Protocol DEE"">"

    ' WBS Group
    x = x & "<group id=""grpWBS"" label=""WBS"">"
    x = x & "<button id=""btnWBSColor"" label=""Apply Colors"" size=""large"" onAction=""btn_WBSColor"" imageMso=""FillColor""/>"
    x = x & "<button id=""btnWBSGroup"" label=""Outline Group"" size=""large"" onAction=""btn_WBSGroup"" imageMso=""Group""/>"
    x = x & "<button id=""btnWBSAggregator"" label=""Aggregator"" size=""normal"" onAction=""btn_WBSAggregator"" imageMso=""MathSymbols""/>"
    x = x & "<button id=""btnWBSColumns"" label=""WBS Columns"" size=""normal"" onAction=""btn_WBSColumns"" imageMso=""TableInsert""/>"
    x = x & "<checkBox id=""chkSkipHeader"" label=""Skip Header Row"" getPressed=""chk_SkipHeader_getPressed"" onAction=""chk_SkipHeader""/>"
    x = x & "<comboBox id=""cmbWBSLevel"" label=""Show Level"" onChange=""cmb_WBSLevelFilter"" sizeString=""Level 1"">"
    x = x & "<item id=""lvlAll"" label=""All Levels""/>"
    x = x & "<item id=""lvl1"" label=""Level 1""/>"
    x = x & "<item id=""lvl2"" label=""Level 2""/>"
    x = x & "<item id=""lvl3"" label=""Level 3""/>"
    x = x & "<item id=""lvl4"" label=""Level 4""/>"
    x = x & "</comboBox>"
    x = x & "</group>"

    ' Timeline Group
    x = x & "<group id=""grpTimeline"" label=""Timeline"">"
    x = x & "<button id=""btnDrawGantt"" label=""Draw Gantt"" size=""large"" onAction=""btn_DrawGantt"" imageMso=""ChartTypeBarInsertClassic""/>"
    x = x & "<button id=""btnManageBars"" label=""Manage Bars"" size=""normal"" onAction=""btn_ManageBars"" imageMso=""ViewFullScreen""/>"
    x = x & "<button id=""btnPrintSetup"" label=""Print Setup"" size=""normal"" onAction=""btn_PrintSetup"" imageMso=""PrintPreviewAndPrint""/>"
    x = x & "</group>"

    ' Progress Group
    x = x & "<group id=""grpProgress"" label=""Progress"">"
    x = x & "<button id=""btnCreatePMS"" label=""Create PMS"" size=""large"" onAction=""btn_CreatePMS"" imageMso=""TableDesign""/>"
    x = x & "<button id=""btnCreatePMSEV"" label=""Create PMS (EV)"" size=""normal"" onAction=""btn_CreatePMSEV"" imageMso=""TableStyleClear""/>"
    x = x & "<button id=""btnLockPeriod"" label=""Lock Period"" size=""normal"" onAction=""btn_LockPeriod"" imageMso=""ProtectSheet""/>"
    x = x & "<button id=""btnSnapshot"" label=""Snapshot"" size=""normal"" onAction=""btn_Snapshot"" imageMso=""Copy""/>"
    x = x & "<button id=""btnDashboard"" label=""Dashboard"" size=""normal"" onAction=""btn_Dashboard"" imageMso=""ChartTypePieInsertClassic""/>"
    x = x & "<button id=""btnSCurveMonthly"" label=""S-Curve (Mo)"" size=""normal"" onAction=""btn_SCurveMonthly"" imageMso=""ChartTypeLineInsertClassic""/>"
    x = x & "<button id=""btnSmartSpread"" label=""Smart Spread"" size=""normal"" onAction=""btn_SmartSpread"" imageMso=""PivotTableInsert""/>"
    x = x & "</group>"

    ' XER Group
    x = x & "<group id=""grpXER"" label=""XER"">"
    x = x & "<button id=""btnLoadSchedule"" label=""Load Schedule"" size=""large"" onAction=""btn_LoadSchedule"" imageMso=""FileOpen""/>"
    x = x & "<button id=""btnExportXER"" label=""Export XER"" size=""large"" onAction=""btn_ExportXER"" imageMso=""ExportTextFile""/>"
    x = x & "</group>"

    ' Audit Group
    x = x & "<group id=""grpAudit"" label=""Audit"">"
    x = x & "<button id=""btnScheduleAudit"" label=""DCMA Audit"" size=""large"" onAction=""btn_ScheduleAudit"" imageMso=""ReviewCheckDocument""/>"
    x = x & "<button id=""btnCPM"" label=""CPM Analysis"" size=""large"" onAction=""btn_CPM"" imageMso=""ChartTypeLineInsertClassic""/>"
    x = x & "<button id=""btnLookAhead2W"" label=""2-Week LA"" size=""normal"" onAction=""btn_LookAhead2W"" imageMso=""CalendarView""/>"
    x = x & "<button id=""btnLookAhead4W"" label=""4-Week LA"" size=""normal"" onAction=""btn_LookAhead4W"" imageMso=""CalendarView""/>"
    x = x & "</group>"

    ' Baseline Group
    x = x & "<group id=""grpBaseline"" label=""Baseline"">"
    x = x & "<button id=""btnLoadBaseline"" label=""Load Baseline"" size=""large"" onAction=""btn_LoadBaseline"" imageMso=""FileOpen""/>"
    x = x & "<button id=""btnCompareBaseline"" label=""Compare"" size=""large"" onAction=""btn_CompareBaseline"" imageMso=""ReviewShowMarkup""/>"
    x = x & "<button id=""btnClearBaseline"" label=""Clear"" size=""normal"" onAction=""btn_ClearBaseline"" imageMso=""Cancel""/>"
    x = x & "</group>"

    ' Version & Audit Group
    x = x & "<group id=""grpVersion"" label=""Version &amp; Audit"">"
    x = x & "<button id=""btnSaveVersion"" label=""Save Version"" size=""large"" onAction=""btn_SaveVersion"" imageMso=""FileSave""/>"
    x = x & "<button id=""btnShowHistory"" label=""History"" size=""large"" onAction=""btn_ShowHistory"" imageMso=""Undo""/>"
    x = x & "<button id=""btnCompareSnapshots"" label=""Diff Snapshots"" size=""normal"" onAction=""btn_CompareSnapshots"" imageMso=""ReviewShowMarkup""/>"
    x = x & "<button id=""btnShowAuditLog"" label=""Audit Log"" size=""normal"" onAction=""btn_ShowAuditLog"" imageMso=""ViewFullScreen""/>"
    x = x & "</group>"

    ' Tools Group
    x = x & "<group id=""grpTools"" label=""Tools"">"
    x = x & "<button id=""btnCopyToLast"" label=""Copy To Last"" size=""normal"" onAction=""btn_CopyToLast"" imageMso=""Copy""/>"
    x = x & "<button id=""btnFillDown"" label=""Fill Down"" size=""normal"" onAction=""btn_FillDown"" imageMso=""FillDown""/>"
    x = x & "<button id=""btnTextCase"" label=""Text Case"" size=""normal"" onAction=""btn_TextCase"" imageMso=""ChangeCaseMenu""/>"
    x = x & "<button id=""btnDataQuality"" label=""Data Quality"" size=""normal"" onAction=""btn_DataQuality"" imageMso=""ReviewCheckDocument""/>"
    x = x & "<button id=""btnSettings"" label=""Settings"" size=""normal"" onAction=""btn_Settings"" imageMso=""ToolsOptions""/>"
    x = x & "<button id=""btnExportLog"" label=""Export Log"" size=""normal"" onAction=""btn_ExportLog"" imageMso=""ExportExcelTable""/>"
    x = x & "</group>"

    ' Governance Group
    x = x & "<group id=""grpGovernance"" label=""Governance"">"
    x = x & "<button id=""btnLockFields"" label=""Lock Fields"" size=""large"" onAction=""btn_LockFields"" imageMso=""ProtectSheet""/>"
    x = x & "<button id=""btnUnlockFields"" label=""Unlock Fields"" size=""normal"" onAction=""btn_UnlockFields"" imageMso=""UnProtectSheet""/>"
    x = x & "<button id=""btnDataDate"" label=""Set Data Date"" size=""large"" onAction=""btn_DataDate"" imageMso=""CalendarView""/>"
    x = x & "<button id=""btnExportAudit"" label=""Export Audit"" size=""normal"" onAction=""btn_ExportAudit"" imageMso=""ExportTextFile""/>"
    x = x & "<button id=""btnForceReset"" label=""Force Reset"" size=""normal"" onAction=""btn_ForceReset"" imageMso=""Refresh""/>"
    x = x & "</group>"

    x = x & "</tab></tabs></ribbon></customUI>"
    GetRibbonXML = x
End Function
