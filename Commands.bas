Attribute VB_Name = "Commands"
'Callback for ttInitButton onAction
Public Sub ttInitSub(control As IRibbonControl)
    IWRatings.Init_Sys
End Sub

'Callback for ttPlayerLibraryButton onAction
Public Sub ttShowHidePlayerLibSub(control As IRibbonControl)
    IWRatings.Plyr_Lib
End Sub

'Callback for ttGetResultsButton onAction
Public Sub ttGetResultsSub(control As IRibbonControl)
    IWRatings.Get_Results
End Sub

'Callback for ttUpdateStatsButton onAction
Public Sub ttUpdateStatsSub(control As IRibbonControl)
    IWRatings.Upd_Stats
End Sub

'Callback for ttGenerateReportsButton onAction
Public Sub ttGenerateReportsSub(control As IRibbonControl)
    IWRatings.Run_Reports
End Sub

'Callback for ttRatingJumpersButton onAction
Public Sub ttShowReportRatingJumpersSub(control As IRibbonControl)
    IWRatings.V_Rpt1
End Sub

'Callback for ttRatingsByPlayerButton onAction
Public Sub ttShowReportRatingsByPlayerSub(control As IRibbonControl)
    IWRatings.V_Rpt2
End Sub

'Callback for ttRatingsByRatingButton onAction
Public Sub ttShowReportRatingsByRatingSub(control As IRibbonControl)
    IWRatings.V_Rpt3
End Sub

'Callback for ttPlayerStatsButton onAction
Public Sub ttShowReportPlayerStatsSub(control As IRibbonControl)
    IWRatings.V_Rpt4
End Sub

