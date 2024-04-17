New-UnifiedAuditLogRetentionPolicy `
-Name "10 Year Retention - Catchall (Powershell)" `
-Priority 199 `
-RetentionDuration TenYears `
-RecordTypes `
    AirAdminActionInvestigation,AirInvestigation,AirManualInvestigation,`
    ApplicationAudit,AttackSim,AzureActiveDirectoryAccountLogon,AzureActiveDirectoryStsLogon,`
    Campaign,ComplianceDLPExchange,ComplianceDLPExchangeClassification,ComplianceDLPSharePointClassification,`
    ComplianceSupervisionExchange,CortanaBriefing,DataCenterSecurityCmdlet,DataInsightsRestApiAudit,HRSignal,`
    HygieneEvent,InformationBarrierPolicyApplication,InformationWorkerProtection,Kaizala,MailSubmission,MCASAlerts,`
    MicrosoftTeamsAdmin,MicrosoftTeamsAnalytics,MicrosoftTeamsDevice,MicrosoftTeamsShifts,MipAutoLabelExchangeItem,`
    MipAutoLabelSharePointItem,MipAutoLabelSharePointPolicyLocation,MIPLabel,OfficeNative,OneDrive,PhysicalBadgingSignal,`
    PowerAppsPlan,Project,Search,SecurityComplianceAlerts,SecurityComplianceInsights,SecurityComplianceRBAC,`
    SensitivityLabelAction,SensitivityLabeledFileAction,SensitivityLabelPolicyMatch,SharePointCommentOperation,`
    SharePointListOperation,SkypeForBusinessPSTNUsage,SkypeForBusinessUsersBlocked,TeamsHealthcare,ThreatFinder,ThreatIntelligence,`
    ThreatIntelligenceAtpContent,ThreatIntelligenceUrl,UserTraining,WDATPAlerts