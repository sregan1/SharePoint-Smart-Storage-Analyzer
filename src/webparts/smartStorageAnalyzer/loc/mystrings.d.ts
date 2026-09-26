declare interface ISmartStorageAnalyzerWebPartStrings {
  // ── Property pane ──
  PropertyPane_HeaderDescription: string;
  PropertyPane_Group_General: string;
  PropertyPane_DefaultView_Label: string;
  PropertyPane_DefaultView_Home: string;
  PropertyPane_DefaultView_Tree: string;
  PropertyPane_DefaultView_List: string;
  PropertyPane_DefaultView_Report: string;

  // ── App shell ──
  App_HeaderTitle: string;
  App_SettingsAriaLabel: string;
  App_ErrorBoundary_Title: string;
  App_ErrorBoundary_BackToHome: string;
  App_ErrorBoundary_TechnicalDetails: string;
  App_SiteOwnerWarning_Title: string;
  App_SiteOwnerWarning_Body: string;

  // ── Home ──
  Home_Subtitle: string;
  Home_Card_Tree_Title: string;
  Home_Card_Tree_Alt: string;
  Home_Card_Tree_Desc: string;
  Home_Card_Tree_Button: string;
  Home_Card_List_Title: string;
  Home_Card_List_Alt: string;
  Home_Card_List_Desc: string;
  Home_Card_List_Button: string;
  Home_Card_Report_Title: string;
  Home_Card_Report_Alt: string;
  Home_Card_Report_Desc: string;
  Home_Card_Report_Button: string;
  Home_DisabledTooltip: string;
  Home_FooterNote_Label: string;
  Home_FooterNote_Body: string;

  // ── Settings ──
  Settings_Back: string;
  Settings_Title: string;
  Settings_Section_Scope: string;
  Settings_IncludeSubsites_Label: string;
  Settings_IncludeSubsites_Tooltip: string;
  Settings_IncludeSubsites_InfoAria: string;
  Settings_IncludeHidden_Label: string;
  Settings_IncludeHidden_Tooltip: string;
  Settings_IncludeHidden_InfoAria: string;
  Settings_Section_Thresholds: string;
  Settings_Thresholds_Hint: string;
  Settings_StaleAfter_Label: string;
  Settings_StaleAfter_Hint: string;
  Settings_VeryStaleAfter_Label: string;
  Settings_VeryStaleAfter_Hint: string;
  Settings_Section_Performance: string;
  Settings_Concurrency_Label: string;
  Settings_Concurrency_Tooltip: string;
  Settings_Concurrency_InfoAria: string;
  Settings_Concurrency_Hint: string;
  Settings_Section_DefaultView: string;
  Settings_DefaultView_Intro: string;
  Settings_Step1_Pre: string;
  Settings_Step1_Bold: string;
  Settings_Step1_Post: string;
  Settings_Step2_Pre: string;
  Settings_Step2_Bold: string;
  Settings_Step2_Post: string;
  Settings_Step3_Pre: string;
  Settings_Step3_Bold: string;
  Settings_Step3_Post: string;
  Settings_Step4_Bold: string;
  Settings_Step4_Post: string;

  // ── Explorer ──
  Explorer_BackToHomeButton: string;
  Explorer_BackToHomeAria: string;
  Explorer_SiteButton: string;
  Explorer_RecycleBinTooltip: string;
  Explorer_Breadcrumb_AllLibraries: string;
  Explorer_Tab_Treemap: string;
  Explorer_Tab_List: string;
  Explorer_RefreshTooltip: string;
  Explorer_RefreshButton: string;
  Explorer_CancelTooltip: string;
  Explorer_CancelButton: string;
  Explorer_CancelingButton: string;
  Explorer_ExportExcelButton: string;
  Explorer_ExportCsvButton: string;
  Explorer_SwitchToListView: string;

  Explorer_LoadingLibraries: string;
  Explorer_NoLibrariesFound: string;
  Explorer_Status_Stopping: string;
  Explorer_Status_Throttled: string;
  Explorer_Status_MeasuringLibraries: string;
  Explorer_Status_ReadingContents: string;
  Explorer_ItemsRead_Singular: string;
  Explorer_ItemsRead_Plural: string;
  Explorer_ElapsedSuffix: string;
  Explorer_LongLoadNote: string;

  Explorer_CanceledBanner: string;

  Explorer_UnknownSizeFolder_Singular: string;
  Explorer_UnknownSizeFolder_Plural: string;
  Explorer_UnknownSizeCause: string;
  Explorer_UnknownSizeAdvice_Throttling: string;
  Explorer_UnknownSizeAdvice_Other: string;
  Explorer_UnknownSizeHoverHint: string;

  Explorer_UnknownRoot_Singular: string;
  Explorer_UnknownRoot_Plural: string;

  Explorer_Reason_Throttling: string;
  Explorer_Reason_PermissionDenied: string;
  Explorer_Reason_ListViewThreshold: string;
  Explorer_Reason_Unrecorded: string;

  Explorer_Col_Name: string;
  Explorer_Col_TotalStorageSize: string;
  Explorer_Col_CurrentFileSize: string;
  Explorer_Col_Items: string;
  Explorer_Col_VersionHistorySize: string;
  Explorer_Col_VersionCount: string;
  Explorer_Col_Modified: string;
  Explorer_Col_Status: string;

  Explorer_SizeUnknownTooltip: string;
  Explorer_SizeUnknownLabel: string;
  Explorer_SizeApproximateTooltip: string;
  Explorer_VersionUnmeasuredTooltip: string;
  Explorer_AgeSuffix: string;

  Explorer_NoLibrariesMeasurable: string;
  Explorer_NoSubfoldersOrFiles: string;
  Explorer_NoLibrariesOnSite: string;
  Explorer_FolderEmpty: string;

  Explorer_TreemapVersionNote: string;
  Explorer_FoldedOther_Singular_Pre: string;
  Explorer_FoldedOther_Plural_Pre: string;
  Explorer_FoldedOther_Post: string;
  Explorer_ListVersionNote: string;

  Explorer_Error_LoadLibraries: string;
  Explorer_Error_LoadFolderContents: string;
  Explorer_Error_LoadFolderFiles: string;
  Explorer_Error_ExcelExport: string;

  // ── Storage Report ──
  Report_Stage_Discovering: string;
  Report_Stage_Items: string;
  Report_Stage_VersionProbe: string;
  Report_Stage_VersionBulk: string;
  Report_Stage_VersionValidate: string;
  Report_Stage_VersionPerFile: string;

  Report_Tooltip_RequestsSoFar: string;
  Report_Tooltip_ItemsRead: string;
  Report_Tooltip_ItemsReadOfExpected: string;
  Report_Tooltip_FilesAdded: string;
  Report_Tooltip_VersionLookupsFailed: string;
  Report_Tooltip_FilesLeftUnmeasured: string;

  Report_InitialStageLabel: string;

  Report_Warning_ListingNotSaved_Quota: string;
  Report_Warning_ListingNotSaved_TooLarge: string;
  Report_Warning_ScanSavedFailed: string;
  Report_Error_ScanFailed: string;
  Report_Error_DeleteFailed: string;
  Report_Error_ExcelExportFailed: string;
  Report_Error_CsvExportFailed: string;
  Report_Error_LoadListingFailed: string;
  Report_Warning_ListingEvicted: string;

  Report_BackAria: string;
  Report_Title: string;
  Report_ViewHistoryButton: string;

  Report_CanceledPrefix: string;
  Report_PartialResultsPrefix: string;
  Report_PartialWarning_SkippedFolders_Singular: string;
  Report_PartialWarning_SkippedFolders_Plural: string;
  Report_PartialWarning_SkippedSites_Singular: string;
  Report_PartialWarning_SkippedSites_Plural: string;
  Report_PartialWarning_SkippedVersions_Singular: string;
  Report_PartialWarning_SkippedVersions_Plural: string;
  Report_PartialWarning_UnmeasuredVersions_Singular: string;
  Report_PartialWarning_UnmeasuredVersions_Plural: string;
  Report_PartialWarning_UnmeasuredVersions_QuickModeSuffix: string;
  Report_PartialWarning_UnmeasuredVersions_DefaultSuffix: string;

  Report_ShowDetailsButton: string;
  Report_HideDetailsButton: string;

  Report_SkippedFoldersHeading: string;
  Report_SkippedFoldersShowingFirst: string;
  Report_CopyToClipboardButton: string;

  Report_IncludeSubsitesLabel: string;
  Report_IncludeHiddenLabel: string;
  Report_IncludeVersionsLabel: string;
  Report_RunScanButton: string;
  Report_ScanningButton: string;
  Report_CancelButton: string;
  Report_CancelingButton: string;

  Report_LibraryPrefix: string;
  Report_StoppingMessage: string;
  Report_ThrottledMessage: string;
  Report_LibrariesProgress: string;
  Report_FilesFound: string;
  Report_EtaSuffix: string;
  Report_ScanDetailsLabel: string;

  Report_TotalStorageSizeLabel: string;
  Report_TotalStorageSizeInfoLabel: string;
  Report_TotalStorageSizeInfo_Incomplete: string;
  Report_TotalStorageSizeInfo_Complete: string;
  Report_CurrentFileSizeLabel: string;
  Report_CurrentFileSizeInfoLabel: string;
  Report_CurrentFileSizeInfo: string;
  Report_VersionHistorySizeLabel: string;
  Report_VersionHistorySizeInfoLabel: string;
  Report_VersionHistorySizeInfo: string;
  Report_RetainedVersions_Singular: string;
  Report_RetainedVersions_Plural: string;
  Report_MeasuredPerFile_Incomplete: string;
  Report_MeasuredPerFile_Complete: string;
  Report_MeasuredPerFile_None: string;

  Report_FilesScannedLabel: string;
  Report_StaleLabel: string;
  Report_VeryStaleLabel: string;
  Report_ShowArchivalOnlyLabel: string;
  Report_ExportExcelButton: string;
  Report_ExportCsvButton: string;
  Report_ExportingButton: string;

  Report_Col_Library: string;
  Report_Col_Name: string;
  Report_Col_CurrentFileSize: string;
  Report_Col_VersionHistorySize: string;
  Report_Col_VersionCount: string;
  Report_Col_Modified: string;
  Report_Col_Author: string;
  Report_Col_Status: string;
  Report_ModifiedAgo: string;

  Report_ViewingSavedScan: string;
  Report_ListingEvictedNote: string;

  Report_ScanHistoryHeading: string;
  Report_NoSavedScans: string;
  Report_SelectTwoToCompare: string;
  Report_ShowOnlyThisSite: string;
  Report_ShowOtherSites: string;
  Report_NoSavedScansForSite: string;
  Report_CompareScanAria: string;
  Report_TotalBadgeSuffix: string;
  Report_CurrentFilesBadgeSuffix: string;
  Report_VersionHistoryBadgeSuffix: string;
  Report_StaleBadgeSuffix: string;
  Report_NoFileListBadge: string;
  Report_NoFileListTooltip: string;
  Report_ViewButton: string;
  Report_DeleteConfirmText: string;
  Report_DeleteButton: string;
  Report_KeepButton: string;
  Report_DeleteScanAria: string;

  Report_ComparisonHeading: string;
  Report_ComparingText: string;
  Report_CrossSiteWarning: string;
  Report_SizeChangeLabel: string;
  Report_NewArchivalCandidates: string;
  Report_ResolvedStale: string;
  Report_FileCountChange: string;
  Report_ListingMissingNote: string;

  // ── Tier badges (shared) ──
  Tier_Active_Label: string;
  Tier_Active_Description: string;
  Tier_Stale_Label: string;
  Tier_Stale_Description: string;
  Tier_VeryStale_Label: string;
  Tier_VeryStale_Description: string;

  // ── SizeBar ──
  SizeBar_AriaLabel: string;

  // ── StorageTable ──
  Table_ShowingRange: string;
  Table_PreviousButton: string;
  Table_PageOf: string;
  Table_NextButton: string;

  // ── Treemap ──
  Treemap_Rendering: string;
  Treemap_Legend_Folder: string;
  Treemap_Legend_Active: string;
  Treemap_Legend_Stale: string;
  Treemap_Legend_VeryStale: string;

  Treemap_Other_SizeCombinedWithVersion: string;
  Treemap_Other_SizeCombined: string;
  Treemap_Other_Tooltip: string;

  Treemap_Folder_UnknownWithError: string;
  Treemap_Folder_UnknownNoError: string;
  Treemap_Folder_FileCount_Singular: string;
  Treemap_Folder_FileCount_Plural: string;
  Treemap_Folder_SizeApproximate: string;
  Treemap_Folder_LastModifiedSuffix: string;
  Treemap_Folder_Tooltip: string;

  Treemap_File_SizeWithVersion: string;
  Treemap_File_SizeUnmeasuredVersion: string;
  Treemap_File_ModifiedSuffix: string;
  Treemap_File_Tooltip: string;

  Treemap_Detail_SizeUnknown: string;
  Treemap_Detail_FileCount_Singular: string;
  Treemap_Detail_FileCount_Plural: string;
  Treemap_Detail_FileWithVersion: string;
  Treemap_Detail_FileWithoutVersion: string;

  // ── Excel/CSV export ──
  Export_ReportTitle: string;
  Export_FolderListingTitle: string;
  Export_SiteUrl: string;
  Export_Generated: string;
  Export_ScanDuration: string;
  Export_FilesScanned: string;
  Export_VersionHistoryCount: string;
  Export_TotalStorageSize: string;
  Export_CurrentFileSize: string;
  Export_VersionHistorySize: string;
  Export_StaleFiles: string;
  Export_StaleSize: string;
  Export_VeryStaleFiles: string;
  Export_VeryStaleSize: string;

  Export_Col_Library: string;
  Export_Col_Path: string;
  Export_Col_Name: string;
  Export_Col_CurrentFileSizeBytes: string;
  Export_Col_VersionHistorySizeBytes: string;
  Export_Col_VersionCount: string;
  Export_Col_Created: string;
  Export_Col_Modified: string;
  Export_Col_AgeDays: string;
  Export_Col_Author: string;
  Export_Col_Tier: string;
  Export_Col_Type: string;
  Export_Col_TotalStorageSizeBytes: string;
  Export_Col_Items: string;
  Export_Col_Status: string;

  Export_TypeFolder: string;
  Export_TypeFile: string;
  Export_Unknown: string;
}

declare module 'SmartStorageAnalyzerWebPartStrings' {
  const strings: ISmartStorageAnalyzerWebPartStrings;
  export = strings;
}
