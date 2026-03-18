'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用|Piggy key value application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 3.8
'* Create Time: 8/5/2021
'* 1.0.2	13/5/2021 Modify New
'* 1.0.3	22/7/2021 Modify GetPigKeyValue
'* 1.0.4	23/7/2021 remove ObjAdoDBLib
'* 1.0.5	4/8/2021 Remove PigSQLSrvLib
'* 1.0.6	5/8/2021 Modify GetPigKeyValue,SavePigKeyValue
'* 1.0.7	7/8/2021 Modify New and add IsUseMemCache
'* 1.0.8	11/8/2021 Add mSavePigKeyValueToShareMem,mSetBodyBytesByShareMem
'* 1.0.9	13/8/2021 Modify mSetBodyBytesByShareMem
'* 1.0.10	13/8/2021 Modify mSaveSMHead,IsUseMemCache
'* 1.0.11	16/8/2021 Modify mSavePigKeyValueToShareMem,mSavePigKeyValueToShareMem,ShareMemRoot,GetPigKeyValue,mSetHeadBytesByShareMem,mSetBodyBytesByShareMem, and add mSaveSMBody
'* 1.0.12	17/8/2021 Add PrintDebugLog,IsPigKeyValueExists,RemovePigKeyValue and modify GetPigKeyValue,SavePigKeyValue
'* 1.0.13	19/8/2021 Modify RemoveExpItems, and add GetStatisticsXml
'* 1.0.14	22/8/2021 Add CacheLevel,ForceRefCacheTime， and modify New,mNew,SavePigKeyValue,GetPigKeyValue,RemovePigKeyValue
'* 1.0.15	23/8/2021 Modify mNew,StruStatistics,New,GetStatisticsXml,IsPigKeyValueExists, and add CacheWorkDir,mIsShareMemExists
'* 1.0.16	23/8/2021 Modify GetPigKeyValue, and Add mGetPigKeyValueByShareMem
'* 1.0.17	25/8/2021 Remove Imports PigToolsLib, change to PigToolsWinLib, and add mIsBytesMatch, mSavePigKeyValueToSM rename to mSavePigKeyValueToShareMem
'* 1.0.18	26/8/2021 Modify RemovePigKeyValue,SavePigKeyValue, and add mClearShareMem
'* 1.0.19	27/8/2021 Modify mGetPigKeyValueByShareMem
'* 1.1		29/8/2021 Chanage PigToolsWinLib to PigToolsLiteLib
'* 1.2		31/8/2021 Modify ForceRefCacheTime
'* 1.3		25/9/2021 Add mSavePigKeyValueToFile,mGetStruFileHead,mSaveHeadToFile,mSaveFileBody
'* 1.4		26/9/2021 Modify mSavePigKeyValueToFile,SavePigKeyValue,GetPigKeyValue, and add mGetPigKeyValueByFile
'* 1.5		2/10/2021 Modify New,mNew,GetPigKeyValue
'* 1.6		3/10/2021 Add StruKeyValueCtrl,mRemoveFile,mGetPigKeyValueByList,mGetPigKeyValueByShareMem, and modify GetPigKeyValue,StruStatistics,GetStatisticsXml
'* 1.7		4/10/2021 Modify GetPigKeyValue,SavePigKeyValue,mAddPigKeyValueToList,mRemoveFile,RemovePigKeyValue, and add mGetPigKeyValueByFile
'* 1.8		21/10/2021 Modify mIsCacheFileExists,SavePigKeyValue,mIsCacheFileExists,GetPigKeyValue,mSavePigKeyValueToShareMem
'* 1.9		24/10/2021 Modify SavePigKeyValue,fSaveValueLen,mGetPigKeyValueByFile,mGetPigKeyValueFromFile,StruValueHead
'* 2.0		28/10/2021 Add fGetSMNameHeadAndBody,mIsCacheFileExists
'* 2.1		30/11/2021 Modify mSetHeadBytesByShareMem,mGetPigKeyValueFromFile, remove mGetStruFileHead
'* 2.2		1/12/2021 Modify mSaveSMHead,mGetPigKeyValueFromFile,mSaveHeadToFile,mSetHeadBytesByShareMem,mGetPigKeyValueFromShareMem, remove mGetSMNamePart
'* 2.3		2/12/2021 Imports System.IO, Add mChkSaveBodyBytes, Modify mSavePigKeyValueToFile,mSetHeadBytesByShareMem
'* 2.5		5/12/2021 Add new SavePigKeyValue,DefaultSaveType,DefaultTextType
'* 2.6		7/12/2021 Modify GetPigKeyValue,mIsCacheFileExists
'* 2.7		8/12/2021 Modify EnmCacheLevel,StruValueHead, add mSetHeadBytesByFile,StruValueHead2PigBytes,mIsCacheFileExists
'* 2.8		9/12/2021 Add PigKeyValue2StruValueHead, remove StruValueHead
'* 3.0		10/12/2021 A large number of code rewrites, the interface is upgraded to 3.0, and the interface is no longer compatible with versions below 3.0
'* 3.1		10/12/2021 Add GetKeyTitle,mGetPigKeyValueByFile,LoadBytesFromFile, modify mRemoveFile,RemovePigKeyValue,mGetPigKeyValueByShareMem,mSavePigKeyValueToShareMem
'* 3.2		12/12/2021 Modify mSavePigKeyValue,SaveBytesToShareMem,SavePigKeyValue,mClearShareMem,mGetPigKeyValueFromShareMem
'* 3.3		14/12/2021 Modify mGetPigKeyValueFromFile,mSavePigKeyValueToFile
'* 3.5		17/12/2021 Modify mRemoveFile
'* 3.6		2/1/2022 Modify mGetPigKeyValueByShareMem,SaveBytesToShareMem,mClearShareMem,GetPigKeyValue,mGetPigKeyValueByFile
'* 3.7		26/7/2022 Imports PigToolsWinLib
'* 3.8		29/7/2022 Modify Imports
'************************************

Imports PigToolsLiteLib
Imports System.IO

Public Class PigKeyValueApp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "3.8.2"
	Private Const SM_HEAD_LEN As Integer = 40
	Private ReadOnly moPigFunc As New PigFunc

	''' <summary>
	''' Value type, non text type, saved in byte array
	''' </summary>
	Public Enum EnmCacheLevel
		Unknow = 0
		''' <summary>
		''' Program for single process multithreading
		''' </summary>
		ToList = 1
		''' <summary>
		''' It is applicable to multi-process and multi-threaded programs under the same user session or IIS application pools.
		''' </summary>
		ToShareMem = 2
		''' <summary>
		''' It is suitable for any multi process and multi thread program on the same host.
		''' </summary>
		ToFile = 3
		''''' <summary>
		''''' It is suitable for multi server, multi process and multi-threaded programs, and has the highest requirements for the availability of cached content, but the writing performance is poor, but the advantage is that it can share the database with the application to reduce the point of failure.
		''''' </summary>
		'ToDB = 4
		''''' <summary>
		''''' It is suitable for multi server, multi process and multi thread programs. The read and write performance is very good, but redis needs to be installed, which needs to increase the cost of managing the high availability of redis.
		''''' </summary>
		'ToRedis = 5
	End Enum

	Public ReadOnly Property PigKeyValues As New PigKeyValues
	Friend Property ShareMemRoot As String = ""

	Private msuStatistics As StruStatistics

	''' <summary>
	''' 统计信息结构
	''' </summary>
	Private Structure StruStatistics
		Dim GetCount As Long
		Dim GetFailCount As Long
		Dim CacheCount As Long
		Dim CacheByListCount As Long
		Dim CacheByShareMemCount As Long
		Dim CacheByFileCount As Long
		Dim CacheByDBCount As Long
		Dim CacheByRedisCount As Long
		Dim SaveCount As Long
		Dim SaveFailCount As Long
		Dim SaveToListCount As Long
		Dim SaveToShareMemCount As Long
		Dim SaveToFileCount As Long
		Dim SaveToDBCount As Long
		Dim SaveToRedisCount As Long
		Dim RemoveCount As Long
		Dim RemoveFailCount As Long
		Dim RemoveExpiredListCount As Long
		Dim RemoveExpiredShareMemCount As Long
		Dim RemoveExpiredFileCount As Long
		Dim RemoveExpiredDBCount As Long
		Dim RemoveExpiredRedisCount As Long
	End Structure


	''' <summary>
	''' 键值控制结构
	''' </summary>
	Private Structure StruKeyValueCtrl
		Dim IsGetByShareMem As Boolean
		Dim IsGetByFile As Boolean
		Dim IsRemoveList As Boolean
		Dim IsClearShareMem As Boolean
		Dim IsRemoveFile As Boolean
		Dim IsRefLastRefCacheTime As Boolean
		Dim ListValueMD5 As String
		Dim ShareMemValueMD5 As String
		Dim IsSaveList As Boolean
		Dim IsSaveShareMem As Boolean
		Dim IsSaveFile As Boolean
	End Structure

	Public Sub New()
		MyBase.New(CLS_VERSION)
		mNew("", EnmCacheLevel.ToList)
	End Sub

	Private Sub mNew(Optional ShareMemRootOrCacheWorkDir As String = "", Optional CacheLevel As EnmCacheLevel = EnmCacheLevel.ToShareMem, Optional ForceRefCacheTime As Integer = 60)
		Try
			Me.CacheLevel = CacheLevel
			Select Case Me.CacheLevel
				Case EnmCacheLevel.ToList
					Me.ShareMemRoot = ""
					Me.CacheWorkDir = ""
				Case EnmCacheLevel.ToShareMem
					If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
					If ShareMemRootOrCacheWorkDir = "" Then ShareMemRootOrCacheWorkDir = Me.AppTitle
					Me.ShareMemRoot = ShareMemRootOrCacheWorkDir
					Me.CacheWorkDir = ""
				Case EnmCacheLevel.ToFile
					If ShareMemRootOrCacheWorkDir = "" Then ShareMemRootOrCacheWorkDir = Me.AppPath
					Me.ShareMemRoot = ShareMemRootOrCacheWorkDir
					Me.CacheWorkDir = ShareMemRootOrCacheWorkDir
				Case Else
					Throw New Exception("Currently unsupported cachelevel")
			End Select
			If Me.ShareMemRoot <> "" Then
				Dim oPigMD5 As PigMD5
				oPigMD5 = New PigMD5(ShareMemRootOrCacheWorkDir, PigMD5.enmTextType.UTF8)
				Me.ShareMemRoot = oPigMD5.PigMD5()
			End If
			If ForceRefCacheTime < 30 Then
				Me.ForceRefCacheTime = 30
			Else
				Me.ForceRefCacheTime = ForceRefCacheTime
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", ex)
			Me.CacheLevel = EnmCacheLevel.Unknow
		End Try
	End Sub
	Public Sub New(ShareMemRoot As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRoot)
	End Sub

	Public Sub New(ShareMemRootOrCacheWorkDir As String, CacheLevel As EnmCacheLevel)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRootOrCacheWorkDir, CacheLevel)
	End Sub


	Private Function mIsCacheFileExists(KeyName As String) As Boolean
		Dim LOG As New PigStepLog("mIsCacheFileExists")
		Try
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(KeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(LOG.Ret)
			End If
			Dim strFilePath As String
			LOG.StepName = "Check File"
			strFilePath = Me.CacheWorkDir & Me.OsPathSep & strHeadTitle
			If IO.File.Exists(strFilePath) = False Then
				LOG.AddStepNameInf(strFilePath)
				Throw New Exception("File not found")
			End If
			strFilePath = Me.CacheWorkDir & Me.OsPathSep & strBodyTitle
			If IO.File.Exists(strFilePath) = False Then
				LOG.AddStepNameInf(strFilePath)
				Throw New Exception("File not found")
			End If
			Return "OK"
		Catch ex As Exception
			Me.PrintDebugLog(LOG.SubName, LOG.StepName, ex.Message.ToString)
			Return False
		End Try
	End Function

	Private Function mIsHeadBytesInit(ByRef AbHead As Byte()) As Boolean
		Try
			If AbHead.Length <> SM_HEAD_LEN Then
				Return False
			Else
				mIsHeadBytesInit = False
				For Each bytAny In AbHead
					If bytAny <> 0 Then
						mIsHeadBytesInit = True
						Exit For
					End If
				Next
			End If
		Catch ex As Exception
			Me.PrintDebugLog("mIsHeadBytesInit", ex.Message.ToString)
			Return False
		End Try
	End Function

	Private Function mGetPigKeyValueFromShareMem(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Dim LOG As New PigStepLog("mGetPigKeyValueFromShareMem")
		Try
			If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
			If OutPigKeyValue IsNot Nothing Then OutPigKeyValue = Nothing
			LOG.StepName = "New PigKeyValue"
			OutPigKeyValue = New PigKeyValue(KeyName)
			If OutPigKeyValue.LastErr <> "" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(OutPigKeyValue.LastErr)
			End If
			'--------
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(KeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Dim pbMain As PigBytes = Nothing
			'--------
			LOG.StepName = "LoadBytesFromShareMem"
			LOG.Ret = Me.LoadBytesFromShareMem(strHeadTitle, SM_HEAD_LEN, pbMain)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName & "." & strHeadTitle)
				Throw New Exception(LOG.Ret)
			End If
			If pbMain Is Nothing Then
				LOG.AddStepNameInf(KeyName & "." & strHeadTitle)
				Throw New Exception("pbMain Is Nothing")
			End If
			If Me.mIsHeadBytesInit(pbMain.Main) = True Then
				LOG.StepName = "OutPigKeyValue.LoadHead"
				LOG.Ret = OutPigKeyValue.LoadHead(pbMain)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf(KeyName)
					Throw New Exception(LOG.Ret)
				End If
				'--------
				pbMain = Nothing
				LOG.StepName = "LoadBytesFromShareMem"
				LOG.Ret = Me.LoadBytesFromShareMem(strBodyTitle, OutPigKeyValue.BodyLen, pbMain)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf(KeyName & "." & strBodyTitle)
					Throw New Exception(LOG.Ret)
				End If
				If pbMain Is Nothing Then
					LOG.AddStepNameInf(KeyName & "." & strBodyTitle)
					Throw New Exception("pbMain Is Nothing")
				End If
				LOG.StepName = "OutPigKeyValue.LoadBody"
				LOG.Ret = OutPigKeyValue.LoadBody(pbMain.Main)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf(KeyName)
					Throw New Exception(LOG.Ret)
				End If
			End If
			pbMain = Nothing
			'--------
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function


	Private Function mClearShareMem(PkvOld As PigKeyValue) As String
		Dim LOG As New PigStepLog("mClearShareMem")
		Try
			If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
			If PkvOld Is Nothing Then Throw New Exception("PkvOld Is Nothing")
			Dim strKeyName As String = PkvOld.KeyName
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(strKeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName)
				Throw New Exception(LOG.Ret)
			End If
			Dim abTmp As Byte()
			ReDim abTmp(SM_HEAD_LEN - 1)
			LOG.StepName = "SaveBytesToShareMem"
			LOG.Ret = Me.SaveBytesToShareMem(strHeadTitle, SM_HEAD_LEN, abTmp)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf("HeadTitle")
				Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
			End If
			Dim lngBodyLen As Long = PkvOld.BodyLen
			If lngBodyLen > 0 Then
				ReDim abTmp(lngBodyLen - 1)
				LOG.Ret = Me.SaveBytesToShareMem(strBodyTitle, lngBodyLen, abTmp)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf("BodyTitle")
					Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
				End If
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function



	Private Function mRemovePigKeyValueFromList(KeyName As String) As String
		Dim LOG As New PigStepLog("mRemovePigKeyValueFromList")
		Try
			If Me.PigKeyValues.IsItemExists(KeyName) = False Then
				Return "OK"
			Else
				LOG.StepName = "PigKeyValues.Remove"
				Me.PigKeyValues.Remove(KeyName)
				If Me.PigKeyValues.LastErr <> "" Then
					LOG.AddStepNameInf(KeyName)
					Throw New Exception(Me.PigKeyValues.LastErr)
				End If
				Return "OK"
			End If
		Catch ex As Exception
			Return Me.GetSubErrInf("mRemovePigKeyValueFromList", LOG.StepName, ex)
		End Try
	End Function


	Private Function mAddPigKeyValueToList(NewItem As PigKeyValue) As String
		Dim LOG As New PigStepLog("mAddPigKeyValueToList")
		Try
			Dim strKeyName As String = NewItem.KeyName
			'If Me.PigKeyValues.IsItemExists(strKeyName) = True Then
			'	LOG.StepName = "mRemovePigKeyValueFromList"
			'	LOG.Ret = Me.mRemovePigKeyValueFromList(strKeyName)
			'	If log.Ret <> "OK" Then
			'		Me.PrintDebugLog(log.SUBNAME, LOG.StepName, LOG.Ret)
			'		If Me.PigKeyValues.IsItemExists(strKeyName) Then
			'			LOG.StepName &= "(" & strKeyName & ")"
			'			Throw New Exception("Cannot remove exists item")
			'		End If
			'	End If
			'End If
			NewItem.LastRefCacheTime = Now
			LOG.StepName = "PigKeyValues.Add"
			Me.PigKeyValues.Add(NewItem)
			If Me.PigKeyValues.LastErr <> "" Then
				LOG.StepName &= "(" & strKeyName & ")"
				Throw New Exception(strKeyName)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mGetPigKeyValueByShareMem(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Dim LOG As New PigStepLog("mGetPigKeyValueByShareMem")
		Try
			If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
			msuStatistics.GetCount += 1
			LOG.StepName = "mGetPigKeyValueFromShareMem"
			LOG.Ret = Me.mGetPigKeyValueFromShareMem(KeyName, OutPigKeyValue)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
			End If
			If OutPigKeyValue IsNot Nothing Then
				If OutPigKeyValue.IsExpired = True Then
					msuStatistics.RemoveCount += 1
					msuStatistics.RemoveExpiredShareMemCount += 1
					LOG.StepName = "mClearShareMem"
					LOG.Ret = Me.mClearShareMem(OutPigKeyValue)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(KeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
						msuStatistics.RemoveFailCount += 1
					End If
					OutPigKeyValue = Nothing
					If Me.PigKeyValues.IsItemExists(KeyName) = True Then
						msuStatistics.RemoveCount += 1
						msuStatistics.RemoveExpiredListCount += 1
						LOG.StepName = "mRemovePigKeyValueFromList"
						LOG.Ret = Me.mRemovePigKeyValueFromList(KeyName)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(KeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							msuStatistics.RemoveFailCount += 1
						End If
					End If
				Else
					msuStatistics.CacheCount += 1
					msuStatistics.CacheByShareMemCount += 1
				End If
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mGetPigKeyValueByFile(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Dim LOG As New PigStepLog("mGetPigKeyValueByFile")
		Try
			msuStatistics.GetCount += 1
			LOG.StepName = "mGetPigKeyValueFromFile"
			LOG.Ret = Me.mGetPigKeyValueFromFile(KeyName, OutPigKeyValue)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
			End If
			If OutPigKeyValue IsNot Nothing Then
				If OutPigKeyValue.IsExpired = True Then
					msuStatistics.RemoveCount += 1
					msuStatistics.RemoveExpiredFileCount += 1
					If OutPigKeyValue.Parent Is Nothing Then OutPigKeyValue.Parent = Me
					LOG.StepName = "mRemoveFile"
					LOG.Ret = Me.mRemoveFile(OutPigKeyValue.KeyName)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(KeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
						msuStatistics.RemoveFailCount += 1
					End If
					OutPigKeyValue = Nothing
					If Me.IsWindows = True Then
						msuStatistics.RemoveCount += 1
						msuStatistics.RemoveExpiredShareMemCount += 1
						LOG.StepName = "mClearShareMem"
						LOG.Ret = Me.mClearShareMem(OutPigKeyValue)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(KeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							msuStatistics.RemoveFailCount += 1
						End If
					End If
					If Me.PigKeyValues.IsItemExists(KeyName) = True Then
						msuStatistics.RemoveCount += 1
						msuStatistics.RemoveExpiredListCount += 1
						LOG.StepName = "mRemovePigKeyValueFromList"
						LOG.Ret = Me.mRemovePigKeyValueFromList(KeyName)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(KeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							msuStatistics.RemoveFailCount += 1
						End If
					End If
				Else
					msuStatistics.CacheCount += 1
					msuStatistics.CacheByFileCount += 1
				End If
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mGetPigKeyValueByList(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Dim LOG As New PigStepLog("mGetPigKeyValueByList")
		Try
			msuStatistics.GetCount += 1
			LOG.StepName = "GetByList"
			OutPigKeyValue = Me.PigKeyValues.Item(KeyName)
			If Me.PigKeyValues.LastErr <> "" Then
				LOG.AddStepNameInf(KeyName)
				Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.PigKeyValues.LastErr)
			End If
			If OutPigKeyValue IsNot Nothing Then
				If OutPigKeyValue.IsExpired = True Then
					msuStatistics.RemoveCount += 1
					msuStatistics.RemoveExpiredListCount += 1
					LOG.StepName = "mRemovePigKeyValueFromList"
					LOG.Ret = Me.mRemovePigKeyValueFromList(KeyName)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(KeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.PigKeyValues.LastErr)
						msuStatistics.RemoveFailCount += 1
					End If
					OutPigKeyValue = Nothing
				Else
					msuStatistics.CacheCount += 1
					msuStatistics.CacheByListCount += 1
				End If
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function
	Public Function GetPigKeyValue(KeyName As String) As PigKeyValue
		Dim LOG As New PigStepLog("GetPigKeyValue")
		Try
			GetPigKeyValue = Nothing
			Dim pkvList As PigKeyValue = Nothing
			LOG.StepName = "mGetPigKeyValueByList"
			LOG.Ret = Me.mGetPigKeyValueByList(KeyName, pkvList)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(LOG.Ret)
			End If
			Select Case Me.CacheLevel
				Case EnmCacheLevel.ToList
					GetPigKeyValue = pkvList
					pkvList = Nothing
				Case EnmCacheLevel.ToShareMem
					If Me.IsWindows = False Then
						LOG.StepName = "Check OS"
						Throw New Exception("This Function can only be used on windows")
					End If
					Dim bolIsGetByShareMem As Boolean = False
					If pkvList Is Nothing Then
						bolIsGetByShareMem = True
					Else
						'If pkvList.Parent Is Nothing Then pkvList.Parent = Me
						If pkvList.fIsForceRefCache = True Then
							bolIsGetByShareMem = True
						End If
					End If
					If bolIsGetByShareMem = True Then
						LOG.StepName = "mGetPigKeyValueByShareMem.ToShareMem"
						LOG.Ret = Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(KeyName)
							Throw New Exception(LOG.Ret)
						End If
						Dim bolIsRemoveList As Boolean = False
						Dim bolIsAddList As Boolean = False
						If GetPigKeyValue Is Nothing Then
							If pkvList IsNot Nothing Then bolIsRemoveList = True
						ElseIf pkvList IsNot Nothing Then
							If GetPigKeyValue.IsMatchAnother(pkvList) = False Then
								bolIsRemoveList = True
								bolIsAddList = True
							Else
								pkvList.LastRefCacheTime = Now
							End If
						Else
							bolIsAddList = True
						End If
						If bolIsRemoveList = True Then
							LOG.StepName = "mClearShareMem"
							LOG.Ret = Me.mClearShareMem(pkvList)
							If LOG.Ret <> "OK" Then
								LOG.AddStepNameInf(KeyName)
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
								msuStatistics.RemoveFailCount += 1
							End If
							pkvList = Nothing
						End If
						If bolIsAddList = True Then
							msuStatistics.SaveToListCount += 1
							LOG.StepName = "mAddPigKeyValueToList"
							LOG.Ret = Me.mAddPigKeyValueToList(GetPigKeyValue)
							If LOG.Ret <> "OK" Then
								LOG.AddStepNameInf(KeyName)
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
								msuStatistics.SaveFailCount += 1
							End If
						End If
					Else
						GetPigKeyValue = pkvList
						pkvList = Nothing
					End If
				Case EnmCacheLevel.ToFile
					Dim bolIsGetByShareMem As Boolean = False
					Dim bolIsGetByFile As Boolean = False
					If pkvList Is Nothing Then
						If Me.IsWindows = True Then
							bolIsGetByShareMem = True
						Else
							bolIsGetByFile = True
						End If
					Else
						If pkvList.Parent Is Nothing Then pkvList.Parent = Me
						If pkvList.fIsForceRefCache = True Then
							bolIsGetByFile = True
						End If
					End If
					Dim pkvShareMem As PigKeyValue = Nothing
					If bolIsGetByFile = True Then
						LOG.StepName = "mGetPigKeyValueByFile.ToFile"
						LOG.Ret = Me.mGetPigKeyValueByFile(KeyName, GetPigKeyValue)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(KeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
						End If
						If GetPigKeyValue Is Nothing Then
							LOG.StepName = "RemovePigKeyValue.ToFile"
							LOG.Ret = Me.RemovePigKeyValue(pkvList, EnmCacheLevel.ToShareMem)
							If LOG.Ret <> "OK" Then
								LOG.AddStepNameInf(KeyName)
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							End If
						Else
							If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
							If Me.IsWindows = True Then
								LOG.StepName = "mGetPigKeyValueByShareMem.ToFile"
								LOG.Ret = Me.mGetPigKeyValueByShareMem(KeyName, pkvShareMem)
								If LOG.Ret <> "OK" Then
									LOG.AddStepNameInf(KeyName)
									Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
								End If
								Dim bolIsSaveShareMem As Boolean = False
								If pkvShareMem IsNot Nothing Then
									If pkvShareMem.IsMatchAnother(GetPigKeyValue) = False Then
										LOG.StepName = "mClearShareMem.ToFile"
										LOG.Ret = Me.mClearShareMem(pkvShareMem)
										If LOG.Ret <> "OK" Then
											LOG.AddStepNameInf(KeyName)
											Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
										End If
										bolIsSaveShareMem = True
									End If
								Else
									bolIsSaveShareMem = True
								End If
								If bolIsSaveShareMem = True Then
									LOG.StepName = "mSavePigKeyValueToShareMem.ToFile"
									LOG.Ret = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
									If LOG.Ret <> "OK" Then
										LOG.AddStepNameInf(KeyName)
										Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
									End If
								End If
							End If
						End If
					ElseIf bolIsGetByShareMem = True Then
						LOG.StepName = "mGetPigKeyValueByShareMem.ToFile"
						LOG.Ret = Me.mGetPigKeyValueByShareMem(KeyName, pkvShareMem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(KeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
						End If
						If pkvShareMem Is Nothing Then
							LOG.StepName = "mGetPigKeyValueByFile.ToFile2"
							LOG.Ret = Me.mGetPigKeyValueByFile(KeyName, GetPigKeyValue)
							If LOG.Ret <> "OK" Then
								LOG.AddStepNameInf(KeyName)
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							End If
							If GetPigKeyValue IsNot Nothing Then
								If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
								msuStatistics.SaveToShareMemCount += 1
								LOG.StepName = "mSavePigKeyValueToShareMem.ToFile2"
								LOG.AddStepNameInf(GetPigKeyValue.StrValue)
								LOG.Ret = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
								If LOG.Ret <> "OK" Then
									LOG.AddStepNameInf(KeyName)
									Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
								End If
							End If
						Else
							If pkvShareMem.Parent Is Nothing Then pkvShareMem.Parent = Me
							msuStatistics.SaveToListCount += 1
							LOG.StepName = "mAddPigKeyValueToList"
							LOG.Ret = Me.mAddPigKeyValueToList(pkvShareMem)
							If LOG.Ret <> "OK" Then
								LOG.AddStepNameInf(KeyName)
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
								msuStatistics.SaveFailCount += 1
							End If
							GetPigKeyValue = pkvShareMem
							pkvList = Nothing
						End If
					Else
						GetPigKeyValue = pkvList
						pkvList = Nothing
					End If
				Case Else
					LOG.StepName = KeyName
					Throw New Exception("Unsupported CacheLevel")
			End Select
			Me.ClearErr()
		Catch ex As Exception
			msuStatistics.GetFailCount += 1
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function


	Public Function SaveBytesToShareMem(SMName As String, SMLen As Long, ByRef AbIn As Byte()) As String
		Dim LOG As New PigStepLog("SaveBytesToShareMem")
		Try
			If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
			Dim smMain As New ShareMem
			LOG.StepName = "Init"
			LOG.Ret = smMain.Init(SMName, SMLen)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(SMName)
				Throw New Exception(LOG.Ret)
			End If
			LOG.StepName = "Write"
			LOG.Ret = smMain.Write(AbIn)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(SMName)
				Throw New Exception(LOG.Ret)
			End If
			smMain = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function



	''' <summary>
	''' 从文件读取一个 PigBytes 对象
	''' </summary>
	''' <param name="FilePath">文件路径</param>
	''' <param name="PbOut">输出的PigBytes</param>
	''' <returns></returns>
	Public Function LoadBytesFromFile(FilePath As String, ByRef PbOut As PigBytes) As String
		Dim LOG As New PigStepLog("LoadBytesFromFile")
		Try
			LOG.StepName = "New PigFile"
			Dim pfHead As New PigFile(FilePath)
			If pfHead.LastErr <> "" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(pfHead.LastErr)
			End If
			LOG.StepName = "Head.LoadFile"
			LOG.Ret = pfHead.LoadFile
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(LOG.Ret)
			End If
			LOG.StepName = "Set PbOut"
			PbOut = pfHead.GbMain
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	''' <summary>
	''' 保存字节数组到文件
	''' </summary>
	''' <param name="FilePath">文件路径</param>
	''' <param name="PbOut">输出的PigBytes</param>
	''' <returns></returns>
	Public Function SaveBytesToFile(FilePath As String, ByRef AbIn As Byte()) As String
		Dim LOG As New PigStepLog("SaveBytesToFile")
		Try
			Dim oPigFile As PigFile
			LOG.StepName = "New PigFile"
			oPigFile = New PigFile(FilePath)
			If oPigFile.LastErr <> "" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(oPigFile.LastErr)
			End If
			LOG.StepName = "New GbMain"
			oPigFile.GbMain = New PigBytes(AbIn)
			If oPigFile.LastErr <> "" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(oPigFile.LastErr)
			End If
			LOG.StepName = "SaveFile"
			oPigFile.SaveFile(False)
			If oPigFile.LastErr <> "" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(oPigFile.LastErr)
			End If
			oPigFile = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	''' <summary>
	''' 从共享内存读取一个 PigBytes 对象
	''' </summary>
	''' <param name="SMName">共享内存名</param>
	''' <param name="SMLen">共享内存长度</param>
	''' <param name="PbOut">输出的PigBytes</param>
	''' <returns></returns>
	Public Function LoadBytesFromShareMem(SMName As String, SMLen As Long, ByRef PbOut As PigBytes) As String
		Dim LOG As New PigStepLog("LoadBytesFromFile")
		Try
			If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
			Dim smMain As New ShareMem
			LOG.StepName = "Init"
			LOG.Ret = smMain.Init(SMName, SMLen)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(SMName)
				Throw New Exception(LOG.Ret)
			End If
			If PbOut IsNot Nothing Then PbOut = Nothing
			LOG.StepName = "New PigBytes"
			PbOut = New PigBytes
			If PbOut.LastErr <> "" Then
				LOG.AddStepNameInf(SMName)
				Throw New Exception(PbOut.LastErr)
			End If
			LOG.StepName = "Read"
			LOG.Ret = smMain.Read(PbOut.Main)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(SMName)
				Throw New Exception(LOG.Ret)
			End If
			smMain = Nothing
			Return "OK"
		Catch ex As Exception
			PbOut = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function




	Private Function mSavePigKeyValueToShareMem(ByRef NewItem As PigKeyValue) As String
		Dim LOG As New PigStepLog("mSavePigKeyValueToShareMem")
		Try
			If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
			Dim strKeyName As String = NewItem.KeyName
			'--------
			LOG.StepName = "Check"
			LOG.Ret = NewItem.Check()
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(strKeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			LOG.StepName = "SaveBytesToShareMem"
			LOG.Ret = Me.SaveBytesToShareMem(strHeadTitle, SM_HEAD_LEN, NewItem.HeadData.Main)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName & "." & strHeadTitle)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			LOG.StepName = "SaveBytesToShareMem"
			LOG.Ret = Me.SaveBytesToShareMem(strBodyTitle, NewItem.BodyLen, NewItem.BodyData.Main)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName & "." & strBodyTitle)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mSavePigKeyValueToFile(ByRef NewItem As PigKeyValue) As String
		Dim LOG As New PigStepLog("mSavePigKeyValueToShareMem")
		Try
			Dim strKeyName As String = NewItem.KeyName
			'--------
			LOG.StepName = "Check"
			LOG.Ret = NewItem.Check()
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(strKeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Dim strFilePath As String
			strFilePath = Me.CacheWorkDir & Me.OsPathSep & strHeadTitle
			LOG.StepName = "SaveBytesToFile"
			LOG.Ret = Me.SaveBytesToFile(strFilePath, NewItem.HeadData.Main)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName & "." & strHeadTitle)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			LOG.StepName = "SaveBytesToFile"
			strFilePath = Me.CacheWorkDir & Me.OsPathSep & strBodyTitle
			LOG.Ret = Me.SaveBytesToFile(strFilePath, NewItem.BodyData.Main)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName & "." & strBodyTitle)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function


	Public Function SavePigKeyValue(KeyName As String, KeyValue As Byte(), ExpTimeSec As Long, Optional IsOverwrite As Boolean = True) As String
		Dim LOG As New PigStepLog("SavePigKeyValue")
		Try
			LOG.StepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, Now.AddSeconds(ExpTimeSec), KeyValue, Me.DefaultSaveType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			LOG.StepName = "mSavePigKeyValue"
			LOG.Ret = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", LOG.StepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(KeyName As String, KeyValue As Byte(), ExpTime As DateTime, Optional IsOverwrite As Boolean = True) As String
		Dim LOG As New PigStepLog("SavePigKeyValue")
		Try
			LOG.StepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, ExpTime, KeyValue, Me.DefaultSaveType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			LOG.StepName = "mSavePigKeyValue"
			LOG.Ret = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", LOG.StepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(KeyName As String, KeyValue As String, ExpTime As DateTime, Optional IsOverwrite As Boolean = True) As String
		Dim LOG As New PigStepLog("SavePigKeyValue")
		Try
			LOG.StepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, ExpTime, KeyValue, Me.DefaultTextType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			LOG.StepName = "mSavePigKeyValue"
			LOG.Ret = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", LOG.StepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(KeyName As String, KeyValue As String, ExpTimeSec As Long, Optional IsOverwrite As Boolean = True) As String
		Dim LOG As New PigStepLog("SavePigKeyValue")
		Try
			LOG.StepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, Now.AddSeconds(ExpTimeSec), KeyValue, Me.DefaultTextType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			LOG.StepName = "mSavePigKeyValue"
			LOG.Ret = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", LOG.StepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True) As String
		Dim strRet As String = ""
		Try
			strRet = Me.mSavePigKeyValue(NewItem, IsOverwrite)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Return strRet
		Catch ex As Exception
			Me.SetSubErrInf("SavePigKeyValue", ex)
			Return strRet
		End Try
	End Function


	Private Function mSavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True) As String
		Dim LOG As New PigStepLog("SavePigKeyValue")
		Try
			Dim strKeyName As String = NewItem.KeyName
			LOG.StepName = "NewItem.Check"
			LOG.Ret = NewItem.Check()
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(strKeyName)
				Throw New Exception(LOG.Ret)
			End If
			Dim pkvOld As PigKeyValue = Nothing
			'获取旧的成员
			Select Case Me.CacheLevel
				Case EnmCacheLevel.ToList
					LOG.StepName = "mGetPigKeyValueByList"
					LOG.Ret = Me.mGetPigKeyValueByList(strKeyName, pkvOld)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(strKeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
					End If
				Case EnmCacheLevel.ToShareMem
					If Me.IsWindows = False Then Throw New Exception("This Function can only be used on windows")
					LOG.StepName = "mGetPigKeyValueByShareMem"
					LOG.Ret = Me.mGetPigKeyValueByShareMem(strKeyName, pkvOld)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(strKeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
					End If
				Case EnmCacheLevel.ToFile
					LOG.StepName = "mGetPigKeyValueByFile"
					LOG.Ret = Me.mGetPigKeyValueByFile(strKeyName, pkvOld)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(strKeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.PigKeyValues.LastErr)
					End If
				Case Else
					LOG.StepName = Me.CacheLevel.ToString
					Throw New Exception("Unsupported CacheLevel")
			End Select
			'确定新增还是更新
			Dim bolIsNew As Boolean = False, bolUpdate As Boolean = False
			If NewItem.Parent Is Nothing Then NewItem.Parent = Me
			If pkvOld Is Nothing Then
				bolIsNew = True
			ElseIf pkvOld.IsMatchAnother(NewItem) = False Then
				If IsOverwrite = False Then
					LOG.StepName = strKeyName
					Throw New Exception("PigKeyValue Exists")
				End If
				bolUpdate = True
			End If

			If bolIsNew = True Then
				msuStatistics.SaveCount += 1
				Select Case Me.CacheLevel
					Case EnmCacheLevel.ToList
						LOG.StepName = "mAddPigKeyValueToList"
						LOG.Ret = Me.mAddPigKeyValueToList(NewItem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Throw New Exception(LOG.Ret)
						End If
						msuStatistics.SaveToListCount += 1
					Case EnmCacheLevel.ToShareMem
						LOG.StepName = "mSavePigKeyValueToShareMem.New"
						LOG.Ret = Me.mSavePigKeyValueToShareMem(NewItem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Throw New Exception(LOG.Ret)
						End If
						msuStatistics.SaveToShareMemCount += 1
					Case EnmCacheLevel.ToFile
						LOG.StepName = "mSavePigKeyValueToFile.New"
						LOG.Ret = Me.mSavePigKeyValueToFile(NewItem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Throw New Exception(LOG.Ret)
						End If
						msuStatistics.SaveToFileCount += 1
				End Select
			ElseIf bolUpdate = True Then
				msuStatistics.SaveCount += 1
				Select Case Me.CacheLevel
					Case EnmCacheLevel.ToList
						LOG.StepName = "CopyToMe.Update.ToList"
						LOG.Ret = pkvOld.CopyToMe(NewItem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Throw New Exception(LOG.Ret)
						End If
						msuStatistics.SaveToListCount += 1
					Case EnmCacheLevel.ToShareMem
						LOG.StepName = "mClearShareMem.Update"
						LOG.Ret = Me.mClearShareMem(pkvOld)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							msuStatistics.RemoveFailCount += 1
						End If
						LOG.StepName = "mSavePigKeyValueToShareMem.Update"
						LOG.Ret = Me.mSavePigKeyValueToShareMem(NewItem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Throw New Exception(LOG.Ret)
						End If
						msuStatistics.SaveToShareMemCount += 1
					Case EnmCacheLevel.ToFile
						If pkvOld.Parent Is Nothing Then pkvOld.Parent = Me
						LOG.StepName = "mRemoveFile.Update"
						LOG.Ret = Me.mRemoveFile(pkvOld.KeyName)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
							msuStatistics.RemoveFailCount += 1
						End If
						If NewItem.Parent Is Nothing Then NewItem.Parent = Me
						LOG.StepName = "mSavePigKeyValueToFile.Update"
						LOG.Ret = Me.mSavePigKeyValueToFile(NewItem)
						If LOG.Ret <> "OK" Then
							LOG.AddStepNameInf(strKeyName)
							Throw New Exception(LOG.Ret)
						End If
						msuStatistics.SaveToFileCount += 1
				End Select
			Else
				pkvOld.LastRefCacheTime = Now
			End If
			Select Case Me.CacheLevel
				Case EnmCacheLevel.ToShareMem, EnmCacheLevel.ToFile
					If Me.CacheLevel = EnmCacheLevel.ToFile Then
						If Me.IsWindows = True Then
							LOG.StepName = "mGetPigKeyValueByShareMem"
							LOG.Ret = Me.mGetPigKeyValueByShareMem(strKeyName, pkvOld)
							If LOG.Ret <> "OK" Then
								LOG.AddStepNameInf(strKeyName)
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.PigKeyValues.LastErr)
							End If
							If pkvOld IsNot Nothing Then
								If pkvOld.Parent Is Nothing Then pkvOld.Parent = Me
								If pkvOld.IsMatchAnother(NewItem) = False Then
									LOG.StepName = "mClearShareMem.ToFile"
									LOG.Ret = Me.mClearShareMem(pkvOld)
									If LOG.Ret <> "OK" Then
										LOG.AddStepNameInf(strKeyName)
										Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
										msuStatistics.RemoveFailCount += 1
									End If
								End If
							End If
						End If
					End If
					If Me.PigKeyValues.IsItemExists(strKeyName) = True Then
						LOG.StepName = "mGetPigKeyValueByList.ToShareMem.ToFile"
						LOG.Ret = Me.mGetPigKeyValueByList(strKeyName, pkvOld)
						If Me.PigKeyValues.LastErr <> "" Then
							LOG.AddStepNameInf(strKeyName)
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.PigKeyValues.LastErr)
						End If
						If pkvOld IsNot Nothing Then
							If pkvOld.Parent Is Nothing Then pkvOld.Parent = Me
							If pkvOld.IsMatchAnother(NewItem) = False Then
								LOG.StepName = "CopyToMe.ToShareMem.ToFile"
								LOG.Ret = pkvOld.CopyToMe(NewItem)
								If LOG.Ret <> "OK" Then
									LOG.AddStepNameInf(strKeyName)
									Throw New Exception(LOG.Ret)
								End If
							End If
						End If
					End If
			End Select
			Return "OK"
		Catch ex As Exception
			msuStatistics.SaveFailCount += 1
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function RemovePigKeyValue(PkvOld As PigKeyValue, CacheLevel As EnmCacheLevel) As String
		Dim LOG As New PigStepLog("RemovePigKeyValue")
		Try
			If PkvOld Is Nothing Then Throw New Exception("PkvOld Is Nothing")
			Dim bolIsToList As Boolean = False, bolIsToShareMem As Boolean = False, bolIsToFile As Boolean = False
			Select Case CacheLevel
				Case EnmCacheLevel.ToList
					bolIsToList = True
				Case EnmCacheLevel.ToShareMem
					bolIsToList = True
					bolIsToShareMem = True
				Case EnmCacheLevel.ToFile
					bolIsToList = True
					bolIsToShareMem = True
					bolIsToFile = True
				Case Else
					LOG.StepName = PkvOld.KeyName
					Throw New Exception("Unsupported CacheLevel")
			End Select
			msuStatistics.RemoveCount += 1
			Dim strErr As String = ""
			If bolIsToFile = True Then
				If Me.mIsCacheFileExists(PkvOld.KeyName) = True Then
					LOG.StepName = "mRemoveFile"
					LOG.Ret = Me.mRemoveFile(PkvOld.KeyName)
					If LOG.Ret <> "OK" Then strErr &= LOG.StepName & ":" & LOG.Ret
				End If
			End If
			If bolIsToShareMem = True Then
				LOG.StepName = "mClearShareMem"
				LOG.Ret = Me.mClearShareMem(PkvOld)
				If LOG.Ret <> "OK" Then strErr &= LOG.StepName & ":" & LOG.Ret
			End If
			If bolIsToList = True Then
				If Me.PigKeyValues.IsItemExists(PkvOld.KeyName) = True Then
					LOG.StepName = "mRemovePigKeyValueFromList"
					LOG.Ret = Me.mRemovePigKeyValueFromList(PkvOld.KeyName)
					If LOG.Ret <> "OK" Then strErr &= LOG.StepName & ":" & LOG.Ret
				End If
			End If
			If strErr <> "" Then
				LOG.AddStepNameInf(PkvOld.KeyName)
				Throw New Exception(strErr)
			End If
			Return "OK"
		Catch ex As Exception
			msuStatistics.RemoveFailCount += 1
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Sub RemoveExpItems()
		Dim LOG As New PigStepLog("RemoveExpItems")
		Try
			Dim intItems As Integer = 0
			Dim astrKeyName(intItems) As String
			LOG.StepName = "For Each"
			For Each oPigKeyValue As PigKeyValue In Me.PigKeyValues
				Dim strKeyName As String = oPigKeyValue.KeyName
				If oPigKeyValue.IsExpired = True Then
					intItems += 1
					ReDim Preserve astrKeyName(intItems)
					astrKeyName(intItems) = strKeyName
				End If
			Next
			If intItems > 0 Then
				For Each pkvAny As PigKeyValue In Me.PigKeyValues
					LOG.StepName = "RemovePigKeyValue"
					LOG.Ret = Me.RemovePigKeyValue(pkvAny, Me.CacheLevel)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf(pkvAny.KeyName)
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
					End If
				Next
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Sub

	Public Function GetStatisticsXml() As String
		Try
			Dim oPigXml As New PigXml(True)
			GetStatisticsXml = ""
			oPigXml.AddEle("PID", Me.fMyPID)
			oPigXml.AddEle("StatisticsTime", Format(Now, "yyyy-MM-dd HH:mm:ss.fff"))
			With msuStatistics
				oPigXml.AddEle("GetCount", .GetCount)
				oPigXml.AddEle("GetFailCount", .GetFailCount)
				'---------
				oPigXml.AddEle("SaveCount", .SaveCount)
				oPigXml.AddEle("SaveFailCount", .SaveFailCount)
				oPigXml.AddEle("SaveToListCount", .SaveToListCount)
				Select Case Me.CacheLevel
					Case EnmCacheLevel.ToShareMem
						oPigXml.AddEle("SaveToShareMemCount", .SaveToShareMemCount)
					Case EnmCacheLevel.ToFile
						oPigXml.AddEle("SaveToShareMemCount", .SaveToShareMemCount)
						oPigXml.AddEle("SaveToFileCount", .SaveToFileCount)
				End Select
				'---------
				oPigXml.AddEle("CacheCount", .CacheCount)
				oPigXml.AddEle("CacheByListCount", .CacheByListCount)
				Select Case Me.CacheLevel
					Case EnmCacheLevel.ToShareMem
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
					Case EnmCacheLevel.ToFile
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
						oPigXml.AddEle("CacheByFileCount", .CacheByFileCount)
				End Select
				'---------
				oPigXml.AddEle("RemoveCount", .RemoveCount)
				oPigXml.AddEle("RemoveFailCount", .RemoveFailCount)
				oPigXml.AddEle("RemoveExpiredListCount", .RemoveExpiredListCount)
				Select Case Me.CacheLevel
					Case EnmCacheLevel.ToShareMem
						oPigXml.AddEle("RemoveExpiredShareMemCount", .RemoveExpiredShareMemCount)
					Case EnmCacheLevel.ToFile
						oPigXml.AddEle("RemoveExpiredShareMemCount", .RemoveExpiredShareMemCount)
						oPigXml.AddEle("RemoveExpiredFileCount", .RemoveExpiredFileCount)
				End Select
			End With
			GetStatisticsXml = oPigXml.MainXmlStr
			oPigXml = Nothing
		Catch ex As Exception
			Me.SetSubErrInf("GetStatisticsXml", ex)
			Return ""
		End Try
	End Function

	Private menmCacheLevel As EnmCacheLevel = EnmCacheLevel.ToList
	Public Property CacheLevel As EnmCacheLevel
		Get
			Return menmCacheLevel
		End Get
		Friend Set(value As EnmCacheLevel)
			menmCacheLevel = value
		End Set
	End Property

	Private mintForceRefCacheTime As Integer = 60
	Public Property ForceRefCacheTime As Integer
		Get
			Return mintForceRefCacheTime
		End Get
		Friend Set(value As Integer)
			mintForceRefCacheTime = value
		End Set
	End Property

	Private mintDefaultSaveType As PigKeyValue.EnmSaveType = PigKeyValue.EnmSaveType.SaveSpace
	Public Property DefaultSaveType As PigKeyValue.EnmSaveType
		Get
			Return mintDefaultSaveType
		End Get
		Friend Set(value As PigKeyValue.EnmSaveType)
			mintDefaultSaveType = value
		End Set
	End Property

	Private mintDefaultTextType As PigText.enmTextType = PigText.enmTextType.UTF8
	Public Property DefaultTextType As PigText.enmTextType
		Get
			Return mintDefaultTextType
		End Get
		Friend Set(value As PigText.enmTextType)
			mintDefaultTextType = value
		End Set
	End Property

	Private mstrCacheWorkDir As String
	Public Property CacheWorkDir As String
		Get
			Return mstrCacheWorkDir
		End Get
		Friend Set(value As String)
			mstrCacheWorkDir = value
		End Set
	End Property
	Private Function mIsBytesMatch(ByRef SrcBytes As Byte(), ByRef MatchBytes As Byte()) As Boolean
		Try
#If NET40_OR_GREATER Then
			Return SrcBytes.SequenceEqual(MatchBytes)
#Else
			Dim i As Long
			If SrcBytes.Length <> MatchBytes.Length Then
				Return False
			Else
				mIsBytesMatch = True
				For i = 0 To SrcBytes.Length - 1
					If SrcBytes(i) <> MatchBytes(i) Then
						mIsBytesMatch = False
						Exit For
					End If
				Next
			End If

#End If
		Catch ex As Exception
			Me.SetSubErrInf("mIsBytesMatch", ex)
			Return False
		End Try
	End Function




	Private Function mGetPigKeyValueFromFile(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Dim LOG As New PigStepLog("mGetPigKeyValueFromFile")
		Try
			If OutPigKeyValue IsNot Nothing Then OutPigKeyValue = Nothing
			LOG.StepName = "New PigKeyValue"
			OutPigKeyValue = New PigKeyValue(KeyName)
			If OutPigKeyValue.LastErr <> "" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(OutPigKeyValue.LastErr)
			End If
			'--------
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(KeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(LOG.Ret)
			End If
			'--------
			Dim pbMain As PigBytes = Nothing
			Dim strFilePath As String
			'--------
			LOG.StepName = "LoadBytesFromFile"
			strFilePath = Me.CacheWorkDir & Me.OsPathSep & strHeadTitle
			LOG.Ret = Me.LoadBytesFromFile(strFilePath, pbMain)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName & "." & strHeadTitle)
				Throw New Exception(LOG.Ret)
			End If
			If pbMain Is Nothing Then
				LOG.AddStepNameInf(KeyName & "." & strHeadTitle)
				Throw New Exception("pbMain Is Nothing")
			End If
			If Me.mIsHeadBytesInit(pbMain.Main) = True Then
				LOG.StepName = "OutPigKeyValue.LoadHead"
				LOG.Ret = OutPigKeyValue.LoadHead(pbMain)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf(KeyName)
					Throw New Exception(LOG.Ret)
				End If
				'--------
				pbMain = Nothing
				strFilePath = Me.CacheWorkDir & Me.OsPathSep & strBodyTitle
				LOG.StepName = "LoadBytesFromFile"
				LOG.Ret = Me.LoadBytesFromFile(strFilePath, pbMain)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf(KeyName & "." & strBodyTitle)
					Throw New Exception(LOG.Ret)
				End If
				If pbMain Is Nothing Then
					LOG.AddStepNameInf(KeyName & "." & strBodyTitle)
					Throw New Exception("pbMain Is Nothing")
				End If
				LOG.StepName = "OutPigKeyValue.LoadBody"
				LOG.Ret = OutPigKeyValue.LoadBody(pbMain.Main)
				If LOG.Ret <> "OK" Then
					LOG.AddStepNameInf(KeyName)
					Throw New Exception(LOG.Ret)
				End If
			End If
			pbMain = Nothing
			'--------
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mRemoveFile(KeyName As String) As String
		Dim LOG As New PigStepLog("mRemoveFile")
		Dim strDelFile As String = ""
		Try
			'--------
			Dim strHeadTitle As String = "", strBodyTitle As String = ""
			LOG.StepName = "GetKeyTitle"
			LOG.Ret = Me.GetKeyTitle(KeyName, strHeadTitle, strBodyTitle)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(KeyName)
				Throw New Exception(LOG.Ret)
			End If
			Dim strDirPath As String = Me.CacheWorkDir & Me.OsPathSep
			strDelFile = strDirPath & strHeadTitle
			If IO.File.Exists(strDelFile) = True Then
				LOG.StepName = "File.Delete"
				IO.File.Delete(strDelFile)
			End If
			strDelFile = strDirPath & strBodyTitle
			If IO.File.Exists(strDelFile) = True Then
				LOG.StepName = "File.Delete"
				IO.File.Delete(strDelFile)
			End If
			Return "OK"
		Catch ex As Exception
			LOG.AddStepNameInf(strDelFile)
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mGetFileLen(FilePath As String, ByRef FileLen As Long) As String
		Dim LOG As New PigStepLog("mGetFileLen")
		Try
			LOG.StepName = "New FileInfo"
			Dim oFileInfo As New FileInfo(FilePath)
			FileLen = oFileInfo.Length
			oFileInfo = Nothing
			Return "OK"
		Catch ex As Exception
			LOG.AddStepNameInf("FilePath")
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function



	''' <summary>
	''' 获取键值文件名
	''' </summary>
	''' <returns></returns>
	Public Function GetKeyTitle(KeyName As String, ByRef HeadTitle As String, ByRef BodyTitle As String) As String
		Dim LOG As New PigStepLog("GetKeyTitle")
		Try
			Dim strFileTitle As String
			Select Case Me.CacheLevel
				Case EnmCacheLevel.ToList
					strFileTitle = KeyName
				Case EnmCacheLevel.ToShareMem
					strFileTitle = "<" & KeyName & "><" & Me.ShareMemRoot & ">"
				Case EnmCacheLevel.ToFile
					strFileTitle = "<" & KeyName & "><" & Me.CacheWorkDir & ">"
				Case Else
					Throw New Exception("Invalid CacheLevel " & Me.CacheLevel)
			End Select
			Dim oPigMD5 As New PigMD5(strFileTitle, PigMD5.enmTextType.UTF8)
			If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
			Dim strPigMD5 As String = oPigMD5.PigMD5
			oPigMD5 = Nothing
			HeadTitle = strPigMD5 & ".h"
			BodyTitle = strPigMD5 & ".b"
			Return "OK"
		Catch ex As Exception
			HeadTitle = ""
			BodyTitle = ""
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

End Class
