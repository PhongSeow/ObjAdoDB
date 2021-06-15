'**********************************
'* Name: ConnSQLSrv
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Connection for SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.9
'* Create Time: 2/5/2021
'* 1.0.2	18/4/2021	Modify OpenOrKeepActive
'* 1.0.3	6/5/2021	Modify CommandTimeout, add IsDBConnReady
'* 1.0.4	16/5/2021	Add SQLSrvDataTypeEnum, Modify OpenOrKeepActive
'* 1.0.5	18/5/2021	Modify ConnStatus,OpenOrKeepActive
'* 1.0.6	12/6/2021	Modify OpenOrKeepActive and add New for Mirror
'* 1.0.8	14/6/2021	Add RefMirrSrvTime,LastRefMirrSrvTime
'* 1.0.9	15/6/2021	Modify OpenOrKeepActive,mIsDBOnline
'**********************************
Public Class ConnSQLSrv
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.9"
	Public Connection As Connection
	Private mcstChkDBStatus As CmdSQLSrvText


	Public Enum SQLSrvDataTypeEnum
		adBigint = 127
		adBinary = 173
		adBit = 104
		adChar = 175
		adDate = 40
		adDatetime = 61
		adDatetime2 = 42
		'adDatetimeoffset = 43
		adDecimal = 106
		adFloat = 62
		'adGeography = 240
		'adGeometry = 240
		'adHierarchyid = 240
		adImage = 34
		adInt = 56
		adMoney = 60
		adNChar = 239
		adNText = 99
		adNumeric = 108
		adNVarchar = 231
		adReal = 59
		adSmallDateTime = 58
		adSmallInt = 52
		adSmallMoney = 122
		adSql_Variant = 98
		adSysname = 231
		adText = 35
		'adTime = 41
		adTimeStamp = 189
		adTinyInt = 48
		adUniqueIdentifier = 36
		adVarBinary = 165
		adVarChar = 167
		'adXml = 241
	End Enum

	Public Enum SQLSrvProviderEnum
		MicrosoftSQLServer = 90
		MicrosoftSQLServer2012NativeClient = 100
	End Enum

	Public Enum RunModeEnum
		Mirror = 10
		StandAlone = 20
	End Enum

	Public Enum ConnStatusEnum
		Unknow = 0
		PrincipalOnline = 10
		MirrorOnline = 20
		Offline = 30
	End Enum

	Private Property mLastConnSQLServer As String

	Private mintRunMode As RunModeEnum
	Public Property RunMode() As RunModeEnum
		Get
			Return mintRunMode
		End Get
		Friend Set(ByVal value As RunModeEnum)
			mintRunMode = value
		End Set
	End Property

	''' <summary>
	''' If Mirror SQL server is not specified, it will run in stand-alone mode.
	''' </summary>
	Private mstrPrincipalSQLServer As String
	Public Property PrincipalSQLServer() As String
		Get
			Return mstrPrincipalSQLServer
		End Get
		Friend Set(ByVal value As String)
			mstrPrincipalSQLServer = value
		End Set
	End Property

	''' <summary>
	''' Time to refresh the mirror database, in seconds.
	''' </summary>
	Private mintRefMirrSrvTime As Integer = 30
	Public Property RefMirrSrvTime() As Integer
		Get
			Return mintRefMirrSrvTime
		End Get
		Set(ByVal value As Integer)
			If value <= 0 Then
				mintRefMirrSrvTime = 30
			Else
				mintRefMirrSrvTime = value
			End If
		End Set
	End Property


	Private Function mIsDBOnline() As Boolean
		Dim strStepName As String = ""
		Try
			If Me.Connection Is Nothing Then Throw New Exception("No connection established")
			If Me.Connection.State <> Connection.ConnStateEnum.adStateOpen Then Throw New Exception("The current connection status is " & Me.Connection.State.ToString)
			If mcstChkDBStatus Is Nothing Then
				strStepName = "New CmdSQLSrvText"
				mcstChkDBStatus = New CmdSQLSrvText("SELECT Convert(varchar(50),DatabasePropertyEx(?,'status')) DBStatus")
				If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
				strStepName = "AddPara(@DBName)"
				mcstChkDBStatus.AddPara("@DBName", SQLSrvDataTypeEnum.adVarChar, 256)
				If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
				strStepName = "Set ActiveConnection"
				mcstChkDBStatus.ActiveConnection = Me.Connection
				If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
			End If
			Dim rsAny As Recordset
			mcstChkDBStatus.ParaValue("@DBName") = Me.CurrDatabase
			strStepName = "Execute"
			rsAny = mcstChkDBStatus.Execute
			If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
			Dim strDBStaus As String = UCase(rsAny.Fields.Item("DBStatus").StrValue)
			rsAny = Nothing
			If strDBStaus = "ONLINE" Then
				Return True
			Else
				Return False
			End If
		Catch ex As Exception
			Me.SetSubErrInf("mIsDBOnline", strStepName, ex)
			Return False
		End Try
	End Function

	''' <summary>
	''' The last time the mirror database was refreshed
	''' </summary>
	Private mdteLastRefMirrSrvTime As DateTime
	Public Property LastRefMirrSrvTime() As DateTime
		Get
			Return mdteLastRefMirrSrvTime
		End Get
		Friend Set(ByVal value As DateTime)
			mdteLastRefMirrSrvTime = value
		End Set
	End Property



	''' <summary>
	''' If Mirror SQL server is specified, it will run in mirror mode and can automatic failover.
	''' </summary>
	Private mstrMirrorSQLServer As String
	Public Property MirrorSQLServer() As String
		Get
			Return mstrMirrorSQLServer
		End Get
		Friend Set(ByVal value As String)
			mstrMirrorSQLServer = value
		End Set
	End Property

	''' <summary>
	''' If running in mirror mode, the current database of the principal server and the mirror server must be the same.
	''' </summary>
	Private mstrCurrDatabase As String
	Public Property CurrDatabase() As String
		Get
			Return mstrCurrDatabase
		End Get
		Friend Set(ByVal value As String)
			mstrCurrDatabase = value
		End Set
	End Property

	''' <summary>
	''' If running in mirror mode, the uid of the principal server and the mirror server must be the same.
	''' </summary>
	Private mstrDBUser As String
	Public Property DBUser() As String
		Get
			Return mstrDBUser
		End Get
		Friend Set(ByVal value As String)
			mstrDBUser = value
		End Set
	End Property

	''' <summary>
	''' If running in mirror mode, the password of the principal server and the mirror server must be the same.
	''' </summary>
	Private mstrDBUserPwd As String
	Public Property DBUserPwd() As String
		Get
			Return mstrDBUserPwd
		End Get
		Friend Set(ByVal value As String)
			mstrDBUserPwd = value
		End Set
	End Property

	''' <summary>
	''' Trusted Connectionst and mirror mode
	''' </summary>
	''' <param name="PrincipalSQLServer">Principal SQLServer hostname or ip</param>
	''' <param name="MirrorSQLServer">Mirror SQLServer hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	''' <param name="Provider">What driver to use</param>
	Public Sub New(PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String, Optional Provider As SQLSrvProviderEnum = SQLSrvProviderEnum.MicrosoftSQLServer)
		MyBase.New(CLS_VERSION)
		Me.MirrorSQLServer = MirrorSQLServer
		Me.mNew(PrincipalSQLServer, CurrDatabase,,, Provider)
	End Sub

	''' <summary>
	''' Trusted Connectionst and stand-alone mode
	''' </summary>
	''' <param name="SQLServer">SQL Server hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	''' <param name="Provider">What driver to use</param>
	Public Sub New(SQLServer As String, CurrDatabase As String, Optional Provider As SQLSrvProviderEnum = SQLSrvProviderEnum.MicrosoftSQLServer)
		MyBase.New(CLS_VERSION)
		Me.mNew(SQLServer, CurrDatabase,,, Provider)
	End Sub

	''' <summary>
	''' Database user password login Connectionst and stand-alone mode
	''' </summary>
	''' <param name="SQLServer">SQL Server hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	''' <param name="DBUser">Database user</param>
	''' <param name="DBUserPwd">Database user password</param>
	''' <param name="Provider">What driver to use</param>
	Public Sub New(SQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String, Optional Provider As SQLSrvProviderEnum = SQLSrvProviderEnum.MicrosoftSQLServer)
		MyBase.New(CLS_VERSION)
		Me.mNew(SQLServer, CurrDatabase, DBUser, DBUserPwd, Provider)
	End Sub

	''' <summary>
	''' Database user password login Connectionst and mirror mode
	''' </summary>
	''' <param name="PrincipalSQLServer">Principal SQLServer hostname or ip</param>
	''' <param name="MirrorSQLServer">Mirror SQLServer hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	''' <param name="DBUser">Database user</param>
	''' <param name="DBUserPwd">Database user password</param>
	''' <param name="Provider">What driver to use</param>
	Public Sub New(PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String, Optional Provider As SQLSrvProviderEnum = SQLSrvProviderEnum.MicrosoftSQLServer)
		MyBase.New(CLS_VERSION)
		Me.MirrorSQLServer = MirrorSQLServer
		Me.mNew(PrincipalSQLServer, CurrDatabase, DBUser, DBUserPwd, Provider)
	End Sub

	Private mbolIsTrustedConnection As Boolean
	Public Property IsTrustedConnection() As Boolean
		Get
			Return mbolIsTrustedConnection
		End Get
		Friend Set(ByVal value As Boolean)
			mbolIsTrustedConnection = value
		End Set
	End Property

	''' <summary>
	''' What driver to use
	''' </summary>
	Private moProvider As SQLSrvProviderEnum
	Public Property Provider() As SQLSrvProviderEnum
		Get
			Return moProvider
		End Get
		Friend Set(ByVal value As SQLSrvProviderEnum)
			moProvider = value
		End Set
	End Property


	Private Sub mNew(PrincipalSQLServer As String, CurrDatabase As String, Optional DBUser As String = "", Optional DBUserPwd As String = "", Optional Provider As SQLSrvProviderEnum = SQLSrvProviderEnum.MicrosoftSQLServer)
		Dim strStepName As String = ""
		Try
			With Me
				.PrincipalSQLServer = PrincipalSQLServer
				.CurrDatabase = CurrDatabase
				If DBUser = "" Then
					.IsTrustedConnection = True
				Else
					.IsTrustedConnection = False
					.DBUser = DBUser
					.DBUserPwd = DBUserPwd
				End If
				If .MirrorSQLServer = "" Then
					.RunMode = RunModeEnum.StandAlone
				Else
					.RunMode = RunModeEnum.Mirror
				End If
				.Provider = Provider
				.ConnectionTimeout = 5
				.CommandTimeout = 60
				Me.Connection = New Connection
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", strStepName, ex)
		End Try
	End Sub

	''' <summary>
	''' Database connection status, including the connection between principal and mirror server.
	''' </summary>
	Private mintConnStatus As ConnStatusEnum = ConnStatusEnum.Unknow
	Public Property ConnStatus() As ConnStatusEnum
		Get
			Return mintConnStatus
		End Get
		Friend Set(ByVal value As ConnStatusEnum)
			mintConnStatus = value
		End Set
	End Property

	''' <summary>
	''' Timeout for executing SQL command
	''' </summary>
	Private mlngCommandTimeout As Long
	Public Property CommandTimeout() As Long
		Get
			Return mlngCommandTimeout
		End Get
		Set(ByVal value As Long)
			mlngCommandTimeout = value
		End Set
	End Property

	''' <summary>
	''' Connection database timeout
	''' </summary>
	Private mlngConnectionTimeout As Long
	Public Property ConnectionTimeout() As Long
		Get
			Return mlngConnectionTimeout
		End Get
		Set(ByVal value As Long)
			mlngConnectionTimeout = value
		End Set
	End Property

	''' <summary>
	''' Open or keep the database connection available
	''' </summary>
	Public Sub OpenOrKeepActive()
		Dim strStepName As String = ""
		Try
			Select Case Me.RunMode
				Case RunModeEnum.StandAlone
					With Me.Connection
						Select Case .State
							Case Connection.ConnStateEnum.adStateClosed
								strStepName = "SetConnSQLServer"
								If Me.IsTrustedConnection = True Then
									.SetConnSQLServer(Me.PrincipalSQLServer, Me.CurrDatabase, Me.Provider)
								Else
									.SetConnSQLServer(Me.PrincipalSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase, Me.Provider)
								End If
								If .LastErr <> "" Then Throw New Exception(.LastErr)
								.ConnectionTimeout = Me.ConnectionTimeout
								.CommandTimeout = Me.CommandTimeout
								strStepName = "Open"
								.Open()
								If .LastErr <> "" Then Throw New Exception(.LastErr)
								Me.ConnStatus = ConnStatusEnum.PrincipalOnline
						End Select
					End With
				Case RunModeEnum.Mirror
					If Me.MirrorSQLServer = "" Then Throw New Exception("Mirror SQLServer is not defined")
					Dim bolIsConn As Boolean = False
					Select Case Me.ConnStatus
						Case ConnStatusEnum.Unknow, ConnStatusEnum.Offline
							If Me.mLastConnSQLServer = "" Or mLastConnSQLServer = Me.MirrorSQLServer Then
								Me.mLastConnSQLServer = Me.PrincipalSQLServer
							Else
								Me.mLastConnSQLServer = Me.MirrorSQLServer
							End If
							bolIsConn = True
						Case Else
							If Math.Abs(DateDiff("s", Me.LastRefMirrSrvTime, Now)) > Me.RefMirrSrvTime Then
								If Me.mIsDBOnline = True Then
									Me.LastRefMirrSrvTime = Now
								Else
									If Me.ConnStatus = ConnStatusEnum.PrincipalOnline Then
										Me.mLastConnSQLServer = Me.MirrorSQLServer
									Else
										Me.mLastConnSQLServer = Me.PrincipalSQLServer
									End If
									bolIsConn = True
								End If
							End If
					End Select
					If bolIsConn = True Then
						If Not Me.Connection Is Nothing Then
							If Me.Connection.State <> Connection.ConnStateEnum.adStateClosed Then
								Me.Connection.Close()
							End If
							Me.Connection = Nothing
						End If
						Me.Connection = New Connection
						With Me.Connection
							strStepName = "SetConnSQLServer2"
							If Me.IsTrustedConnection = True Then
								.SetConnSQLServer(Me.mLastConnSQLServer, Me.CurrDatabase, Me.Provider)
							Else
								.SetConnSQLServer(Me.mLastConnSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase, Me.Provider)
							End If
							If .LastErr <> "" Then Throw New Exception(.LastErr)
							.ConnectionTimeout = Me.ConnectionTimeout
							.CommandTimeout = Me.CommandTimeout
							strStepName = "Open2"
							.Open()
							If .LastErr = "" Then
								If Me.mIsDBOnline = True Then
									If Me.mLastConnSQLServer = Me.PrincipalSQLServer Then
										Me.ConnStatus = ConnStatusEnum.PrincipalOnline
									Else
										Me.ConnStatus = ConnStatusEnum.MirrorOnline
									End If
									Me.LastRefMirrSrvTime = Now
								End If
								bolIsConn = False
							End If
						End With
						If bolIsConn = True Then
							If Me.mLastConnSQLServer = "" Or mLastConnSQLServer = Me.MirrorSQLServer Then
								Me.mLastConnSQLServer = Me.PrincipalSQLServer
							Else
								Me.mLastConnSQLServer = Me.MirrorSQLServer
							End If
							With Me.Connection
								strStepName = "SetConnSQLServer3"
								If Me.IsTrustedConnection = True Then
									.SetConnSQLServer(Me.mLastConnSQLServer, Me.CurrDatabase, Me.Provider)
								Else
									.SetConnSQLServer(Me.mLastConnSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase, Me.Provider)
								End If
								If .LastErr <> "" Then Throw New Exception(.LastErr)
								.ConnectionTimeout = Me.ConnectionTimeout
								.CommandTimeout = Me.CommandTimeout
								strStepName = "Open3"
								.Open()
								If .LastErr = "" Then
									If Me.mIsDBOnline = True Then
										If Me.mLastConnSQLServer = Me.PrincipalSQLServer Then
											Me.ConnStatus = ConnStatusEnum.PrincipalOnline
										Else
											Me.ConnStatus = ConnStatusEnum.MirrorOnline
										End If
										Me.LastRefMirrSrvTime = Now
									Else
										Me.ConnStatus = ConnStatusEnum.Offline
									End If
								Else
									Me.ConnStatus = ConnStatusEnum.Offline
								End If
							End With
						End If
					End If
				Case Else
					Throw New Exception("Unknow run mode")
			End Select
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("OpenOrKeepActive", strStepName, ex)
			Me.ConnStatus = ConnStatusEnum.Unknow
		End Try
	End Sub

	Public ReadOnly Property IsDBConnReady() As Boolean
		Get
			Try
				Select Case Me.ConnStatus
					Case ConnStatusEnum.PrincipalOnline, ConnStatusEnum.MirrorOnline
						Return True
					Case Else
						Return False
				End Select
			Catch ex As Exception
				Me.SetSubErrInf("IsDBConnReady", ex)
				Return False
			End Try
		End Get
	End Property

End Class
