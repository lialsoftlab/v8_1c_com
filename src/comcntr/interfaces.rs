#![allow(non_camel_case_types, non_snake_case, unused)]

use std::any::Any;
use std::convert::{TryFrom, TryInto};
use std::error::Error;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{PUINT, UINT};
use winapi::shared::ntdef::{INT, PULONG, ULONG};
use winapi::shared::wtypes::{BSTR, DATE, VARIANT_BOOL};
use winapi::um::oaidl::{
    IDispatch, IDispatchVtbl, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH, LPVARIANT, SAFEARRAY, VARIANT,
};
use winapi::um::oleauto::{SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl};
use winapi::um::winnt::{HRESULT, LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::RIDL;

// 1C ComConnector (comcntr.dll) class
RIDL! {#[uuid(0x181E893D, 0x73A4, 0x4722, 0xB6, 0x1D, 0xD6, 0x04, 0xB3, 0xD6, 0x7D, 0x47)]
class V8COMConnectorClass;
}
pub type LPV8COMCONNECTORCLASS = *mut V8COMConnectorClass;

// 1C COM Connector (comcntr.dll) interfaces
RIDL! {#[uuid(0x2ff2245e, 0xc604, 0x45bd, 0xac, 0x16, 0x19, 0xb1, 0xf6, 0x4b, 0xd9, 0xa4)]
interface IV8COMConnector3(IV8COMConnector3Vtbl): IV8COMConnector2(IV8COMConnector2Vtbl) {
    fn get_MaxConnections(
        pVal: *mut u32,
    ) -> HRESULT,
    fn put_MaxConnections(
        pVal: u32,
    ) -> HRESULT,
    fn ConnectWorkingProcess(
        serverName: BSTR,
        conn: *mut *mut IWorkingProcessConnection,
    ) -> HRESULT,
    fn ConnectAgent(
        serverName: BSTR,
        conn: *mut *mut IServerAgentConnection,
    ) -> HRESULT,
    fn get_RAgentPortDefault(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_RMngrPortDefault(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_LowBoundDefault(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_HighBoundDefault(
        pVal: *mut INT,
    ) -> HRESULT,
}}
pub type LPV8COMCONNECTOR3 = *mut IV8COMConnector3;

RIDL! {#[uuid(0x687cb41e, 0x3fbc, 0x4096, 0x9b, 0xaa, 0x90, 0x65, 0xf2, 0x54, 0x6d, 0x8f)]
interface IV8COMConnector2(IV8COMConnector2Vtbl): IV8COMConnector(IV8COMConnectorVtbl) {
    fn get_PoolCapacity(
        pVal: *mut u32,
    ) -> HRESULT,
    fn put_PoolCapacity(
        pVal: u32,
    ) -> HRESULT,
    fn get_PoolTimeout(
        pVal: *mut u32,
    ) -> HRESULT,
    fn put_PoolTimeout(
        pVal: u32,
    ) -> HRESULT,
}}
type LPV8COMCONNECTOR2 = *mut IV8COMConnector2;

RIDL! {#[uuid(0xba4e52bd, 0xdcb2, 0x4bf7, 0xbb, 0x29, 0x84, 0xc1, 0xca, 0x45, 0x6a, 0x8f)]
interface IV8COMConnector(IV8COMConnectorVtbl): IDispatch(IDispatchVtbl) {
    fn Connect(
        connectString: BSTR,
        conn: *mut LPDISPATCH,
    ) -> HRESULT,
}}
pub type LPV8COMCONNECTOR = *mut IV8COMConnector;

RIDL! {#[uuid(0xf097a4b8, 0x28db, 0x4162, 0x89, 0x04, 0x77, 0x2d, 0x6d, 0x8b, 0xcc, 0x76)]
interface IWorkingProcessConnection(IWorkingProcessConnectionVtbl): IDispatch(IDispatchVtbl) {
    fn AddAuthentication(
        userName: BSTR,
        userPassword: BSTR,
    ) -> HRESULT,
    fn GetInfoBases(
        infoBases: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetInfoBaseConnections(
        infoBase: *const IInfoBaseInfo,
        connections: *mut SAFEARRAY,
    ) -> HRESULT,
    fn CreateInfoBaseInfo(
        infoBase: *mut *mut IInfoBaseInfo,
    ) -> HRESULT,
    fn CreateInfoBase(
        infoBase: *const IInfoBaseInfo,
        mode: INT,
    ) -> HRESULT,
    fn DropInfoBase(
        infoBase: *const IInfoBaseInfo,
        mode: INT,
    ) -> HRESULT,
    fn Disconnect(
        connection: *const IInfoBaseConnectionInfo,
    ) -> HRESULT,
    fn Connect(
        infoBase: *const IInfoBaseInfo,
        userName: BSTR,
        userPassword: BSTR,
        conn: *mut LPDISPATCH,
    ) -> HRESULT,
    fn AuthenticateAdmin(
        srvrUserName: BSTR,
        srvrUserPassword: BSTR,
    ) -> HRESULT,
    fn UpdateInfoBase(
        infoBase: *const IInfoBaseInfo,
    ) -> HRESULT,
}}
pub type LPWORKINGPROCESSCONNECTION = *mut IWorkingProcessConnection;

RIDL! {#[uuid(0x94ffc9f2, 0x286c, 0x480c, 0xbb, 0x80, 0xa2, 0x0d, 0x8e, 0x8e, 0x14, 0x64)]
interface IInfoBaseInfo(IInfoBaseInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DBMS(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DBMS(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DBServerName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DBServerName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DBName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DBName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DBUser(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DBUser(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DBPassword(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DBPassword(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DateOffset(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_DateOffset(
        pVal: INT,
    ) -> HRESULT,
    fn get_Locale(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Locale(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_SecurityLevel(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_SecurityLevel(
        pVal: INT,
    ) -> HRESULT,
    fn get_ConnectDenied(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_ConnectDenied(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_DeniedFrom(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn put_DeniedFrom(
        pVal: DATE,
    ) -> HRESULT,
    fn get_DeniedTo(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn put_DeniedTo(
        pVal: DATE,
    ) -> HRESULT,
    fn get_DeniedMessage(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DeniedMessage(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_PermissionCode(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_PermissionCode(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_DeniedParameter(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_DeniedParameter(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_ScheduledJobsDenied(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_ScheduledJobsDenied(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_SessionsDenied(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_SessionsDenied(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_LicenseDistributionAllowed(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_LicenseDistributionAllowed(
        pVal: INT,
    ) -> HRESULT,
    fn get_ExternalSessionManagerConnectionString(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ExternalSessionManagerConnectionString(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_ExternalSessionManagerRequired(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_ExternalSessionManagerRequired(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_SecurityProfileName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_SecurityProfileName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_SafeModeSecurityProfileName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_SafeModeSecurityProfileName(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPINFOBASEINFO = *mut IInfoBaseInfo;

RIDL! {#[uuid(0x7df710d1, 0xba3f, 0x4714, 0x8d, 0xdc, 0x63, 0x4c, 0x3c, 0x4b, 0xb1, 0x38)]
interface IInfoBaseConnectionInfo(IInfoBaseConnectionInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_userName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_HostName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_AppID(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_ConnID(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_ConnectedAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_IBConnMode(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_DBConnMode(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_dbProcInfo(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_dbProcTookAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_dbProcTook(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationAllDBMS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationLast5MinDBMS(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_durationCurrentDBMS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationAll(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_durationCurrent(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_callsAll(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_callsLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_bytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_bytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_blockedByDBMS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_dbmsBytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_dbmsBytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_MemoryCurrent(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_MemoryLast5Min(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_MemoryAll(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_InBytesCurrent(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_InBytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_InBytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_OutBytesCurrent(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_OutBytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_OutBytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
}}
pub type LPINFOBASECONNECTIONINFO = *mut IInfoBaseConnectionInfo;

RIDL! {#[uuid(0x0433d6e5, 0xc99a, 0x4fbc, 0xaa, 0xa6, 0x7b, 0x20, 0xad, 0xd1, 0x34, 0xd0)]
interface IServerAgentConnection(IServerAgentConnectionVtbl): IDispatch(IDispatchVtbl) {
    fn get_ConnectionString(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn GetClusters(
        registries: *mut SAFEARRAY,
    ) -> HRESULT,
    fn CreateClusterInfo(
        registry: *mut *mut IClusterInfo,
    ) -> HRESULT,
    fn RegCluster(
        registry: *const IClusterInfo,
    ) -> HRESULT,
    fn UnregCluster(
        registry: *const IClusterInfo,
    ) -> HRESULT,
    fn GetWorkingProcesses(
        registry: *const IClusterInfo,
        processes: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetInfoBases(
        registry: *const IClusterInfo,
        infoBases: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetConnections(
        registry: *const IClusterInfo,
        connections: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetInfoBaseConnections(
        registry: *const IClusterInfo,
        infoBase: *const IInfoBaseShort,
        connections: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetLocks(
        registry: *const IClusterInfo,
        locks: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetInfoBaseLocks(
        registry: *const IClusterInfo,
        infoBase: *const IInfoBaseShort,
        locks: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetConnectionLocks(
        registry: *const IClusterInfo,
        client: *const IConnectionShort,
        locks: *mut SAFEARRAY,
    ) -> HRESULT,
    fn Authenticate(
        registry: *const IClusterInfo,
        userName: BSTR,
        userPswd: BSTR,
    ) -> HRESULT,
    fn GetClusterAdmins(
        registry: *const IClusterInfo,
        users: *mut SAFEARRAY,
    ) -> HRESULT,
    fn CreateClusterAdminInfo(
        user: *mut *mut IRegUserInfo,
    ) -> HRESULT,
    fn RegClusterAdmin(
        registry: *const IClusterInfo,
        user: *const IRegUserInfo,
    ) -> HRESULT,
    fn UnregClusterAdmin(
        registry: *const IClusterInfo,
        userName: BSTR,
    ) -> HRESULT,
    fn GetWorkingServers(
        registry: *const IClusterInfo,
        servers: *mut SAFEARRAY,
    ) -> HRESULT,
    fn CreateWorkingServerInfo(
        server: *mut *mut IWorkingServerInfo,
    ) -> HRESULT,
    fn RegWorkingServer(
        registry: *const IClusterInfo,
        server: *const IWorkingServerInfo,
    ) -> HRESULT,
    fn UnregWorkingServer(
        registry: *const IClusterInfo,
        server: *const IWorkingServerInfo,
    ) -> HRESULT,
    fn GetServerWorkingProcesses(
        registry: *const IClusterInfo,
        server: *const IWorkingServerInfo,
        processes: *mut SAFEARRAY,
    ) -> HRESULT,
    fn UpdateWorkingServer(
        registry: *const IClusterInfo,
        server: *const IWorkingServerInfo,
    ) -> HRESULT,
    fn SetClusterMultiProcess(
        registry: *const IClusterInfo,
        MultiProcess: VARIANT_BOOL,
    ) -> HRESULT,
    fn AuthenticateAgent(
        userName: BSTR,
        userPswd: BSTR,
    ) -> HRESULT,
    fn GetAgentAdmins(
        users: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegAgentAdmin(
        user: *const IRegUserInfo,
    ) -> HRESULT,
    fn UnregAgentAdmin(
        userName: BSTR,
    ) -> HRESULT,
    fn SetClusterSecurityLevel(
        registry: *const IClusterInfo,
        secLevel: INT,
    ) -> HRESULT,
    fn SetClusterDescription(
        registry: *const IClusterInfo,
        Descr: BSTR,
    ) -> HRESULT,
    fn UpdateInfoBase(
        registry: *const IClusterInfo,
        infoBase: *const IInfoBaseShort,
    ) -> HRESULT,
    fn SetClusterRecycling(
        registry: *const IClusterInfo,
        LifeTimeLimit: INT,
        ExpirationTimeout: INT,
    ) -> HRESULT,
    fn GetClusterManagers(
        registry: *const IClusterInfo,
        managers: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetClusterServices(
        registry: *const IClusterInfo,
        services: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetSessions(
        registry: *const IClusterInfo,
        sessions: *mut SAFEARRAY,
    ) -> HRESULT,
    fn GetInfoBaseSessions(
        registry: *const IClusterInfo,
        infoBase: *const IInfoBaseShort,
        sessions: *mut SAFEARRAY,
    ) -> HRESULT,
    fn TerminateSession(
        registry: *const IClusterInfo,
        session: *const ISessionInfo,
    ) -> HRESULT,
    fn SetClusterRecyclingByMemory(
        registry: *const IClusterInfo,
        MaxMemorySize: INT,
        MaxMemoryTimeLimit: INT,
    ) -> HRESULT,
    fn SetClusterRecyclingByTime(
        registry: *const IClusterInfo,
        LifeTimeLimit: INT,
    ) -> HRESULT,
    fn SetClusterRecyclingExpirationTimeout(
        registry: *const IClusterInfo,
        ExpirationTimeout: INT,
    ) -> HRESULT,
    fn SetClusterRecyclingErrorsCountThreshold(
        registry: *const IClusterInfo,
        errorsCountThreshold: INT,
    ) -> HRESULT,
    fn SetClusterRecyclingKillProblemProcesses(
        registry: *const IClusterInfo,
        killProblemProcesses: VARIANT_BOOL,
    ) -> HRESULT,
    fn GetSessionLocks(
        registry: *const IClusterInfo,
        session: *const ISessionInfo,
        locks: *mut SAFEARRAY,
    ) -> HRESULT,
    fn ApplyAssignmentRules(
        cluster: *const IClusterInfo,
        full: INT,
    ) -> HRESULT,
    fn RegAssignmentRule(
        cluster: *const IClusterInfo,
        workingServer: *const IWorkingServerInfo,
        AssgnRuleInfo: *const IAssignmentRuleInfo,
        position: UINT,
    ) -> HRESULT,
    fn UnregAssignmentRule(
        cluster: *const IClusterInfo,
        workingServer: *const IWorkingServerInfo,
        assignmentRule: *const IAssignmentRuleInfo,
    ) -> HRESULT,
    fn GetAssignmentRules(
        cluster: *const IClusterInfo,
        workingServer: *const IWorkingServerInfo,
        assignmentRules: *mut SAFEARRAY,
    ) -> HRESULT,
    fn CreateAssignmentRule(
        AssgnRuleInfo: *mut *mut IAssignmentRuleInfo,
    ) -> HRESULT,
    fn GetSecurityProfiles(
        cluster: *const IClusterInfo,
        securityProfiles: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfile(
        cluster: *const IClusterInfo,
        securityProfile: *const ISecurityProfile,
    ) -> HRESULT,
    fn UnregSecurityProfile(
        cluster: *const IClusterInfo,
        spName: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfile(
        hint: *mut *mut ISecurityProfile,
    ) -> HRESULT,
    fn GetSecurityProfileVirtualDirectories(
        cluster: *const IClusterInfo,
        spName: BSTR,
        virtualDirectories: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfileVirtualDirectory(
        cluster: *const IClusterInfo,
        spName: BSTR,
        virtualDirectory: *const ISecurityProfileVirtualDirectory,
    ) -> HRESULT,
    fn UnregSecurityProfileVirtualDirectory(
        cluster: *const IClusterInfo,
        spName: BSTR,
        Name: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfileVirtualDirectory(
        virtualDirectory: *mut *mut ISecurityProfileVirtualDirectory,
    ) -> HRESULT,
    fn GetSecurityProfileCOMClasses(
        cluster: *const IClusterInfo,
        spName: BSTR,
        comClasses: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfileCOMClass(
        cluster: *const IClusterInfo,
        spName: BSTR,
        comClass: *const ISecurityProfileCOMClass,
    ) -> HRESULT,
    fn UnregSecurityProfileCOMClass(
        cluster: *const IClusterInfo,
        spName: BSTR,
        Name: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfileCOMClass(
        comClass: *mut *mut ISecurityProfileCOMClass,
    ) -> HRESULT,
    fn GetSecurityProfileAddIns(
        cluster: *const IClusterInfo,
        spName: BSTR,
        addIns: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfileAddIn(
        cluster: *const IClusterInfo,
        spName: BSTR,
        addIn: *const ISecurityProfileAddIn,
    ) -> HRESULT,
    fn UnregSecurityProfileAddIn(
        cluster: *const IClusterInfo,
        spName: BSTR,
        Name: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfileAddIn(
        addIn: *mut *mut ISecurityProfileAddIn,
    ) -> HRESULT,
    fn GetSecurityProfileUnSafeExternalModules(
        cluster: *const IClusterInfo,
        spName: BSTR,
        modules: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfileUnSafeExternalModule(
        cluster: *const IClusterInfo,
        spName: BSTR,
        module: *const ISecurityProfileExternalModule,
    ) -> HRESULT,
    fn UnregSecurityProfileUnSafeExternalModule(
        cluster: *const IClusterInfo,
        spName: BSTR,
        Name: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfileUnSafeExternalModule(
        module: *mut *mut ISecurityProfileExternalModule,
    ) -> HRESULT,
    fn GetSecurityProfileApplications(
        cluster: *const IClusterInfo,
        spName: BSTR,
        appls: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfileApplication(
        cluster: *const IClusterInfo,
        spName: BSTR,
        appl: *const ISecurityProfileApplication,
    ) -> HRESULT,
    fn UnregSecurityProfileApplication(
        cluster: *const IClusterInfo,
        spName: BSTR,
        Name: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfileApplication(
        appl: *mut *mut ISecurityProfileApplication,
    ) -> HRESULT,
    fn GetSecurityProfileInternetResources(
        cluster: *const IClusterInfo,
        spName: BSTR,
        appls: *mut SAFEARRAY,
    ) -> HRESULT,
    fn RegSecurityProfileInternetResource(
        cluster: *const IClusterInfo,
        spName: BSTR,
        ir: *const ISecurityProfileInternetResource,
    ) -> HRESULT,
    fn UnregSecurityProfileInternetResource(
        cluster: *const IClusterInfo,
        spName: BSTR,
        Name: BSTR,
    ) -> HRESULT,
    fn CreateSecurityProfileInternetResource(
        ir: *mut *mut ISecurityProfileInternetResource,
    ) -> HRESULT,
}}
pub type LPSERVERAGENTCONNECTION = *mut IServerAgentConnection;

RIDL! {#[uuid(0x2a0dc852, 0x5ab2, 0x4ddf, 0xa3, 0x73, 0x6d, 0x7c, 0x8b, 0xcc, 0x65, 0x35)]
interface IClusterInfo(IClusterInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_ClusterName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ClusterName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_HostName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_HostName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_MainPort(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_MainPort(
        pVal: INT,
    ) -> HRESULT,
    fn get_MultiProcess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_MultiProcess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_SecurityLevel(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_SecurityLevel(
        pVal: INT,
    ) -> HRESULT,
    fn get_LifeTimeLimit(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_LifeTimeLimit(
        pVal: INT,
    ) -> HRESULT,
    fn get_ExpirationTimeout(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_ExpirationTimeout(
        pVal: INT,
    ) -> HRESULT,
    fn get_MaxMemorySize(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_MaxMemorySize(
        pVal: INT,
    ) -> HRESULT,
    fn get_MaxMemoryTimeLimit(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_MaxMemoryTimeLimit(
        pVal: INT,
    ) -> HRESULT,
    fn get_SessionFaultToleranceLevel(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_SessionFaultToleranceLevel(
        pVal: INT,
    ) -> HRESULT,
    fn get_LoadBalancingMode(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_LoadBalancingMode(
        pVal: INT,
    ) -> HRESULT,
}}
pub type LPCLUSTERINFO = *mut IClusterInfo;

RIDL! {#[uuid(0x358d4db6, 0x0771, 0x465c, 0xa8, 0xc0, 0x74, 0x3d, 0x07, 0x29, 0xc2, 0x5d)]
interface IInfoBaseShort(IInfoBaseShortVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPINFOBASESHORT = *mut IInfoBaseShort;

RIDL! {#[uuid(0x43ae2f7e, 0xa98f, 0x4a46, 0xa8, 0x68, 0xd3, 0xe1, 0x93, 0x11, 0x6a, 0x65)]
interface IConnectionShort(IConnectionShortVtbl): IDispatch(IDispatchVtbl) {
    fn get_infoBase(
        infoBase: *mut *mut IInfoBaseShort,
    ) -> HRESULT,
    fn get_process(
        process: *mut *mut IWorkingProcessInfo,
    ) -> HRESULT,
    fn get_Host(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_Application(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_ConnID(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_ConnectedAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_blockedByLM(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_SessionID(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_blockedByLS(
        pVal: *mut UINT,
    ) -> HRESULT,
}}
pub type LPCONNECTIONSHORT = *mut IConnectionShort;

RIDL! {#[uuid(0x719a6f6a, 0x0b91, 0x4d55, 0xb5, 0x7a, 0x67, 0xc8, 0xe4, 0xd6, 0xf7, 0x00)]
interface IWorkingProcessInfo(IWorkingProcessInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_HostName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_HostName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_MainPort(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_Enable(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_Enable(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_Running(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_connections(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_StartedAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_AvgCallTime(
        pVal: *mut f64,
    ) -> HRESULT,
    fn get_AvgServerCallTime(
        pVal: *mut f64,
    ) -> HRESULT,
    fn get_AvgDBCallTime(
        pVal: *mut f64,
    ) -> HRESULT,
    fn get_AvgBackCallTime(
        pVal: *mut f64,
    ) -> HRESULT,
    fn get_AvgLockCallTime(
        pVal: *mut f64,
    ) -> HRESULT,
    fn get_SelectionSize(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_AvgThreads(
        pVal: *mut f64,
    ) -> HRESULT,
    fn get_Capacity(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_Capacity(
        pVal: INT,
    ) -> HRESULT,
    fn get_MemorySize(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_MemoryExcessTime(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_AvailablePerfomance(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_PID(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_Use(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_Use(
        pVal: INT,
    ) -> HRESULT,
    fn get_IsEnable(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn get_License(
        process: *mut VARIANT,
    ) -> HRESULT,
}}
pub type LPWORKINGPROCESSINFO = *mut IWorkingProcessInfo;

RIDL! {#[uuid(0x31d10916, 0xda55, 0x4f7d, 0xb9, 0x34, 0x05, 0x41, 0x05, 0x3e, 0x6c, 0x52)]
interface IRegUserInfo(IRegUserInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
    fn put_Password(
        value: BSTR,
    ) -> HRESULT,
    fn get_SysUserName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_SysUserName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_PasswordAuthAllowed(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_PasswordAuthAllowed(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_SysAuthAllowed(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_SysAuthAllowed(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
}}
pub type LPREGUSERINFO = *mut IRegUserInfo;

RIDL! {#[uuid(0x234b5c3f, 0x4b51, 0x49bc, 0x8d, 0x95, 0xe1, 0xfa, 0x19, 0x24, 0x04, 0xac)]
interface IWorkingServerInfo(IWorkingServerInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_HostName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_HostName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_MainPort(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_MainPort(
        pVal: INT,
    ) -> HRESULT,
    fn GetPortRanges(
        ranges: *mut SAFEARRAY,
    ) -> HRESULT,
    fn CreatePortRange(
        range: *mut *mut IPortRangeInfo,
    ) -> HRESULT,
    fn InsertPortRange(
        range: *const IPortRangeInfo,
    ) -> HRESULT,
    fn ErasePortRange(
        range: *const IPortRangeInfo,
    ) -> HRESULT,
    fn get_MainServer(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_MainServer(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_InfoBasesPerWorkingProcessLimit(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn put_InfoBasesPerWorkingProcessLimit(
        pVal: UINT,
    ) -> HRESULT,
    fn put_WorkingProcessMemoryLimit(
        pVal: i64,
    ) -> HRESULT,
    fn get_WorkingProcessMemoryLimit(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_ConnectionsPerWorkingProcessLimit(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn put_ConnectionsPerWorkingProcessLimit(
        pVal: UINT,
    ) -> HRESULT,
    fn get_ClusterMainPort(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_ClusterMainPort(
        pVal: INT,
    ) -> HRESULT,
    fn get_DedicatedManagers(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_DedicatedManagers(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_SafeWorkingProcessesMemoryLimit(
        pVal: *mut i64,
    ) -> HRESULT,
    fn put_SafeWorkingProcessesMemoryLimit(
        pVal: i64,
    ) -> HRESULT,
    fn get_SafeCallMemoryLimit(
        pVal: *mut i64,
    ) -> HRESULT,
    fn put_SafeCallMemoryLimit(
        pVal: i64,
    ) -> HRESULT,
}}
pub type LPWORKINGSERVERINFO = *mut IWorkingServerInfo;

RIDL! {#[uuid(0xcbb703d2, 0xd3c7, 0x40e1, 0x9f, 0xac, 0x7f, 0x86, 0x98, 0x95, 0xa7, 0x2e)]
interface IPortRangeInfo(IPortRangeInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_LowBound(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_LowBound(
        pVal: INT,
    ) -> HRESULT,
    fn get_HighBound(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_HighBound(
        pVal: INT,
    ) -> HRESULT,
}}
pub type LPPORTRANGEINFO = *mut IPortRangeInfo;

RIDL! {#[uuid(0xda4a0529, 0x1b5a, 0x4b25, 0x92, 0x30, 0xb9, 0xb2, 0x2d, 0x78, 0xc4, 0x83)]
interface ISessionInfo(ISessionInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_infoBase(
        infoBase: *mut *mut IInfoBaseShort,
    ) -> HRESULT,
    fn get_SessionID(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_StartedAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_LastActiveAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_Host(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_AppID(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_userName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_Locale(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_process(
        process: *mut VARIANT,
    ) -> HRESULT,
    fn get_connection(
        pVal: *mut VARIANT,
    ) -> HRESULT,
    fn get_dbProcInfo(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_dbProcTookAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_dbProcTook(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_blockedByDBMS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_blockedByLS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationCurrentDBMS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationLast5MinDBMS(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_durationAllDBMS(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_dbmsBytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_dbmsBytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_durationCurrent(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_durationLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_durationAll(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_callsLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_callsAll(
        pVal: *mut UINT,
    ) -> HRESULT,
    fn get_bytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_bytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_MemoryCurrent(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_MemoryLast5Min(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_MemoryAll(
        pVal: *mut i64,
    ) -> HRESULT,
    fn get_InBytesCurrent(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_InBytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_InBytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_OutBytesCurrent(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_OutBytesLast5Min(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_OutBytesAll(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_License(
        process: *mut VARIANT,
    ) -> HRESULT,
    fn get_Hibernate(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn get_PassiveSessionHibernateTime(
        pVal: *mut u64,
    ) -> HRESULT,
    fn get_HibernateSessionTerminateTime(
        pVal: *mut u64,
    ) -> HRESULT,
}}
pub type LPSESSIONINFO = *mut ISessionInfo;

RIDL! {#[uuid(0xdf99d0e1, 0x0ad6, 0x4dde, 0xa0, 0x52, 0x37, 0x86, 0x53, 0x12, 0x8e, 0xce)]
interface IAssignmentRuleInfo(IAssignmentRuleInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_ObjectType(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ObjectType(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_InfoBaseName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_InfoBaseName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_RuleType(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_RuleType(
        pVal: INT,
    ) -> HRESULT,
    fn get_ApplicationExt(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ApplicationExt(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Priority(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_Priority(
        pVal: INT,
    ) -> HRESULT,
}}
pub type LPASSIGNMENTRULEINFO = *mut IAssignmentRuleInfo;

RIDL! {#[uuid(0x62ca5bad, 0x94bf, 0x4e8f, 0x88, 0x8b, 0xf8, 0x25, 0xf0, 0x10, 0xa7, 0x95)]
interface ISecurityProfile(ISecurityProfileVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_FileSystemFullAccess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_FileSystemFullAccess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_COMFullAccess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_COMFullAccess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_AddInFullAccess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_AddInFullAccess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_ExternalAppFullAccess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_ExternalAppFullAccess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_InternetFullAccess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_InternetFullAccess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_SafeModeProfile(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_SafeModeProfile(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_PrivilegedModeInSafeModeAllowed(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_PrivilegedModeInSafeModeAllowed(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_UnSafeExternalModuleFullAccess(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_UnSafeExternalModuleFullAccess(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_CryptographyAllowed(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_CryptographyAllowed(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_RightExtension(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_RightExtension(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_RightExtensionDefinitionRoles(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_RightExtensionDefinitionRoles(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILE = *mut ISecurityProfile;

RIDL! {#[uuid(0xd982e86d, 0xf4d1, 0x45fa, 0x80, 0x19, 0x25, 0x6e, 0x31, 0x38, 0x5a, 0x5e)]
interface ISecurityProfileVirtualDirectory(ISecurityProfileVirtualDirectoryVtbl): IDispatch(IDispatchVtbl) {
    fn get_Alias(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Alias(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_PhysicalPath(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_PhysicalPath(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_AllowedRead(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_AllowedRead(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
    fn get_AllowedWrite(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn put_AllowedWrite(
        pVal: VARIANT_BOOL,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILEVIRTUALDIRECTORY = *mut ISecurityProfileVirtualDirectory;

RIDL! {#[uuid(0xce320f0f, 0x7181, 0x4c21, 0x94, 0x07, 0x7a, 0x23, 0xfa, 0x72, 0x00, 0x41)]
interface ISecurityProfileCOMClass(ISecurityProfileCOMClassVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_FileName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_FileName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_ObjectUUID(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ObjectUUID(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_ComputerName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ComputerName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILECOMCLASS = *mut ISecurityProfileCOMClass;

RIDL! {#[uuid(0x01e92b89, 0xa348, 0x4aa2, 0x9b, 0x71, 0x53, 0xef, 0xf4, 0x3d, 0x28, 0xf7)]
interface ISecurityProfileAddIn(ISecurityProfileAddInVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_AddInHash(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_AddInHash(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILEADDIN = *mut ISecurityProfileAddIn;

RIDL! {#[uuid(0x594f531e, 0x5767, 0x45ba, 0x87, 0x85, 0xc0, 0xbb, 0x5d, 0x2a, 0x45, 0x7e)]
interface ISecurityProfileExternalModule(ISecurityProfileExternalModuleVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_ExternalModuleHash(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_ExternalModuleHash(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILEEXTERNALMODULE = *mut ISecurityProfileExternalModule;

RIDL! {#[uuid(0x2629138f, 0x448d, 0x4cf9, 0xb6, 0xd3, 0x94, 0x23, 0x82, 0x6f, 0xc1, 0xd7)]
interface ISecurityProfileApplication(ISecurityProfileApplicationVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_CommandMask(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_CommandMask(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILEAPPLICATION = *mut ISecurityProfileApplication;

RIDL! {#[uuid(0xc7c0e781, 0x11f3, 0x4d69, 0x9a, 0x93, 0xce, 0xdb, 0x43, 0x9b, 0x9b, 0xd0)]
interface ISecurityProfileInternetResource(ISecurityProfileInternetResourceVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Name(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Address(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Address(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Port(
        pVal: *mut INT,
    ) -> HRESULT,
    fn put_Port(
        pVal: INT,
    ) -> HRESULT,
    fn get_Protocol(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Protocol(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
}}
pub type LPSECURITYPROFILEINTERNETRESOURCE = *mut ISecurityProfileInternetResource;

RIDL! {#[uuid(0x915d5ba9, 0x08ba, 0x40d4, 0xbc, 0x8c, 0x9b, 0x79, 0xe0, 0x4b, 0xf2, 0x6d)]
interface IObjectLock(IObjectLockVtbl): IDispatch(IDispatchVtbl) {
    fn get_connection(
        connection: *mut VARIANT,
    ) -> HRESULT,
    fn get_Object(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_LockedAt(
        pVal: *mut DATE,
    ) -> HRESULT,
    fn get_LockDescr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_session(
        session: *mut VARIANT,
    ) -> HRESULT,
}}
pub type LPOBJECTLOCK = *mut IObjectLock;

RIDL! {#[uuid(0x8474fb87, 0x8a0e, 0x4f97, 0x8f, 0x3e, 0x37, 0x9a, 0x6c, 0x95, 0x54, 0x97)]
interface ILicenseInfo(ILicenseInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_IssuedByServer(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn get_RMngrAddress(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_RMngrPort(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_RMngrPID(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_LicenseType(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_Series(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_Net(
        pVal: *mut VARIANT_BOOL,
    ) -> HRESULT,
    fn get_MaxUsersCur(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_MaxUsersAll(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_FileName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_FullPresentation(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_ShortPresentation(
        pVal: *mut BSTR,
    ) -> HRESULT,
}}
pub type LPLICENSEINFO = *mut ILicenseInfo;

RIDL! {#[uuid(0xe3dc5c5b, 0x7290, 0x4ace, 0xa6, 0xd8, 0x96, 0xdd, 0x5d, 0xba, 0xa9, 0xcb)]
interface IClusterServiceInfo(IClusterServiceInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_Name(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_MainOnly(
        pVal: *mut i8,
    ) -> HRESULT,
    fn get_ClusterManagers(
        pVal: *mut SAFEARRAY,
    ) -> HRESULT,
}}
pub type LPCLUSTERSERVICEINFO = *mut IClusterServiceInfo;

RIDL! {#[uuid(0x223a0210, 0x5059, 0x41be, 0x8d, 0x59, 0x8e, 0xea, 0xfe, 0x5a, 0x7a, 0xca)]
interface IClusterManagerInfo(IClusterManagerInfoVtbl): IDispatch(IDispatchVtbl) {
    fn get_HostName(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_HostName(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_MainPort(
        pVal: *mut INT,
    ) -> HRESULT,
    fn get_Descr(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn put_Descr(
        pVal: BSTR,
    ) -> HRESULT,
    fn get_PID(
        pVal: *mut BSTR,
    ) -> HRESULT,
    fn get_MainManager(
        pVal: *mut i8,
    ) -> HRESULT,
}}
pub type LPCLUSTERMANAGERINFO = *mut IClusterManagerInfo;
