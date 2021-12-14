unit uComAdmin;

// https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections

interface

uses
  System.Generics.Collections,
  System.SysUtils,
  COMAdmin_TLB;

type
  TCOMAdminAccessChecksLevelOptions = (COMAdminAccessChecksApplicationLevel, COMAdminAccessChecksApplicationComponentLevel);
  TCOMAdminApplicationActivation = (COMAdminActivationInproc, COMAdminActivationLocal);
  TCOMAdminAuthenticationCapability = (COMAdminAuthenticationCapabilitiesNone, COMAdminAuthenticationCapabilitiesSecureReference,
                                       COMAdminAuthenticationCapabilitiesStaticCloaking, COMAdminAuthenticationCapabilitiesDynamicCloaking);
  TCOMAdminAuthenticationLevel = (COMAdminAuthenticationDefault, COMAdminAuthenticationNone, COMAdminAuthenticationConnect,
                                  COMAdminAuthenticationCall, COMAdminAuthenticationPacket, COMAdminAuthenticationIntegrity,
                                  COMAdminAuthenticationPrivacy);
  TCOMAdminComponentType = (COMAdmin32BitComponent = $1, COMAdmin64BitComponent = $2);
  TCOMAdminImpersonationLevel = (COMAdminImpersonationAnonymous, COMAdminImpersonationIdentify, COMAdminImpersonationImpersonate, COMAdminImpersonationDelegate);
  TCOMAdminOperatingSystem = (COMAdminOSNotInitialized, COMAdminOSWindows3_1, COMAdminOSWindows9x, COMAdminOSWindows2000,
                              COMAdminOSWindows2000AdvancedServer, COMAdminOSWindows2000Unknown, COMAdminOSUnknown, COMAdminOSWindowsXPPersonal,
                              COMAdminOSWindowsXPProfessional, COMAdminOSWindowsNETStandardServer, COMAdminOSWindowsNETEnterpriseServer,
                              COMAdminOSWindowsNETDatacenterServer, COMAdminOSWindowsNETWebServer, COMAdminOSWindowsLonghornPersonal,
                              COMAdminOSWindowsLonghornProfessional, COMAdminOSWindowsLonghornStandardServer, COMAdminOSWindowsLonghornEnterpriseServer,
                              COMAdminOSWindowsLonghornDatacenterServer, COMAdminOSWindowsLonghornWebServer, COMAdminOSWindows7Personal,
                              COMAdminOSWindows7Professional, COMAdminOSWindows7StandardServer, COMAdminOSWindows7EnterpriseServer,
                              COMAdminOSWindows7DatacenterServer, COMAdminOSWindows7WebServer, COMAdminOSWindows8Personal,
                              COMAdminOSWindows8Professional, COMAdminOSWindows8StandardServer, COMAdminOSWindows8EnterpriseServer,
                              COMAdminOSWindows8DatacenterServer, COMAdminOSWindows8WebServer, COMAdminOSWindowsBluePersonal,
                              COMAdminOSWindowsBlueProfessional, COMAdminOSWindowsBlueStandardServer, COMAdminOSWindowsBlueEnterpriseServer,
                              COMAdminOSWindowsBlueDatacenterServer, COMAdminOSWindowsBlueWebServer);
  TCOMAdminQCAuthenticateMsgs = (COMAdminQCMessageAuthenticateSecureApps, COMAdminQCMessageAuthenticateOff, COMAdminQCMessageAuthenticateOn);
  TCOMAdminSRPTrustLevel = (COMAdminSRPDisallow = $0, COMAminSRPFullyTrusted = $40000);

  // forward declaration of internally used classes
  TComAdminBaseList = class;

  // generic base class for all single objects
  TComAdminBaseObject = class(TObject)
  private
    FCatalogObject: ICatalogObject;
    FCatalogCollection: ICatalogCollection;
    FKey: string;
    FName: string;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    property Name: string read FName write FName;
    property Key: string read FKey write FKey;
  end;

  // generic list class for all collections of objects
  TComAdminBaseList = class(TObjectList<TComAdminBaseObject>)
  strict private
    FCatalogCollection: ICatalogCollection;
  public
    constructor Create(ACatalogCollection: ICatalogCollection); reintroduce;
    property CatalogCollection: ICatalogCollection read FCatalogCollection write FCatalogCollection;
  end;

  TComAdminUser = class(TComAdminBaseObject);

  TComAdminUserList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminUser;
  public
    property Items[Index: Integer]: TComAdminUser read GetItem; default;
  end;

  TComAdminRole = class(TComAdminBaseObject)
  strict private
    FDescription: string;
    FUsers: TComAdminUserList;
    procedure GetUsers;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Description: string read FDescription write FDescription;
    property Users: TComAdminUserList read FUsers write FUsers;
  end;

  TComAdminRoleList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminRole;
  public
    property Items[Index: Integer]: TComAdminRole read GetItem; default;
  end;

  TComAdminInstance = class(TComAdminBaseObject)
  private
    FProcessID: Cardinal;
    FHasRecycled: Boolean;
    FIsPaused: Boolean;
  public
    property HasRecycled: Boolean read FHasRecycled;
    property IsPaused: Boolean read FIsPaused;
    property ProcessID: Cardinal read FProcessID;
  end;

  TComAdminInstanceList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminInstance;
  public
    property Items[Index: Integer]: TComAdminInstance read GetItem; default;
  end;

  TComAdminPartition = class(TComAdminBaseObject)
  strict private
    FDescription: string;
    FChangeable: Boolean;
    FDeleteable: Boolean;
    procedure ReadExtendedProperties;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    property Changeable: Boolean read FChangeable write FChangeable default True;
    property Deleteable: Boolean read FDeleteable write FDeleteable default True;
    property Description: string read FDescription write FDescription;
  end;

  TComAdminPartitionList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminPartition;
  public
    property Items[Index: Integer]: TComAdminPartition read GetItem; default;
  end;

  TCOMAdminComponent = class(TComAdminBaseObject)
  strict private
    procedure ReadExtendedProperties;
  private
    FAllowInprocSubscribers: Boolean;
    FApplicationID: string;
    FBitness: TCOMAdminComponentType;
    FComponentAccessChecksEnabled: Boolean;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    property AllowInprocSubscribers: Boolean read FAllowInprocSubscribers write FAllowInprocSubscribers default True;
    property ApplicationID: string read FApplicationID write FApplicationID;
    property Bitness: TCOMAdminComponentType read FBitness write FBitness;
    property ComponentAccessChecksEnabled: Boolean read FComponentAccessChecksEnabled write FComponentAccessChecksEnabled default False;
  end;

  TCOMAdminComponentList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TCOMAdminComponent;
  public
    property Items[Index: Integer]: TCOMAdminComponent read GetItem; default;
  end;

  TComAdminApplication = class(TComAdminBaseObject)
  strict private
    FRoles: TComAdminRoleList;
    FInstances: TComAdminInstanceList;
    FComponents: TCOMAdminComponentList;
    FDescription: string;
    FGig3SupportEnabled: Boolean;
    FAccessChecksLevel: TCOMAdminAccessChecksLevelOptions;
    FActivation: TCOMAdminApplicationActivation;
    FAccessChecksEnabled: Boolean;
    FDirectory: string;
    FProxy: Boolean;
    FProxyServerName: string;
    FPartitionID: string;
    FAuthenticationLevel: TCOMAdminAuthenticationLevel;
    FAuthenticationCapability: TCOMAdminAuthenticationCapability;
    FChangeable: Boolean;
    FCommandLine: string;
    FConcurrentApps: Cardinal;
    FCreatedBy: string;
    FCRMEnabled: Boolean;
    FCRMLogFile: string;
    FDeleteable: Boolean;
    FDumpEnabled: Boolean;
    FDumpOnException: Boolean;
    FDumpOnFailFast: Boolean;
    FDumpPath: string;
    FEventsEnabled: Boolean;
    FIdentity: string;
    FImpersonationLevel: TCOMAdminImpersonationLevel;
    FIsSystem: Boolean;
    FIsEnabled: Boolean;
    FMaxDumpCount: Cardinal;
    FPassword: string;
    FQCAuthenticateMsgs: TCOMAdminQCAuthenticateMsgs;
    FQCListenerMaxThreads: Cardinal;
    FQueueListenerEnabled: Boolean;
    FQueuingEnabled: Boolean;
    FRecycleCallLimit: Cardinal;
    FRecycleActivationLimit: Cardinal;
    FRecycleMemoryLimit: Cardinal;
    FRecycleExpirationTimeout: Cardinal;
    FRecycleLifetimeLimit: Cardinal;
    FRunForever: Boolean;
    FReplicable: Boolean;
    FServiceName: string;
    FShutdownAfter: Cardinal;
    FSRPEnabled: Boolean;
    FSoapActivated: Boolean;
    FSRPTrustLevel: TCOMAdminSRPTrustLevel;
    FSoapBaseUrl: string;
    FSoapVRoot: string;
    FSoapMailTo: string;
    procedure GetComponents;
    procedure GetRoles;
    procedure ReadExtendedProperties;
    procedure SetConcurrentApps(const Value: Cardinal);
    procedure SetMaxDumpCount(const Value: Cardinal);
    procedure SetQCListenerMaxThreads(const Value: Cardinal);
    procedure SetRecycleActivationLimit(const Value: Cardinal);
    procedure SetRecycleCallLimit(const Value: Cardinal);
    procedure SetRecycleExpirationTimeout(const Value: Cardinal);
    procedure SetRecycleLifetimeLimit(const Value: Cardinal);
    procedure SetRecycleMemoryLimit(const Value: Cardinal);
    procedure SetShutdownAfter(const Value: Cardinal);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function GetInstances: TComAdminInstanceList;
    property Roles: TComAdminRoleList read FRoles;
    property Instances: TComAdminInstanceList read FInstances;
    property Gig3SupportEnabled: Boolean read FGig3SupportEnabled write FGig3SupportEnabled default False;
    property AccessChecksLevel: TCOMAdminAccessChecksLevelOptions read FAccessChecksLevel write FAccessChecksLevel default COMAdminAccessChecksApplicationComponentLevel;
    property Activation: TCOMAdminApplicationActivation read FActivation write FActivation default COMAdminActivationLocal;
    property AccessChecksEnabled: Boolean read FAccessChecksEnabled write FAccessChecksEnabled default True;
    property Directory: string read FDirectory write FDirectory;
    property Proxy: Boolean read FProxy write FProxy default False;
    property ProxyServerName: string read FProxyServerName write FProxyServerName;
    property PartitionID: string read FPartitionID write FPartitionID;
    property AuthenticationLevel: TCOMAdminAuthenticationLevel read FAuthenticationLevel write FAuthenticationLevel default COMAdminAuthenticationDefault;
    property AuthenticationCapability: TCOMAdminAuthenticationCapability read FAuthenticationCapability write FAuthenticationCapability default COMAdminAuthenticationCapabilitiesDynamicCloaking;
    property Changeable: Boolean read FChangeable write FChangeable default True;
    property CommandLine: string read FCommandLine write FCommandLine;
    property ConcurrentApps: Cardinal read FConcurrentApps write SetConcurrentApps default 1;
    property CreatedBy: string read FCreatedBy write FCreatedBy;
    property CRMEnabled: Boolean read FCRMEnabled write FCRMEnabled default False;
    property CRMLogFile: string read FCRMLogFile write FCRMLogFile;
    property Deleteable: Boolean read FDeleteable write FDeleteable default True;
    property Description: string read FDescription write FDescription;
    property DumpEnabled: Boolean read FDumpEnabled write FDumpEnabled default False;
    property DumpOnException: Boolean read FDumpOnException write FDumpOnException default False;
    property DumpOnFailFast: Boolean read FDumpOnFailFast write FDumpOnFailFast default False;
    property DumpPath: string read FDumpPath write FDumpPath;
    property EventsEnabled: Boolean read FEventsEnabled write FEventsEnabled default True;
    property Identity: string read FIdentity write FIdentity;
    property ImpersonationLevel: TCOMAdminImpersonationLevel read FImpersonationLevel write FImpersonationLevel default COMAdminImpersonationImpersonate;
    property IsEnabled: Boolean read FIsEnabled write FIsEnabled default True;
    property IsSystem: Boolean read FIsSystem default False;
    property MaxDumpCount: Cardinal read FMaxDumpCount write SetMaxDumpCount default 5;
    property Password: string write FPassword;
    property QCAuthenticateMsgs: TCOMAdminQCAuthenticateMsgs read FQCAuthenticateMsgs write FQCAuthenticateMsgs default COMAdminQCMessageAuthenticateSecureApps;
    property QCListenerMaxThreads: Cardinal read FQCListenerMaxThreads write SetQCListenerMaxThreads default 0;
    property QueueListenerEnabled: Boolean read FQueueListenerEnabled write FQueueListenerEnabled default False;
    property QueuingEnabled: Boolean read FQueuingEnabled write FQueuingEnabled default False;
    property RecycleActivationLimit: Cardinal read FRecycleActivationLimit write SetRecycleActivationLimit default 0;
    property RecycleCallLimit: Cardinal read FRecycleCallLimit write SetRecycleCallLimit default 0;
    property RecycleExpirationTimeout: Cardinal read FRecycleExpirationTimeout write SetRecycleExpirationTimeout default 15;
    property RecycleLifetimeLimit: Cardinal read FRecycleLifetimeLimit write SetRecycleLifetimeLimit default 0;
    property RecycleMemoryLimit: Cardinal read FRecycleMemoryLimit write SetRecycleMemoryLimit default 0;
    property Replicable: Boolean read FReplicable write FReplicable default True;
    property RunForever: Boolean read FRunForever write FRunForever default False;
    property ServiceName: string read FServiceName write FServiceName;
    property ShutdownAfter: Cardinal read FShutdownAfter write SetShutdownAfter default 3;
    property SoapActivated: Boolean read FSoapActivated write FSoapActivated default False;
    property SoapBaseUrl: string read FSoapBaseUrl write FSoapBaseUrl;
    property SoapMailTo: string read FSoapMailTo write FSoapMailTo;
    property SoapVRoot: string read FSoapVRoot write FSoapVRoot;
    property SRPEnabled: Boolean read FSRPEnabled write FSRPEnabled default False;
    property SRPTrustLevel: TCOMAdminSRPTrustLevel read FSRPTrustLevel write FSRPTrustLevel default COMAminSRPFullyTrusted;
  end;

  TComAdminApplicationList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminApplication;
  public
    function Find(const AName: string; var AApplication: TComAdminApplication): Boolean;
    property Items[Index: Integer]: TComAdminApplication read GetItem; default;
  end;

  TComAdminComputer = class(TComAdminBaseObject)
  strict private
    FApplicationProxyRSN: string;
    FDCOMEnabled: Boolean;
    FCISEnabled: Boolean;
    FImpersonationLevel: TCOMAdminImpersonationLevel;
    FAuthenticationLevel: TCOMAdminAuthenticationLevel;
    FDefaultToInternetPorts: Boolean;
    FDescription: string;
    FDSPartitionLookupEnabled: Boolean;
    FIsRouter: Boolean;
    FInternetPortsListed: Boolean;
    FLoadBalancingCLSID: string;
    FLocalPartitionLookupEnabled: Boolean;
    FOperatingSystem: TCOMAdminOperatingSystem;
    FPartitionsEnabled: Boolean;
    FPorts: string;
    FResourcePoolingEnabled: Boolean;
    FRPCProxyEnabled: Boolean;
    procedure ReadExtendedProperties;
  private
    FSecureReferencesEnabled: Boolean;
    FSecurityTrackingEnabled: Boolean;
    FSRPActivateAsActivatorChecks: Boolean;
    FSRPRunningObjectChecks: Boolean;
    FTransactionTimeout: Cardinal;
    procedure SetTransactionTimeout(const Value: Cardinal);
  public
    constructor Create(ACatalogCollection: ICatalogCollection); reintroduce;
    property ApplicationProxyRSN: string read FApplicationProxyRSN write FApplicationProxyRSN;
    property CISEnabled: Boolean read FCISEnabled write FCISEnabled default False;
    property DCOMEnabled: Boolean read FDCOMEnabled write FDCOMEnabled default True;
    property DefaultAuthenticationLevel: TCOMAdminAuthenticationLevel read FAuthenticationLevel write FAuthenticationLevel default COMAdminAuthenticationConnect;
    property DefaultImpersonationLevel: TCOMAdminImpersonationLevel read FImpersonationLevel write FImpersonationLevel default COMAdminImpersonationIdentify;
    property DefaultToInternetPorts: Boolean read FDefaultToInternetPorts write FDefaultToInternetPorts default False;
    property Description: string read FDescription write FDescription;
    property DSPartitionLookupEnabled: Boolean read FDSPartitionLookupEnabled write FDSPartitionLookupEnabled default True;
    property InternetPortsListed: Boolean read FInternetPortsListed write FInternetPortsListed default False;
    property IsRouter: Boolean read FIsRouter write FIsRouter default False;
    property LoadBalancingCLSID: string read FLoadBalancingCLSID write FLoadBalancingCLSID;
    property LocalPartitionLookupEnabled: Boolean read FLocalPartitionLookupEnabled write FLocalPartitionLookupEnabled default True;
    property OperatingSystem: TCOMAdminOperatingSystem read FOperatingSystem write FOperatingSystem default COMAdminOSNotInitialized;
    property PartitionsEnabled: Boolean read FPartitionsEnabled write FPartitionsEnabled default False;
    property Ports: string read FPorts write FPorts;
    property ResourcePoolingEnabled: Boolean read FResourcePoolingEnabled write FResourcePoolingEnabled default True;
    property RPCProxyEnabled: Boolean read FRPCProxyEnabled write FRPCProxyEnabled default False;
    property SecureReferencesEnabled: Boolean read FSecureReferencesEnabled write FSecureReferencesEnabled default False;
    property SecurityTrackingEnabled: Boolean read FSecurityTrackingEnabled write FSecurityTrackingEnabled default True;
    property SRPActivateAsActivatorChecks: Boolean read FSRPActivateAsActivatorChecks write FSRPActivateAsActivatorChecks default True;
    property SRPRunningObjectChecks: Boolean read FSRPRunningObjectChecks write FSRPRunningObjectChecks default True;
    property TransactionTimeout: Cardinal read FTransactionTimeout write SetTransactionTimeout default 60;
  end;

  TComAdminCatalog = class(TObject)
  strict private
    FCatalog: ICOMAdminCatalog2;
    FApplications: TComAdminApplicationList;
    FPartitions: TComAdminPartitionList;
    FComputer: TComAdminComputer;
    FFilter: string;
    procedure GetApplications;
    procedure GetPartitions;
    procedure SetFilter(const Value: string);
  public
    constructor Create(const AServer: string; const AFilter: string = ''); reintroduce;
    destructor Destroy; override;
    procedure ExportApplication(AIndex: Integer; const AFilename: string);
    procedure ExportApplicationByKey(const AKey, AFilename: string);
    procedure ExportApplicationByName(const AName, AFilename: string);
    property Applications: TComAdminApplicationList read FApplications;
    property Computer: TComAdminComputer read FComputer write FComputer;
    property Filter: string read FFilter write SetFilter;
  end;

  EItemNotFoundException = Exception;

implementation

uses
  System.Masks,
  System.Variants,
  Winapi.Windows;

const
  COLLECTION_NAME_APPS = 'Applications';
  COLLECTION_NAME_COMPONENTS = 'Components';
  COLLECTION_NAME_COMPUTER = 'LocalComputer';
  COLLECTION_NAME_INSTANCES = 'ApplicationInstances';
  COLLECTION_NAME_PARTITIONS = 'Partitions';
  COLLECTION_NAME_ROLES = 'Roles';
  COLLECTION_NAME_USERS = 'UsersInRole';
  DEFAULT_APP_FILTER = '*';
  PROPERTY_NAME_3GIG = '3GigSupportEnabled';
  PROPERTY_NAME_ACCESS_CHECK_LEVEL = 'AccessChecksLevel';
  PROPERTY_NAME_ACCESS_CHECKS = 'ApplicationAccessChecksEnabled';
  PROPERTY_NAME_ACTIVATION = 'Activation';
  PROPERTY_NAME_ALLOW_SUBSCRIBERS = 'AllowInprocSubscribers';
  PROPERTY_NAME_APPLICATION_DIRECTORY = 'ApplicationDirectory';
  PROPERTY_NAME_APPLICATION_ID = 'ApplicationID';
  PROPERTY_NAME_APPLICATION_PROXY = 'ApplicationProxy';
  PROPERTY_NAME_AUTH_CAPABILITY = 'AuthenticationCapability';
  PROPERTY_NAME_AUTHENTICATION = 'Authentication';
  PROPERTY_NAME_BITNESS = 'Bitness';
  PROPERTY_NAME_CHANGEABLE = 'Changeable';
  PROPERTY_NAME_CIS_ENABLED = 'CISEnabled';
  PROPERTY_NAME_COMMAND_LINE = 'CommandLine';
  PROPERTY_NAME_COMPONENT_ACCESS_CHECKS = 'ComponentAccessChecksEnabled';
  PROPERTY_NAME_CONCURRENT_APPS = 'ConcurrentApps';
  PROPERTY_NAME_CREATED_BY = 'CreatedBy';
  PROPERTY_NAME_CRM_ENABLED = 'CRMEnabled';
  PROPERTY_NAME_CRM_LOGFILE = 'CRMLogFile';
  PROPERTY_NAME_DCOM_ENABLED = 'DCOMEnabled';
  PROPERTY_NAME_DEFAULT_AUTHENTICATION = 'DefaultAuthenticationLevel';
  PROPERTY_NAME_DEFAULT_IMPERSONATION = 'DefaultImpersonationLevel';
  PROPERTY_NAME_DEFAULT_TO_INTERNET = 'DefaultToInternetPorts';
  PROPERTY_NAME_DELETEABLE = 'Deleteable';
  PROPERTY_NAME_DESCRIPTION = 'Description';
  PROPERTY_NAME_DS_PARTITION_LOOKUP = 'DSPartitionLookupEnabled';
  PROPERTY_NAME_DUMP_ENABLED = 'DumpEnabled';
  PROPERTY_NAME_DUMP_EXCEPTION = 'DumpOnException';
  PROPERTY_NAME_DUMP_FAILFAST = 'DumpOnFailfast';
  PROPERTY_NAME_DUMP_PATH = 'DumpPath';
  PROPERTY_NAME_ENABLED = 'IsEnabled';
  PROPERTY_NAME_EVENTS_ENABLED = 'EventsEnabled';
  PROPERTY_NAME_IDENTITY = 'Identity';
  PROPERTY_NAME_IMPERSONATION = 'ImpersonationLevel';
  PROPERTY_NAME_INTERNET_PORTS = 'InternetPortsListed';
  PROPERTY_NAME_IS_ROUTER = 'IsRouter';
  PROPERTY_NAME_LOAD_BALANCING_ID = 'LoadBalancingCLSID';
  PROPERTY_NAME_MAX_DUMPS = 'MaxDumpCount';
  PROPERTY_NAME_OPERATING_SYSTEM = 'OperatingSystem';
  PROPERTY_NAME_PARTITION_ID = 'AppPartitionID';
  PROPERTY_NAME_PARTITION_LOOKUP = 'LocalPartitionLookupEnabled';
  PROPERTY_NAME_PARTITIONS_ENABLED = 'PartitionsEnabled';
  PROPERTY_NAME_PASSWORD = 'Password';
  PROPERTY_NAME_PAUSED = 'IsPaused';
  PROPERTY_NAME_PORTS = 'Ports';
  PROPERTY_NAME_PROCESSID = 'ProcessID';
  PROPERTY_NAME_PROXY_RSN = 'ApplicationProxyRSN';
  PROPERTY_NAME_PROXY_SERVER_NAME = 'ApplicationProxyServerName';
  PROPERTY_NAME_QC_AUTHENTICATE = 'QCAuthenticateMsgs';
  PROPERTY_NAME_QC_MAXTHREADS = 'QCListenerMaxThreads';
  PROPERTY_NAME_QUEUE_LISTENER = 'QueueListenerEnabled';
  PROPERTY_NAME_QUEUING_ENABLED = 'QueuingEnabled';
  PROPERTY_NAME_RECYCLE_ACTIVATION = 'RecycleActivationLimit';
  PROPERTY_NAME_RECYCLE_CALL_LIMIT = 'RecycleCallLimit';
  PROPERTY_NAME_RECYCLE_EXPIRATION = 'RecycleExpirationTimeout';
  PROPERTY_NAME_RECYCLE_LIFETIME_LIMIT = 'RecycleLifetimeLimit';
  PROPERTY_NAME_RECYCLE_MEMORY_LIMIT = 'RecycleMemoryLimit';
  PROPERTY_NAME_RECYCLED = 'HasRecycled';
  PROPERTY_NAME_REPLICABLE = 'Replicable';
  PROPERTY_NAME_RESOURCE_POOLING = 'ResourcePoolingEnabled';
  PROPERTY_NAME_RPC_PROXY_ENABLED = 'RPCProxyEnabled';
  PROPERTY_NAME_RUN_FOREVER = 'RunForever';
  PROPERTY_NAME_SECURE_REFERENCES = 'SecureReferencesEnabled';
  PROPERTY_NAME_SECURE_TRACKING = 'SecurityTrackingEnabled';
  PROPERTY_NAME_SERVICE_NAME = 'ServiceName';
  PROPERTY_NAME_SHUTDOWN = 'ShutdownAfter';
  PROPERTY_NAME_SOAP_ACTIVATED = 'SoapActivated';
  PROPERTY_NAME_SOAP_BASE_URL = 'SoapBaseUrl';
  PROPERTY_NAME_SOAP_MAILTO = 'SoapMailTo';
  PROPERTY_NAME_SOAP_VROOT = 'SoapVRoot';
  PROPERTY_NAME_SRP_ACTIVATE_CHECKS = 'SRPActivateAsActivatorChecks';
  PROPERTY_NAME_SRP_ENABLED = 'SRPEnabled';
  PROPERTY_NAME_SRP_OBJECTS_CHECK = 'SRPRunningObjectChecks';
  PROPERTY_NAME_SRP_TRUSTLEVEL = 'SRPTrustLevel';
  PROPERTY_NAME_SYSTEM = 'IsSystem';
  PROPERTY_NAME_TRANSACTION_TIMEOUT = 'TransactionTimeout';
  ERROR_NOT_FOUND = 'Element %s could not be found in this collection';
  ERROR_OUT_OF_RANGE = 'Value out of range';

{ TComAdminBaseObject }

constructor TComAdminBaseObject.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create;
  FCatalogObject := ACatalogObject;
  if Assigned(ACollection) then
    FCatalogCollection := ACollection.CatalogCollection;
  FKey := FCatalogObject.Key;
  FName := FCatalogObject.Name;
end;

{ TComAdminBaseList }

constructor TComAdminBaseList.Create(ACatalogCollection: ICatalogCollection);
begin
  inherited Create(True);
  FCatalogCollection := ACatalogCollection;
  FCatalogCollection.Populate;
end;

{ TComAdminUserList }

function TComAdminUserList.GetItem(Index: Integer): TComAdminUser;
begin
  Result := inherited Items[Index] as TComAdminUser;
end;

{ TComAdminRole }

constructor TComAdminRole.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  FDescription := VarToStrDef(FCatalogObject.Value[PROPERTY_NAME_DESCRIPTION], '');
  FUsers := TComAdminUserList.Create(ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_USERS, FKey) as ICatalogCollection);
  GetUsers;
end;

destructor TComAdminRole.Destroy;
begin
  FUsers.Free;
  inherited;
end;

procedure TComAdminRole.GetUsers;
var
  i: Integer;
begin
  for i := 0 to FUsers.CatalogCollection.Count - 1 do
    FUsers.Add(TComAdminUser.Create(FUsers, FUsers.CatalogCollection.Item[i] as ICatalogObject));
end;

{ TComAdminRoleList }

function TComAdminRoleList.GetItem(Index: Integer): TComAdminRole;
begin
  Result := inherited Items[Index] as TComAdminRole;
end;

{ TComAdminInstanceList }

function TComAdminInstanceList.GetItem(Index: Integer): TComAdminInstance;
begin
  Result := inherited Items[Index] as TComAdminInstance;
end;

{ TComAdminPartition }

constructor TComAdminPartition.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
end;

procedure TComAdminPartition.ReadExtendedProperties;
begin
  FChangeable := VarAsType(FCatalogObject.Value[PROPERTY_NAME_CHANGEABLE], varBoolean);
  FDeleteable := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DELETEABLE], varBoolean);
  FDescription := VarToStr(FCatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
end;

{ TComAdminPartitionList }

function TComAdminPartitionList.GetItem(Index: Integer): TComAdminPartition;
begin
  Result := inherited Items[Index] as TComAdminPartition;
end;

{ TCOMAdminComponent }

constructor TCOMAdminComponent.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
end;

procedure TCOMAdminComponent.ReadExtendedProperties;
begin
  FAllowInprocSubscribers := VarAsType(FCatalogObject.Value[PROPERTY_NAME_ALLOW_SUBSCRIBERS], varBoolean);
  FApplicationID := VarToStr(FCatalogObject.Value[PROPERTY_NAME_APPLICATION_ID]);
  FBitness := VarAsType(FCatalogObject.Value[PROPERTY_NAME_BITNESS], varLongWord);
  FComponentAccessChecksEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_COMPONENT_ACCESS_CHECKS], varBoolean);
end;

{ TCOMAdminComponentList }

function TCOMAdminComponentList.GetItem(Index: Integer): TCOMAdminComponent;
begin
  Result := inherited Items[Index] as TCOMAdminComponent;
end;

{ TComAdminApplication }

constructor TComAdminApplication.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
  // Create List objects
  FInstances := TComAdminInstanceList.Create(ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, FKey) as ICatalogCollection);
  FRoles := TComAdminRoleList.Create(ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_ROLES, FKey) as ICatalogCollection);
  GetRoles;
  FComponents := TCOMAdminComponentList.Create(ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_COMPONENTS, FKey) as ICatalogCollection);
  GetComponents;
end;

destructor TComAdminApplication.Destroy;
begin
  FRoles.Free;
  FInstances.Free;
  FComponents.Free;
  inherited;
end;

procedure TComAdminApplication.GetComponents;
var
  i: Integer;
begin
  for i := 0 to FComponents.CatalogCollection.Count - 1 do
    FComponents.Add(TCOMAdminComponent.Create(FComponents, FComponents.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminApplication.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

function TComAdminApplication.GetInstances: TComAdminInstanceList;
var
  Collection: ICatalogCollection;
  Instance: TComAdminInstance;
  i: Integer;
begin
  FInstances.Clear;
  Collection := FCatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, FKey) as ICatalogCollection;
  Collection.Populate;
  for i := 0 to Collection.Count - 1 do
  begin
    Instance := TComAdminInstance.Create(FInstances, Collection.Item[i] as ICatalogObject);
    Instance.FHasRecycled := (Collection.Item[i] as ICatalogObject).Value[PROPERTY_NAME_RECYCLED];
    Instance.FIsPaused := (Collection.Item[i] as ICatalogObject).Value[PROPERTY_NAME_PAUSED];
    Instance.FProcessID := VarAsType((Collection.Item[i] as ICatalogObject).Value[PROPERTY_NAME_PROCESSID], varLongWord);
    FInstances.Add(Instance);
  end;
  Result := FInstances;
end;

procedure TComAdminApplication.ReadExtendedProperties;
begin
  FGig3SupportEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_3GIG], varBoolean);
  FAccessChecksLevel := VarAsType(FCatalogObject.Value[PROPERTY_NAME_ACCESS_CHECK_LEVEL], varLongWord);
  FActivation := VarAsType(FCatalogObject.Value[PROPERTY_NAME_ACTIVATION], varLongWord);
  FAccessChecksEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_ACCESS_CHECKS], varBoolean);
  FDirectory := VarToStr(FCatalogObject.Value[PROPERTY_NAME_APPLICATION_DIRECTORY]);
  FProxy := VarAsType(FCatalogObject.Value[PROPERTY_NAME_APPLICATION_PROXY], varBoolean);
  FProxyServerName := VarToStr(FCatalogObject.Value[PROPERTY_NAME_PROXY_SERVER_NAME]);
  FPartitionID := VarToStr(FCatalogObject.Value[PROPERTY_NAME_PARTITION_ID]);
  FAuthenticationLevel := VarAsType(FCatalogObject.Value[PROPERTY_NAME_AUTHENTICATION], varLongWord);
  FAuthenticationCapability := VarAsType(FCatalogObject.Value[PROPERTY_NAME_AUTH_CAPABILITY], varLongWord);
  FChangeable := VarAsType(FCatalogObject.Value[PROPERTY_NAME_CHANGEABLE], varBoolean);
  FCommandLine := VarToStr(FCatalogObject.Value[PROPERTY_NAME_COMMAND_LINE]);
  FConcurrentApps := VarAsType(FCatalogObject.Value[PROPERTY_NAME_CONCURRENT_APPS], varLongWord);
  FCreatedBy := VarToStr(FCatalogObject.Value[PROPERTY_NAME_CREATED_BY]);
  FCRMEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_CRM_ENABLED], varBoolean);
  FCRMLogFile := VarToStr(FCatalogObject.Value[PROPERTY_NAME_CRM_LOGFILE]);
  FDeleteable := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DELETEABLE], varBoolean);
  FDescription := VarToStr(FCatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FDumpEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DUMP_ENABLED], varBoolean);
  FDumpOnException := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DUMP_EXCEPTION], varBoolean);
  FDumpOnFailFast := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DUMP_FAILFAST], varBoolean);
  FDumpPath := VarToStr(FCatalogObject.Value[PROPERTY_NAME_DUMP_PATH]);
  FEventsEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_EVENTS_ENABLED], varBoolean);
  FIdentity := VarToStr(FCatalogObject.Value[PROPERTY_NAME_IDENTITY]);
  FImpersonationLevel := VarAsType(FCatalogObject.Value[PROPERTY_NAME_IMPERSONATION], varLongWord);
  FIsEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_ENABLED], varBoolean);
  FIsSystem := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SYSTEM], varBoolean);
  FMaxDumpCount := VarAsType(FCatalogObject.Value[PROPERTY_NAME_MAX_DUMPS], varLongWord);
  FQCAuthenticateMsgs := VarAsType(FCatalogObject.Value[PROPERTY_NAME_QC_AUTHENTICATE], varLongWord);
  FQCListenerMaxThreads := VarAsType(FCatalogObject.Value[PROPERTY_NAME_QC_MAXTHREADS], varLongWord);
  FQueueListenerEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_QUEUE_LISTENER], varBoolean);
  FQueuingEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_QUEUING_ENABLED], varBoolean);
  FRecycleActivationLimit := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RECYCLE_ACTIVATION], varLongWord);
  FRecycleCallLimit := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RECYCLE_CALL_LIMIT], varLongWord);
  FRecycleExpirationTimeout := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RECYCLE_EXPIRATION], varLongWord);
  FRecycleLifetimeLimit := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RECYCLE_LIFETIME_LIMIT], varLongWord);
  FRecycleMemoryLimit := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RECYCLE_MEMORY_LIMIT], varLongWord);
  FReplicable := VarAsType(FCatalogObject.Value[PROPERTY_NAME_REPLICABLE], varBoolean);
  FRunForever := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RUN_FOREVER], varBoolean);
  FServiceName := VarToStr(FCatalogObject.Value[PROPERTY_NAME_SERVICE_NAME]);
  FShutdownAfter := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SHUTDOWN], varLongWord);
  FSoapActivated := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SOAP_ACTIVATED], varBoolean);
  FSoapBaseUrl := VarToStr(FCatalogObject.Value[PROPERTY_NAME_SOAP_BASE_URL]);
  FSoapMailTo := VarToStr(FCatalogObject.Value[PROPERTY_NAME_SOAP_MAILTO]);
  FSoapVRoot := VarToStr(FCatalogObject.Value[PROPERTY_NAME_SOAP_VROOT]);
  FSRPEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SRP_ENABLED], varBoolean);
  FSRPTrustLevel := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SRP_TRUSTLEVEL], varLongWord);
end;

procedure TComAdminApplication.SetConcurrentApps(const Value: Cardinal);
begin
  case Value of
    1..1048576: FConcurrentApps := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetMaxDumpCount(const Value: Cardinal);
begin
  case Value of
    1..200: FMaxDumpCount := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetQCListenerMaxThreads(const Value: Cardinal);
begin
  case Value of
    1..1000: FQCListenerMaxThreads := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetRecycleActivationLimit(const Value: Cardinal);
begin
  case Value of
    1..1048576: FRecycleActivationLimit := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetRecycleCallLimit(const Value: Cardinal);
begin
  case Value of
    1..1048576: FRecycleCallLimit := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetRecycleExpirationTimeout(const Value: Cardinal);
begin
  case Value of
    1..1440: FRecycleExpirationTimeout := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetRecycleLifetimeLimit(const Value: Cardinal);
begin
  case Value of
    1..30240: FRecycleLifetimeLimit := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetRecycleMemoryLimit(const Value: Cardinal);
begin
  case Value of
    1..1048576: FRecycleMemoryLimit := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

procedure TComAdminApplication.SetShutdownAfter(const Value: Cardinal);
begin
  case Value of
    1..1440: FShutdownAfter := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

{ TComAdminApplicationList }

function TComAdminApplicationList.Find(const AName: string; var AApplication: TComAdminApplication): Boolean;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if Items[i].Name.Equals(AName) then
    begin
      AApplication := Items[i];
      Exit(True);
    end;
  end;
  Result := False;
end;

function TComAdminApplicationList.GetItem(Index: Integer): TComAdminApplication;
begin
  Result := inherited Items[Index] as TComAdminApplication;
end;

{ TComAdminComputer }

constructor TComAdminComputer.Create(ACatalogCollection: ICatalogCollection);
begin
  ACatalogCollection.Populate;
  inherited Create(nil, ACatalogCollection.Item[0] as ICatalogObject);
  ReadExtendedProperties;
end;

procedure TComAdminComputer.ReadExtendedProperties;
begin
  FApplicationProxyRSN := VarToStr(FCatalogObject.Value[PROPERTY_NAME_PROXY_RSN]);
  FCISEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_CIS_ENABLED], varBoolean);
  FDCOMEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DCOM_ENABLED], varBoolean);
  FAuthenticationLevel := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DEFAULT_AUTHENTICATION], varLongWord);
  FImpersonationLevel := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DEFAULT_IMPERSONATION], varLongWord);
  FDefaultToInternetPorts := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DEFAULT_TO_INTERNET], varBoolean);
  FDescription := VarToStr(FCatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FDSPartitionLookupEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_DS_PARTITION_LOOKUP], varBoolean);
  FInternetPortsListed := VarAsType(FCatalogObject.Value[PROPERTY_NAME_INTERNET_PORTS], varBoolean);
  FIsRouter := VarAsType(FCatalogObject.Value[PROPERTY_NAME_IS_ROUTER], varBoolean);
  FLoadBalancingCLSID := VarToStr(FCatalogObject.Value[PROPERTY_NAME_LOAD_BALANCING_ID]);
  FLocalPartitionLookupEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_IS_ROUTER], varBoolean);
  FOperatingSystem := VarAsType(FCatalogObject.Value[PROPERTY_NAME_OPERATING_SYSTEM], varLongWord);
  FPartitionsEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_PARTITIONS_ENABLED], varBoolean);
  FPorts := VarToStr(FCatalogObject.Value[PROPERTY_NAME_PORTS]);
  FResourcePoolingEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RESOURCE_POOLING], varBoolean);
  if FCISEnabled then // property only available if CIS is enabled
    FRPCProxyEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_RPC_PROXY_ENABLED], varBoolean);
  FSecureReferencesEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SECURE_REFERENCES], varBoolean);
  FSecurityTrackingEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SECURE_TRACKING], varBoolean);
  FSRPActivateAsActivatorChecks := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SRP_ACTIVATE_CHECKS], varBoolean);
  FSRPRunningObjectChecks := VarAsType(FCatalogObject.Value[PROPERTY_NAME_SRP_OBJECTS_CHECK], varBoolean);
  FTransactionTimeout := VarAsType(FCatalogObject.Value[PROPERTY_NAME_TRANSACTION_TIMEOUT], varLongWord);
end;

procedure TComAdminComputer.SetTransactionTimeout(const Value: Cardinal);
begin
  case Value of
    1..3600: FTransactionTimeout := Value;
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
  end;
end;

{ TComAdminCatalog }

constructor TComAdminCatalog.Create(const AServer: string; const AFilter: string);
begin
  inherited Create;
  if AFilter.IsEmpty then
    FFilter := DEFAULT_APP_FILTER
  else
    FFilter := AFilter;
  FCatalog := CoCOMAdminCatalog.Create;
  FApplications := TComAdminApplicationList.Create(FCatalog.GetCollection(COLLECTION_NAME_APPS) as ICatalogCollection);
  FComputer := TComAdminComputer.Create(FCatalog.GetCollection(COLLECTION_NAME_COMPUTER) as ICatalogCollection);
  FPartitions := TComAdminPartitionList.Create(FCatalog.GetCollection(COLLECTION_NAME_PARTITIONS) as ICatalogCollection);
  if not AServer.IsEmpty then
    FCatalog.Connect(AServer);
  GetApplications;
  GetPartitions;
end;

destructor TComAdminCatalog.Destroy;
begin
  FApplications.Free;
  FComputer.Free;
  FPartitions.Free;
  inherited;
end;

procedure TComAdminCatalog.ExportApplication(AIndex: Integer; const AFilename: string);
begin
  FCatalog.ExportApplication(FApplications.Items[AIndex].Key, AFilename, COMAdminExportUsers and COMAdminExportForceOverwriteOfFiles);
end;

procedure TComAdminCatalog.ExportApplicationByKey(const AKey, AFilename: string);
begin
  FCatalog.ExportApplication(AKey, AFilename, COMAdminExportUsers and COMAdminExportForceOverwriteOfFiles);
end;

procedure TComAdminCatalog.ExportApplicationByName(const AName, AFilename: string);
var
  Application: TComAdminApplication;
begin
  if FApplications.Find(AName, Application) then
    ExportApplicationByKey(Application.Key, AFilename)
  else
    raise EItemNotFoundException.CreateFmt(ERROR_NOT_FOUND, [QuotedStr(AName)]);
end;

procedure TComAdminCatalog.GetApplications;
var
  i: Integer;
  LMask: TMask;
begin
  for i := 0 to FApplications.CatalogCollection.Count - 1 do
  begin
    LMask := TMask.Create(FFilter);
    try
      if LMask.Matches((FApplications.CatalogCollection.Item[i] as ICatalogObject).Name) then
        FApplications.Add(TComAdminApplication.Create(FApplications, FApplications.CatalogCollection.Item[i] as ICatalogObject));
    finally
      LMask.Free;
    end;
  end;
end;

procedure TComAdminCatalog.GetPartitions;
var
  i: Integer;
begin
  for i := 0 to FPartitions.CatalogCollection.Count - 1 do
  begin
    FPartitions.Add(TComAdminPartition.Create(FPartitions, FPartitions.CatalogCollection.Item[i] as ICatalogObject));
  end;
end;

procedure TComAdminCatalog.SetFilter(const Value: string);
begin
  if not FFilter.Equals(Value) then
  begin
    FFilter := Value;
    FApplications.Clear;
    GetApplications;
  end;
end;

end.
