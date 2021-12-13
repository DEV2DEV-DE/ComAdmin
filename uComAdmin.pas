unit uComAdmin;

// https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections

interface

uses
  System.Generics.Collections,
  System.SysUtils,
  COMAdmin_TLB;

type
  TCOMAdminAccessChecksLevelOptions = (aclApplicationLevel, aclComponentLevel);
  TCOMAdminApplicationActivation = (aaActivationInproc, aaActivationLocal);
  TCOMAdminAuthenticationLevel = (alDefault, alNone, alConnect, alCall, alPacket, alIntegrity, alPrivacy);
  TCOMAdminAuthenticationCapability = (acNone = $0, acSecureReference = $2, acStaticCloaking = $20, acDynamicCloaking = $40);
  TCOMAdminImpersonationLevel = (ilAnonymous = 1, ilIdentify = 2, ilImpersonate = 3, ilDelegate = 4);
  TCOMAdminQCAuthenticateMsgs = (qcSecureApps, qcOff, qcOn);
  TCOMAdminSRPTrustLevel = (tlDisallow = $0, tlFullyTrusted = $40000);

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

  TComAdminInstanceList = class(TComAdminBaseList);

  TComAdminApplication = class(TComAdminBaseObject)
  strict private
    FRoles: TComAdminRoleList;
    FInstances: TComAdminInstanceList;
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
    property AccessChecksLevel: TCOMAdminAccessChecksLevelOptions read FAccessChecksLevel write FAccessChecksLevel default aclComponentLevel;
    property Activation: TCOMAdminApplicationActivation read FActivation write FActivation default aaActivationLocal;
    property AccessChecksEnabled: Boolean read FAccessChecksEnabled write FAccessChecksEnabled default True;
    property Directory: string read FDirectory write FDirectory;
    property Proxy: Boolean read FProxy write FProxy default False;
    property ProxyServerName: string read FProxyServerName write FProxyServerName;
    property PartitionID: string read FPartitionID write FPartitionID;
    property AuthenticationLevel: TCOMAdminAuthenticationLevel read FAuthenticationLevel write FAuthenticationLevel default alDefault;
    property AuthenticationCapability: TCOMAdminAuthenticationCapability read FAuthenticationCapability write FAuthenticationCapability default acDynamicCloaking;
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
    property ImpersonationLevel: TCOMAdminImpersonationLevel read FImpersonationLevel write FImpersonationLevel default ilImpersonate;
    property IsEnabled: Boolean read FIsEnabled write FIsEnabled default True;
    property IsSystem: Boolean read FIsSystem default False;
    property MaxDumpCount: Cardinal read FMaxDumpCount write SetMaxDumpCount default 5;
    property Password: string write FPassword;
    property QCAuthenticateMsgs: TCOMAdminQCAuthenticateMsgs read FQCAuthenticateMsgs write FQCAuthenticateMsgs default qcSecureApps;
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
    property SRPTrustLevel: TCOMAdminSRPTrustLevel read FSRPTrustLevel write FSRPTrustLevel default tlFullyTrusted;

  end;

  TComAdminApplicationList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminApplication;
  public
    function Find(const AName: string; var AApplication: TComAdminApplication): Boolean;
    property Items[Index: Integer]: TComAdminApplication read GetItem; default;
  end;

  TComAdminCatalog = class(TObject)
  strict private
    FCatalog: ICOMAdminCatalog2;
    FApplications: TComAdminApplicationList;
    FFilter: string;
    procedure GetApplications;
    procedure SetFilter(const Value: string);
  public
    constructor Create(const AServer: string); reintroduce;
    destructor Destroy; override;
    procedure ExportApplication(AIndex: Integer; const AFilename: string);
    procedure ExportApplicationByKey(const AKey, AFilename: string);
    procedure ExportApplicationByName(const AName, AFilename: string);
    property Applications: TComAdminApplicationList read FApplications;
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
  COLLECTION_NAME_ROLES = 'Roles';
  COLLECTION_NAME_USERS = 'UsersInRole';
  COLLECTION_NAME_INSTANCES = 'ApplicationInstances';
  DEFAULT_APP_FILTER = 'ProdLog-*';
  PROPERTY_NAME_RECYCLED = 'HasRecycled';
  PROPERTY_NAME_PAUSED = 'IsPaused';
  PROPERTY_NAME_PROCESSID = 'ProcessID';
  PROPERTY_NAME_3GIG = '3GigSupportEnabled';
  PROPERTY_NAME_ACCESS_CHECK_LEVEL = 'AccessChecksLevel';
  PROPERTY_NAME_ACTIVATION = 'Activation';
  PROPERTY_NAME_ACCESS_CHECKS = 'ApplicationAccessChecksEnabled';
  PROPERTY_NAME_APPLICATION_DIRECTORY = 'ApplicationDirectory';
  PROPERTY_NAME_APPLICATION_PROXY = 'ApplicationProxy';
  PROPERTY_NAME_PROXY_SERVER_NAME = 'ApplicationProxyServerName';
  PROPERTY_NAME_PARTITION_ID = 'AppPartitionID';
  PROPERTY_NAME_AUTHENTICATION = 'Authentication';
  PROPERTY_NAME_AUTH_CAPABILITY = 'AuthenticationCapability';
  PROPERTY_NAME_CHANGEABLE = 'Changeable';
  PROPERTY_NAME_COMMAND_LINE = 'CommandLine';
  PROPERTY_NAME_CONCURRENT_APPS = 'ConcurrentApps';
  PROPERTY_NAME_CREATED_BY = 'CreatedBy';
  PROPERTY_NAME_CRM_ENABLED = 'CRMEnabled';
  PROPERTY_NAME_CRM_LOGFILE = 'CRMLogFile';
  PROPERTY_NAME_DELETEABLE = 'Deleteable';
  PROPERTY_NAME_DESCRIPTION = 'Description';
  PROPERTY_NAME_DUMP_ENABLED = 'DumpEnabled';
  PROPERTY_NAME_DUMP_EXCEPTION = 'DumpOnException';
  PROPERTY_NAME_DUMP_FAILFAST = 'DumpOnFailfast';
  PROPERTY_NAME_DUMP_PATH = 'DumpPath';
  PROPERTY_NAME_EVENTS_ENABLED = 'EventsEnabled';
  PROPERTY_NAME_IDENTITY = 'Identity';
  PROPERTY_NAME_IMPERSONATION = 'ImpersonationLevel';
  PROPERTY_NAME_ENABLED = 'IsEnabled';
  PROPERTY_NAME_SYSTEM = 'IsSystem';
  PROPERTY_NAME_MAX_DUMPS = 'MaxDumpCount';
  PROPERTY_NAME_PASSWORD = 'Password';
  PROPERTY_NAME_QC_AUTHENTICATE = 'QCAuthenticateMsgs';
  PROPERTY_NAME_QC_MAXTHREADS = 'QCListenerMaxThreads';
  PROPERTY_NAME_QUEUE_LISTENER = 'QueueListenerEnabled';
  PROPERTY_NAME_QUEUING_ENABLED = 'QueuingEnabled';
  PROPERTY_NAME_RECYCLE_ACTIVATION = 'RecycleActivationLimit';
  PROPERTY_NAME_RECYCLE_CALL_LIMIT = 'RecycleCallLimit';
  PROPERTY_NAME_RECYCLE_EXPIRATION = 'RecycleExpirationTimeout';
  PROPERTY_NAME_RECYCLE_LIFETIME_LIMIT = 'RecycleLifetimeLimit';
  PROPERTY_NAME_RECYCLE_MEMORY_LIMIT = 'RecycleMemoryLimit';
  PROPERTY_NAME_REPLICABLE = 'Replicable';
  PROPERTY_NAME_RUN_FOREVER = 'RunForever';
  PROPERTY_NAME_SERVICE_NAME = 'ServiceName';
  PROPERTY_NAME_SHUTDOWN = 'ShutdownAfter';
  PROPERTY_NAME_SOAP_ACTIVATED = 'SoapActivated';
  PROPERTY_NAME_SOAP_BASE_URL = 'SoapBaseUrl';
  PROPERTY_NAME_SOAP_MAILTO = 'SoapMailTo';
  PROPERTY_NAME_SOAP_VROOT = 'SoapVRoot';
  PROPERTY_NAME_SRP_ENABLED = 'SRPEnabled';
  PROPERTY_NAME_SRP_TRUSTLEVEL = 'SRPTrustLevel';
  ERROR_NOT_FOUND = 'Element %s could not be found in this collection';
  ERROR_OUT_OF_RANGE = 'Value out of range';

{ TComAdminBaseObject }

constructor TComAdminBaseObject.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create;
  FCatalogObject := ACatalogObject;
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

{ TComAdminApplication }

constructor TComAdminApplication.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
  // Create List objects
  FRoles := TComAdminRoleList.Create(ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_ROLES, FKey) as ICatalogCollection);
  FInstances := TComAdminInstanceList.Create(ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, FKey) as ICatalogCollection);
  GetRoles;
end;

destructor TComAdminApplication.Destroy;
begin
  FRoles.Free;
  FInstances.Free;
  inherited;
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
  FGig3SupportEnabled := VarAsType(FCatalogObject.Value[PROPERTY_NAME_ACCESS_CHECK_LEVEL], varBoolean);
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

procedure TComAdminApplication.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
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

{ TComAdminCatalog }

constructor TComAdminCatalog.Create(const AServer: string);
begin
  inherited Create;
  FFilter := DEFAULT_APP_FILTER;
  FCatalog := CoCOMAdminCatalog.Create;
  FApplications := TComAdminApplicationList.Create(FCatalog.GetCollection(COLLECTION_NAME_APPS) as ICatalogCollection);
  if not AServer.IsEmpty then
    FCatalog.Connect(AServer);
  GetApplications;
end;

destructor TComAdminCatalog.Destroy;
begin
  FApplications.Free;
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
