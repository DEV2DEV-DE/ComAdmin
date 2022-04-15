unit uComAdmin;

// Compile with typeinfo for RTTI support
{$TYPEINFO ON}

// https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections

interface

uses
  System.Generics.Collections,
  System.SysUtils,
  COMAdmin_TLB;

type
  TCOMAdminAccessChecksLevelOptions = (COMAdminAccessChecksApplicationLevel, COMAdminAccessChecksApplicationComponentLevel);
  TCOMAdminApplicationActivation = (COMAdminActivationInproc, COMAdminActivationLocal);
  TCOMAdminAuthenticationCapability = (COMAdminAuthenticationCapabilitiesNone = $0, COMAdminAuthenticationCapabilitiesSecureReference = $2,
                                       COMAdminAuthenticationCapabilitiesStaticCloaking = $20, COMAdminAuthenticationCapabilitiesDynamicCloaking = $40);
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
  TCOMAdminProtocol = (COMAdminProtocolTCP, COMAdminProtocolHTTP, COMAdminProtocolSPX);

  TComAdminReadEvent = procedure (const AObjectType, AObjectName: string) of object;
  TComAdminDebug = procedure(const AMessage: string) of object;

const
  COM_ADMIN_PROTOCOLS: array[COMAdminProtocolTCP..COMAdminProtocolSPX] of string = ('ncacn_ip_tcp', 'ncacn_http', 'ncacn_spx');

  DEFAULT_CREATION_TIMEOUT = 60000;
  DEFAULT_MAX_DUMP = 5;
  DEFAULT_MAX_POOL = 1048576;
  DEFAULT_RECYCLE_TIMEOUT = 15;
  DEFAULT_SHUTDOWN = 3;
  DEFAULT_TRANSACTION_TIMEOUT = 60;

  MAX_DUMP_COUNT = 200;
  MAX_LIFETIME_LIMIT = 30240;
  MAX_POOL_SIZE = DEFAULT_MAX_POOL;
  MAX_RECYCLE_TIMEOUT = 1440;
  MAX_THREADS = 1000;
  MAX_TIMEOUT = 3600;

  IID_IUserCollection: TGUID = '{C29ADAEE-CB81-4D36-BEDF-9F131094D9A5}';

type
  // Interface for objects that have a users collection
  IUserCollection = Interface(IInterface)
    ['{C29ADAEE-CB81-4D36-BEDF-9F131094D9A5}']
    function GetUsersCollectionName: string;
  end;

  // forward declaration of internally used classes
  TComAdminBaseList = class;
  TComAdminCatalog = class;

  // generic base class for all single objects
  TComAdminBaseObject = class(TInterfacedObject)
  strict private
    FCatalogCollection: ICatalogCollection;
    FCatalogObject: ICatalogObject;
    FCollection: TComAdminBaseList;
    FKey: string;
    FName: string;
    FSupportsUsers: Boolean;
  private
    function InternalCheckRange(AMinValue, AMaxValue, AValue: Cardinal): Boolean;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    procedure CopyProperties(ASourceObject, ATargetObject: TComAdminBaseObject);
    property CatalogCollection: ICatalogCollection read FCatalogCollection;
    property CatalogObject: ICatalogObject read FCatalogObject;
    property Collection: TComAdminBaseList read FCollection;
    property Key: string read FKey write FKey;
    property SupportsUsers: Boolean read FSupportsUsers write FSupportsUsers default False;
  published
    property Name: string read FName write FName;
  end;

  // generic list class for all collections of objects
  TComAdminBaseList = class(TObjectList<TComAdminBaseObject>)
  strict private
    FCatalog: TComAdminCatalog;
    FCatalogCollection: ICatalogCollection;
    FName: string;
    FOwner: TComAdminBaseObject;
    function GetIndexByKey(const AKey: string): Integer;
  public
    constructor Create(AOwner: TComAdminBaseObject; ACatalog: TComAdminCatalog; ACatalogCollection: ICatalogCollection); reintroduce;
    function Contains(const AItemName: string): Boolean;
    function Delete(Index: Integer): Integer;
    function SaveChanges: Integer;
    property Catalog: TComAdminCatalog read FCatalog write FCatalog;
    property CatalogCollection: ICatalogCollection read FCatalogCollection write FCatalogCollection;
    property Name: string read FName;
    property Owner: TComAdminBaseObject read FOwner;
  end;

  TComAdminUser = class(TComAdminBaseObject)
  end;

  TComAdminUserList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminUser;
  public
    function Append(ASourceUser: TComAdminUser): TComAdminUser;
    function Find(const AName: string; out AUser: TComAdminUser): Boolean;
    property Items[Index: Integer]: TComAdminUser read GetItem; default;
  end;

  TComAdminRole = class(TComAdminBaseObject)
  strict private
    FDescription: string;
    FUsers: TComAdminUserList;
    procedure GetUsers;
    procedure SetDescription(const Value: string);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function CopyProperties(ASourceRole: TComAdminRole): Integer;
    procedure SyncUsers(ASourceRole: TComAdminRole);
    property Users: TComAdminUserList read FUsers write FUsers;
  published
    property Description: string read FDescription write SetDescription;
  end;

  TComAdminRoleList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminRole;
  public
    function Append(ASourceRole: TComAdminRole): TComAdminRole;
    function Find(const AName: string; out ARole: TComAdminRole): Boolean;
    property Items[Index: Integer]: TComAdminRole read GetItem; default;
  end;

  TComAdminInstance = class(TComAdminBaseObject)
  strict private
    FProcessID: Cardinal;
    FHasRecycled: Boolean;
    FIsPaused: Boolean;
    procedure ReadExtendedProperties;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
  published
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

  TComAdminPartition = class(TComAdminBaseObject, IUserCollection)
  strict private
    FDescription: string;
    FChangeable: Boolean;
    FDeleteable: Boolean;
    FID: string;
    FRoles: TComAdminRoleList;
    procedure GetRoles;
    procedure ReadExtendedProperties;
    procedure SetChangeable(const Value: Boolean);
    procedure SetDeleteable(const Value: Boolean);
    procedure SetDescription(const Value: string);
    procedure SetID(const Value: string);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function GetUsersCollectionName: string;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property Changeable: Boolean read FChangeable write SetChangeable default True;
    property Deleteable: Boolean read FDeleteable write SetDeleteable default True;
    property Description: string read FDescription write SetDescription;
    property ID: string read FID write SetID;
  end;

  TComAdminPartitionList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminPartition;
  public
    property Items[Index: Integer]: TComAdminPartition read GetItem; default;
  end;

  TComAdminMethod = class(TComAdminBaseObject)
  strict private
    FAutoComplete: Boolean;
    FCLSID: string;
    FDescription: string;
    FIID: string;
    FIndex: Cardinal;
    FRoles: TComAdminRoleList;
    procedure GetRoles;
    procedure ReadExtendedProperties;
    procedure SetAutoComplete(const Value: Boolean);
    procedure SetDescription(const Value: string);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property AutoComplete: Boolean read FAutoComplete write SetAutoComplete default False;
    property Description: string read FDescription write SetDescription;
    property CLSID: string read FCLSID;
    property IID: string read FIID;
    property Index: Cardinal read FIndex;
  end;

  TComAdminMethodList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminMethod;
  public
    property Items[Index: Integer]: TComAdminMethod read GetItem; default;
  end;

  TComAdminInterface = class(TComAdminBaseObject)
  strict private
    FCLSID: string;
    FDescription: string;
    FIID: string;
    FRoles: TComAdminRoleList;
    FMethods: TComAdminMethodList;
    FQueuingEnabled: Boolean;
    FQueuingSupported: Boolean;
    procedure GetRoles;
    procedure GetMethods;
    procedure ReadExtendedProperties;
    procedure SetQueuingEnabled(const Value: Boolean);
    procedure SetFDescription(const Value: string);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Methods: TComAdminMethodList read FMethods write FMethods;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property CLSID: string read FCLSID;
    property Description: string read FDescription write SetFDescription;
    property IID: string read FIID;
    property QueuingEnabled: Boolean read FQueuingEnabled write SetQueuingEnabled;
    property QueuingSupported: Boolean read FQueuingSupported;
  end;

  TComAdminInterfaceList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminInterface;
  public
    property Items[Index: Integer]: TComAdminInterface read GetItem; default;
  end;

  TCOMAdminComponent = class(TComAdminBaseObject)
  strict private
    FAllowInprocSubscribers: Boolean;
    FApplicationID: string;
    FBitness: TCOMAdminComponentType;
    FComponentAccessChecksEnabled: Boolean;
    FComponentTransactionTimeout: Cardinal;
    FComponentTransactionTimeoutEnabled: Boolean;
    FCOMTIIntrinsics: Boolean;
    FConstructionEnabled: Boolean;
    FConstructorString: string;
    FCreationTimeout: Cardinal;
    FDescription: string;
    FDll: string;
    FEventTrackingEnabled: Boolean;
    FExceptionClass: string;
    FFireInParallel: Boolean;
    FIISIntrinsics: Boolean;
    FInitializeServerApplication: Boolean;
    FIsEnabled: Boolean;
    FIsInstalled: Boolean;
    FJustInTimeActivation: Boolean;
    FLoadBalancingSupported: Boolean;
    FInterfaces: TComAdminInterfaceList;
    FIsEventClass: Boolean;
    FIsPrivateComponent: Boolean;
    FMinPoolSize: Cardinal;
    FMaxPoolSize: Cardinal;
    FMultiInterfacePublisherFilterCLSID: string;
    FMustRunInDefaultContext: Boolean;
    FMustRunInClientContext: Boolean;
    FObjectPoolingEnabled: Boolean;
    FProgID: string;
    FRoles: TComAdminRoleList;
    procedure GetInterfaces;
    procedure GetRoles;
    procedure ReadExtendedProperties;
    procedure SetAllowInprocSubscribers(const Value: Boolean);
    procedure SetApplicationID(const Value: string);
    procedure SetBitness(const Value: TCOMAdminComponentType);
    procedure SetComponentAccessChecksEnabled(const Value: Boolean);
    procedure SetComponentTransactionTimeout(const Value: Cardinal);
    procedure SetComponentTransactionTimeoutEnabled(const Value: Boolean);
    procedure SetCOMTIIntrinsics(const Value: Boolean);
    procedure SetConstructionEnabled(const Value: Boolean);
    procedure SetConstructorString(const Value: string);
    procedure SetCreationTimeout(const Value: Cardinal);
    procedure SetDescription(const Value: string);
    procedure SetDll(const Value: string);
    procedure SetEventTrackingEnabled(const Value: Boolean);
    procedure SetExceptionClass(const Value: string);
    procedure SetFireInParallel(const Value: Boolean);
    procedure SetIISIntrinsics(const Value: Boolean);
    procedure SetInitializeServerApplication(const Value: Boolean);
    procedure SetIsEnabled(const Value: Boolean);
    procedure SetIsEventClass(const Value: Boolean);
    procedure SetIsInstalled(const Value: Boolean);
    procedure SetIsPrivateComponent(const Value: Boolean);
    procedure SetJustInTimeActivation(const Value: Boolean);
    procedure SetLoadBalancingSupported(const Value: Boolean);
    procedure SetMaxPoolSize(const Value: Cardinal);
    procedure SetMinPoolSize(const Value: Cardinal);
    procedure SetMultiInterfacePublisherFilterCLSID(const Value: string);
    procedure SetMustRunInClientContext(const Value: Boolean);
    procedure SetMustRunInDefaultContext(const Value: Boolean);
    procedure SetObjectPoolingEnabled(const Value: Boolean);
    procedure SetProgID(const Value: string);
    procedure SyncRoles(ASourceComponent: TCOMAdminComponent);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function CopyProperties(ASourceComponent: TCOMAdminComponent): Integer;
    property Roles: TComAdminRoleList read FRoles write FRoles;
    property Interfaces: TComAdminInterfaceList read FInterfaces write FInterfaces;
  published
    property AllowInprocSubscribers: Boolean read FAllowInprocSubscribers write SetAllowInprocSubscribers default True;
    property ApplicationID: string read FApplicationID write SetApplicationID;
    property Bitness: TCOMAdminComponentType read FBitness write SetBitness;
    property ComponentAccessChecksEnabled: Boolean read FComponentAccessChecksEnabled write SetComponentAccessChecksEnabled default False;
    property ComponentTransactionTimeout: Cardinal read FComponentTransactionTimeout write SetComponentTransactionTimeout default DEFAULT_TRANSACTION_TIMEOUT;
    property ComponentTransactionTimeoutEnabled: Boolean read FComponentTransactionTimeoutEnabled write SetComponentTransactionTimeoutEnabled default False;
    property COMTIIntrinsics: Boolean read FCOMTIIntrinsics write SetCOMTIIntrinsics default False;
    property ConstructionEnabled: Boolean read FConstructionEnabled write SetConstructionEnabled default False;
    property ConstructorString: string read FConstructorString write SetConstructorString;
    property CreationTimeout: Cardinal read FCreationTimeout write SetCreationTimeout default DEFAULT_CREATION_TIMEOUT;
    property Description: string read FDescription write SetDescription;
    property Dll: string read FDll write SetDll;
    property EventTrackingEnabled: Boolean read FEventTrackingEnabled write SetEventTrackingEnabled default True;
    property ExceptionClass: string read FExceptionClass write SetExceptionClass;
    property FireInParallel: Boolean read FFireInParallel write SetFireInParallel default False;
    property IISIntrinsics: Boolean read FIISIntrinsics write SetIISIntrinsics default False;
    property InitializeServerApplication: Boolean read FInitializeServerApplication write SetInitializeServerApplication default False;
    property IsEnabled: Boolean read FIsEnabled write SetIsEnabled default True;
    property IsEventClass: Boolean read FIsEventClass write SetIsEventClass default False;
    property IsInstalled: Boolean read FIsInstalled write SetIsInstalled default False;
    property IsPrivateComponent: Boolean read FIsPrivateComponent write SetIsPrivateComponent default False;
    property JustInTimeActivation: Boolean read FJustInTimeActivation write SetJustInTimeActivation default False;
    property LoadBalancingSupported: Boolean read FLoadBalancingSupported write SetLoadBalancingSupported default False;
    property MaxPoolSize: Cardinal read FMaxPoolSize write SetMaxPoolSize default DEFAULT_MAX_POOL;
    property MinPoolSize: Cardinal read FMinPoolSize write SetMinPoolSize default 0;
    property MultiInterfacePublisherFilterCLSID: string read FMultiInterfacePublisherFilterCLSID write SetMultiInterfacePublisherFilterCLSID;
    property MustRunInClientContext: Boolean read FMustRunInClientContext write SetMustRunInClientContext default False;
    property MustRunInDefaultContext: Boolean read FMustRunInDefaultContext write SetMustRunInDefaultContext default False;
    property ObjectPoolingEnabled: Boolean read FObjectPoolingEnabled write SetObjectPoolingEnabled default False;
    property ProgID: string read FProgID write SetProgID;
  end;

  TCOMAdminComponentList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TCOMAdminComponent;
    function BuildTargetLibraryName(ASourceComponent: TCOMAdminComponent): string;
  public
    function Append(ASourceComponent: TCOMAdminComponent): TCOMAdminComponent;
    function CopyLibrary(ASourceComponent: TCOMAdminComponent; AOverwrite: Boolean = False): Boolean;
    function Find(const AName: string; out AComponent: TCOMAdminComponent): Boolean;
    property Items[Index: Integer]: TCOMAdminComponent read GetItem; default;
  end;

  TComAdminApplication = class(TComAdminBaseObject, IUserCollection)
  strict private
    FAccessChecksEnabled: Boolean;
    FAccessChecksLevel: TCOMAdminAccessChecksLevelOptions;
    FActivation: TCOMAdminApplicationActivation;
    FAuthenticationCapability: TCOMAdminAuthenticationCapability;
    FAuthenticationLevel: TCOMAdminAuthenticationLevel;
    FChangeable: Boolean;
    FCommandLine: string;
    FComponents: TCOMAdminComponentList;
    FConcurrentApps: Cardinal;
    FCreatedBy: string;
    FCRMEnabled: Boolean;
    FCRMLogFile: string;
    FDeleteable: Boolean;
    FDescription: string;
    FDirectory: string;
    FDumpEnabled: Boolean;
    FDumpOnException: Boolean;
    FDumpOnFailFast: Boolean;
    FDumpPath: string;
    FEventsEnabled: Boolean;
    FGig3SupportEnabled: Boolean;
    FIdentity: string;
    FImpersonationLevel: TCOMAdminImpersonationLevel;
    FInstances: TComAdminInstanceList;
    FIsEnabled: Boolean;
    FIsSystem: Boolean;
    FMaxDumpCount: Cardinal;
    FPartitionID: string;
    FProxy: Boolean;
    FProxyServerName: string;
    FQCAuthenticateMsgs: TCOMAdminQCAuthenticateMsgs;
    FQCListenerMaxThreads: Cardinal;
    FQueueListenerEnabled: Boolean;
    FQueuingEnabled: Boolean;
    FRecycleActivationLimit: Cardinal;
    FRecycleCallLimit: Cardinal;
    FRecycleExpirationTimeout: Cardinal;
    FRecycleLifetimeLimit: Cardinal;
    FRecycleMemoryLimit: Cardinal;
    FReplicable: Boolean;
    FRoles: TComAdminRoleList;
    FRunForever: Boolean;
    FServiceName: string;
    FShutdownAfter: Cardinal;
    FSoapActivated: Boolean;
    FSoapBaseUrl: string;
    FSoapMailTo: string;
    FSoapVRoot: string;
    FSRPEnabled: Boolean;
    FSRPTrustLevel: TCOMAdminSRPTrustLevel;
    function BuildInstallFileName: string;
    procedure GetComponents;
    function GetInstances: TComAdminInstanceList;
    procedure GetRoles;
    function GetUsersCollectionName: string;
    procedure ReadExtendedProperties;
    procedure SetAccessChecksEnabled(const Value: Boolean);
    procedure SetAccessChecksLevel(const Value: TCOMAdminAccessChecksLevelOptions);
    procedure SetActivation(const Value: TCOMAdminApplicationActivation);
    procedure SetAuthenticationCapability(const Value: TCOMAdminAuthenticationCapability);
    procedure SetAuthenticationLevel(const Value: TCOMAdminAuthenticationLevel);
    procedure SetChangeable(const Value: Boolean);
    procedure SetCommandLine(const Value: string);
    procedure SetConcurrentApps(const Value: Cardinal);
    procedure SetCreatedBy(const Value: string);
    procedure SetCRMEnabled(const Value: Boolean);
    procedure SetCRMLogFile(const Value: string);
    procedure SetDeleteable(const Value: Boolean);
    procedure SetDescription(const Value: string);
    procedure SetDirectory(const Value: string);
    procedure SetDumpEnabled(const Value: Boolean);
    procedure SetDumpOnException(const Value: Boolean);
    procedure SetDumpOnFailFast(const Value: Boolean);
    procedure SetDumpPath(const Value: string);
    procedure SetEventsEnabled(const Value: Boolean);
    procedure SetGig3SupportEnabled(const Value: Boolean);
    procedure SetIdentity(const Value: string);
    procedure SetImpersonationLevel(const Value: TCOMAdminImpersonationLevel);
    procedure SetIsEnabled(const Value: Boolean);
    procedure SetMaxDumpCount(const Value: Cardinal);
    procedure SetPartitionID(const Value: string);
    procedure SetPassword(const Value: string);
    procedure SetProxy(const Value: Boolean);
    procedure SetProxyServerName(const Value: string);
    procedure SetQCAuthenticateMsgs(const Value: TCOMAdminQCAuthenticateMsgs);
    procedure SetQCListenerMaxThreads(const Value: Cardinal);
    procedure SetQueueListenerEnabled(const Value: Boolean);
    procedure SetQueuingEnabled(const Value: Boolean);
    procedure SetRecycleActivationLimit(const Value: Cardinal);
    procedure SetRecycleCallLimit(const Value: Cardinal);
    procedure SetRecycleExpirationTimeout(const Value: Cardinal);
    procedure SetRecycleLifetimeLimit(const Value: Cardinal);
    procedure SetRecycleMemoryLimit(const Value: Cardinal);
    procedure SetReplicable(const Value: Boolean);
    procedure SetRunForever(const Value: Boolean);
    procedure SetServiceName(const Value: string);
    procedure SetShutdownAfter(const Value: Cardinal);
    procedure SetSoapActivated(const Value: Boolean);
    procedure SetSoapBaseUrl(const Value: string);
    procedure SetSoapMailTo(const Value: string);
    procedure SetSoapVRoot(const Value: string);
    procedure SetSRPEnabled(const Value: Boolean);
    procedure SetSRPTrustLevel(const Value: TCOMAdminSRPTrustLevel);
    procedure SyncComponents(ASourceApplication: TCOMAdminApplication);
    procedure SyncRoles(ASourceApplication: TCOMAdminApplication);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function CopyProperties(ASourceApplication: TCOMAdminApplication; const APassword: string): Integer;
    function CopyToServer(ATargetServer: TComAdminCatalog; AOptions: Integer): Boolean;
    function InstallComponent(const ALibraryName: string): TCOMAdminComponent;
    function IsIntegratedIdentity: Boolean;
    property Components: TCOMAdminComponentList read FComponents;
    property Instances: TComAdminInstanceList read GetInstances;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property AccessChecksEnabled: Boolean read FAccessChecksEnabled write SetAccessChecksEnabled default True;
    property AccessChecksLevel: TCOMAdminAccessChecksLevelOptions read FAccessChecksLevel write SetAccessChecksLevel default COMAdminAccessChecksApplicationComponentLevel;
    property Activation: TCOMAdminApplicationActivation read FActivation write SetActivation default COMAdminActivationLocal;
    property AuthenticationCapability: TCOMAdminAuthenticationCapability read FAuthenticationCapability write SetAuthenticationCapability default COMAdminAuthenticationCapabilitiesDynamicCloaking;
    property AuthenticationLevel: TCOMAdminAuthenticationLevel read FAuthenticationLevel write SetAuthenticationLevel default COMAdminAuthenticationDefault;
    property Changeable: Boolean read FChangeable write SetChangeable default True;
    property CommandLine: string read FCommandLine write SetCommandLine;
    property ConcurrentApps: Cardinal read FConcurrentApps write SetConcurrentApps default 1;
    property CreatedBy: string read FCreatedBy write SetCreatedBy;
    property CRMEnabled: Boolean read FCRMEnabled write SetCRMEnabled default False;
    property CRMLogFile: string read FCRMLogFile write SetCRMLogFile;
    property Deleteable: Boolean read FDeleteable write SetDeleteable default True;
    property Description: string read FDescription write SetDescription;
    property Directory: string read FDirectory write SetDirectory;
    property DumpEnabled: Boolean read FDumpEnabled write SetDumpEnabled default False;
    property DumpOnException: Boolean read FDumpOnException write SetDumpOnException default False;
    property DumpOnFailFast: Boolean read FDumpOnFailFast write SetDumpOnFailFast default False;
    property DumpPath: string read FDumpPath write SetDumpPath;
    property EventsEnabled: Boolean read FEventsEnabled write SetEventsEnabled default True;
    property Gig3SupportEnabled: Boolean read FGig3SupportEnabled write SetGig3SupportEnabled default False;
    property Identity: string read FIdentity write SetIdentity;
    property ImpersonationLevel: TCOMAdminImpersonationLevel read FImpersonationLevel write SetImpersonationLevel default COMAdminImpersonationImpersonate;
    property IsEnabled: Boolean read FIsEnabled write SetIsEnabled default True;
    property IsSystem: Boolean read FIsSystem default False;
    property MaxDumpCount: Cardinal read FMaxDumpCount write SetMaxDumpCount default DEFAULT_MAX_DUMP;
    property PartitionID: string read FPartitionID write SetPartitionID;
    property Password: string write SetPassword;
    property Proxy: Boolean read FProxy write SetProxy default False;
    property ProxyServerName: string read FProxyServerName write SetProxyServerName;
    property QCAuthenticateMsgs: TCOMAdminQCAuthenticateMsgs read FQCAuthenticateMsgs write SetQCAuthenticateMsgs default COMAdminQCMessageAuthenticateSecureApps;
    property QCListenerMaxThreads: Cardinal read FQCListenerMaxThreads write SetQCListenerMaxThreads default 0;
    property QueueListenerEnabled: Boolean read FQueueListenerEnabled write SetQueueListenerEnabled default False;
    property QueuingEnabled: Boolean read FQueuingEnabled write SetQueuingEnabled default False;
    property RecycleActivationLimit: Cardinal read FRecycleActivationLimit write SetRecycleActivationLimit default 0;
    property RecycleCallLimit: Cardinal read FRecycleCallLimit write SetRecycleCallLimit default 0;
    property RecycleExpirationTimeout: Cardinal read FRecycleExpirationTimeout write SetRecycleExpirationTimeout default DEFAULT_RECYCLE_TIMEOUT;
    property RecycleLifetimeLimit: Cardinal read FRecycleLifetimeLimit write SetRecycleLifetimeLimit default 0;
    property RecycleMemoryLimit: Cardinal read FRecycleMemoryLimit write SetRecycleMemoryLimit default 0;
    property Replicable: Boolean read FReplicable write SetReplicable default True;
    property RunForever: Boolean read FRunForever write SetRunForever default False;
    property ServiceName: string read FServiceName write SetServiceName;
    property ShutdownAfter: Cardinal read FShutdownAfter write SetShutdownAfter default DEFAULT_SHUTDOWN;
    property SoapActivated: Boolean read FSoapActivated write SetSoapActivated default False;
    property SoapBaseUrl: string read FSoapBaseUrl write SetSoapBaseUrl;
    property SoapMailTo: string read FSoapMailTo write SetSoapMailTo;
    property SoapVRoot: string read FSoapVRoot write SetSoapVRoot;
    property SRPEnabled: Boolean read FSRPEnabled write SetSRPEnabled default False;
    property SRPTrustLevel: TCOMAdminSRPTrustLevel read FSRPTrustLevel write SetSRPTrustLevel default COMAminSRPFullyTrusted;
  end;

  TComAdminApplicationList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminApplication;
  public
    function Append(ASourceApplication: TComAdminApplication; const ACreatorString: string = ''; const APassword: string = ''): TComAdminApplication;
    function Find(const AName: string; out AApplication: TComAdminApplication): Boolean;
    property Items[Index: Integer]: TComAdminApplication read GetItem; default;
  end;

  TComAdminComputer = class(TComAdminBaseObject)
  strict private
    FApplicationProxyRSN: string;
    FAuthenticationLevel: TCOMAdminAuthenticationLevel;
    FCISEnabled: Boolean;
    FDCOMEnabled: Boolean;
    FDefaultToInternetPorts: Boolean;
    FDescription: string;
    FDSPartitionLookupEnabled: Boolean;
    FImpersonationLevel: TCOMAdminImpersonationLevel;
    FInternetPortsListed: Boolean;
    FIsRouter: Boolean;
    FLoadBalancingCLSID: string;
    FLocalPartitionLookupEnabled: Boolean;
    FOperatingSystem: TCOMAdminOperatingSystem;
    FPartitionsEnabled: Boolean;
    FPorts: string;
    FResourcePoolingEnabled: Boolean;
    FRPCProxyEnabled: Boolean;
    FSecureReferencesEnabled: Boolean;
    FSecurityTrackingEnabled: Boolean;
    FSRPActivateAsActivatorChecks: Boolean;
    FSRPRunningObjectChecks: Boolean;
    FTransactionTimeout: Cardinal;
    procedure ReadExtendedProperties;
    procedure SetTransactionTimeout(const Value: Cardinal);
  public
    constructor Create(ACatalogCollection: ICatalogCollection); reintroduce;
  published
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
    property TransactionTimeout: Cardinal read FTransactionTimeout write SetTransactionTimeout default DEFAULT_TRANSACTION_TIMEOUT;
  end;

  TComAdminApplicationServer = class(TComAdminBaseObject);

  TComAdminApplicationCluster = class(TComAdminBaseList);

  TComAdminDCOMProtocol = class(TComAdminBaseObject)
  strict private
    FOrder: Cardinal;
    FProtocolCode: TCOMAdminProtocol;
    procedure ReadExtendedProperties;
    procedure SetOrder(const Value: Cardinal);
    procedure SetProtocolCode(const Value: TCOMAdminProtocol);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
  published
    property Order: Cardinal read FOrder write SetOrder default 0;
    property ProtocolCode: TCOMAdminProtocol read FProtocolCode write SetProtocolCode;
  end;

  TComAdminDCOMProtocolList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminDCOMProtocol;
  public
    property Items[Index: Integer]: TComAdminDCOMProtocol read GetItem; default;
  end;

  TComAdminEventClass = class(TComAdminBaseObject)
  strict private
    FBitness: TCOMAdminComponentType;
    FApplication: string;
    FProgID: string;
    FDescription: string;
    FCLSID: string;
    FIsPrivateComponent: Boolean;
    procedure ReadExtendedProperties;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
  published
    property Application: string read FApplication;
    property Bitness: TCOMAdminComponentType read FBitness;
    property CLSID: string read FCLSID;
    property Description: string read FDescription;
    property IsPrivateComponent: Boolean read FIsPrivateComponent default False;
    property ProgID: string read FProgID;
  end;

  TComAdminEventClassList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TComAdminEventClass;
  public
    property Items[Index: Integer]: TComAdminEventClass read GetItem; default;
  end;

  // List class to hold passwords for different identities
  TPasswordList = TList<TPair<string,string>>;

  TComAdminCatalog = class(TObject)
  strict private
    FApplicationCluster: TComAdminApplicationCluster;
    FApplications: TComAdminApplicationList;
    FCatalog: ICOMAdminCatalog2;
    FChangeCount: Integer;
    FComputer: TComAdminComputer;
    FCopiedLibraries: TList<string>;
    FCopyLibraries: Boolean;
    FDebug: Boolean;
    FEventClasses: TComAdminEventClassList;
    FFilter: string;
    FLibraryPath: string;
    FOnDebug: TComAdminDebug;
    FOnReadObject: TComAdminReadEvent;
    FPasswords: TPasswordList;
    FPartitions: TComAdminPartitionList;
    FProtocols: TComAdminDCOMProtocolList;
    FServer: string;
    procedure GetApplications;
    procedure GetEventClasses;
    procedure GetPartitions;
    procedure GetProtocols;
    procedure GetServers;
    procedure SetFilter(const Value: string);
  public
    constructor Create(const AServer: string; const AFilter: string; AOnReadEvent: TComAdminReadEvent); reintroduce; overload;
    destructor Destroy; override;
    procedure AddPasswordForIdentity(const AIdentity, APassword: string);
    procedure DebugMessage(const AMessage: string); overload;
    procedure DebugMessage(const AMessage: string; AParams: array of const); overload;
    procedure ExportApplication(AIndex: Integer; const AFilename: string);
    procedure ExportApplicationByKey(const AKey, AFilename: string);
    procedure ExportApplicationByName(const AName, AFilename: string);
    function GetPasswordForIdentity(const AIdentity: string): string;
    function SyncToServer(const ATargetServer, ACreatorString: string): Integer; overload;
    property ApplicationCluster: TComAdminApplicationCluster read FApplicationCluster write FApplicationCluster;
    property Applications: TComAdminApplicationList read FApplications;
    property Catalog: ICOMAdminCatalog2 read FCatalog;
    property ChangeCount: Integer read FChangeCount write FChangeCount;
    property Computer: TComAdminComputer read FComputer write FComputer;
    property CopiedLibraries: TList<string> read FCopiedLibraries write FCopiedLibraries;
    property CopyLibraries: Boolean read FCopyLibraries write FCopyLibraries;
    property DCOMProtocols: TComAdminDCOMProtocolList read FProtocols;
    property Debug: Boolean read FDebug write FDebug default False;
    property EventClasses: TComAdminEventClassList read FEventClasses;
    property Filter: string read FFilter write SetFilter;
    property LibraryPath: string read FLibraryPath write FLibraryPath;
    property OnDebug: TComAdminDebug read FOnDebug write FOnDebug;
    property OnReadObject: TComAdminReadEvent read FOnReadObject write FOnReadObject;
    property Server: string read FServer;
  end;

  EItemNotFoundException = Exception;
  EUnknownProtocolException = Exception;

implementation

uses
  System.Masks,
  System.Variants,
  Winapi.Windows, System.Rtti, System.TypInfo, System.IOUtils;

{$I uComAdminConst.inc}

{ TComAdminBaseObject }

constructor TComAdminBaseObject.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create;
  FCollection := ACollection;
  FCatalogObject := ACatalogObject;
  if Assigned(ACollection) then
  begin
    FCatalogCollection := ACollection.CatalogCollection;
    if Assigned(ACollection.Catalog.OnReadObject) then
      ACollection.Catalog.OnReadObject(ACollection.Name, ACatalogObject.Name);
  end;
  FKey := FCatalogObject.Key;
  FName := FCatalogObject.Name;
end;

procedure TComAdminBaseObject.CopyProperties(ASourceObject, ATargetObject: TComAdminBaseObject);
var
  LRttiContext: TRttiContext;
  LType: TRttiType;
  LProperty: TRttiProperty;
  AValue, ASource, ATarget: TValue;
begin
  LRttiContext := TRttiContext.Create;
  try
    LType := LRttiContext.GetType(ASourceObject.ClassInfo);
    ASource := TValue.From<TComAdminBaseObject>(ASourceObject);
    ATarget := TValue.From<TComAdminBaseObject>(ATargetObject);

    for LProperty in LType.GetProperties do
    begin
      if (LProperty.IsReadable) and (LProperty.IsWritable) and (LProperty.Visibility = mvPublished) then
      begin
        { TODO -odev2dev : also copy the properties from the source catalog object to the target }
        AValue := LProperty.GetValue(ASource.AsObject);
        FCollection.Catalog.DebugMessage(DEBUG_MESSAGE_SET_PROPERTY, [LProperty.Name, AValue.ToString]);
        LProperty.SetValue(ATarget.AsObject, AValue);
      end;
    end;

  finally
    LRttiContext.Free;
  end;
end;

function TComAdminBaseObject.InternalCheckRange(AMinValue, AMaxValue, AValue: Cardinal): Boolean;
begin
  Result := (AValue >= AMinValue) and (AValue <= AMaxValue);
  if not Result then
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

{ TComAdminBaseList }

function TComAdminBaseList.Contains(const AItemName: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Count - 1 do
    if Items[i].Name.Equals(AItemName) then
      Exit(True);
end;

constructor TComAdminBaseList.Create(AOwner: TComAdminBaseObject; ACatalog: TComAdminCatalog; ACatalogCollection: ICatalogCollection);
begin
  inherited Create(True);
  if Assigned(ACatalogCollection) then
  begin
    FOwner := AOwner;
    FCatalog := ACatalog;
    FCatalogCollection := ACatalogCollection;
    FName := FCatalogCollection.Name;
    FCatalogCollection.Populate;
  end;
end;

function TComAdminBaseList.Delete(Index: Integer): Integer;
begin
  CatalogCollection.Remove(GetIndexByKey(Items[Index].Key));
  Result := SaveChanges;
  inherited Delete(Index);
end;

function TComAdminBaseList.GetIndexByKey(const AKey: string): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to CatalogCollection.Count - 1 do
  begin
    if AKey.Equals((CatalogCollection.Item[i] as ICatalogObject).Key) then
      Exit(i);
  end;
end;

function TComAdminBaseList.SaveChanges: Integer;
begin
  Result := 0;
  try
    if Assigned(FCatalogCollection) then
      Result := FCatalogCollection.SaveChanges;
  except
    on E:Exception do
      FCatalog.DebugMessage('Error in %s.TComAdminBaseList.SaveChanges: ', [Self.Name, E.Message]);
  end;
end;

{ TComAdminUser }

{ TComAdminUserList }

function TComAdminUserList.Append(ASourceUser: TComAdminUser): TComAdminUser;
var
  LUser: ICatalogObject;
begin
  LUser := CatalogCollection.Add as ICatalogObject;
  LUser.Value[PROPERTY_NAME_USER] := ASourceUser.Name;
  Result := TComAdminUser.Create(Self, LUser);
  Result.Name := ASourceUser.Name;
  Catalog.ChangeCount := Catalog.ChangeCount + SaveChanges;
  Self.Add(Result);
end;

function TComAdminUserList.Find(const AName: string; out AUser: TComAdminUser): Boolean;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if Items[i].Name.Equals(AName) then
    begin
      AUser := Items[i];
      Exit(True);
    end;
  end;
  Result := False;
end;

function TComAdminUserList.GetItem(Index: Integer): TComAdminUser;
begin
  Result := inherited Items[Index] as TComAdminUser;
end;

{ TComAdminRole }

constructor TComAdminRole.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
var
  LObject: IUserCollection;
begin
  inherited Create(ACollection, ACatalogObject);


  if (ACollection.Owner.QueryInterface(IID_IUserCollection, LObject) = S_OK) then
  begin
    FDescription := VarToStrDef(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION], '');
    FUsers := TComAdminUserList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(LObject.GetUsersCollectionName, Key) as ICatalogCollection);
    GetUsers;
  end;

end;

destructor TComAdminRole.Destroy;
begin
  if Assigned(FUsers) then
    FUsers.Free;
  inherited;
end;

function TComAdminRole.CopyProperties(ASourceRole: TComAdminRole): Integer;
begin
  Result := 0;
  if Collection.Owner.SupportsUsers then
  try
    inherited CopyProperties(ASourceRole, Self);

    // Changes must be saved before any sub-collections can be updated
    Result := Collection.SaveChanges;

    // Synchronize users from source role
    SyncUsers(ASourceRole);
  except
    on E:Exception do
      Collection.Catalog.DebugMessage('TComAdminRole.CopyProperties: ', [E.Message]);
  end;
end;

procedure TComAdminRole.GetUsers;
var
  i: Integer;
begin
  for i := 0 to FUsers.CatalogCollection.Count - 1 do
    FUsers.Add(TComAdminUser.Create(FUsers, FUsers.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminRole.SetDescription(const Value: string);
begin
  FDescription := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DESCRIPTION) then
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := Value;
end;

procedure TComAdminRole.SyncUsers(ASourceRole: TComAdminRole);
var
  i: Integer;
begin
  for i := 0 to ASourceRole.Users.Count - 1 do
  begin
    Collection.Catalog.DebugMessage(DEBUG_MESSAGE_SYNC_USER, [ASourceRole.Users[i].Name]);
    if not FUsers.Contains(ASourceRole.Users[i].Name) then
      FUsers.Append(ASourceRole.Users[i]); // User does not exists in target role ==> create & copy
  end;
end;

{ TComAdminRoleList }

function TComAdminRoleList.Append(ASourceRole: TComAdminRole): TComAdminRole;
var
  LRole: ICatalogObject;
begin
  LRole := CatalogCollection.Add as ICatalogObject;
  LRole.Value[PROPERTY_NAME_NAME] := ASourceRole.Name;
  Result := TComAdminRole.Create(Self, LRole);
  Result.CopyProperties(ASourceRole);
  Assert(Assigned(Catalog), 'Invalid catalog');
  Catalog.ChangeCount := Catalog.ChangeCount + SaveChanges;
  Self.Add(Result);
end;

function TComAdminRoleList.Find(const AName: string; out ARole: TComAdminRole): Boolean;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if Items[i].Name.Equals(AName) then
    begin
      ARole := Items[i];
      Exit(True);
    end;
  end;
  Result := False;
end;

function TComAdminRoleList.GetItem(Index: Integer): TComAdminRole;
begin
  Result := inherited Items[Index] as TComAdminRole;
end;

{ TComAdminInstance }

constructor TComAdminInstance.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
end;

procedure TComAdminInstance.ReadExtendedProperties;
begin
  FHasRecycled := CatalogObject.Value[PROPERTY_NAME_RECYCLED];
  FIsPaused := CatalogObject.Value[PROPERTY_NAME_PAUSED];
  FProcessID := VarAsType(CatalogObject.Value[PROPERTY_NAME_PROCESSID], varLongWord);
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
  SupportsUsers := True;
  FRoles := TComAdminRoleList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_PARTITION_ROLES, Key) as ICatalogCollection);
  ReadExtendedProperties;
  GetRoles;
end;

destructor TComAdminPartition.Destroy;
begin
  FRoles.Free;
  inherited;
end;

function TComAdminPartition.GetUsersCollectionName: string;
begin
  Result := COLLECTION_NAME_USERS_PARTITION;
end;

procedure TComAdminPartition.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminPartition.ReadExtendedProperties;
begin
  FChangeable := VarAsType(CatalogObject.Value[PROPERTY_NAME_CHANGEABLE], varBoolean);
  FDeleteable := VarAsType(CatalogObject.Value[PROPERTY_NAME_DELETEABLE], varBoolean);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FID := Key;
end;

procedure TComAdminPartition.SetChangeable(const Value: Boolean);
begin
  FChangeable := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CHANGEABLE) then
    CatalogObject.Value[PROPERTY_NAME_CHANGEABLE] := Value;
end;

procedure TComAdminPartition.SetDeleteable(const Value: Boolean);
begin
  FDeleteable := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DELETEABLE) then
    CatalogObject.Value[PROPERTY_NAME_DELETEABLE] := Value;
end;

procedure TComAdminPartition.SetDescription(const Value: string);
begin
  FDescription := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DESCRIPTION) then
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := Value;
end;

procedure TComAdminPartition.SetID(const Value: string);
begin
  FID := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ID) then
    CatalogObject.Value[PROPERTY_NAME_ID] := Value;
end;

{ TComAdminPartitionList }

function TComAdminPartitionList.GetItem(Index: Integer): TComAdminPartition;
begin
  Result := inherited Items[Index] as TComAdminPartition;
end;

{ TComAdminMethod }

constructor TComAdminMethod.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  FRoles := TComAdminRoleList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_METHOD_ROLES, Key) as ICatalogCollection);
  ReadExtendedProperties;
  GetRoles;
end;

destructor TComAdminMethod.Destroy;
begin
  FRoles.Free;
  inherited;
end;

procedure TComAdminMethod.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminMethod.ReadExtendedProperties;
begin
  FAutoComplete := VarAsType(CatalogObject.Value[PROPERTY_NAME_AUTO_COMPLETE], varBoolean);
  FCLSID := VarToStr(CatalogObject.Value[PROPERTY_NAME_CLSID]);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FIID := VarToStr(CatalogObject.Value[PROPERTY_NAME_IID]);
  FIndex := VarAsType(CatalogObject.Value[PROPERTY_NAME_INDEX], varLongWord);
end;

procedure TComAdminMethod.SetAutoComplete(const Value: Boolean);
begin
  FAutoComplete := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_AUTO_COMPLETE) then
    CatalogObject.Value[PROPERTY_NAME_AUTO_COMPLETE] := Value;
end;

procedure TComAdminMethod.SetDescription(const Value: string);
begin
  FDescription := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DESCRIPTION) then
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := Value;
end;

{ TComAdminMethodList }

function TComAdminMethodList.GetItem(Index: Integer): TComAdminMethod;
begin
  Result := inherited Items[Index] as TComAdminMethod;
end;

{ TComAdminInterface }

constructor TComAdminInterface.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
  FRoles := TComAdminRoleList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_INTERFACE_ROLES, Key) as ICatalogCollection);
  FMethods := TComAdminMethodList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_METHODS, Key) as ICatalogCollection);
  GetRoles;
  GetMethods;
end;

destructor TComAdminInterface.Destroy;
begin
  FRoles.Free;
  FMethods.Free;
  inherited;
end;

procedure TComAdminInterface.GetMethods;
var
  i: Integer;
begin
  for i := 0 to FMethods.CatalogCollection.Count - 1 do
    FMethods.Add(TComAdminMethod.Create(FMethods, FMethods.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminInterface.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminInterface.ReadExtendedProperties;
begin
  FCLSID := VarToStr(CatalogObject.Value[PROPERTY_NAME_CLSID]);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FIID := VarToStr(CatalogObject.Value[PROPERTY_NAME_IID]);
  FQueuingEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_QUEUING_ENABLED], varBoolean);
  FQueuingSupported := VarAsType(CatalogObject.Value[PROPERTY_NAME_QUEUING_SUPPORTED], varBoolean);
end;

procedure TComAdminInterface.SetFDescription(const Value: string);
begin
  FDescription := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DESCRIPTION) then
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := Value;
end;

procedure TComAdminInterface.SetQueuingEnabled(const Value: Boolean);
begin
  FQueuingEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_QUEUING_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_QUEUING_ENABLED] := Value;
end;

{ TComAdminInterfaceList }

function TComAdminInterfaceList.GetItem(Index: Integer): TComAdminInterface;
begin
  Result := inherited Items[Index] as TComAdminInterface;
end;

{ TCOMAdminComponent }

function TCOMAdminComponent.CopyProperties(ASourceComponent: TCOMAdminComponent): Integer;
begin
  inherited CopyProperties(ASourceComponent, Self);

  // Changes must be saved before any sub-collections can be updated
  Result := CatalogCollection.SaveChanges;

  // Synchronize roles from source component
  SyncRoles(ASourceComponent);
end;

constructor TCOMAdminComponent.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
  FRoles := TComAdminRoleList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_COMPONENT_ROLES, Key) as ICatalogCollection);
  FInterfaces := TComAdminInterfaceList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_INTERFACES, Key) as ICatalogCollection);
  GetRoles;
  GetInterfaces;
end;

destructor TCOMAdminComponent.Destroy;
begin
  FRoles.Free;
  FInterfaces.Free;
  inherited;
end;

procedure TCOMAdminComponent.GetInterfaces;
var
  i: Integer;
begin
  for i := 0 to FInterfaces.CatalogCollection.Count - 1 do
    FInterfaces.Add(TComAdminInterface.Create(FInterfaces, FInterfaces.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TCOMAdminComponent.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TCOMAdminComponent.ReadExtendedProperties;
begin
  FAllowInprocSubscribers := VarAsType(CatalogObject.Value[PROPERTY_NAME_ALLOW_SUBSCRIBERS], varBoolean);
  FApplicationID := VarToStr(CatalogObject.Value[PROPERTY_NAME_APPLICATION_ID]);
  FBitness := VarAsType(CatalogObject.Value[PROPERTY_NAME_BITNESS], varLongWord);
  FComponentAccessChecksEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_COMPONENT_ACCESS_CHECKS], varBoolean);
  FComponentTransactionTimeout := VarAsType(CatalogObject.Value[PROPERTY_NAME_COMPONENT_TIMEOUT], varLongWord);
  FComponentTransactionTimeoutEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_COMPONENT_TIMEOUT_ENABLED], varBoolean);
  FCOMTIIntrinsics := VarAsType(CatalogObject.Value[PROPERTY_NAME_COM_TIINTRINSICS], varBoolean);
  FConstructionEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_CONSTRUCTION_ENABLED], varBoolean);
  FConstructorString := VarToStr(CatalogObject.Value[PROPERTY_NAME_CONSTRUCTOR_STRING]);
  FCreationTimeout := VarAsType(CatalogObject.Value[PROPERTY_NAME_CREATION_TIMEOUT], varLongWord);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FDll := VarToStr(CatalogObject.Value[PROPERTY_NAME_DLL]);
  FEventTrackingEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_EVENT_TRACKING], varBoolean);
  FExceptionClass := VarToStr(CatalogObject.Value[PROPERTY_NAME_EXCEPTION_CLASS]);
  FFireInParallel := VarAsType(CatalogObject.Value[PROPERTY_NAME_FIRE_IN_PARALLEL], varBoolean);
  FIISIntrinsics := VarAsType(CatalogObject.Value[PROPERTY_NAME_IIS_INTRINSICS], varBoolean);
//  FInitializeServerApplication := VarAsType(CatalogObject.Value[PROPERTY_NAME_INIT_SERVER_APPLICATION], varBoolean);
  FIsEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_ENABLED], varBoolean);
  FIsEventClass := VarAsType(CatalogObject.Value[PROPERTY_NAME_IS_EVENT_CLASS], varBoolean);
  FIsInstalled := VarAsType(CatalogObject.Value[PROPERTY_NAME_IS_INSTALLED], varBoolean);
  FIsPrivateComponent := VarAsType(CatalogObject.Value[PROPERTY_NAME_IS_PRIVATE_COMPONENT], varBoolean);
  FJustInTimeActivation := VarAsType(CatalogObject.Value[PROPERTY_NAME_JUST_IN_TIME], varBoolean);
  FLoadBalancingSupported := VarAsType(CatalogObject.Value[PROPERTY_NAME_LOAD_BALANCING], varBoolean);
  FMaxPoolSize := VarAsType(CatalogObject.Value[PROPERTY_NAME_MAX_POOL_SIZE], varLongWord);
  FMinPoolSize := VarAsType(CatalogObject.Value[PROPERTY_NAME_MIN_POOL_SIZE], varLongWord);
  FMultiInterfacePublisherFilterCLSID := VarToStr(CatalogObject.Value[PROPERTY_NAME_MI_FILTER_CLSID]);
  FMustRunInClientContext := VarAsType(CatalogObject.Value[PROPERTY_NAME_MUST_RUN_CLIENT_CONTEXT], varBoolean);
  FMustRunInDefaultContext := VarAsType(CatalogObject.Value[PROPERTY_NAME_MUST_RUN_DEFAULT_CONTEXT], varBoolean);
  FObjectPoolingEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_OBJECT_POOLING], varBoolean);
  FProgID := VarToStr(CatalogObject.Value[PROPERTY_NAME_PROG_ID]);
end;

procedure TCOMAdminComponent.SetAllowInprocSubscribers(const Value: Boolean);
begin
  FAllowInprocSubscribers := Value;
end;

procedure TCOMAdminComponent.SetApplicationID(const Value: string);
begin
  FApplicationID := Value;
//  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_APPLICATION_ID) then
//    CatalogObject.Value[PROPERTY_NAME_APPLICATION_ID] := Value;
end;

procedure TCOMAdminComponent.SetBitness(const Value: TCOMAdminComponentType);
begin
  FBitness := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_BITNESS) then
    CatalogObject.Value[PROPERTY_NAME_BITNESS] := Value;
end;

procedure TCOMAdminComponent.SetComponentAccessChecksEnabled(const Value: Boolean);
begin
  FComponentAccessChecksEnabled := Value;
end;

procedure TCOMAdminComponent.SetComponentTransactionTimeout(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_TIMEOUT, Value) then
  begin
    FComponentTransactionTimeout := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_COMPONENT_TIMEOUT) then
      CatalogObject.Value[PROPERTY_NAME_COMPONENT_TIMEOUT] := Value;
  end;
end;

procedure TCOMAdminComponent.SetComponentTransactionTimeoutEnabled(const Value: Boolean);
begin
  FComponentTransactionTimeoutEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_COMPONENT_TIMEOUT_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_COMPONENT_TIMEOUT_ENABLED] := Value;
end;

procedure TCOMAdminComponent.SetCOMTIIntrinsics(const Value: Boolean);
begin
  FCOMTIIntrinsics := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_COM_TIINTRINSICS) then
    CatalogObject.Value[PROPERTY_NAME_COM_TIINTRINSICS] := Value;
end;

procedure TCOMAdminComponent.SetConstructionEnabled(const Value: Boolean);
begin
  FConstructionEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CONSTRUCTION_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_CONSTRUCTION_ENABLED] := Value;
end;

procedure TCOMAdminComponent.SetConstructorString(const Value: string);
begin
  FConstructorString := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CONSTRUCTOR_STRING) then
    CatalogObject.Value[PROPERTY_NAME_CONSTRUCTOR_STRING] := Value;
end;

procedure TCOMAdminComponent.SetCreationTimeout(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAXLONG, Value) then
  begin
    FCreationTimeout := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CREATION_TIMEOUT) then
      CatalogObject.Value[PROPERTY_NAME_CREATION_TIMEOUT] := Value;
  end;
end;

procedure TCOMAdminComponent.SetDescription(const Value: string);
begin
  FDescription := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DESCRIPTION) then
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := Value;
end;

procedure TCOMAdminComponent.SetDll(const Value: string);
begin
  FDll := Value;
end;

procedure TCOMAdminComponent.SetEventTrackingEnabled(const Value: Boolean);
begin
  FEventTrackingEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_EVENT_TRACKING) then
    CatalogObject.Value[PROPERTY_NAME_EVENT_TRACKING] := Value;
end;

procedure TCOMAdminComponent.SetExceptionClass(const Value: string);
begin
  FExceptionClass := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_EXCEPTION_CLASS) then
    CatalogObject.Value[PROPERTY_NAME_EXCEPTION_CLASS] := Value;
end;

procedure TCOMAdminComponent.SetFireInParallel(const Value: Boolean);
begin
  FFireInParallel := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_FIRE_IN_PARALLEL) then
    CatalogObject.Value[PROPERTY_NAME_FIRE_IN_PARALLEL] := Value;
end;

procedure TCOMAdminComponent.SetIISIntrinsics(const Value: Boolean);
begin
  FIISIntrinsics := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_IIS_INTRINSICS) then
    CatalogObject.Value[PROPERTY_NAME_IIS_INTRINSICS] := Value;
end;

procedure TCOMAdminComponent.SetInitializeServerApplication(const Value: Boolean);
begin
  FInitializeServerApplication := Value;
// This property is documented but obviously not implemented
//  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_INIT_SERVER_APPLICATION) then
//    CatalogObject.Value[PROPERTY_NAME_INIT_SERVER_APPLICATION] := Value;
end;

procedure TCOMAdminComponent.SetIsEnabled(const Value: Boolean);
begin
  FIsEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_ENABLED] := Value;
end;

procedure TCOMAdminComponent.SetIsEventClass(const Value: Boolean);
begin
  FIsEventClass := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_IS_EVENT_CLASS) then
    CatalogObject.Value[PROPERTY_NAME_IS_EVENT_CLASS] := Value;
end;

procedure TCOMAdminComponent.SetIsInstalled(const Value: Boolean);
begin
  FIsInstalled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_IS_INSTALLED) then
    CatalogObject.Value[PROPERTY_NAME_IS_INSTALLED] := Value;
end;

procedure TCOMAdminComponent.SetIsPrivateComponent(const Value: Boolean);
begin
  FIsPrivateComponent := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_IS_PRIVATE_COMPONENT) then
    CatalogObject.Value[PROPERTY_NAME_IS_PRIVATE_COMPONENT] := Value;
end;

procedure TCOMAdminComponent.SetJustInTimeActivation(const Value: Boolean);
begin
  FJustInTimeActivation := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_JUST_IN_TIME) then
    CatalogObject.Value[PROPERTY_NAME_JUST_IN_TIME] := Value;
end;

procedure TCOMAdminComponent.SetLoadBalancingSupported(const Value: Boolean);
begin
  FLoadBalancingSupported := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_LOAD_BALANCING) then
    CatalogObject.Value[PROPERTY_NAME_LOAD_BALANCING] := Value;
end;

procedure TCOMAdminComponent.SetMaxPoolSize(const Value: Cardinal);
begin
  if InternalCheckRange(1, MAX_POOL_SIZE, Value) then
  begin
    FMaxPoolSize := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_MAX_POOL_SIZE) then
      CatalogObject.Value[PROPERTY_NAME_MAX_POOL_SIZE] := Value;
  end;
end;

procedure TCOMAdminComponent.SetMinPoolSize(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_POOL_SIZE, Value) then
  begin
    FMinPoolSize := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_MIN_POOL_SIZE) then
      CatalogObject.Value[PROPERTY_NAME_MIN_POOL_SIZE] := Value;
  end;
end;

procedure TCOMAdminComponent.SetMultiInterfacePublisherFilterCLSID(const Value: string);
begin
  FMultiInterfacePublisherFilterCLSID := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_MI_FILTER_CLSID) then
    CatalogObject.Value[PROPERTY_NAME_MI_FILTER_CLSID] := Value;
end;

procedure TCOMAdminComponent.SetMustRunInClientContext(const Value: Boolean);
begin
  FMustRunInClientContext := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_MUST_RUN_CLIENT_CONTEXT) then
    CatalogObject.Value[PROPERTY_NAME_MUST_RUN_CLIENT_CONTEXT] := Value;
end;

procedure TCOMAdminComponent.SetMustRunInDefaultContext(const Value: Boolean);
begin
  FMustRunInDefaultContext := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_MUST_RUN_DEFAULT_CONTEXT) then
    CatalogObject.Value[PROPERTY_NAME_MUST_RUN_DEFAULT_CONTEXT] := Value;
end;

procedure TCOMAdminComponent.SetObjectPoolingEnabled(const Value: Boolean);
begin
  FObjectPoolingEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_OBJECT_POOLING) then
    CatalogObject.Value[PROPERTY_NAME_OBJECT_POOLING] := Value;
end;

procedure TCOMAdminComponent.SetProgID(const Value: string);
begin
  FProgID := Value;
end;

procedure TCOMAdminComponent.SyncRoles(ASourceComponent: TCOMAdminComponent);
var
  i: Integer;
  LRole: TComAdminRole;
begin
    // sync roles in source component to target component
  for i := 0 to ASourceComponent.Roles.Count - 1 do
  begin
    Collection.Catalog.DebugMessage(DEBUG_MESSAGE_SYNC_ROLE, [ASourceComponent.Roles[i].Name]);
    if FRoles.Find(ASourceComponent.Roles[i].Name, LRole) then
      LRole.CopyProperties(ASourceComponent.Roles[i])
    else
      FRoles.Append(ASourceComponent.Roles[i]); // Role does not exists in target component ==> create & copy
  end;
  // delete all roles in target component that not exists in source component
  for i := ASourceComponent.Roles.Count - 1 downto 0 do
  begin
    if not ASourceComponent.Roles.Contains(ASourceComponent.Roles[i].Name) then
      Collection.Catalog.ChangeCount := Collection.Catalog.ChangeCount + ASourceComponent.Roles.Delete(i);
  end;
end;

{ TCOMAdminComponentList }

function TCOMAdminComponentList.Append(ASourceComponent: TCOMAdminComponent): TCOMAdminComponent;
var
  LLibraryName: string;
begin
  if CopyLibrary(ASourceComponent) then
  begin
    LLibraryName := TPath.Combine(Catalog.LibraryPath, ExtractFileName(ASourceComponent.Dll));
    Result := (Owner as TComAdminApplication).InstallComponent(LLibraryName);
    Result.CopyProperties(ASourceComponent);
    Catalog.ChangeCount := Catalog.ChangeCount + SaveChanges;
  end else
    raise Exception.Create(ERROR_COPY_LIBRARY);
end;

function TCOMAdminComponentList.BuildTargetLibraryName(ASourceComponent: TCOMAdminComponent): string;
begin
  if Catalog.LibraryPath.IsEmpty then
    raise Exception.CreateFmt(ERROR_INVALID_LIBRARY_PATH, [Catalog.Server]);
  if Catalog.Server.IsEmpty then
    Result := Format('%s', [TPath.Combine(Catalog.LibraryPath, ExtractFileName(ASourceComponent.Dll))])
  else
    Result := Format('\\%s\%s', [Catalog.Server, TPath.Combine(Catalog.LibraryPath, ExtractFileName(ASourceComponent.Dll)).Replace(':','$')]);
end;

function TCOMAdminComponentList.CopyLibrary(ASourceComponent: TCOMAdminComponent; AOverwrite: Boolean): Boolean;
var
  LTargetLibrary: string;
begin
  LTargetLibrary := BuildTargetLibraryName(ASourceComponent);
  if AOverwrite or not FileExists(LTargetLibrary) then
  try
    if (Owner as TComAdminApplication).Instances.Count > 0 then
    begin
      Catalog.Catalog.ShutdownApplication(Owner.Key);
      if (Owner as TComAdminApplication).Instances.Count = 0 then
        TFile.Copy(ASourceComponent.Dll, LTargetLibrary, AOverwrite)
      else
        raise Exception.Create(ERROR_APPLICATION_NOT_DOWN);
    end else
      TFile.Copy(ASourceComponent.Dll, LTargetLibrary, AOverwrite);
    Owner.Collection.Catalog.DebugMessage(DEBUG_MESSAGE_COPY_LIBRARY, [LTargetLibrary]);
  except
    on E:Exception do
      Owner.Collection.Catalog.DebugMessage(DEBUG_MESSAGE_COPY_ERROR, [LTargetLibrary, E.Message]);
  end;
  Result := FileExists(LTargetLibrary);
end;

function TCOMAdminComponentList.Find(const AName: string; out AComponent: TCOMAdminComponent): Boolean;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if Items[i].Name.Equals(AName) then
    begin
      AComponent := Items[i];
      Exit(True);
    end;
  end;
  Result := False;
end;

function TCOMAdminComponentList.GetItem(Index: Integer): TCOMAdminComponent;
begin
  Result := inherited Items[Index] as TCOMAdminComponent;
end;

{ TComAdminApplication }

function TComAdminApplication.BuildInstallFileName: string;
begin
  if Collection.Catalog.LibraryPath.IsEmpty then
    raise Exception.CreateFmt(ERROR_INVALID_LIBRARY_PATH, [Collection.Catalog.Server]);
  Result := Format('\\%s\%s\%s.msi', [Collection.Catalog.Server, Collection.Catalog.LibraryPath, Name]);
end;

function TComAdminApplication.CopyProperties(ASourceApplication: TCOMAdminApplication; const APassword: string): Integer;
begin

  inherited CopyProperties(ASourceApplication, Self);
  Password := APassword;

  // Changes must be saved before any sub-collections can be updated
  Result := CatalogCollection.SaveChanges;

  // Synchronize roles from source application
  SyncRoles(ASourceApplication);
  SyncComponents(ASourceApplication);

end;

function TComAdminApplication.CopyToServer(ATargetServer: TComAdminCatalog; AOptions: Integer): Boolean;
var
  i: Integer;
  LInstallFileName: string;
begin
  for i := 0 to ATargetServer.Applications.Count - 1 do
  begin
    if ATargetServer.Applications[i].Name.Equals(Name) then
    begin
      // create filename for installation file
      LInstallFileName := BuildInstallFileName;
      // export application to file
      Collection.Catalog.ExportApplicationByKey(Key, LInstallFileName);
      // shutdown application on target server
      ATargetServer.Catalog.ShutdownApplication(ATargetServer.Applications[i].Key);
      // delete existing application on target server
      ATargetServer.Applications.Delete(i);
      // install application on target server
      ATargetServer.Catalog.InstallApplication(LInstallFileName, ATargetServer.LibraryPath, AOptions, '', '', '');
      // delete installation files as they are no longer needed
      TFile.Delete(LInstallFileName);
      TFile.Delete(Format('%s.cab', [LInstallFileName]));
      Exit(True);
    end;
  end;
  Result := False;
end;

constructor TComAdminApplication.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);

  SupportsUsers := True;
  ReadExtendedProperties;

  // Create List objects
  FInstances := TComAdminInstanceList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, Key) as ICatalogCollection);
  FRoles := TComAdminRoleList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_ROLES, Key) as ICatalogCollection);
  FComponents := TCOMAdminComponentList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_COMPONENTS, Key) as ICatalogCollection);

  GetRoles;
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

function TComAdminApplication.GetUsersCollectionName: string;
begin
  Result := COLLECTION_NAME_USERS;
end;

procedure TComAdminApplication.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

function TComAdminApplication.InstallComponent(const ALibraryName: string): TCOMAdminComponent;
begin
  Collection.Catalog.Catalog.InstallComponent(Name, ALibraryName, '', '');
  FComponents.CatalogCollection.Populate; // muste be populated to retrieve the newly installed component
  FComponents.Add(TCOMAdminComponent.Create(FComponents, (FComponents.CatalogCollection.Item[FComponents.CatalogCollection.Count - 1]) as ICatalogObject));
  Result := FComponents.Items[FComponents.Count - 1] as TCOMAdminComponent;
end;

function TComAdminApplication.IsIntegratedIdentity: Boolean;
begin
  Result := FIdentity.ToLower.Equals(IDENTITY_STRING_INTERACTIVE)
         or FIdentity.ToLower.Equals(IDENTITY_STRING_LOCALSERVICE)
         or FIdentity.ToLower.Equals(IDENTITY_STRING_NETWORKSERVICE)
         or FIdentity.ToLower.Equals(IDENTITY_STRING_SYSTEM);
end;

function TComAdminApplication.GetInstances: TComAdminInstanceList;
var
  Collection: ICatalogCollection;
  i: Integer;
begin
  // As the instances collection must be really up to date, the getter reads it fresh from the catalog
  Collection := CatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, Key) as ICatalogCollection;
  Collection.Populate;
  FInstances.Clear;
  for i := 0 to Collection.Count - 1 do
    FInstances.Add(TComAdminInstance.Create(FInstances, Collection.Item[i] as ICatalogObject));
  Result := FInstances;
end;

procedure TComAdminApplication.ReadExtendedProperties;
begin
  FGig3SupportEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_3GIG], varBoolean);
  FAccessChecksLevel := VarAsType(CatalogObject.Value[PROPERTY_NAME_ACCESS_CHECK_LEVEL], varLongWord);
  FActivation := VarAsType(CatalogObject.Value[PROPERTY_NAME_ACTIVATION], varLongWord);
  FAccessChecksEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_ACCESS_CHECKS], varBoolean);
  FDirectory := VarToStr(CatalogObject.Value[PROPERTY_NAME_APPLICATION_DIRECTORY]);
  FProxy := VarAsType(CatalogObject.Value[PROPERTY_NAME_APPLICATION_PROXY], varBoolean);
  FProxyServerName := VarToStr(CatalogObject.Value[PROPERTY_NAME_PROXY_SERVER_NAME]);
  FPartitionID := VarToStr(CatalogObject.Value[PROPERTY_NAME_PARTITION_ID]);
  FAuthenticationLevel := VarAsType(CatalogObject.Value[PROPERTY_NAME_AUTHENTICATION], varLongWord);
  FAuthenticationCapability := VarAsType(CatalogObject.Value[PROPERTY_NAME_AUTH_CAPABILITY], varLongWord);
  FChangeable := VarAsType(CatalogObject.Value[PROPERTY_NAME_CHANGEABLE], varBoolean);
  FCommandLine := VarToStr(CatalogObject.Value[PROPERTY_NAME_COMMAND_LINE]);
  FConcurrentApps := VarAsType(CatalogObject.Value[PROPERTY_NAME_CONCURRENT_APPS], varLongWord);
  FCreatedBy := VarToStr(CatalogObject.Value[PROPERTY_NAME_CREATED_BY]);
  FCRMEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_CRM_ENABLED], varBoolean);
  FCRMLogFile := VarToStr(CatalogObject.Value[PROPERTY_NAME_CRM_LOGFILE]);
  FDeleteable := VarAsType(CatalogObject.Value[PROPERTY_NAME_DELETEABLE], varBoolean);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FDumpEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_DUMP_ENABLED], varBoolean);
  FDumpOnException := VarAsType(CatalogObject.Value[PROPERTY_NAME_DUMP_EXCEPTION], varBoolean);
  FDumpOnFailFast := VarAsType(CatalogObject.Value[PROPERTY_NAME_DUMP_FAILFAST], varBoolean);
  FDumpPath := VarToStr(CatalogObject.Value[PROPERTY_NAME_DUMP_PATH]);
  FEventsEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_EVENTS_ENABLED], varBoolean);
  FIdentity := VarToStr(CatalogObject.Value[PROPERTY_NAME_IDENTITY]);
  FImpersonationLevel := VarAsType(CatalogObject.Value[PROPERTY_NAME_IMPERSONATION], varLongWord);
  FIsEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_ENABLED], varBoolean);
  FIsSystem := VarAsType(CatalogObject.Value[PROPERTY_NAME_SYSTEM], varBoolean);
  FMaxDumpCount := VarAsType(CatalogObject.Value[PROPERTY_NAME_MAX_DUMPS], varLongWord);
  FQCAuthenticateMsgs := VarAsType(CatalogObject.Value[PROPERTY_NAME_QC_AUTHENTICATE], varLongWord);
  FQCListenerMaxThreads := VarAsType(CatalogObject.Value[PROPERTY_NAME_QC_MAXTHREADS], varLongWord);
  FQueueListenerEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_QUEUE_LISTENER], varBoolean);
  FQueuingEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_QUEUING_ENABLED], varBoolean);
  FRecycleActivationLimit := VarAsType(CatalogObject.Value[PROPERTY_NAME_RECYCLE_ACTIVATION], varLongWord);
  FRecycleCallLimit := VarAsType(CatalogObject.Value[PROPERTY_NAME_RECYCLE_CALL_LIMIT], varLongWord);
  FRecycleExpirationTimeout := VarAsType(CatalogObject.Value[PROPERTY_NAME_RECYCLE_EXPIRATION], varLongWord);
  FRecycleLifetimeLimit := VarAsType(CatalogObject.Value[PROPERTY_NAME_RECYCLE_LIFETIME_LIMIT], varLongWord);
  FRecycleMemoryLimit := VarAsType(CatalogObject.Value[PROPERTY_NAME_RECYCLE_MEMORY_LIMIT], varLongWord);
  FReplicable := VarAsType(CatalogObject.Value[PROPERTY_NAME_REPLICABLE], varBoolean);
  FRunForever := VarAsType(CatalogObject.Value[PROPERTY_NAME_RUN_FOREVER], varBoolean);
  FServiceName := VarToStr(CatalogObject.Value[PROPERTY_NAME_SERVICE_NAME]);
  FShutdownAfter := VarAsType(CatalogObject.Value[PROPERTY_NAME_SHUTDOWN], varLongWord);
  FSoapActivated := VarAsType(CatalogObject.Value[PROPERTY_NAME_SOAP_ACTIVATED], varBoolean);
  FSoapBaseUrl := VarToStr(CatalogObject.Value[PROPERTY_NAME_SOAP_BASE_URL]);
  FSoapMailTo := VarToStr(CatalogObject.Value[PROPERTY_NAME_SOAP_MAILTO]);
  FSoapVRoot := VarToStr(CatalogObject.Value[PROPERTY_NAME_SOAP_VROOT]);
  FSRPEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_SRP_ENABLED], varBoolean);
  FSRPTrustLevel := VarAsType(CatalogObject.Value[PROPERTY_NAME_SRP_TRUSTLEVEL], varLongWord);
end;

procedure TComAdminApplication.SetAccessChecksEnabled(const Value: Boolean);
begin
  FAccessChecksEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ACCESS_CHECKS) then
    CatalogObject.Value[PROPERTY_NAME_ACCESS_CHECKS] := Value;
end;

procedure TComAdminApplication.SetAccessChecksLevel(const Value: TCOMAdminAccessChecksLevelOptions);
begin
  FAccessChecksLevel := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ACCESS_CHECK_LEVEL) then
    CatalogObject.Value[PROPERTY_NAME_ACCESS_CHECK_LEVEL] := Value;
end;

procedure TComAdminApplication.SetActivation(const Value: TCOMAdminApplicationActivation);
begin
  FActivation := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ACTIVATION) then
    CatalogObject.Value[PROPERTY_NAME_ACTIVATION] := Value;
end;

procedure TComAdminApplication.SetAuthenticationCapability(const Value: TCOMAdminAuthenticationCapability);
begin
  FAuthenticationCapability := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_AUTH_CAPABILITY) then
    CatalogObject.Value[PROPERTY_NAME_AUTH_CAPABILITY] := Value;
end;

procedure TComAdminApplication.SetAuthenticationLevel(const Value: TCOMAdminAuthenticationLevel);
begin
  FAuthenticationLevel := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_AUTHENTICATION) then
    CatalogObject.Value[PROPERTY_NAME_AUTHENTICATION] := Value;
end;

procedure TComAdminApplication.SetChangeable(const Value: Boolean);
begin
  FChangeable := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CHANGEABLE) then
    CatalogObject.Value[PROPERTY_NAME_CHANGEABLE] := Value;
end;

procedure TComAdminApplication.SetCommandLine(const Value: string);
begin
  FCommandLine := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_COMMAND_LINE) then
    CatalogObject.Value[PROPERTY_NAME_COMMAND_LINE] := Value;
end;

procedure TComAdminApplication.SetConcurrentApps(const Value: Cardinal);
begin
  if InternalCheckRange(1, MAX_POOL_SIZE, Value) then
  begin
    FConcurrentApps := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CONCURRENT_APPS) then
      CatalogObject.Value[PROPERTY_NAME_CONCURRENT_APPS] := Value;
  end;
end;

procedure TComAdminApplication.SetCreatedBy(const Value: string);
begin
  FCreatedBy := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CREATED_BY) then
    CatalogObject.Value[PROPERTY_NAME_CREATED_BY] := Value;
end;

procedure TComAdminApplication.SetCRMEnabled(const Value: Boolean);
begin
  FCRMEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CRM_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_CRM_ENABLED] := Value;
end;

procedure TComAdminApplication.SetCRMLogFile(const Value: string);
begin
  FCRMLogFile := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_CRM_LOGFILE) then
    CatalogObject.Value[PROPERTY_NAME_CRM_LOGFILE] := Value;
end;

procedure TComAdminApplication.SetDeleteable(const Value: Boolean);
begin
  FDeleteable := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DELETEABLE) then
    CatalogObject.Value[PROPERTY_NAME_DELETEABLE] := Value;
end;

procedure TComAdminApplication.SetDescription(const Value: string);
begin
  FDescription := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DESCRIPTION) then
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := Value;
end;

procedure TComAdminApplication.SetDirectory(const Value: string);
begin
  FDirectory := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_APPLICATION_DIRECTORY) then
    CatalogObject.Value[PROPERTY_NAME_APPLICATION_DIRECTORY] := Value;
end;

procedure TComAdminApplication.SetDumpEnabled(const Value: Boolean);
begin
  FDumpEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DUMP_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_DUMP_ENABLED] := Value;
end;

procedure TComAdminApplication.SetDumpOnException(const Value: Boolean);
begin
  FDumpOnException := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DUMP_EXCEPTION) then
    CatalogObject.Value[PROPERTY_NAME_DUMP_EXCEPTION] := Value;
end;

procedure TComAdminApplication.SetDumpOnFailFast(const Value: Boolean);
begin
  FDumpOnFailFast := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DUMP_FAILFAST) then
    CatalogObject.Value[PROPERTY_NAME_DUMP_FAILFAST] := Value;
end;

procedure TComAdminApplication.SetDumpPath(const Value: string);
begin
  FDumpPath := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_DUMP_PATH) then
    CatalogObject.Value[PROPERTY_NAME_DUMP_PATH] := Value;
end;

procedure TComAdminApplication.SetEventsEnabled(const Value: Boolean);
begin
  FEventsEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_EVENTS_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_EVENTS_ENABLED] := Value;
end;

procedure TComAdminApplication.SetGig3SupportEnabled(const Value: Boolean);
begin
  FGig3SupportEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_3GIG) then
    CatalogObject.Value[PROPERTY_NAME_3GIG] := Value;
end;

procedure TComAdminApplication.SetIdentity(const Value: string);
begin
  FIdentity := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_IDENTITY) then
    CatalogObject.Value[PROPERTY_NAME_IDENTITY] := Value;
end;

procedure TComAdminApplication.SetImpersonationLevel(const Value: TCOMAdminImpersonationLevel);
begin
  FImpersonationLevel := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_IMPERSONATION) then
    CatalogObject.Value[PROPERTY_NAME_IMPERSONATION] := Value;
end;

procedure TComAdminApplication.SetIsEnabled(const Value: Boolean);
begin
  FIsEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_ENABLED] := Value;
end;

procedure TComAdminApplication.SetMaxDumpCount(const Value: Cardinal);
begin
  if InternalCheckRange(1, MAX_DUMP_COUNT, Value) then
  begin
    FMaxDumpCount := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_MAX_DUMPS) then
      CatalogObject.Value[PROPERTY_NAME_MAX_DUMPS] := Value;
  end;
end;

procedure TComAdminApplication.SetPartitionID(const Value: string);
begin
  FPartitionID := Value;
end;

procedure TComAdminApplication.SetPassword(const Value: string);
begin
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_PASSWORD) then
    CatalogObject.Value[PROPERTY_NAME_PASSWORD] := Value;
end;

procedure TComAdminApplication.SetProxy(const Value: Boolean);
begin
  FProxy := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_APPLICATION_PROXY) then
    CatalogObject.Value[PROPERTY_NAME_APPLICATION_PROXY] := Value;
end;

procedure TComAdminApplication.SetProxyServerName(const Value: string);
begin
  FProxyServerName := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_PROXY_SERVER_NAME) then
    CatalogObject.Value[PROPERTY_NAME_PROXY_SERVER_NAME] := Value;
end;

procedure TComAdminApplication.SetQCAuthenticateMsgs(const Value: TCOMAdminQCAuthenticateMsgs);
begin
  FQCAuthenticateMsgs := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_QC_AUTHENTICATE) then
    CatalogObject.Value[PROPERTY_NAME_QC_AUTHENTICATE] := Value;
end;

procedure TComAdminApplication.SetQCListenerMaxThreads(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_THREADS, Value) then
  begin
    FQCListenerMaxThreads := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_QC_MAXTHREADS) then
      CatalogObject.Value[PROPERTY_NAME_QC_MAXTHREADS] := Value;
  end;
end;

procedure TComAdminApplication.SetQueueListenerEnabled(const Value: Boolean);
begin
  FQueueListenerEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_QUEUE_LISTENER) then
    CatalogObject.Value[PROPERTY_NAME_QUEUE_LISTENER] := Value;
end;

procedure TComAdminApplication.SetQueuingEnabled(const Value: Boolean);
begin
  FQueuingEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_QUEUING_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_QUEUING_ENABLED] := Value;
end;

procedure TComAdminApplication.SetRecycleActivationLimit(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_POOL_SIZE, Value) then
  begin
    FRecycleActivationLimit := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_RECYCLE_ACTIVATION) then
      CatalogObject.Value[PROPERTY_NAME_RECYCLE_ACTIVATION] := Value;
  end;
end;

procedure TComAdminApplication.SetRecycleCallLimit(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_POOL_SIZE, Value) then
  begin
    FRecycleCallLimit := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_RECYCLE_CALL_LIMIT) then
      CatalogObject.Value[PROPERTY_NAME_RECYCLE_CALL_LIMIT] := Value;
  end;
end;

procedure TComAdminApplication.SetRecycleExpirationTimeout(const Value: Cardinal);
begin
  if InternalCheckRange(1, MAX_RECYCLE_TIMEOUT, Value) then
  begin
    FRecycleExpirationTimeout := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_RECYCLE_EXPIRATION) then
      CatalogObject.Value[PROPERTY_NAME_RECYCLE_EXPIRATION] := Value;
  end;
end;

procedure TComAdminApplication.SetRecycleLifetimeLimit(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_LIFETIME_LIMIT, Value) then
  begin
    FRecycleLifetimeLimit := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_RECYCLE_LIFETIME_LIMIT) then
      CatalogObject.Value[PROPERTY_NAME_RECYCLE_LIFETIME_LIMIT] := Value;
  end;
end;

procedure TComAdminApplication.SetRecycleMemoryLimit(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_POOL_SIZE, Value) then
  begin
    FRecycleMemoryLimit := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_RECYCLE_MEMORY_LIMIT) then
      CatalogObject.Value[PROPERTY_NAME_RECYCLE_MEMORY_LIMIT] := Value;
  end;
end;

procedure TComAdminApplication.SetReplicable(const Value: Boolean);
begin
  FReplicable := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_REPLICABLE) then
    CatalogObject.Value[PROPERTY_NAME_REPLICABLE] := Value;
end;

procedure TComAdminApplication.SetRunForever(const Value: Boolean);
begin
  FRunForever := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_RUN_FOREVER) then
    CatalogObject.Value[PROPERTY_NAME_RUN_FOREVER] := Value;
end;

procedure TComAdminApplication.SetServiceName(const Value: string);
begin
  FServiceName := Value;
end;

procedure TComAdminApplication.SetShutdownAfter(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_RECYCLE_TIMEOUT, Value) then
  begin
    FShutdownAfter := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SHUTDOWN) then
      CatalogObject.Value[PROPERTY_NAME_SHUTDOWN] := Value;
  end;
end;

procedure TComAdminApplication.SetSoapActivated(const Value: Boolean);
begin
  FSoapActivated := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SOAP_ACTIVATED) then
    CatalogObject.Value[PROPERTY_NAME_SOAP_ACTIVATED] := Value;
end;

procedure TComAdminApplication.SetSoapBaseUrl(const Value: string);
begin
  FSoapBaseUrl := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SOAP_BASE_URL) then
    CatalogObject.Value[PROPERTY_NAME_SOAP_BASE_URL] := Value;
end;

procedure TComAdminApplication.SetSoapMailTo(const Value: string);
begin
  FSoapMailTo := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SOAP_MAILTO) then
    CatalogObject.Value[PROPERTY_NAME_SOAP_MAILTO] := Value;
end;

procedure TComAdminApplication.SetSoapVRoot(const Value: string);
begin
  FSoapVRoot := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SOAP_VROOT) then
    CatalogObject.Value[PROPERTY_NAME_SOAP_VROOT] := Value;
end;

procedure TComAdminApplication.SetSRPEnabled(const Value: Boolean);
begin
  FSRPEnabled := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SRP_ENABLED) then
    CatalogObject.Value[PROPERTY_NAME_SRP_ENABLED] := Value;
end;

procedure TComAdminApplication.SetSRPTrustLevel(const Value: TCOMAdminSRPTrustLevel);
begin
  FSRPTrustLevel := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_SRP_TRUSTLEVEL) then
    CatalogObject.Value[PROPERTY_NAME_SRP_TRUSTLEVEL] := Value;
end;

procedure TComAdminApplication.SyncComponents(ASourceApplication: TCOMAdminApplication);
var
  i: Integer;
  LComponent: TCOMAdminComponent;
begin
    // sync components in source application to target application
  for i := 0 to ASourceApplication.Components.Count - 1 do
  begin
    Collection.Catalog.DebugMessage(DEBUG_MESSAGE_SYNC_COMPONENT, [ASourceApplication.Components[i].Name]);
    if FComponents.Find(ASourceApplication.Components[i].Name, LComponent) then
    begin
      LComponent.CopyProperties(ASourceApplication.Components[i]);
      if Collection.Catalog.CopyLibraries then
      begin
        if (FComponents.Owner as TComAdminApplication).Instances.Count > 0 then
          FComponents.Catalog.Catalog.ShutdownApplication(FComponents.Owner.Key);
        if not Collection.Catalog.CopiedLibraries.Contains(ASourceApplication.Components[i].Dll) then
        begin
          FComponents.CopyLibrary(ASourceApplication.Components[i], True);
          Collection.Catalog.CopiedLibraries.Add(ASourceApplication.Components[i].Dll);
        end;
      end;
    end else
      FComponents.Append(ASourceApplication.Components[i]); // Component does not exists in target application ==> create & copy
  end;
  // delete all components in target application that not exists in source application
  for i := ASourceApplication.Components.Count - 1 downto 0 do
  begin
    if not ASourceApplication.Components.Contains(ASourceApplication.Components[i].Name) then
      Collection.Catalog.ChangeCount := Collection.Catalog.ChangeCount + ASourceApplication.Components.Delete(i);
  end;
end;

procedure TComAdminApplication.SyncRoles(ASourceApplication: TCOMAdminApplication);
var
  i: Integer;
  LRole: TComAdminRole;
begin
    // sync roles in source application to target application
  for i := 0 to ASourceApplication.Roles.Count - 1 do
  begin
    Collection.Catalog.DebugMessage(DEBUG_MESSAGE_SYNC_ROLE, [ASourceApplication.Roles[i].Name]);
    if FRoles.Find(ASourceApplication.Roles[i].Name, LRole) then
      LRole.CopyProperties(ASourceApplication.Roles[i])
    else
      FRoles.Append(ASourceApplication.Roles[i]); // Role does not exists in target application ==> create & copy
  end;
  // delete all roles in target application that not exists in source application
  for i := ASourceApplication.Roles.Count - 1 downto 0 do
  begin
    if not ASourceApplication.Roles.Contains(ASourceApplication.Roles[i].Name) then
    begin
      Collection.Catalog.DebugMessage(DEBUG_MESSAGE_DELETE_ROLE, [ASourceApplication.Roles[i].Name, ASourceApplication.Name, Collection.Catalog.Server]);
      Collection.Catalog.ChangeCount := Collection.Catalog.ChangeCount + ASourceApplication.Roles.Delete(i);
    end;
  end;
end;

{ TComAdminApplicationList }

function TComAdminApplicationList.Append(ASourceApplication: TComAdminApplication; const ACreatorString: string; const APassword: string): TComAdminApplication;
var
  LApplication: ICatalogObject;
begin
  LApplication := CatalogCollection.Add as ICatalogObject;
  LApplication.Value[PROPERTY_NAME_NAME] := ASourceApplication.Name;
  Result := TComAdminApplication.Create(Self, LApplication);
  Result.CopyProperties(ASourceApplication, APassword);
  if not ACreatorString.IsEmpty then
    Result.CreatedBy := ACreatorString;
  Catalog.ChangeCount := Catalog.ChangeCount + SaveChanges;
  Result.Key := LApplication.Key;
  Self.Add(Result);
end;

function TComAdminApplicationList.Find(const AName: string; out AApplication: TComAdminApplication): Boolean;
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
  FApplicationProxyRSN := VarToStr(CatalogObject.Value[PROPERTY_NAME_PROXY_RSN]);
  FCISEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_CIS_ENABLED], varBoolean);
  FDCOMEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_DCOM_ENABLED], varBoolean);
  FAuthenticationLevel := VarAsType(CatalogObject.Value[PROPERTY_NAME_DEFAULT_AUTHENTICATION], varLongWord);
  FImpersonationLevel := VarAsType(CatalogObject.Value[PROPERTY_NAME_DEFAULT_IMPERSONATION], varLongWord);
  FDefaultToInternetPorts := VarAsType(CatalogObject.Value[PROPERTY_NAME_DEFAULT_TO_INTERNET], varBoolean);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FDSPartitionLookupEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_DS_PARTITION_LOOKUP], varBoolean);
  FInternetPortsListed := VarAsType(CatalogObject.Value[PROPERTY_NAME_INTERNET_PORTS], varBoolean);
  FIsRouter := VarAsType(CatalogObject.Value[PROPERTY_NAME_IS_ROUTER], varBoolean);
  FLoadBalancingCLSID := VarToStr(CatalogObject.Value[PROPERTY_NAME_LOAD_BALANCING_ID]);
  FLocalPartitionLookupEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_IS_ROUTER], varBoolean);
  FOperatingSystem := VarAsType(CatalogObject.Value[PROPERTY_NAME_OPERATING_SYSTEM], varLongWord);
  FPartitionsEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_PARTITIONS_ENABLED], varBoolean);
  FPorts := VarToStr(CatalogObject.Value[PROPERTY_NAME_PORTS]);
  FResourcePoolingEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_RESOURCE_POOLING], varBoolean);
  if FCISEnabled then // property only available if CIS is enabled
    FRPCProxyEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_RPC_PROXY_ENABLED], varBoolean);
  FSecureReferencesEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_SECURE_REFERENCES], varBoolean);
  FSecurityTrackingEnabled := VarAsType(CatalogObject.Value[PROPERTY_NAME_SECURE_TRACKING], varBoolean);
  FSRPActivateAsActivatorChecks := VarAsType(CatalogObject.Value[PROPERTY_NAME_SRP_ACTIVATE_CHECKS], varBoolean);
  FSRPRunningObjectChecks := VarAsType(CatalogObject.Value[PROPERTY_NAME_SRP_OBJECTS_CHECK], varBoolean);
  FTransactionTimeout := VarAsType(CatalogObject.Value[PROPERTY_NAME_TRANSACTION_TIMEOUT], varLongWord);
end;

procedure TComAdminComputer.SetTransactionTimeout(const Value: Cardinal);
begin
  if InternalCheckRange(1, MAX_TIMEOUT, Value) then
  begin
    FTransactionTimeout := Value;
    if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_TRANSACTION_TIMEOUT) then
      CatalogObject.Value[PROPERTY_NAME_TRANSACTION_TIMEOUT] := Value;
  end;
end;

{ TComAdminDCOMProtocol }

constructor TComAdminDCOMProtocol.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
end;

procedure TComAdminDCOMProtocol.ReadExtendedProperties;
var
  LProtocol: string;
  i: Integer;
begin
  FOrder := VarAsType(CatalogObject.Value[PROPERTY_NAME_ORDER], varLongWord);
  LProtocol := VarToStr(CatalogObject.Value[PROPERTY_NAME_PROTOCOL_CODE]);
  if LProtocol.Equals(COM_ADMIN_PROTOCOLS[COMAdminProtocolTCP]) then
    FProtocolCode := COMAdminProtocolTCP
  else if LProtocol.Equals(COM_ADMIN_PROTOCOLS[COMAdminProtocolHTTP]) then
    FProtocolCode := COMAdminProtocolHTTP
  else if LProtocol.Equals(COM_ADMIN_PROTOCOLS[COMAdminProtocolSPX]) then
    FProtocolCode := COMAdminProtocolSPX
  else
    raise EUnknownProtocolException.Create(ERROR_UNKNOWN_PROTOCOL);
end;

procedure TComAdminDCOMProtocol.SetOrder(const Value: Cardinal);
begin
  FOrder := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_ORDER) then
    CatalogObject.Value[PROPERTY_NAME_ORDER] := Value;
end;

procedure TComAdminDCOMProtocol.SetProtocolCode(const Value: TCOMAdminProtocol);
begin
  FProtocolCode := Value;
  if not CatalogObject.IsPropertyReadOnly(PROPERTY_NAME_PROTOCOL_CODE) then
    CatalogObject.Value[PROPERTY_NAME_PROTOCOL_CODE] := COM_ADMIN_PROTOCOLS[Value];
end;

{ TComAdminDCOMProtocolList }

function TComAdminDCOMProtocolList.GetItem(Index: Integer): TComAdminDCOMProtocol;
begin
  Result := inherited Items[Index] as TComAdminDCOMProtocol;
end;

{ TComAdminEventClass }

constructor TComAdminEventClass.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
end;

procedure TComAdminEventClass.ReadExtendedProperties;
begin
  FBitness := VarAsType(CatalogObject.Value[PROPERTY_NAME_BITNESS], varLongWord);
  FApplication := VarToStr(CatalogObject.Value[PROPERTY_NAME_APPLICATION]);
  FProgID := VarToStr(CatalogObject.Value[PROPERTY_NAME_PROG_ID]);
  FDescription := VarToStr(CatalogObject.Value[PROPERTY_NAME_DESCRIPTION]);
  FCLSID := VarToStr(CatalogObject.Value[PROPERTY_NAME_CLSID]);
  FIsPrivateComponent := VarAsType(CatalogObject.Value[PROPERTY_NAME_IS_PRIVATE_COMPONENT], varBoolean);
end;

{ TComAdminEventClassList }

function TComAdminEventClassList.GetItem(Index: Integer): TComAdminEventClass;
begin
  Result := inherited Items[Index] as TComAdminEventClass;
end;

{ TComAdminCatalog }

constructor TComAdminCatalog.Create(const AServer, AFilter: string; AOnReadEvent: TComAdminReadEvent);
begin
  inherited Create;

  FCatalog := CoCOMAdminCatalog.Create;
  FCopiedLibraries := TList<string>.Create;
  FPasswords := TPasswordList.Create;
  FServer := AServer;

  FOnReadObject := AOnReadEvent;

  if not AServer.IsEmpty then
    FCatalog.Connect(AServer);

  if AFilter.IsEmpty then
    FFilter := DEFAULT_APP_FILTER
  else
    FFilter := AFilter;

  FApplicationCluster := TComAdminApplicationCluster.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_APPLICATION_CLUSTER) as ICatalogCollection);
  FApplications := TComAdminApplicationList.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_APPS) as ICatalogCollection);
  FEventClasses := TComAdminEventClassList.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_EVENT_CLASSES) as ICatalogCollection);
  FPartitions := TComAdminPartitionList.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_PARTITIONS) as ICatalogCollection);
  FProtocols := TComAdminDCOMProtocolList.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_DCOM_PROTOCOLS) as ICatalogCollection);
  FComputer := TComAdminComputer.Create(FCatalog.GetCollection(COLLECTION_NAME_COMPUTER) as ICatalogCollection);

  GetApplications;
  GetEventClasses;
  GetPartitions;
  GetProtocols;
  GetServers;

end;

destructor TComAdminCatalog.Destroy;
begin
  FPasswords.Clear;
  FPasswords.Free;
  FCopiedLibraries.Free;
  FApplications.Free;
  FComputer.Free;
  FEventClasses.Free;
  FPartitions.Free;
  FProtocols.Free;
  FApplicationCluster.Free;
  inherited;
end;

procedure TComAdminCatalog.DebugMessage(const AMessage: string);
begin
  if Assigned(FOnDebug) then
    FOnDebug(AMessage);
end;

procedure TComAdminCatalog.DebugMessage(const AMessage: string; AParams: array of const);
begin
  DebugMessage(Format(AMessage, AParams));
end;

procedure TComAdminCatalog.AddPasswordForIdentity(const AIdentity, APassword: string);
begin
  FPasswords.Add(TPair<string,string>.Create(AIdentity.ToLower, APassword));
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
  LApplication: TComAdminApplication;
begin
  if FApplications.Find(AName, LApplication) then
    ExportApplicationByKey(LApplication.Key, AFilename)
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

procedure TComAdminCatalog.GetEventClasses;
var
  i: Integer;
begin
  for i := 0 to FEventClasses.CatalogCollection.Count - 1 do
    FEventClasses.Add(TComAdminEventClass.Create(FEventClasses, FEventClasses.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminCatalog.GetPartitions;
var
  i: Integer;
begin
  for i := 0 to FPartitions.CatalogCollection.Count - 1 do
    FPartitions.Add(TComAdminPartition.Create(FPartitions, FPartitions.CatalogCollection.Item[i] as ICatalogObject));
end;

function TComAdminCatalog.GetPasswordForIdentity(const AIdentity: string): string;
var
  i: Integer;
begin
  for i := 0 to FPasswords.Count - 1 do
    if FPasswords[i].Key.Equals(AIdentity.ToLower) then
      Exit(FPasswords[i].Value);
  raise EItemNotFoundException.CreateFmt(ERROR_NOT_FOUND, [AIdentity]);
end;

procedure TComAdminCatalog.GetProtocols;
var
  i: Integer;
begin
  for i := 0 to FProtocols.CatalogCollection.Count - 1 do
    FProtocols.Add(TComAdminDCOMProtocol.Create(FProtocols, FProtocols.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminCatalog.GetServers;
var
  i: Integer;
begin
  for i := 0 to FApplicationCluster.CatalogCollection.Count - 1 do
    FApplicationCluster.Add(TComAdminApplicationServer.Create(FApplicationCluster, FApplicationCluster.CatalogCollection.Item[i] as ICatalogObject));
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

function TComAdminCatalog.SyncToServer(const ATargetServer, ACreatorString: string): Integer;
var
  LTargetServerCatalog: TComAdminCatalog;
  LApplication: TComAdminApplication;
  LPassword: string;
  i: Integer;
begin
  DebugMessage(DEBUG_MESSAGE_START_SYNC, [FServer, ATargetServer]);
  LTargetServerCatalog := TComAdminCatalog.Create(ATargetServer, FFilter, FOnReadObject);
  try
    FCopiedLibraries.Clear;
    LTargetServerCatalog.ChangeCount := 0;
    LTargetServerCatalog.CopyLibraries := FCopyLibraries;
    LTargetServerCatalog.LibraryPath := FLibraryPath;
    LTargetServerCatalog.OnDebug := FOnDebug;

    // sync applications from main server to target server
    for i := 0 to FApplications.Count - 1 do
    begin
      DebugMessage(DEBUG_MESSAGE_SYNC_APPLICATION, [FApplications[i].Name]);

      if not FApplications[i].IsIntegratedIdentity then
      begin
        LPassword := GetPasswordForIdentity(FApplications[i].Identity);
        DebugMessage(DEBUG_MESSAGE_GET_PASSWORD, [FApplications[i].Identity]);
      end;

      if LTargetServerCatalog.Applications.Find(FApplications[i].Name, LApplication) then
      begin
        LTargetServerCatalog.Catalog.ShutdownApplication(LApplication.Key);
        LApplication.CopyProperties(FApplications[i], LPassword); // Application exists on target server ==> copy properties
        if not ACreatorString.IsEmpty then
          LApplication.CreatedBy := ACreatorString;
      end else
        LTargetServerCatalog.Applications.Append(FApplications[i], ACreatorString, LPassword); // Application does not exists on target server ==> create & copy
    end;

    // delete all applications on target server that not exists on main server
    for i := LTargetServerCatalog.Applications.Count - 1 downto 0 do
    begin
      if not FApplications.Contains(LTargetServerCatalog.Applications[i].Name) then
        LTargetServerCatalog.ChangeCount := LTargetServerCatalog.ChangeCount + LTargetServerCatalog.Applications.Delete(i);
    end;

    LTargetServerCatalog.Applications.SaveChanges;
    Result := LTargetServerCatalog.ChangeCount;

  finally
    LTargetServerCatalog.Free;
  end;

  DebugMessage(DEBUG_MESSAGE_SYNC_COMPLETED);

end;

end.
