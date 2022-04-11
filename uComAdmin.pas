unit uComAdmin;

// Compile with RTTI
{$M+}

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

  TComAdminReadEvent = procedure (const AObjectType, AObjectName: string) of object;

const
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
  // Interface to check for objects that have a roles collection
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
  private
    function InternalCheckRange(AMinValue, AMaxValue, AValue: Cardinal): Boolean;
    procedure CopyObject(ASourceObject, ATargetObject: TComAdminBaseObject);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    procedure CopyProperties(ABaseClass: TComAdminBaseObject);
    property CatalogCollection: ICatalogCollection read FCatalogCollection;
    property CatalogObject: ICatalogObject read FCatalogObject;
    property Collection: TComAdminBaseList read FCollection;
  published
    property Key: string read FKey write FKey;
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
  private
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
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function GetUsersCollectionName: string;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property Changeable: Boolean read FChangeable write FChangeable default True;
    property Deleteable: Boolean read FDeleteable write FDeleteable default True;
    property Description: string read FDescription write FDescription;
    property ID: string read FID write FID;
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
    procedure ReadExtendedProperties;
  private
    procedure GetRoles;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property AutoComplete: Boolean read FAutoComplete write FAutoComplete default False;
    property Description: string read FDescription write FDescription;
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
    procedure ReadExtendedProperties;
  private
    procedure GetRoles;
    procedure GetMethods;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property CLSID: string read FCLSID;
    property Description: string read FDescription write FDescription;
    property IID: string read FIID;
    property QueuingEnabled: Boolean read FQueuingEnabled write FQueuingEnabled;
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
    procedure SetComponentTransactionTimeout(const Value: Cardinal);
    procedure SetCreationTimeout(const Value: Cardinal);
    procedure SetMaxPoolSize(const Value: Cardinal);
    procedure SetMinPoolSize(const Value: Cardinal);
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Roles: TComAdminRoleList read FRoles write FRoles;
  published
    property AllowInprocSubscribers: Boolean read FAllowInprocSubscribers write FAllowInprocSubscribers default True;
    property ApplicationID: string read FApplicationID write FApplicationID;
    property Bitness: TCOMAdminComponentType read FBitness write FBitness;
    property ComponentAccessChecksEnabled: Boolean read FComponentAccessChecksEnabled write FComponentAccessChecksEnabled default False;
    property ComponentTransactionTimeout: Cardinal read FComponentTransactionTimeout write SetComponentTransactionTimeout default DEFAULT_TRANSACTION_TIMEOUT;
    property ComponentTransactionTimeoutEnabled: Boolean read FComponentTransactionTimeoutEnabled write FComponentTransactionTimeoutEnabled default False;
    property COMTIIntrinsics: Boolean read FCOMTIIntrinsics write FCOMTIIntrinsics default False;
    property ConstructionEnabled: Boolean read FConstructionEnabled write FConstructionEnabled default False;
    property ConstructorString: string read FConstructorString write FConstructorString;
    property CreationTimeout: Cardinal read FCreationTimeout write SetCreationTimeout default DEFAULT_CREATION_TIMEOUT;
    property Description: string read FDescription write FDescription;
    property Dll: string read FDll write FDll;
    property EventTrackingEnabled: Boolean read FEventTrackingEnabled write FEventTrackingEnabled default True;
    property ExceptionClass: string read FExceptionClass write FExceptionClass;
    property FireInParallel: Boolean read FFireInParallel write FFireInParallel default False;
    property IISIntrinsics: Boolean read FIISIntrinsics write FIISIntrinsics default False;
    property InitializeServerApplication: Boolean read FInitializeServerApplication write FInitializeServerApplication default False;
    property IsEnabled: Boolean read FIsEnabled write FIsEnabled default True;
    property IsEventClass: Boolean read FIsEventClass write FIsEventClass default False;
    property IsInstalled: Boolean read FIsInstalled write FIsInstalled default False;
    property IsPrivateComponent: Boolean read FIsPrivateComponent write FIsPrivateComponent default False;
    property JustInTimeActivation: Boolean read FJustInTimeActivation write FJustInTimeActivation default False;
    property LoadBalancingSupported: Boolean read FLoadBalancingSupported write FLoadBalancingSupported default False;
    property MaxPoolSize: Cardinal read FMaxPoolSize write SetMaxPoolSize default DEFAULT_MAX_POOL;
    property MinPoolSize: Cardinal read FMinPoolSize write SetMinPoolSize default 0;
    property MultiInterfacePublisherFilterCLSID: string read FMultiInterfacePublisherFilterCLSID write FMultiInterfacePublisherFilterCLSID;
    property MustRunInClientContext: Boolean read FMustRunInClientContext write FMustRunInClientContext default False;
    property MustRunInDefaultContext: Boolean read FMustRunInDefaultContext write FMustRunInDefaultContext default False;
    property ObjectPoolingEnabled: Boolean read FObjectPoolingEnabled write FObjectPoolingEnabled default False;
    property ProgID: string read FProgID write FProgID;
  end;

  TCOMAdminComponentList = class(TComAdminBaseList)
  strict private
    function GetItem(Index: Integer): TCOMAdminComponent;
    function BuildTargetLibraryName(ASourceComponent: TCOMAdminComponent): string;
  public
    function Append(ASourceComponent: TCOMAdminComponent): TCOMAdminComponent;
    function Find(const AName: string; out AComponent: TCOMAdminComponent): Boolean;
    property Items[Index: Integer]: TCOMAdminComponent read GetItem; default;
  end;

  TComAdminApplication = class(TComAdminBaseObject, IUserCollection)
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
    FComponents: TCOMAdminComponentList;
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
    function GetInstances: TComAdminInstanceList;
    function CopyProperties(ASourceApplication: TCOMAdminApplication): Integer;
    function GetUsersCollectionName: string;
    function InstallComponent(const ALibraryName: string): TCOMAdminComponent;
    property Roles: TComAdminRoleList read FRoles;
  published
    property AccessChecksEnabled: Boolean read FAccessChecksEnabled write SetAccessChecksEnabled default True;
    property AccessChecksLevel: TCOMAdminAccessChecksLevelOptions read FAccessChecksLevel write SetAccessChecksLevel default COMAdminAccessChecksApplicationComponentLevel;
    property Activation: TCOMAdminApplicationActivation read FActivation write SetActivation default COMAdminActivationLocal;
    property AuthenticationCapability: TCOMAdminAuthenticationCapability read FAuthenticationCapability write SetAuthenticationCapability default COMAdminAuthenticationCapabilitiesDynamicCloaking;
    property AuthenticationLevel: TCOMAdminAuthenticationLevel read FAuthenticationLevel write SetAuthenticationLevel default COMAdminAuthenticationDefault;
    property Changeable: Boolean read FChangeable write SetChangeable default True;
    property CommandLine: string read FCommandLine write SetCommandLine;
    property Components: TCOMAdminComponentList read FComponents;
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
    property Instances: TComAdminInstanceList read FInstances;
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
    function Append(ASourceApplication: TComAdminApplication; const ACreatorString: string = ''): TComAdminApplication;
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

  TComAdminCatalog = class(TObject)
  strict private
    FApplications: TComAdminApplicationList;
    FCatalog: ICOMAdminCatalog2;
    FChangeCount: Integer;
    FComputer: TComAdminComputer;
    FDebug: Boolean;
    FFilter: string;
    FLibraryPath: string;
    FOnReadObject: TComAdminReadEvent;
    FPartitions: TComAdminPartitionList;
    FServer: string;
    procedure GetApplications;
    procedure GetPartitions;
    procedure SetFilter(const Value: string);
  public
    constructor Create(const AServer: string; const AFilter: string; AOnReadEvent: TComAdminReadEvent); reintroduce; overload;
    destructor Destroy; override;
    procedure ExportApplication(AIndex: Integer; const AFilename: string);
    procedure ExportApplicationByKey(const AKey, AFilename: string);
    procedure ExportApplicationByName(const AName, AFilename: string);
    function SyncToServer(const ATargetServer: string; const ACreatorString: string = ''): Integer;
    property Applications: TComAdminApplicationList read FApplications;
    property Catalog: ICOMAdminCatalog2 read FCatalog;
    property ChangeCount: Integer read FChangeCount write FChangeCount;
    property Computer: TComAdminComputer read FComputer write FComputer;
    property Debug: Boolean read FDebug write FDebug default False;
    property Filter: string read FFilter write SetFilter;
    property LibraryPath: string read FLibraryPath write FLibraryPath;
    property OnReadObject: TComAdminReadEvent read FOnReadObject write FOnReadObject;
    property Server: string read FServer;
  end;

  EItemNotFoundException = Exception;

implementation

uses
  System.Masks,
  System.Variants,
  Winapi.Windows, System.Rtti, System.TypInfo, System.IOUtils;

const
  COLLECTION_NAME_APPS = 'Applications';
  COLLECTION_NAME_COMPONENTS = 'Components';
  COLLECTION_NAME_COMPONENT_ROLES = 'RolesForComponent';
  COLLECTION_NAME_COMPUTER = 'LocalComputer';
  COLLECTION_NAME_INSTANCES = 'ApplicationInstances';
  COLLECTION_NAME_INTERFACES = 'InterfacesForComponent';
  COLLECTION_NAME_INTERFACE_ROLES = 'RolesForInterface';
  COLLECTION_NAME_METHODS = 'MethodsForInterface';
  COLLECTION_NAME_METHOD_ROLES = 'RolesForMethod';
  COLLECTION_NAME_PARTITIONS = 'Partitions';
  COLLECTION_NAME_PARTITION_ROLES = 'RolesForPartition';
  COLLECTION_NAME_ROLES = 'Roles';
  COLLECTION_NAME_USERS = 'UsersInRole';
  COLLECTION_NAME_USERS_PARTITION = 'UsersInPartitionRole';

  DEFAULT_APP_FILTER = '*';

  ERROR_INVALID_LIBRARY_PATH = 'Invalid library path for server %s';
  ERROR_NOT_FOUND = 'Element %s could not be found in this collection';
  ERROR_OUT_OF_RANGE = 'Value out of range';

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
  PROPERTY_NAME_AUTO_COMPLETE = 'AutoComplete';
  PROPERTY_NAME_BITNESS = 'Bitness';
  PROPERTY_NAME_CHANGEABLE = 'Changeable';
  PROPERTY_NAME_CIS_ENABLED = 'CISEnabled';
  PROPERTY_NAME_CLSID = 'CLSID';
  PROPERTY_NAME_COMMAND_LINE = 'CommandLine';
  PROPERTY_NAME_COMPONENT_ACCESS_CHECKS = 'ComponentAccessChecksEnabled';
  PROPERTY_NAME_COMPONENT_TIMEOUT = 'ComponentTransactionTimeout';
  PROPERTY_NAME_COMPONENT_TIMEOUT_ENABLED = 'ComponentTransactionTimeoutEnabled';
  PROPERTY_NAME_COM_TIINTRINSICS = 'COMTIIntrinsics';
  PROPERTY_NAME_CONCURRENT_APPS = 'ConcurrentApps';
  PROPERTY_NAME_CONSTRUCTION_ENABLED = 'ConstructionEnabled';
  PROPERTY_NAME_CONSTRUCTOR_STRING = 'ConstructorString';
  PROPERTY_NAME_CREATED_BY = 'CreatedBy';
  PROPERTY_NAME_CREATION_TIMEOUT = 'CreationTimeout';
  PROPERTY_NAME_CRM_ENABLED = 'CRMEnabled';
  PROPERTY_NAME_CRM_LOGFILE = 'CRMLogFile';
  PROPERTY_NAME_DCOM_ENABLED = 'DCOMEnabled';
  PROPERTY_NAME_DEFAULT_AUTHENTICATION = 'DefaultAuthenticationLevel';
  PROPERTY_NAME_DEFAULT_IMPERSONATION = 'DefaultImpersonationLevel';
  PROPERTY_NAME_DEFAULT_TO_INTERNET = 'DefaultToInternetPorts';
  PROPERTY_NAME_DELETEABLE = 'Deleteable';
  PROPERTY_NAME_DESCRIPTION = 'Description';
  PROPERTY_NAME_DLL = 'DLL';
  PROPERTY_NAME_DS_PARTITION_LOOKUP = 'DSPartitionLookupEnabled';
  PROPERTY_NAME_DUMP_ENABLED = 'DumpEnabled';
  PROPERTY_NAME_DUMP_EXCEPTION = 'DumpOnException';
  PROPERTY_NAME_DUMP_FAILFAST = 'DumpOnFailfast';
  PROPERTY_NAME_DUMP_PATH = 'DumpPath';
  PROPERTY_NAME_ENABLED = 'IsEnabled';
  PROPERTY_NAME_EVENTS_ENABLED = 'EventsEnabled';
  PROPERTY_NAME_EVENT_TRACKING = 'EventTrackingEnabled';
  PROPERTY_NAME_EXCEPTION_CLASS = 'ExceptionClass';
  PROPERTY_NAME_FIRE_IN_PARALLEL = 'FireInParallel';
  PROPERTY_NAME_IDENTITY = 'Identity';
  PROPERTY_NAME_IID = 'IID';
  PROPERTY_NAME_IIS_INTRINSICS = 'IISIntrinsics';
  PROPERTY_NAME_IMPERSONATION = 'ImpersonationLevel';
  PROPERTY_NAME_INDEX = 'Index';
  PROPERTY_NAME_INIT_SERVER_APPLICATION = 'InitializeServerApplication';
  PROPERTY_NAME_INTERNET_PORTS = 'InternetPortsListed';
  PROPERTY_NAME_IS_EVENT_CLASS = 'IsEventClass';
  PROPERTY_NAME_IS_INSTALLED = 'IsInstalled';
  PROPERTY_NAME_IS_PRIVATE_COMPONENT = 'IsPrivateComponent';
  PROPERTY_NAME_IS_ROUTER = 'IsRouter';
  PROPERTY_NAME_JUST_IN_TIME = 'JustInTimeActivation';
  PROPERTY_NAME_KEY = 'Key';
  PROPERTY_NAME_LOAD_BALANCING = 'LoadBalancingSupported';
  PROPERTY_NAME_LOAD_BALANCING_ID = 'LoadBalancingCLSID';
  PROPERTY_NAME_MAX_DUMPS = 'MaxDumpCount';
  PROPERTY_NAME_MAX_POOL_SIZE = 'MaxPoolSize';
  PROPERTY_NAME_MIN_POOL_SIZE = 'MinPoolSize';
  PROPERTY_NAME_MI_FILTER_CLSID = 'MultiInterfacePublisherFilterCLSID';
  PROPERTY_NAME_MUST_RUN_CLIENT_CONTEXT = 'MustRunInClientContext';
  PROPERTY_NAME_MUST_RUN_DEFAULT_CONTEXT = 'MustRunInDefaultContext';
  PROPERTY_NAME_NAME = 'Name';
  PROPERTY_NAME_OBJECT_POOLING = 'ObjectPoolingEnabled';
  PROPERTY_NAME_OPERATING_SYSTEM = 'OperatingSystem';
  PROPERTY_NAME_PARTITION_ID = 'AppPartitionID';
  PROPERTY_NAME_PARTITION_LOOKUP = 'LocalPartitionLookupEnabled';
  PROPERTY_NAME_PARTITIONS_ENABLED = 'PartitionsEnabled';
  PROPERTY_NAME_PASSWORD = 'Password';
  PROPERTY_NAME_PAUSED = 'IsPaused';
  PROPERTY_NAME_PORTS = 'Ports';
  PROPERTY_NAME_PROCESSID = 'ProcessID';
  PROPERTY_NAME_PROG_ID = 'ProgID';
  PROPERTY_NAME_PROXY_RSN = 'ApplicationProxyRSN';
  PROPERTY_NAME_PROXY_SERVER_NAME = 'ApplicationProxyServerName';
  PROPERTY_NAME_QC_AUTHENTICATE = 'QCAuthenticateMsgs';
  PROPERTY_NAME_QC_MAXTHREADS = 'QCListenerMaxThreads';
  PROPERTY_NAME_QUEUE_LISTENER = 'QueueListenerEnabled';
  PROPERTY_NAME_QUEUING_ENABLED = 'QueuingEnabled';
  PROPERTY_NAME_QUEUING_SUPPORTED = 'QueuingSupported';
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
  PROPERTY_NAME_USER = 'User';

{ TComAdminBaseObject }

procedure TComAdminBaseObject.CopyObject(ASourceObject, ATargetObject: TComAdminBaseObject);
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
        AValue := LProperty.GetValue(ASource.AsObject);
        LProperty.SetValue(ATarget.AsObject, AValue);
      end;
    end;

  finally
    LRttiContext.Free;
  end;
end;

procedure TComAdminBaseObject.CopyProperties(ABaseClass: TComAdminBaseObject);
begin
  CopyObject(ABaseClass, Self);
end;

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

function TComAdminBaseObject.InternalCheckRange(AMinValue, AMaxValue, AValue: Cardinal): Boolean;
begin
  Result := (AValue >= AMinValue) and (AValue <= AMaxValue);
  if not Result then
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

{ TComAdminBaseList }

constructor TComAdminBaseList.Create(AOwner: TComAdminBaseObject; ACatalog: TComAdminCatalog; ACatalogCollection: ICatalogCollection);
begin
  inherited Create(True);
  FOwner := AOwner;
  FCatalog := ACatalog;
  FCatalogCollection := ACatalogCollection;
  if Assigned(FCatalogCollection) then
  begin
    FName := FCatalogCollection.Name;
    FCatalogCollection.Populate;
  end;
end;

function TComAdminBaseList.Delete(Index: Integer): Integer;
begin
  CatalogCollection.Remove(GetIndexByKey(Items[Index].Key));
  Result := CatalogCollection.SaveChanges;
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
  Result := FCatalogCollection.SaveChanges;
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
  Catalog.ChangeCount := Catalog.ChangeCount + CatalogCollection.SaveChanges;
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

function TComAdminRole.CopyProperties(ASourceRole: TComAdminRole): Integer;
begin
  inherited CopyProperties(ASourceRole);

  // Changes must be saved before any sub-collections can be updated
  Result := CatalogCollection.SaveChanges;

  // Synchronize users from source role
  SyncUsers(ASourceRole);
end;

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
  try
    FUsers.Free;
  except
    on E:Exception do
      OutputDebugString(PChar('#Error in destructor for ' + Self.ClassName + ' (' + Collection.Owner.Name + '.' + Collection.Name + '.' + Self.Name + '): ' + E.Message));
  end;
  inherited;
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
    CatalogObject.Value[PROPERTY_NAME_DESCRIPTION] := FDescription;
end;

procedure TComAdminRole.SyncUsers(ASourceRole: TComAdminRole);
var
  i: Integer;
  LUser: TComAdminUser;
begin
  for i := 0 to ASourceRole.Users.Count - 1 do
  begin
    if not FUsers.Find(ASourceRole.Users[i].Name, LUser) then
      LUser := FUsers.Append(ASourceRole.Users[i]); // User does not exists in target role ==> create & copy
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
  Catalog.ChangeCount := Catalog.ChangeCount + CatalogCollection.SaveChanges;
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
    FMethods.Add(TComAdminRole.Create(FMethods, FMethods.CatalogCollection.Item[i] as ICatalogObject));
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

{ TComAdminInterfaceList }

function TComAdminInterfaceList.GetItem(Index: Integer): TComAdminInterface;
begin
  Result := inherited Items[Index] as TComAdminInterface;
end;

{ TCOMAdminComponent }

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

procedure TCOMAdminComponent.SetComponentTransactionTimeout(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_TIMEOUT, Value) then
    FComponentTransactionTimeout := Value
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

procedure TCOMAdminComponent.SetCreationTimeout(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAXLONG, Value) then
    FCreationTimeout := Value
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

procedure TCOMAdminComponent.SetMaxPoolSize(const Value: Cardinal);
begin
  if InternalCheckRange(1, MAX_POOL_SIZE, Value) then
    FMaxPoolSize := Value
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

procedure TCOMAdminComponent.SetMinPoolSize(const Value: Cardinal);
begin
  if InternalCheckRange(0, MAX_POOL_SIZE, Value) then
    FMinPoolSize := Value
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

{ TCOMAdminComponentList }

function TCOMAdminComponentList.Append(ASourceComponent: TCOMAdminComponent): TCOMAdminComponent;
var
  LLibraryName, LTargetLibrary: string;
begin
  LTargetLibrary := BuildTargetLibraryName(ASourceComponent);
  if not FileExists(LTargetLibrary) then
    TFile.Copy(ASourceComponent.Dll, LTargetLibrary);
  LLibraryName := TPath.Combine(Catalog.LibraryPath, ExtractFileName(ASourceComponent.Dll));
  Result := (Owner as TComAdminApplication).InstallComponent(LLibraryName);
  Result.CopyProperties(ASourceComponent);
  Catalog.ChangeCount := Catalog.ChangeCount + CatalogCollection.SaveChanges;
end;

function TCOMAdminComponentList.BuildTargetLibraryName(ASourceComponent: TCOMAdminComponent): string;
begin
  if Catalog.LibraryPath.IsEmpty then
    raise Exception.CreateFmt(ERROR_INVALID_LIBRARY_PATH, [Catalog.Server]);
  Result := Format('\\%s\%s', [Catalog.Server, TPath.Combine(Catalog.LibraryPath, ExtractFileName(ASourceComponent.Dll)).Replace(':','$')]);
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

function TComAdminApplication.CopyProperties(ASourceApplication: TCOMAdminApplication): Integer;
begin

  inherited CopyProperties(ASourceApplication);

  // Changes must be saved before any sub-collections can be updated
  Result := CatalogCollection.SaveChanges;

  // Synchronize roles from source application
  SyncRoles(ASourceApplication);
  SyncComponents(ASourceApplication);

end;

constructor TComAdminApplication.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create(ACollection, ACatalogObject);
  ReadExtendedProperties;
  // Create List objects
  FInstances := TComAdminInstanceList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, Key) as ICatalogCollection);
  GetInstances;
  FRoles := TComAdminRoleList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_ROLES, Key) as ICatalogCollection);
  GetRoles;
  FComponents := TCOMAdminComponentList.Create(Self, ACollection.Catalog, ACollection.CatalogCollection.GetCollection(COLLECTION_NAME_COMPONENTS, Key) as ICatalogCollection);
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

function TComAdminApplication.GetInstances: TComAdminInstanceList;
var
  Collection: ICatalogCollection;
  Instance: TComAdminInstance;
  i: Integer;
begin
  FInstances.Clear;
  Collection := CatalogCollection.GetCollection(COLLECTION_NAME_INSTANCES, Key) as ICatalogCollection;
  Collection.Populate;
  for i := 0 to Collection.Count - 1 do
  begin
    Instance := TComAdminInstance.Create(FInstances, Collection.Item[i] as ICatalogObject);
    FInstances.Add(Instance);
  end;
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
    if FComponents.Find(ASourceApplication.Components[i].Name, LComponent) then
    begin
      LComponent.CopyProperties(ASourceApplication.Components[i]);
    end else
      LComponent := FComponents.Append(ASourceApplication.Components[i]); // Component does not exists in target application ==> create & copy
  end;
  // delete all components in target application that not exists in source application
  for i := ASourceApplication.Components.Count - 1 downto 0 do
  begin
    if not ASourceApplication.Components.Find(ASourceApplication.Components[i].Name, LComponent) then
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
    if FRoles.Find(ASourceApplication.Roles[i].Name, LRole) then
      LRole.CopyProperties(ASourceApplication.Roles[i])
    else
      LRole := FRoles.Append(ASourceApplication.Roles[i]); // Role does not exists in target application ==> create & copy
  end;
  // delete all roles in target application that not exists in source application
  for i := ASourceApplication.Roles.Count - 1 downto 0 do
  begin
    if not ASourceApplication.Roles.Find(ASourceApplication.Roles[i].Name, LRole) then
      Collection.Catalog.ChangeCount := Collection.Catalog.ChangeCount + ASourceApplication.Roles.Delete(i);
  end;
end;

{ TComAdminApplicationList }

function TComAdminApplicationList.Append(ASourceApplication: TComAdminApplication; const ACreatorString: string): TComAdminApplication;
var
  LApplication: ICatalogObject;
begin
  LApplication := CatalogCollection.Add as ICatalogObject;
  LApplication.Value[PROPERTY_NAME_NAME] := ASourceApplication.Name;
  Result := TComAdminApplication.Create(Self, LApplication);
  Result.CopyProperties(ASourceApplication);
  if not ACreatorString.IsEmpty then
    Result.CreatedBy := ACreatorString;
  Catalog.ChangeCount := Catalog.ChangeCount + CatalogCollection.SaveChanges;
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
    FTransactionTimeout := Value
  else
    raise EArgumentOutOfRangeException.Create(ERROR_OUT_OF_RANGE);
end;

{ TComAdminCatalog }

constructor TComAdminCatalog.Create(const AServer, AFilter: string; AOnReadEvent: TComAdminReadEvent);
begin
  inherited Create;

  FCatalog := CoCOMAdminCatalog.Create;
  FServer := AServer;

  FOnReadObject := AOnReadEvent;

  if not AServer.IsEmpty then
    FCatalog.Connect(AServer);

  if AFilter.IsEmpty then
    FFilter := DEFAULT_APP_FILTER
  else
    FFilter := AFilter;

  FApplications := TComAdminApplicationList.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_APPS) as ICatalogCollection);
  FPartitions := TComAdminPartitionList.Create(nil, Self, FCatalog.GetCollection(COLLECTION_NAME_PARTITIONS) as ICatalogCollection);
  FComputer := TComAdminComputer.Create(FCatalog.GetCollection(COLLECTION_NAME_COMPUTER) as ICatalogCollection);

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

function TComAdminCatalog.SyncToServer(const ATargetServer: string; const ACreatorString: string): Integer;
var
  LTargetServerCatalog: TComAdminCatalog;
  LApplication: TComAdminApplication;
  i: Integer;
begin
  LTargetServerCatalog := TComAdminCatalog.Create(ATargetServer, FFilter, FOnReadObject);
  try
    LTargetServerCatalog.LibraryPath := FLibraryPath;
    LTargetServerCatalog.ChangeCount := 0;
    // sync applications from main server to target server
    for i := 0 to FApplications.Count - 1 do
    begin
      if LTargetServerCatalog.Applications.Find(FApplications[i].Name, LApplication) then
      begin
        LApplication.CopyProperties(FApplications[i]); // Application exists on target server ==> copy properties
        if not ACreatorString.IsEmpty then
          LApplication.CreatedBy := ACreatorString;
      end else
        LTargetServerCatalog.Applications.Append(FApplications[i], ACreatorString); // Application does not exists on target server ==> create & copy
    end;
    // delete all applications on target server that not exists on main server
    for i := LTargetServerCatalog.Applications.Count - 1 downto 0 do
    begin
      if not FApplications.Find(LTargetServerCatalog.Applications[i].Name, LApplication) then
        LTargetServerCatalog.ChangeCount := LTargetServerCatalog.ChangeCount + LTargetServerCatalog.Applications.Delete(i);
    end;
    LTargetServerCatalog.Applications.CatalogCollection.SaveChanges;
    Result := LTargetServerCatalog.ChangeCount;
  finally
    LTargetServerCatalog.Free;
  end;
end;

end.
