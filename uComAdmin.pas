unit uComAdmin;

// https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections

interface

uses
  System.Generics.Collections,
  System.SysUtils,
  COMAdmin_TLB;

type
  // forward declaration of internally used classes
  TComAdminBaseList = class;

  // generic base class for all single objects
  TComAdminBaseObject = class(TObject)
  private
    FCatalogObject: ICatalogObject;
    FCatalogCollection: ICatalogCollection;
    FKey: string;
    FName: string;
    FDescription: string;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    property Name: string read FName write FName;
    property Key: string read FKey write FKey;
    property Descrition: string read FDescription write FDescription;
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
    FUsers: TComAdminUserList;
    procedure GetUsers;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
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
    procedure GetRoles;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    function GetInstances: TComAdminInstanceList;
    procedure Shutdown;
    property Roles: TComAdminRoleList read FRoles;
    property Instances: TComAdminInstanceList read FInstances;
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

  EItemNotFound = Exception;

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
  PROPERTY_NAME_DESCRIPTION = 'Description';
  PROPERTY_NAME_RECYCLED = 'HasRecycled';
  PROPERTY_NAME_PAUSED = 'IsPaused';
  PROPERTY_NAME_PROCESSID = 'ProcessID';
  ERROR_NOT_FOUND = 'Das Element %s wurde in der Auflistung nicht gefunden.';

{ TComAdminBaseObject }

constructor TComAdminBaseObject.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create;
  FCatalogObject := ACatalogObject;
  FCatalogCollection := ACollection.CatalogCollection;
  if (FCatalogCollection.Name <> COLLECTION_NAME_USERS) and
     (FCatalogCollection.Name <> COLLECTION_NAME_INSTANCES)  then // some objects do not contain a description
    FDescription := VarToStrDef(FCatalogObject.Value[PROPERTY_NAME_DESCRIPTION], '');
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

procedure TComAdminApplication.GetRoles;
var
  i: Integer;
begin
  for i := 0 to FRoles.CatalogCollection.Count - 1 do
    FRoles.Add(TComAdminRole.Create(FRoles, FRoles.CatalogCollection.Item[i] as ICatalogObject));
end;

procedure TComAdminApplication.Shutdown;
var
  ProcessHandle: THandle;
  Instance: TComAdminBaseObject;
begin
  GetInstances;
  for Instance in FInstances do
  begin
    ProcessHandle := OpenProcess(PROCESS_TERMINATE, False, (Instance as TComAdminInstance).FProcessID);
    TerminateProcess(ProcessHandle, 0);
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
    raise EItemNotFound.CreateFmt(ERROR_NOT_FOUND, [AName]);
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
