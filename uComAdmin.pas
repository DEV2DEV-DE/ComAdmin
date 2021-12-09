unit uComAdmin;

// https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections

interface

uses
  System.SysUtils,
  System.Generics.Collections,
  COMAdmin_TLB;

type
  TComAdminBaseList = class;

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

  TComAdminBaseList = class(TObjectList<TComAdminBaseObject>)
  private
    FCatalogCollection: ICatalogCollection;
  public
    constructor Create(ACatalogCollection: ICatalogCollection); reintroduce;
    property CatalogCollection: ICatalogCollection read FCatalogCollection write FCatalogCollection;
  end;

  TComAdminUser = class(TComAdminBaseObject);

  TComAdminUserList = class(TComAdminBaseList)
  private
    function GetItem(Index: Integer): TComAdminUser;
  public
    property Items[Index: Integer]: TComAdminUser read GetItem; default;
  end;

  TComAdminRole = class(TComAdminBaseObject)
  private
    FUsers: TComAdminUserList;
    procedure GetUsers;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Users: TComAdminUserList read FUsers write FUsers;
  end;

  TComAdminRoleList = class(TComAdminBaseList)
  private
    function GetItem(Index: Integer): TComAdminRole;
  public
    property Items[Index: Integer]: TComAdminRole read GetItem; default;
  end;

  TComAdminApplication = class(TComAdminBaseObject)
  private
    FRoles: TComAdminRoleList;
    procedure GetRoles;
  public
    constructor Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject); reintroduce;
    destructor Destroy; override;
    property Roles: TComAdminRoleList read FRoles;
  end;

  TComAdminApplicationList = class(TComAdminBaseList)
  private
    function GetItem(Index: Integer): TComAdminApplication;
  public
    function Find(const AName: string; var AApplication: TComAdminApplication): Boolean;
    property Items[Index: Integer]: TComAdminApplication read GetItem; default;
  end;

  TComAdminCatalog = class(TObject)
  private
    FCatalog: ICOMAdminCatalog2;
    FApplications: TComAdminApplicationList;
    FFilter: string;
    procedure GetApplications;
  public
    constructor Create(const AServer: string); reintroduce;
    destructor Destroy; override;
    procedure ExportApplication(AIndex: Integer; const AFilename: string);
    procedure ExportApplicationByKey(const AKey, AFilename: string);
    procedure ExportApplicationByName(const AName, AFilename: string);
    property Applications: TComAdminApplicationList read FApplications;
    property Filter: string read FFilter write FFilter;
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
  DEFAULT_APP_FILTER = 'ProdLog-*';
  PROPERTY_NAME_DESCRIPTION = 'Description';
  ERROR_NOT_FOUND = 'Das Element %s wurde in der Auflistung nicht gefunden.';

{ TComAdminBaseObject }

constructor TComAdminBaseObject.Create(ACollection: TComAdminBaseList; ACatalogObject: ICatalogObject);
begin
  inherited Create;
  FCatalogObject := ACatalogObject;
  FCatalogCollection := ACollection.CatalogCollection;
  if FCatalogCollection.Name <> COLLECTION_NAME_USERS then // Die User-Objekte enthalten keine Beschreibung
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
  GetRoles;
end;

destructor TComAdminApplication.Destroy;
begin
  FRoles.Free;
  inherited;
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

end.
