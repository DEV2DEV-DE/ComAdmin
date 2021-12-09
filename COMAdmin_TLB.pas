unit COMAdmin_TLB;

// ************************************************************************ //
// WARNUNG
// -------
// Die in dieser Datei deklarierten Typen wurden aus Daten einer Typbibliothek
// generiert. Wenn diese Typbibliothek explizit oder indirekt (über eine
// andere Typbibliothek) reimportiert wird oder wenn der Befehl
// 'Aktualisieren' im Typbibliotheks-Editor während des Bearbeitens der
// Typbibliothek aktiviert ist, wird der Inhalt dieser Datei neu generiert und
// alle manuell vorgenommenen Änderungen gehen verloren.                                        
// ************************************************************************ //

// $Rev: 98336 $
// Datei am 08.12.2021 20:01:41 erzeugt aus der unten beschriebenen Typbibliothek.

// ************************************************************************  //
// Typbib.: C:\Windows\SysWOW64\Com\comadmin.dll (1)
// LIBID: {F618C513-DFB8-11D1-A2CF-00805FC79235}
// LCID: 0
// Hilfedatei: 
// Hilfe-String: COM + 1.0 Admin Type Library
// Liste der Abhäng.: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit muss ohne Typüberprüfung für Zeiger compiliert werden.  
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;
  
// *********************************************************************//
// In der Typbibliothek deklarierte GUIDS. Die folgenden Präfixe werden verwendet:        
//   Typbibliotheken      : LIBID_xxxx                                      
//   CoClasses            : CLASS_xxxx                                      
//   DISPInterfaces       : DIID_xxxx                                       
//   Nicht-DISP-Interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // Haupt- und Nebenversionen der Typbibliothek
  COMAdminMajorVersion = 1;
  COMAdminMinorVersion = 0;

  LIBID_COMAdmin: TGUID = '{F618C513-DFB8-11D1-A2CF-00805FC79235}';

  IID_ICOMAdminCatalog: TGUID = '{DD662187-DFC2-11D1-A2CF-00805FC79235}';
  IID_ICOMAdminCatalog2: TGUID = '{790C6E0B-9194-4CC9-9426-A48A63185696}';
  CLASS_COMAdminCatalog: TGUID = '{F618C514-DFB8-11D1-A2CF-00805FC79235}';
  IID_ICatalogObject: TGUID = '{6EB22871-8A19-11D0-81B6-00A0C9231C29}';
  CLASS_COMAdminCatalogObject: TGUID = '{F618C515-DFB8-11D1-A2CF-00805FC79235}';
  IID_ICatalogCollection: TGUID = '{6EB22872-8A19-11D0-81B6-00A0C9231C29}';
  CLASS_COMAdminCatalogCollection: TGUID = '{F618C516-DFB8-11D1-A2CF-00805FC79235}';

// *********************************************************************//
// Deklaration von in der Typbibliothek definierten Aufzählungen                     
// *********************************************************************//
// Konstanten für enum COMAdminInUse
type
  COMAdminInUse = TOleEnum;
const
  COMAdminNotInUse = $00000000;
  COMAdminInUseByCatalog = $00000001;
  COMAdminInUseByRegistryUnknown = $00000002;
  COMAdminInUseByRegistryProxyStub = $00000003;
  COMAdminInUseByRegistryTypeLib = $00000004;
  COMAdminInUseByRegistryClsid = $00000005;

// Konstanten für enum COMAdminComponentType
type
  COMAdminComponentType = TOleEnum;
const
  COMAdmin32BitComponent = $00000001;
  COMAdmin64BitComponent = $00000002;

// Konstanten für enum COMAdminApplicationInstallOptions
type
  COMAdminApplicationInstallOptions = TOleEnum;
const
  COMAdminInstallNoUsers = $00000000;
  COMAdminInstallUsers = $00000001;
  COMAdminInstallForceOverwriteOfFiles = $00000002;

// Konstanten für enum COMAdminApplicationExportOptions
type
  COMAdminApplicationExportOptions = TOleEnum;
const
  COMAdminExportNoUsers = $00000000;
  COMAdminExportUsers = $00000001;
  COMAdminExportApplicationProxy = $00000002;
  COMAdminExportForceOverwriteOfFiles = $00000004;
  COMAdminExportIn10Format = $00000010;

// Konstanten für enum COMAdminThreadingModels
type
  COMAdminThreadingModels = TOleEnum;
const
  COMAdminThreadingModelApartment = $00000000;
  COMAdminThreadingModelFree = $00000001;
  COMAdminThreadingModelMain = $00000002;
  COMAdminThreadingModelBoth = $00000003;
  COMAdminThreadingModelNeutral = $00000004;
  COMAdminThreadingModelNotSpecified = $00000005;

// Konstanten für enum COMAdminTransactionOptions
type
  COMAdminTransactionOptions = TOleEnum;
const
  COMAdminTransactionIgnored = $00000000;
  COMAdminTransactionNone = $00000001;
  COMAdminTransactionSupported = $00000002;
  COMAdminTransactionRequired = $00000003;
  COMAdminTransactionRequiresNew = $00000004;

// Konstanten für enum COMAdminTxIsolationLevelOptions
type
  COMAdminTxIsolationLevelOptions = TOleEnum;
const
  COMAdminTxIsolationLevelAny = $00000000;
  COMAdminTxIsolationLevelReadUnCommitted = $00000001;
  COMAdminTxIsolationLevelReadCommitted = $00000002;
  COMAdminTxIsolationLevelRepeatableRead = $00000003;
  COMAdminTxIsolationLevelSerializable = $00000004;

// Konstanten für enum COMAdminSynchronizationOptions
type
  COMAdminSynchronizationOptions = TOleEnum;
const
  COMAdminSynchronizationIgnored = $00000000;
  COMAdminSynchronizationNone = $00000001;
  COMAdminSynchronizationSupported = $00000002;
  COMAdminSynchronizationRequired = $00000003;
  COMAdminSynchronizationRequiresNew = $00000004;

// Konstanten für enum COMAdminActivationOptions
type
  COMAdminActivationOptions = TOleEnum;
const
  COMAdminActivationInproc = $00000000;
  COMAdminActivationLocal = $00000001;

// Konstanten für enum COMAdminAccessChecksLevelOptions
type
  COMAdminAccessChecksLevelOptions = TOleEnum;
const
  COMAdminAccessChecksApplicationLevel = $00000000;
  COMAdminAccessChecksApplicationComponentLevel = $00000001;

// Konstanten für enum COMAdminAuthenticationLevelOptions
type
  COMAdminAuthenticationLevelOptions = TOleEnum;
const
  COMAdminAuthenticationDefault = $00000000;
  COMAdminAuthenticationNone = $00000001;
  COMAdminAuthenticationConnect = $00000002;
  COMAdminAuthenticationCall = $00000003;
  COMAdminAuthenticationPacket = $00000004;
  COMAdminAuthenticationIntegrity = $00000005;
  COMAdminAuthenticationPrivacy = $00000006;

// Konstanten für enum COMAdminImpersonationLevelOptions
type
  COMAdminImpersonationLevelOptions = TOleEnum;
const
  COMAdminImpersonationAnonymous = $00000001;
  COMAdminImpersonationIdentify = $00000002;
  COMAdminImpersonationImpersonate = $00000003;
  COMAdminImpersonationDelegate = $00000004;

// Konstanten für enum COMAdminAuthenticationCapabilitiesOptions
type
  COMAdminAuthenticationCapabilitiesOptions = TOleEnum;
const
  COMAdminAuthenticationCapabilitiesNone = $00000000;
  COMAdminAuthenticationCapabilitiesSecureReference = $00000002;
  COMAdminAuthenticationCapabilitiesStaticCloaking = $00000020;
  COMAdminAuthenticationCapabilitiesDynamicCloaking = $00000040;

// Konstanten für enum COMAdminOS
type
  COMAdminOS = TOleEnum;
const
  COMAdminOSNotInitialized = $00000000;
  COMAdminOSWindows3_1 = $00000001;
  COMAdminOSWindows9x = $00000002;
  COMAdminOSWindows2000 = $00000003;
  COMAdminOSWindows2000AdvancedServer = $00000004;
  COMAdminOSWindows2000Unknown = $00000005;
  COMAdminOSUnknown = $00000006;
  COMAdminOSWindowsXPPersonal = $0000000B;
  COMAdminOSWindowsXPProfessional = $0000000C;
  COMAdminOSWindowsNETStandardServer = $0000000D;
  COMAdminOSWindowsNETEnterpriseServer = $0000000E;
  COMAdminOSWindowsNETDatacenterServer = $0000000F;
  COMAdminOSWindowsNETWebServer = $00000010;
  COMAdminOSWindowsLonghornPersonal = $00000011;
  COMAdminOSWindowsLonghornProfessional = $00000012;
  COMAdminOSWindowsLonghornStandardServer = $00000013;
  COMAdminOSWindowsLonghornEnterpriseServer = $00000014;
  COMAdminOSWindowsLonghornDatacenterServer = $00000015;
  COMAdminOSWindowsLonghornWebServer = $00000016;
  COMAdminOSWindows7Personal = $00000017;
  COMAdminOSWindows7Professional = $00000018;
  COMAdminOSWindows7StandardServer = $00000019;
  COMAdminOSWindows7EnterpriseServer = $0000001A;
  COMAdminOSWindows7DatacenterServer = $0000001B;
  COMAdminOSWindows7WebServer = $0000001C;
  COMAdminOSWindows8Personal = $0000001D;
  COMAdminOSWindows8Professional = $0000001E;
  COMAdminOSWindows8StandardServer = $0000001F;
  COMAdminOSWindows8EnterpriseServer = $00000020;
  COMAdminOSWindows8DatacenterServer = $00000021;
  COMAdminOSWindows8WebServer = $00000022;
  COMAdminOSWindowsBluePersonal = $00000023;
  COMAdminOSWindowsBlueProfessional = $00000024;
  COMAdminOSWindowsBlueStandardServer = $00000025;
  COMAdminOSWindowsBlueEnterpriseServer = $00000026;
  COMAdminOSWindowsBlueDatacenterServer = $00000027;
  COMAdminOSWindowsBlueWebServer = $00000028;

// Konstanten für enum COMAdminServiceOptions
type
  COMAdminServiceOptions = TOleEnum;
const
  COMAdminServiceLoadBalanceRouter = $00000001;

// Konstanten für enum COMAdminServiceStatusOptions
type
  COMAdminServiceStatusOptions = TOleEnum;
const
  COMAdminServiceStopped = $00000000;
  COMAdminServiceStartPending = $00000001;
  COMAdminServiceStopPending = $00000002;
  COMAdminServiceRunning = $00000003;
  COMAdminServiceContinuePending = $00000004;
  COMAdminServicePausePending = $00000005;
  COMAdminServicePaused = $00000006;
  COMAdminServiceUnknownState = $00000007;

// Konstanten für enum COMAdminQCMessageAuthenticateOptions
type
  COMAdminQCMessageAuthenticateOptions = TOleEnum;
const
  COMAdminQCMessageAuthenticateSecureApps = $00000000;
  COMAdminQCMessageAuthenticateOff = $00000001;
  COMAdminQCMessageAuthenticateOn = $00000002;

// Konstanten für enum COMAdminFileFlags
type
  COMAdminFileFlags = TOleEnum;
const
  COMAdminFileFlagLoadable = $00000001;
  COMAdminFileFlagCOM = $00000002;
  COMAdminFileFlagContainsPS = $00000004;
  COMAdminFileFlagContainsComp = $00000008;
  COMAdminFileFlagContainsTLB = $00000010;
  COMAdminFileFlagSelfReg = $00000020;
  COMAdminFileFlagSelfUnReg = $00000040;
  COMAdminFileFlagUnloadableDLL = $00000080;
  COMAdminFileFlagDoesNotExist = $00000100;
  COMAdminFileFlagAlreadyInstalled = $00000200;
  COMAdminFileFlagBadTLB = $00000400;
  COMAdminFileFlagGetClassObjFailed = $00000800;
  COMAdminFileFlagClassNotAvailable = $00001000;
  COMAdminFileFlagRegistrar = $00002000;
  COMAdminFileFlagNoRegistrar = $00004000;
  COMAdminFileFlagDLLRegsvrFailed = $00008000;
  COMAdminFileFlagRegTLBFailed = $00010000;
  COMAdminFileFlagRegistrarFailed = $00020000;
  COMAdminFileFlagError = $00040000;

// Konstanten für enum COMAdminComponentFlags
type
  COMAdminComponentFlags = TOleEnum;
const
  COMAdminCompFlagTypeInfoFound = $00000001;
  COMAdminCompFlagCOMPlusPropertiesFound = $00000002;
  COMAdminCompFlagProxyFound = $00000004;
  COMAdminCompFlagInterfacesFound = $00000008;
  COMAdminCompFlagAlreadyInstalled = $00000010;
  COMAdminCompFlagNotInApplication = $00000020;

// Konstanten für enum COMAdminErrorCodes
type
  COMAdminErrorCodes = TOleEnum;
const
  COMAdminErrObjectErrors = $80110401;
  COMAdminErrObjectInvalid = $80110402;
  COMAdminErrKeyMissing = $80110403;
  COMAdminErrAlreadyInstalled = $80110404;
  COMAdminErrAppFileWriteFail = $80110407;
  COMAdminErrAppFileReadFail = $80110408;
  COMAdminErrAppFileVersion = $80110409;
  COMAdminErrBadPath = $8011040A;
  COMAdminErrApplicationExists = $8011040B;
  COMAdminErrRoleExists = $8011040C;
  COMAdminErrCantCopyFile = $8011040D;
  COMAdminErrNoUser = $8011040F;
  COMAdminErrInvalidUserids = $80110410;
  COMAdminErrNoRegistryCLSID = $80110411;
  COMAdminErrBadRegistryProgID = $80110412;
  COMAdminErrAuthenticationLevel = $80110413;
  COMAdminErrUserPasswdNotValid = $80110414;
  COMAdminErrCLSIDOrIIDMismatch = $80110418;
  COMAdminErrRemoteInterface = $80110419;
  COMAdminErrDllRegisterServer = $8011041A;
  COMAdminErrNoServerShare = $8011041B;
  COMAdminErrDllLoadFailed = $8011041D;
  COMAdminErrBadRegistryLibID = $8011041E;
  COMAdminErrAppDirNotFound = $8011041F;
  COMAdminErrRegistrarFailed = $80110423;
  COMAdminErrCompFileDoesNotExist = $80110424;
  COMAdminErrCompFileLoadDLLFail = $80110425;
  COMAdminErrCompFileGetClassObj = $80110426;
  COMAdminErrCompFileClassNotAvail = $80110427;
  COMAdminErrCompFileBadTLB = $80110428;
  COMAdminErrCompFileNotInstallable = $80110429;
  COMAdminErrNotChangeable = $8011042A;
  COMAdminErrNotDeletable = $8011042B;
  COMAdminErrSession = $8011042C;
  COMAdminErrCompMoveLocked = $8011042D;
  COMAdminErrCompMoveBadDest = $8011042E;
  COMAdminErrRegisterTLB = $80110430;
  COMAdminErrSystemApp = $80110433;
  COMAdminErrCompFileNoRegistrar = $80110434;
  COMAdminErrCoReqCompInstalled = $80110435;
  COMAdminErrServiceNotInstalled = $80110436;
  COMAdminErrPropertySaveFailed = $80110437;
  COMAdminErrObjectExists = $80110438;
  COMAdminErrComponentExists = $80110439;
  COMAdminErrRegFileCorrupt = $8011043B;
  COMAdminErrPropertyOverflow = $8011043C;
  COMAdminErrNotInRegistry = $8011043E;
  COMAdminErrObjectNotPoolable = $8011043F;
  COMAdminErrApplidMatchesClsid = $80110446;
  COMAdminErrRoleDoesNotExist = $80110447;
  COMAdminErrStartAppNeedsComponents = $80110448;
  COMAdminErrRequiresDifferentPlatform = $80110449;
  COMAdminErrQueuingServiceNotAvailable = $80110602;
  COMAdminErrObjectParentMissing = $80110808;
  COMAdminErrObjectDoesNotExist = $80110809;
  COMAdminErrCanNotExportAppProxy = $8011044A;
  COMAdminErrCanNotStartApp = $8011044B;
  COMAdminErrCanNotExportSystemApp = $8011044C;
  COMAdminErrCanNotSubscribeToComponent = $8011044D;
  COMAdminErrAppNotRunning = $8011080A;
  COMAdminErrEventClassCannotBeSubscriber = $8011044E;
  COMAdminErrLibAppProxyIncompatible = $8011044F;
  COMAdminErrBasePartitionOnly = $80110450;
  COMAdminErrDuplicatePartitionName = $80110457;
  COMAdminErrPartitionInUse = $80110459;
  COMAdminErrImportedComponentsNotAllowed = $8011045B;
  COMAdminErrRegdbNotInitialized = $80110472;
  COMAdminErrRegdbNotOpen = $80110473;
  COMAdminErrRegdbSystemErr = $80110474;
  COMAdminErrRegdbAlreadyRunning = $80110475;
  COMAdminErrMigVersionNotSupported = $80110480;
  COMAdminErrMigSchemaNotFound = $80110481;
  COMAdminErrCatBitnessMismatch = $80110482;
  COMAdminErrCatUnacceptableBitness = $80110483;
  COMAdminErrCatWrongAppBitnessBitness = $80110484;
  COMAdminErrCatPauseResumeNotSupported = $80110485;
  COMAdminErrCatServerFault = $80110486;
  COMAdminErrCantRecycleLibraryApps = $8011080F;
  COMAdminErrCantRecycleServiceApps = $80110811;
  COMAdminErrProcessAlreadyRecycled = $80110812;
  COMAdminErrPausedProcessMayNotBeRecycled = $80110813;
  COMAdminErrInvalidPartition = $8011080B;
  COMAdminErrPartitionMsiOnly = $80110819;
  COMAdminErrStartAppDisabled = $80110451;
  COMAdminErrCompMoveSource = $8011081C;
  COMAdminErrCompMoveDest = $8011081D;
  COMAdminErrCompMovePrivate = $8011081E;
  COMAdminErrCannotCopyEventClass = $80110820;

type

// *********************************************************************//
// Forward-Deklaration von in der Typbibliothek definierten Typen                     
// *********************************************************************//
  ICOMAdminCatalog = interface;
  ICOMAdminCatalogDisp = dispinterface;
  ICOMAdminCatalog2 = interface;
  ICOMAdminCatalog2Disp = dispinterface;
  ICatalogObject = interface;
  ICatalogObjectDisp = dispinterface;
  ICatalogCollection = interface;
  ICatalogCollectionDisp = dispinterface;

// *********************************************************************//
// Deklaration von in der Typbibliothek definierten CoClasses
// (HINWEIS: Hier wird jede CoClass ihrem Standard-Interface zugewiesen)              
// *********************************************************************//
  COMAdminCatalog = ICOMAdminCatalog2;
  COMAdminCatalogObject = ICatalogObject;
  COMAdminCatalogCollection = ICatalogCollection;


// *********************************************************************//
// Deklaration von Strukturen, Unions und Aliasen.                          
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}
  PPSafeArray1 = ^PSafeArray; {*}


// *********************************************************************//
// Interface: ICOMAdminCatalog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DD662187-DFC2-11D1-A2CF-00805FC79235}
// *********************************************************************//
  ICOMAdminCatalog = interface(IDispatch)
    ['{DD662187-DFC2-11D1-A2CF-00805FC79235}']
    function GetCollection(const bstrCollName: WideString): IDispatch; safecall;
    function Connect(const bstrCatalogServerName: WideString): IDispatch; safecall;
    function Get_MajorVersion: Integer; safecall;
    function Get_MinorVersion: Integer; safecall;
    function GetCollectionByQuery(const bstrCollName: WideString; var ppsaVarQuery: PSafeArray): IDispatch; safecall;
    procedure ImportComponent(const bstrApplIDOrName: WideString; 
                              const bstrCLSIDOrProgID: WideString); safecall;
    procedure InstallComponent(const bstrApplIDOrName: WideString; const bstrDLL: WideString; 
                               const bstrTLB: WideString; const bstrPSDLL: WideString); safecall;
    procedure ShutdownApplication(const bstrApplIDOrName: WideString); safecall;
    procedure ExportApplication(const bstrApplIDOrName: WideString; 
                                const bstrApplicationFile: WideString; lOptions: Integer); safecall;
    procedure InstallApplication(const bstrApplicationFile: WideString; 
                                 const bstrDestinationDirectory: WideString; lOptions: Integer; 
                                 const bstrUserId: WideString; const bstrPassword: WideString; 
                                 const bstrRSN: WideString); safecall;
    procedure StopRouter; safecall;
    procedure RefreshRouter; safecall;
    procedure StartRouter; safecall;
    procedure Reserved1; safecall;
    procedure Reserved2; safecall;
    procedure InstallMultipleComponents(const bstrApplIDOrName: WideString; 
                                        var ppsaVarFileNames: PSafeArray; 
                                        var ppsaVarCLSIDs: PSafeArray); safecall;
    procedure GetMultipleComponentsInfo(const bstrApplIDOrName: WideString; 
                                        var ppsaVarFileNames: PSafeArray; 
                                        out ppsaVarCLSIDs: PSafeArray; 
                                        out ppsaVarClassNames: PSafeArray; 
                                        out ppsaVarFileFlags: PSafeArray; 
                                        out ppsaVarComponentFlags: PSafeArray); safecall;
    procedure RefreshComponents; safecall;
    procedure BackupREGDB(const bstrBackupFilePath: WideString); safecall;
    procedure RestoreREGDB(const bstrBackupFilePath: WideString); safecall;
    procedure QueryApplicationFile(const bstrApplicationFile: WideString; 
                                   out pbstrApplicationName: WideString; 
                                   out pbstrApplicationDescription: WideString; 
                                   out pbHasUsers: WordBool; out pbIsProxy: WordBool; 
                                   out ppsaVarFileNames: PSafeArray); safecall;
    procedure StartApplication(const bstrApplIDOrName: WideString); safecall;
    function ServiceCheck(lService: Integer): Integer; safecall;
    procedure InstallMultipleEventClasses(const bstrApplIDOrName: WideString; 
                                          var ppsaVarFileNames: PSafeArray; 
                                          var ppsaVarCLSIDs: PSafeArray); safecall;
    procedure InstallEventClass(const bstrApplIDOrName: WideString; const bstrDLL: WideString; 
                                const bstrTLB: WideString; const bstrPSDLL: WideString); safecall;
    procedure GetEventClassesForIID(const bstrIID: WideString; out ppsaVarCLSIDs: PSafeArray; 
                                    out ppsaVarProgIDs: PSafeArray; 
                                    out ppsaVarDescriptions: PSafeArray); safecall;
    property MajorVersion: Integer read Get_MajorVersion;
    property MinorVersion: Integer read Get_MinorVersion;
  end;

// *********************************************************************//
// DispIntf:  ICOMAdminCatalogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {DD662187-DFC2-11D1-A2CF-00805FC79235}
// *********************************************************************//
  ICOMAdminCatalogDisp = dispinterface
    ['{DD662187-DFC2-11D1-A2CF-00805FC79235}']
    function GetCollection(const bstrCollName: WideString): IDispatch; dispid 1;
    function Connect(const bstrCatalogServerName: WideString): IDispatch; dispid 2;
    property MajorVersion: Integer readonly dispid 3;
    property MinorVersion: Integer readonly dispid 4;
    function GetCollectionByQuery(const bstrCollName: WideString; 
                                  var ppsaVarQuery: {NOT_OLEAUTO(PSafeArray)}OleVariant): IDispatch; dispid 5;
    procedure ImportComponent(const bstrApplIDOrName: WideString; 
                              const bstrCLSIDOrProgID: WideString); dispid 6;
    procedure InstallComponent(const bstrApplIDOrName: WideString; const bstrDLL: WideString; 
                               const bstrTLB: WideString; const bstrPSDLL: WideString); dispid 7;
    procedure ShutdownApplication(const bstrApplIDOrName: WideString); dispid 8;
    procedure ExportApplication(const bstrApplIDOrName: WideString; 
                                const bstrApplicationFile: WideString; lOptions: Integer); dispid 9;
    procedure InstallApplication(const bstrApplicationFile: WideString; 
                                 const bstrDestinationDirectory: WideString; lOptions: Integer; 
                                 const bstrUserId: WideString; const bstrPassword: WideString; 
                                 const bstrRSN: WideString); dispid 10;
    procedure StopRouter; dispid 11;
    procedure RefreshRouter; dispid 12;
    procedure StartRouter; dispid 13;
    procedure Reserved1; dispid 14;
    procedure Reserved2; dispid 15;
    procedure InstallMultipleComponents(const bstrApplIDOrName: WideString; 
                                        var ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        var ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 16;
    procedure GetMultipleComponentsInfo(const bstrApplIDOrName: WideString; 
                                        var ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarClassNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarFileFlags: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarComponentFlags: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 17;
    procedure RefreshComponents; dispid 18;
    procedure BackupREGDB(const bstrBackupFilePath: WideString); dispid 19;
    procedure RestoreREGDB(const bstrBackupFilePath: WideString); dispid 20;
    procedure QueryApplicationFile(const bstrApplicationFile: WideString; 
                                   out pbstrApplicationName: WideString; 
                                   out pbstrApplicationDescription: WideString; 
                                   out pbHasUsers: WordBool; out pbIsProxy: WordBool; 
                                   out ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 21;
    procedure StartApplication(const bstrApplIDOrName: WideString); dispid 22;
    function ServiceCheck(lService: Integer): Integer; dispid 23;
    procedure InstallMultipleEventClasses(const bstrApplIDOrName: WideString; 
                                          var ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                          var ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 24;
    procedure InstallEventClass(const bstrApplIDOrName: WideString; const bstrDLL: WideString; 
                                const bstrTLB: WideString; const bstrPSDLL: WideString); dispid 25;
    procedure GetEventClassesForIID(const bstrIID: WideString; 
                                    out ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                    out ppsaVarProgIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                    out ppsaVarDescriptions: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 26;
  end;

// *********************************************************************//
// Interface: ICOMAdminCatalog2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {790C6E0B-9194-4CC9-9426-A48A63185696}
// *********************************************************************//
  ICOMAdminCatalog2 = interface(ICOMAdminCatalog)
    ['{790C6E0B-9194-4CC9-9426-A48A63185696}']
    function GetCollectionByQuery2(const bstrCollectionName: WideString; 
                                   const pVarQueryStrings: OleVariant): IDispatch; safecall;
    function GetApplicationInstanceIDFromProcessID(lProcessID: Integer): WideString; safecall;
    procedure ShutdownApplicationInstances(const pVarApplicationInstanceID: OleVariant); safecall;
    procedure PauseApplicationInstances(const pVarApplicationInstanceID: OleVariant); safecall;
    procedure ResumeApplicationInstances(const pVarApplicationInstanceID: OleVariant); safecall;
    procedure RecycleApplicationInstances(const pVarApplicationInstanceID: OleVariant; 
                                          lReasonCode: Integer); safecall;
    function AreApplicationInstancesPaused(const pVarApplicationInstanceID: OleVariant): WordBool; safecall;
    function DumpApplicationInstance(const bstrApplicationInstanceID: WideString; 
                                     const bstrDirectory: WideString; lMaxImages: Integer): WideString; safecall;
    function Get_IsApplicationInstanceDumpSupported: WordBool; safecall;
    procedure CreateServiceForApplication(const bstrApplicationIDOrName: WideString; 
                                          const bstrServiceName: WideString; 
                                          const bstrStartType: WideString; 
                                          const bstrErrorControl: WideString; 
                                          const bstrDependencies: WideString; 
                                          const bstrRunAs: WideString; 
                                          const bstrPassword: WideString; bDesktopOk: WordBool); safecall;
    procedure DeleteServiceForApplication(const bstrApplicationIDOrName: WideString); safecall;
    function GetPartitionID(const bstrApplicationIDOrName: WideString): WideString; safecall;
    function GetPartitionName(const bstrApplicationIDOrName: WideString): WideString; safecall;
    procedure Set_CurrentPartition(const Param1: WideString); safecall;
    function Get_CurrentPartitionID: WideString; safecall;
    function Get_CurrentPartitionName: WideString; safecall;
    function Get_GlobalPartitionID: WideString; safecall;
    procedure FlushPartitionCache; safecall;
    procedure CopyApplications(const bstrSourcePartitionIDOrName: WideString; 
                               const pVarApplicationID: OleVariant; 
                               const bstrDestinationPartitionIDOrName: WideString); safecall;
    procedure CopyComponents(const bstrSourceApplicationIDOrName: WideString; 
                             const pVarCLSIDOrProgID: OleVariant; 
                             const bstrDestinationApplicationIDOrName: WideString); safecall;
    procedure MoveComponents(const bstrSourceApplicationIDOrName: WideString; 
                             const pVarCLSIDOrProgID: OleVariant; 
                             const bstrDestinationApplicationIDOrName: WideString); safecall;
    procedure AliasComponent(const bstrSrcApplicationIDOrName: WideString; 
                             const bstrCLSIDOrProgID: WideString; 
                             const bstrDestApplicationIDOrName: WideString; 
                             const bstrNewProgId: WideString; const bstrNewClsid: WideString); safecall;
    function IsSafeToDelete(const bstrDllName: WideString): COMAdminInUse; safecall;
    procedure ImportUnconfiguredComponents(const bstrApplicationIDOrName: WideString; 
                                           const pVarCLSIDOrProgID: OleVariant; 
                                           const pVarComponentType: OleVariant); safecall;
    procedure PromoteUnconfiguredComponents(const bstrApplicationIDOrName: WideString; 
                                            const pVarCLSIDOrProgID: OleVariant; 
                                            const pVarComponentType: OleVariant); safecall;
    procedure ImportComponents(const bstrApplicationIDOrName: WideString; 
                               const pVarCLSIDOrProgID: OleVariant; 
                               const pVarComponentType: OleVariant); safecall;
    function Get_Is64BitCatalogServer: WordBool; safecall;
    procedure ExportPartition(const bstrPartitionIDOrName: WideString; 
                              const bstrPartitionFileName: WideString; lOptions: Integer); safecall;
    procedure InstallPartition(const bstrFileName: WideString; const bstrDestDirectory: WideString; 
                               lOptions: Integer; const bstrUserId: WideString; 
                               const bstrPassword: WideString; const bstrRSN: WideString); safecall;
    function QueryApplicationFile2(const bstrApplicationFile: WideString): IDispatch; safecall;
    function GetComponentVersionCount(const bstrCLSIDOrProgID: WideString): Integer; safecall;
    property IsApplicationInstanceDumpSupported: WordBool read Get_IsApplicationInstanceDumpSupported;
    property CurrentPartition: WideString write Set_CurrentPartition;
    property CurrentPartitionID: WideString read Get_CurrentPartitionID;
    property CurrentPartitionName: WideString read Get_CurrentPartitionName;
    property GlobalPartitionID: WideString read Get_GlobalPartitionID;
    property Is64BitCatalogServer: WordBool read Get_Is64BitCatalogServer;
  end;

// *********************************************************************//
// DispIntf:  ICOMAdminCatalog2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {790C6E0B-9194-4CC9-9426-A48A63185696}
// *********************************************************************//
  ICOMAdminCatalog2Disp = dispinterface
    ['{790C6E0B-9194-4CC9-9426-A48A63185696}']
    function GetCollectionByQuery2(const bstrCollectionName: WideString; 
                                   const pVarQueryStrings: OleVariant): IDispatch; dispid 27;
    function GetApplicationInstanceIDFromProcessID(lProcessID: Integer): WideString; dispid 28;
    procedure ShutdownApplicationInstances(const pVarApplicationInstanceID: OleVariant); dispid 29;
    procedure PauseApplicationInstances(const pVarApplicationInstanceID: OleVariant); dispid 30;
    procedure ResumeApplicationInstances(const pVarApplicationInstanceID: OleVariant); dispid 31;
    procedure RecycleApplicationInstances(const pVarApplicationInstanceID: OleVariant; 
                                          lReasonCode: Integer); dispid 32;
    function AreApplicationInstancesPaused(const pVarApplicationInstanceID: OleVariant): WordBool; dispid 33;
    function DumpApplicationInstance(const bstrApplicationInstanceID: WideString; 
                                     const bstrDirectory: WideString; lMaxImages: Integer): WideString; dispid 34;
    property IsApplicationInstanceDumpSupported: WordBool readonly dispid 35;
    procedure CreateServiceForApplication(const bstrApplicationIDOrName: WideString; 
                                          const bstrServiceName: WideString; 
                                          const bstrStartType: WideString; 
                                          const bstrErrorControl: WideString; 
                                          const bstrDependencies: WideString; 
                                          const bstrRunAs: WideString; 
                                          const bstrPassword: WideString; bDesktopOk: WordBool); dispid 36;
    procedure DeleteServiceForApplication(const bstrApplicationIDOrName: WideString); dispid 37;
    function GetPartitionID(const bstrApplicationIDOrName: WideString): WideString; dispid 38;
    function GetPartitionName(const bstrApplicationIDOrName: WideString): WideString; dispid 39;
    property CurrentPartition: WideString writeonly dispid 40;
    property CurrentPartitionID: WideString readonly dispid 41;
    property CurrentPartitionName: WideString readonly dispid 42;
    property GlobalPartitionID: WideString readonly dispid 43;
    procedure FlushPartitionCache; dispid 44;
    procedure CopyApplications(const bstrSourcePartitionIDOrName: WideString; 
                               const pVarApplicationID: OleVariant; 
                               const bstrDestinationPartitionIDOrName: WideString); dispid 45;
    procedure CopyComponents(const bstrSourceApplicationIDOrName: WideString; 
                             const pVarCLSIDOrProgID: OleVariant; 
                             const bstrDestinationApplicationIDOrName: WideString); dispid 46;
    procedure MoveComponents(const bstrSourceApplicationIDOrName: WideString; 
                             const pVarCLSIDOrProgID: OleVariant; 
                             const bstrDestinationApplicationIDOrName: WideString); dispid 47;
    procedure AliasComponent(const bstrSrcApplicationIDOrName: WideString; 
                             const bstrCLSIDOrProgID: WideString; 
                             const bstrDestApplicationIDOrName: WideString; 
                             const bstrNewProgId: WideString; const bstrNewClsid: WideString); dispid 48;
    function IsSafeToDelete(const bstrDllName: WideString): COMAdminInUse; dispid 49;
    procedure ImportUnconfiguredComponents(const bstrApplicationIDOrName: WideString; 
                                           const pVarCLSIDOrProgID: OleVariant; 
                                           const pVarComponentType: OleVariant); dispid 50;
    procedure PromoteUnconfiguredComponents(const bstrApplicationIDOrName: WideString; 
                                            const pVarCLSIDOrProgID: OleVariant; 
                                            const pVarComponentType: OleVariant); dispid 51;
    procedure ImportComponents(const bstrApplicationIDOrName: WideString; 
                               const pVarCLSIDOrProgID: OleVariant; 
                               const pVarComponentType: OleVariant); dispid 52;
    property Is64BitCatalogServer: WordBool readonly dispid 53;
    procedure ExportPartition(const bstrPartitionIDOrName: WideString; 
                              const bstrPartitionFileName: WideString; lOptions: Integer); dispid 54;
    procedure InstallPartition(const bstrFileName: WideString; const bstrDestDirectory: WideString; 
                               lOptions: Integer; const bstrUserId: WideString; 
                               const bstrPassword: WideString; const bstrRSN: WideString); dispid 55;
    function QueryApplicationFile2(const bstrApplicationFile: WideString): IDispatch; dispid 56;
    function GetComponentVersionCount(const bstrCLSIDOrProgID: WideString): Integer; dispid 57;
    function GetCollection(const bstrCollName: WideString): IDispatch; dispid 1;
    function Connect(const bstrCatalogServerName: WideString): IDispatch; dispid 2;
    property MajorVersion: Integer readonly dispid 3;
    property MinorVersion: Integer readonly dispid 4;
    function GetCollectionByQuery(const bstrCollName: WideString; 
                                  var ppsaVarQuery: {NOT_OLEAUTO(PSafeArray)}OleVariant): IDispatch; dispid 5;
    procedure ImportComponent(const bstrApplIDOrName: WideString; 
                              const bstrCLSIDOrProgID: WideString); dispid 6;
    procedure InstallComponent(const bstrApplIDOrName: WideString; const bstrDLL: WideString; 
                               const bstrTLB: WideString; const bstrPSDLL: WideString); dispid 7;
    procedure ShutdownApplication(const bstrApplIDOrName: WideString); dispid 8;
    procedure ExportApplication(const bstrApplIDOrName: WideString; 
                                const bstrApplicationFile: WideString; lOptions: Integer); dispid 9;
    procedure InstallApplication(const bstrApplicationFile: WideString; 
                                 const bstrDestinationDirectory: WideString; lOptions: Integer; 
                                 const bstrUserId: WideString; const bstrPassword: WideString; 
                                 const bstrRSN: WideString); dispid 10;
    procedure StopRouter; dispid 11;
    procedure RefreshRouter; dispid 12;
    procedure StartRouter; dispid 13;
    procedure Reserved1; dispid 14;
    procedure Reserved2; dispid 15;
    procedure InstallMultipleComponents(const bstrApplIDOrName: WideString; 
                                        var ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        var ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 16;
    procedure GetMultipleComponentsInfo(const bstrApplIDOrName: WideString; 
                                        var ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarClassNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarFileFlags: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                        out ppsaVarComponentFlags: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 17;
    procedure RefreshComponents; dispid 18;
    procedure BackupREGDB(const bstrBackupFilePath: WideString); dispid 19;
    procedure RestoreREGDB(const bstrBackupFilePath: WideString); dispid 20;
    procedure QueryApplicationFile(const bstrApplicationFile: WideString; 
                                   out pbstrApplicationName: WideString; 
                                   out pbstrApplicationDescription: WideString; 
                                   out pbHasUsers: WordBool; out pbIsProxy: WordBool; 
                                   out ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 21;
    procedure StartApplication(const bstrApplIDOrName: WideString); dispid 22;
    function ServiceCheck(lService: Integer): Integer; dispid 23;
    procedure InstallMultipleEventClasses(const bstrApplIDOrName: WideString; 
                                          var ppsaVarFileNames: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                          var ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 24;
    procedure InstallEventClass(const bstrApplIDOrName: WideString; const bstrDLL: WideString; 
                                const bstrTLB: WideString; const bstrPSDLL: WideString); dispid 25;
    procedure GetEventClassesForIID(const bstrIID: WideString; 
                                    out ppsaVarCLSIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                    out ppsaVarProgIDs: {NOT_OLEAUTO(PSafeArray)}OleVariant; 
                                    out ppsaVarDescriptions: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 26;
  end;

// *********************************************************************//
// Interface: ICatalogObject
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6EB22871-8A19-11D0-81B6-00A0C9231C29}
// *********************************************************************//
  ICatalogObject = interface(IDispatch)
    ['{6EB22871-8A19-11D0-81B6-00A0C9231C29}']
    function Get_Value(const bstrPropName: WideString): OleVariant; safecall;
    procedure Set_Value(const bstrPropName: WideString; pvarRetVal: OleVariant); safecall;
    function Get_Key: OleVariant; safecall;
    function Get_Name: OleVariant; safecall;
    function IsPropertyReadOnly(const bstrPropName: WideString): WordBool; safecall;
    function Get_Valid: WordBool; safecall;
    function IsPropertyWriteOnly(const bstrPropName: WideString): WordBool; safecall;
    property Value[const bstrPropName: WideString]: OleVariant read Get_Value write Set_Value;
    property Key: OleVariant read Get_Key;
    property Name: OleVariant read Get_Name;
    property Valid: WordBool read Get_Valid;
  end;

// *********************************************************************//
// DispIntf:  ICatalogObjectDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6EB22871-8A19-11D0-81B6-00A0C9231C29}
// *********************************************************************//
  ICatalogObjectDisp = dispinterface
    ['{6EB22871-8A19-11D0-81B6-00A0C9231C29}']
    property Value[const bstrPropName: WideString]: OleVariant dispid 1;
    property Key: OleVariant readonly dispid 2;
    property Name: OleVariant readonly dispid 3;
    function IsPropertyReadOnly(const bstrPropName: WideString): WordBool; dispid 4;
    property Valid: WordBool readonly dispid 5;
    function IsPropertyWriteOnly(const bstrPropName: WideString): WordBool; dispid 6;
  end;

// *********************************************************************//
// Interface: ICatalogCollection
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6EB22872-8A19-11D0-81B6-00A0C9231C29}
// *********************************************************************//
  ICatalogCollection = interface(IDispatch)
    ['{6EB22872-8A19-11D0-81B6-00A0C9231C29}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Item(lIndex: Integer): IDispatch; safecall;
    function Get_Count: Integer; safecall;
    procedure Remove(lIndex: Integer); safecall;
    function Add: IDispatch; safecall;
    procedure Populate; safecall;
    function SaveChanges: Integer; safecall;
    function GetCollection(const bstrCollName: WideString; varObjectKey: OleVariant): IDispatch; safecall;
    function Get_Name: OleVariant; safecall;
    function Get_AddEnabled: WordBool; safecall;
    function Get_RemoveEnabled: WordBool; safecall;
    function GetUtilInterface: IDispatch; safecall;
    function Get_DataStoreMajorVersion: Integer; safecall;
    function Get_DataStoreMinorVersion: Integer; safecall;
    procedure PopulateByKey(psaKeys: PSafeArray); safecall;
    procedure PopulateByQuery(const bstrQueryString: WideString; lQueryType: Integer); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Item[lIndex: Integer]: IDispatch read Get_Item;
    property Count: Integer read Get_Count;
    property Name: OleVariant read Get_Name;
    property AddEnabled: WordBool read Get_AddEnabled;
    property RemoveEnabled: WordBool read Get_RemoveEnabled;
    property DataStoreMajorVersion: Integer read Get_DataStoreMajorVersion;
    property DataStoreMinorVersion: Integer read Get_DataStoreMinorVersion;
  end;

// *********************************************************************//
// DispIntf:  ICatalogCollectionDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6EB22872-8A19-11D0-81B6-00A0C9231C29}
// *********************************************************************//
  ICatalogCollectionDisp = dispinterface
    ['{6EB22872-8A19-11D0-81B6-00A0C9231C29}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Item[lIndex: Integer]: IDispatch readonly dispid 1;
    property Count: Integer readonly dispid 1610743810;
    procedure Remove(lIndex: Integer); dispid 1610743811;
    function Add: IDispatch; dispid 1610743812;
    procedure Populate; dispid 2;
    function SaveChanges: Integer; dispid 3;
    function GetCollection(const bstrCollName: WideString; varObjectKey: OleVariant): IDispatch; dispid 4;
    property Name: OleVariant readonly dispid 6;
    property AddEnabled: WordBool readonly dispid 7;
    property RemoveEnabled: WordBool readonly dispid 8;
    function GetUtilInterface: IDispatch; dispid 9;
    property DataStoreMajorVersion: Integer readonly dispid 10;
    property DataStoreMinorVersion: Integer readonly dispid 11;
    procedure PopulateByKey(psaKeys: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 12;
    procedure PopulateByQuery(const bstrQueryString: WideString; lQueryType: Integer); dispid 13;
  end;

// *********************************************************************//
// Die Klasse CoCOMAdminCatalog stellt die Methoden Create und CreateRemote zur
// Verfügung, um Instanzen des Standard-Interface ICOMAdminCatalog2, dargestellt
// von CoClass COMAdminCatalog, zu erzeugen. Diese Funktionen können
// von einem Client verwendet werden, der die CoClasses automatisieren
// will, die von dieser Typbibliothek dargestellt werden.                                           
// *********************************************************************//
  CoCOMAdminCatalog = class
    class function Create: ICOMAdminCatalog2;
    class function CreateRemote(const MachineName: string): ICOMAdminCatalog2;
  end;

// *********************************************************************//
// Die Klasse CoCOMAdminCatalogObject stellt die Methoden Create und CreateRemote zur
// Verfügung, um Instanzen des Standard-Interface ICatalogObject, dargestellt
// von CoClass COMAdminCatalogObject, zu erzeugen. Diese Funktionen können
// von einem Client verwendet werden, der die CoClasses automatisieren
// will, die von dieser Typbibliothek dargestellt werden.                                           
// *********************************************************************//
  CoCOMAdminCatalogObject = class
    class function Create: ICatalogObject;
    class function CreateRemote(const MachineName: string): ICatalogObject;
  end;

// *********************************************************************//
// Die Klasse CoCOMAdminCatalogCollection stellt die Methoden Create und CreateRemote zur
// Verfügung, um Instanzen des Standard-Interface ICatalogCollection, dargestellt
// von CoClass COMAdminCatalogCollection, zu erzeugen. Diese Funktionen können
// von einem Client verwendet werden, der die CoClasses automatisieren
// will, die von dieser Typbibliothek dargestellt werden.                                           
// *********************************************************************//
  CoCOMAdminCatalogCollection = class
    class function Create: ICatalogCollection;
    class function CreateRemote(const MachineName: string): ICatalogCollection;
  end;

implementation

uses System.Win.ComObj;

class function CoCOMAdminCatalog.Create: ICOMAdminCatalog2;
begin
  Result := CreateComObject(CLASS_COMAdminCatalog) as ICOMAdminCatalog2;
end;

class function CoCOMAdminCatalog.CreateRemote(const MachineName: string): ICOMAdminCatalog2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_COMAdminCatalog) as ICOMAdminCatalog2;
end;

class function CoCOMAdminCatalogObject.Create: ICatalogObject;
begin
  Result := CreateComObject(CLASS_COMAdminCatalogObject) as ICatalogObject;
end;

class function CoCOMAdminCatalogObject.CreateRemote(const MachineName: string): ICatalogObject;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_COMAdminCatalogObject) as ICatalogObject;
end;

class function CoCOMAdminCatalogCollection.Create: ICatalogCollection;
begin
  Result := CreateComObject(CLASS_COMAdminCatalogCollection) as ICatalogCollection;
end;

class function CoCOMAdminCatalogCollection.CreateRemote(const MachineName: string): ICatalogCollection;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_COMAdminCatalogCollection) as ICatalogCollection;
end;

end.
