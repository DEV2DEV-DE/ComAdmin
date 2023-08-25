# ComAdmin
[![GitHub license](https://img.shields.io/github/license/DEV2DEV-DE/ComAdmin)](https://github.com/DEV2DEV-DE/ComAdmin/blob/main/LICENSE)
![GitHub repo size](https://img.shields.io/github/repo-size/DEV2DEV-DE/ComAdmin)
![GitHub commit activity](https://img.shields.io/github/commit-activity/m/DEV2DEV-DE/ComAdmin)
![GitHub last commit](https://img.shields.io/github/last-commit/DEV2DEV-DE/ComAdmin)
![GitHub issues](https://img.shields.io/github/issues/DEV2DEV-DE/ComAdmin)

Delphi wrapper for the COM+ administration for Windows

Links to MS documentation:

https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections#collection-hierarchy

## Usage
```pascal
var
  ComCatalog: TComAdminCatalog;
  i: Integer;
begin
  ComCatalog := TComAdminCatalog.Create('MyServerName', '*', nil, nil);
  try
    for i := 0 to ComCatalog.Applications.Count - 1 do
      Memo1.Lines.Add(ComCatalog.Applications[i].Name);
  finally
    ComCatalog.Free;
  end;
end;
```
Just provide an empty string as the first parameter of the constructor if you want to access the catalog of the local computer.

## Class-Hierarchy:
```
● Catalog
  ∞ Applications (+/-)
    ● Application
      ∞ Instances (+/-)
        ● Instance
      ∞ Roles (+/-)
        ● Role
          ∞ Users (+/-) 
            ● User
      ∞ Components (-)
        ● Component
          ∞ Roles (+/-)
            ● Role
          ∞ Interfaces
            ● Interface
              ∞ Roles (+/-)
                ● Role
              ∞ Methods
                ● Method
                  ∞ Roles (+/-)
                    ● Role
  ● Computer
  ∞ Partitions (+/-)
    ● Partition
      ∞ Roles
        ● Role
          ∞ Users (+/-) 
            ● User
  ∞ ApplicationCluster (+/-)
    ● Server
  ∞ Protocols (+/-)
    ● Protocol
  ∞ EventClasses (+/-)
    ● EventClass
  ∞ InprocServers (-)
    ● InprocServer
  ∞ TransientSubscriptions (+/-)
    ● TransientSubscription
      ∞ Publishers (+/-)
        ● Values
      ∞ Subscribers (+/-)
        ● Values
```
\+ This collection supports an Add method

\- This collection supports a Remove method

The Components collection of an Application does not support an Add method. To install or import components into an application, use methods on the Catalog object.
