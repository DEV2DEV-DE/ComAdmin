# ComAdmin
![](https://tokei.rs/b1/github/DEV2DEV-DE/ComAdmin?category=code)
![](https://tokei.rs/b1/github/DEV2DEV-DE/ComAdmin?category=files)

Delphi wrapper for the COM+ administration for Windows

Links to MS documentation:

https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections#collection-hierarchy

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
  ∞ Protocols (+/-)
    ● Protocol
  ∞ EventClasses (+/-)
    ● EventClass
```
\+ This collection supports an Add method

\- This collection supports a Remove method

The Components collection of an Application does not support an Add method. To install or import components into an application, use methods on the Catalog object.
