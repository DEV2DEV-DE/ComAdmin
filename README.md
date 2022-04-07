# ComAdmin
Delphi wrapper around the COM+ administration for Windows

Links to MS documentation:

https://docs.microsoft.com/en-us/windows/win32/cossdk/com--administration-collections#collection-hierarchy

Class-Hierarchy:

Catalog
 - Applications(*)
   - Application
     - Components(*)
       - Component (wip)
     - Instances(*)
       - Instance
     - Roles(*)
       - Role
         - Users(*) 
           - User
 - Computer
 - Partitions(*)
   - Partition
