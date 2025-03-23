# 365-Tenant-Assessment

Performs a full assessment of a Microsoft 365 tenant and exports the data to an Excel sheet.

This script is based on Sean McAvinue's original [script](https://practical365.com/microsoft-365-tenant-to-tenant-migration-assessment-version-2/)

The original script works well overall, but I encountered issues when using Entra modules. It requires the `MSGraph` module, which conflicts with the authentication mechanisms when Entra modules are in use.  
This is a known issue that Microsoft has not yet resolved.

## What I've Changed

- Heavily refactored the codebase by condensing logic into functions, resulting in a cleaner script and making it more extensible and maintainable.
- Created dedicated functions for data collection using `Invoke-MgGraphRequest` to ensure compatibility, as I'm personally using the Entra module.

# ToDo
The script that Sean wrote was extensive and impressive so it took a while to go through it all.
I may have some functions I'm not using here while I was conceptualizing how to change it.
