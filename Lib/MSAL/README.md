# MSAL Library

This folder is for the Microsoft.Identity.Client DLL files if you don't have MSAL.PS installed.

## Option 1: Install MSAL.PS (Recommended)

```powershell
Install-Module MSAL.PS -Scope CurrentUser
```

## Option 2: Manual DLL Download

Download Microsoft.Identity.Client from NuGet:
https://www.nuget.org/packages/Microsoft.Identity.Client/

Place the following files here:
- Microsoft.Identity.Client.dll

The application will automatically detect and load them.
