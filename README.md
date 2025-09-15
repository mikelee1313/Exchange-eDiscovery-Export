# Exchange-eDiscovery-Export

**Exchange-eDiscovery-Export** is a PowerShell tool for exporting calendar items, contacts, and tasks from Exchange Online mailboxes using the Microsoft Graph API. It leverages Microsoft Purview eDiscovery APIs to perform exports without requiring legacy Exchange Online PowerShell modules. This script supports both certificate and client secret-based authentication, robust logging, and direct REST API calls for advanced eDiscovery workflows.

---

## Features

- Export **calendar**, **contacts**, and **tasks** from Exchange Online mailboxes.
- Uses **Microsoft Graph API** and Purview **eDiscovery** endpoints.
- Supports **App-Only** authentication (certificate or client secret).
- Automates eDiscovery case and search creation, export initiation, and download handling.
- Cleans up and refreshes authentication tokens for secure operations.
- Verbose debug logging and error diagnostics.
- Modular functions for folder ID retrieval, Graph API invocation, export workflows, and more.

---

## Prerequisites

- **Microsoft Graph PowerShell module**  
  Install via:  
  ```powershell
  Install-Module Microsoft.Graph
  ```
- **MSAL.PS module** (for official download method)  
  Install via:  
  ```powershell
  Install-Module MSAL.PS
  ```

- **Exchange Online PowerShell module**  
  Install via:  
  ```powershell
  Install-Module -Name ExchangeOnlineManagement
  ```
  
- **App Registration in Azure AD** with the following permissions:
  - `eDiscovery.ReadWrite.All` (Application permission, admin consent granted)
  - `MicrosoftPurviewEDiscovery` API permissions for downloads
- **Certificate or Client Secret** configured for App-Only authentication

For setup guides, see:  
- [Getting Started with eDiscovery APIs](https://techcommunity.microsoft.com/blog/microsoft-security-blog/getting-started-with-the-ediscovery-apis/4407597)  
- [Microsoft Graph eDiscovery App Auth Setup](https://learn.microsoft.com/en-us/graph/security-ediscovery-appauthsetup)

---

## Parameters

| Parameter              | Type     | Description                                                                                     | Default Value                                               |
|------------------------|----------|-------------------------------------------------------------------------------------------------|-------------------------------------------------------------|
| `ClearTokenCache`      | Switch   | Clears all Microsoft Graph authentication tokens and cached states.                             | None                                                        |
| `VerboseTokenClearing` | Switch   | Shows detailed info when clearing authentication tokens.                                        | None                                                        |
| `EmailAddress`         | String   | Email address of the mailbox to export.                                                         | `NestorW@M365x61250205.OnMicrosoft.com`                     |
| `AppID`                | String   | Application (client) ID for Azure AD app registration.                                          | `5baa1427-1e90-4501-831d-a8e67465f0d9`                      |
| `TenantId`             | String   | Azure AD tenant ID.                                                                             | `85612ccb-4c28-4a34-88df-a538cc139a51`                      |
| `AuthMode`             | String   | Authentication mode: `Cert` or `Secret`.                                                        | `Cert`                                                      |
| `certlocation`         | String   | Certificate location for authentication.                                                        | `Cert:\LocalMachine\My\`                                    |
| `CertificateThumbprint`| String   | Thumbprint of the certificate for authentication.                                               | `B696FDCFE1453F3FBC6031F54DE988DA0ED905A9`                  |
| `ClientSecret`         | String   | Client secret (if AuthMode is `Secret`).                                                        | None                                                        |
| `Version`              | String   | Microsoft Graph API version: `prod` (v1.0) or `beta`.                                          | `beta`                                                      |
| `ContentType`          | String   | What to export: `all`, `calendar`, `contacts`, or `tasks`.                                     | `all`                                                       |
| `Operation`            | String   | Operation: `menu`, `create`, `export`, or `download`.                                          | `menu`                                                      |
| `CaseId`               | String   | ID of an existing eDiscovery case (if reusing).                                                | None                                                        |
| `SearchId`             | String   | ID of an existing eDiscovery search (if reusing).                                              | None                                                        |
| `ExportId`             | String   | ID of an existing export to download.                                                          | None                                                        |
| `DownloadPath`         | String   | Path to save exported files.                                                                   | `C:\temp`                                                   |
| `DebugOutput`          | Switch   | Enables verbose debug output.                                                                  | None                                                        |

---

## Usage

### Export All Content Types (Calendar, Contacts, Tasks)
```powershell
.\exchange-eDiscovery-export.ps1 -EmailAddress user@domain.com
```

### Export Only Calendar Items
```powershell
.\exchange-eDiscovery-export.ps1 -EmailAddress user@domain.com -ContentType calendar
```

### Export Only Tasks
```powershell
.\exchange-eDiscovery-export.ps1 -EmailAddress user@domain.com -ContentType tasks
```

### Export Only Contacts
```powershell
.\exchange-eDiscovery-export.ps1 -EmailAddress user@domain.com -ContentType contacts
```

### Use v1.0 (prod) Graph API Endpoints
```powershell
.\exchange-eDiscovery-export.ps1 -Version prod
```

### Clear Authentication Tokens
```powershell
.\exchange-eDiscovery-export.ps1 -ClearTokenCache
```
Add `-VerboseTokenClearing` for more details:
```powershell
.\exchange-eDiscovery-export.ps1 -ClearTokenCache -VerboseTokenClearing
```

### Specify Certificate or Secret Authentication
```powershell
.\exchange-eDiscovery-export.ps1 -AuthMode Cert -CertificateThumbprint ABC123...
.\exchange-eDiscovery-export.ps1 -AuthMode Secret -ClientSecret 'your-client-secret'
```

### Specify Export Download Path
```powershell
.\exchange-eDiscovery-export.ps1 -DownloadPath 'D:\Exports'
```

---

## Troubleshooting

If you encounter issues running the **Exchange-eDiscovery-Export** script, use the following guidance based on common problems and error handling built into the script:

### 1. Authentication Failures

- **Error:** _Authentication failed with Microsoft Graph API_
  - **Solution:**  
    - Ensure the correct `AppID`, `TenantId`, and authentication mode (`Cert` or `Secret`) are specified.
    - If using certificate authentication, verify the certificate thumbprint and location.
    - If using client secret authentication, ensure the client secret is valid and not expired.
    - Make sure your app registration has **admin consent** for all required permissions.

- **Error:** _Received 401 Unauthorized_
  - **Solution:**  
    - Confirm that your Azure AD app registration has the required permissions (`eDiscovery.ReadWrite.All`, `MicrosoftPurviewEDiscovery`, etc.).
    - Double-check that admin consent has been granted.
    - Use `-ClearTokenCache` to refresh authentication tokens.

### 2. Module Import Errors

- **Error:** _Error importing Microsoft Graph modules_
  - **Solution:**  
    - Install the required modules:
      ```powershell
      Install-Module Microsoft.Graph -Scope CurrentUser
      Install-Module MSAL.PS -Scope CurrentUser
      ```
    - Run PowerShell as administrator if you encounter permission issues.

### 3. Permissions and Consent Issues

- **Error:** _Permission issue detected! The app may be missing required permissions or admin consent._
  - **Solution:**  
    - Visit [Microsoft Graph eDiscovery App Auth Setup](https://learn.microsoft.com/en-us/graph/security-ediscovery-appauthsetup) and ensure all permissions are assigned with admin consent.
    - Ensure your app registration is not missing any required API permissions.

### 4. eDiscovery API Limitations

- **Error:** _Microsoft Graph eDiscovery export APIs are returning errors. This is common with beta endpoints that aren't fully implemented yet._
  - **Solution:**  
    - Try switching the `-Version` parameter between `beta` and `prod` to see if one works better.
    - Some endpoints may only be available in `beta`. The script will attempt fallback methods automatically where possible.

### 5. Folder ID Retrieval Issues

- **Error:** _Unable to retrieve folder IDs from Exchange. Please check connection and permissions._
  - **Solution:**  
    - Ensure your account has permission to access mailbox statistics via Exchange Online.
    - Verify network connectivity and Exchange Online service availability.

### 6. Export Fails or Data Missing

- **Error:** _Export operation did not complete in time_
  - **Solution:**  
    - Increase the timeout by modifying script parameters if necessary.
    - Check Microsoft 365 service health for issues affecting eDiscovery or Graph endpoints.

- **Error:** _No valid eDiscovery case ID available. Exiting function._
  - **Solution:**  
    - Ensure you have permissions to create eDiscovery cases.
    - Try using a different mailbox or verify the mailbox exists.

### 7. General Debugging Advice

- Run the script with `-DebugOutput` for verbose logs.
- Check the log file created in your `%TEMP%` directory (e.g., `Exchange_Calendar_Export_<timestamp>.log`).
- Use `-VerboseTokenClearing` with `-ClearTokenCache` for more detailed authentication cleanup output.
- If you see unexpected errors, review the script output and log file for details and error codes.

If your issue persists, check the official [Microsoft Graph eDiscovery documentation](https://learn.microsoft.com/en-us/graph/security-ediscovery-appauthsetup), or open an issue in this repository with relevant error messages and logs.

---

## Credits

**Authors:**  
Mike Lee | Dempsey Dunkin | Ranjit Sharma  
**Date:** September 15, 2025
