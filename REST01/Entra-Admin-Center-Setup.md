# Entra Admin Center Setup for REST01 Fibonacci API

This guide explains exactly what to change in Microsoft Entra Admin Center so the new REST API in REST01 is protected with the same SSO token model as `/randommobile`.

## Current design (what was implemented)

- The new REST API validates the same token audience values already used by `src/middle-tier/ssoauth-helper.ts`.
- The add-in calls the REST API with the same header pattern:
  - `Authorization: Bearer <middletierToken>`
- Because of that, in most cases **no Entra app registration changes are required**.

## When NO Entra changes are required

No change is needed if all of these are true:

1. You keep using this existing app registration:
   - App (client) ID: `3f2ae881-21e0-44c7-b806-6b09d7a6e16d`
2. The token obtained by Office SSO still has audience equal to one of:
   - `3f2ae881-21e0-44c7-b806-6b09d7a6e16d`
   - `api://3f2ae881-21e0-44c7-b806-6b09d7a6e16d`
   - `api://localhost:3000/3f2ae881-21e0-44c7-b806-6b09d7a6e16d`
   - `api://app-InfoGuard-backend-dev01.azurewebsites.net/3f2ae881-21e0-44c7-b806-6b09d7a6e16d`
3. `access_as_user` scope remains enabled under Expose an API.
4. Pre-authorized Office app entry remains present for app ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` with the same delegated permission ID.

## Verification steps in Entra Admin Center

1. Open https://entra.microsoft.com.
2. Go to **Identity > Applications > App registrations**.
3. Open app registration **InfoGuard** (Application ID `3f2ae881-21e0-44c7-b806-6b09d7a6e16d`).
4. Open **Expose an API**.
5. Confirm Application ID URI contains:
   - `api://localhost:3000/3f2ae881-21e0-44c7-b806-6b09d7a6e16d`
6. Confirm scope `access_as_user` exists and is enabled.
7. Confirm pre-authorized client includes:
   - Client ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`
   - Delegated permission is `access_as_user`
8. Open **Authentication**.
9. Confirm redirect URIs still include your add-in pages:
   - `https://localhost:3000/commands.html`
   - `https://localhost:3000/taskpane.html`
   - `https://localhost:3000/`
   - SPA: `https://localhost:3000/fallbackauthdialog.html`

## If you decide to use a NEW dedicated API App Registration (optional)

Use this only if you want the Fibonacci API to have its own app identity.

1. Create a new app registration (single or multi-tenant per your policy).
2. In **Expose an API**:
   - Set Application ID URI, for example: `api://localhost:7079/<new-api-client-id>`
   - Add scope `access_as_user`
3. In the add-in app registration (InfoGuard), add API permission to the new API scope.
4. Grant admin consent for the tenant.
5. Add your Office client app to pre-authorized applications for the new API, if needed.
6. Update token acquisition resource/scope in add-in flow so the SSO token audience matches the new API.
7. In REST01 API config, replace `Entra:AllowedAudiences` with the new API audience values.

## Local run notes

1. Start REST API:
   - `cd REST01/InfoGuard.RestApi`
   - `dotnet run`
2. By default, development HTTPS URL is `https://localhost:7079`.
3. The taskpane Fibonacci client currently calls `https://localhost:7079`.
4. Ensure ASP.NET development certificate is trusted on your machine.
