# MDTA Cup

Tracks Minnesota Debate Tournament Association Cup points for Lincoln-Douglas, Classic, Public Forum, Policy, and Congressional Debate. Results come from the Tabroom results API. A public standings page needs no login. A few named users import tournaments, confirm student names, and publish.

## What you create in Azure

Create these in one resource group. This repository does not create them.

- An App Service plan and a .NET 8 Web App.
- An empty Azure SQL database. The app creates the tables on startup.
- Web App settings:
  - `ConnectionStrings__Cup` — the SQL connection string
  - `Admin__Username` and `Admin__Password` — the first account, created only when the user table is empty

Allow Azure services through the SQL firewall. Later users are added from the Users page after login.

## Run the site

From this folder, with the connection string set in `MDTAPoints.Api/appsettings.Development.json` or the environment:

```bash
dotnet run --project MDTAPoints.Api
```

Open the standings at the site root. Log in to import a Tabroom tournament id, confirm any close names, and publish. Publish stays blocked while a name is still a possible match.

## Score one tournament from the console

```bash
dotnet run --project MDTAPoints -- score 41328
```

That prints Cup points for the varsity and open debate events at that tournament. The season, name links, and award standings are kept by the web app.

## Publish to the Web App

Publish `MDTAPoints.Api` from Visual Studio, or with `dotnet publish` and a zip deploy, onto the Web App you already created. Publishing does not create the App Service or the database.
