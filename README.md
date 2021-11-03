# Restxcel
## _Simple restful API to create excel files_

Simple ASP.NET Web API using EPPlusFree NuGet package to expose a restful interface for creating excel documents.

## Dependencies:
- .NET 5.0
- EPPlusFree (GNU Lesser General Public License)
- Newtonsoft.Json (MIT License)
- Swashbuckle.AspNetCore (MIT License)

## How to use

### Use hosted version on Azure
https://nvm-restxcel.azurewebsites.net/swagger

### Use your own installation
Download a release from the [release](https://github.com/nvm-uli/restxcel/releases) page.
  or
Download the source code and run with (.NET 5.0 SDK required)
`dotnet run`

Open swagger endpoint on `[yourscheme]://[yourendpoint]:[yourport]/swagger`

#### Configure template directory
The template directory for permanent templates can be configured in appsettings.json file.
```
  "Restxcel": {
    "TemplatesDirectory": "C:\\mypath\\templates"
  }
  or
  "Restxcel": {
    "TemplatesDirectory": "/home/my/path/templates"
  }
```

## Features

- Use xltx Excel file templates as a base or start with an empty file
- Add and modify worksheets, data and format to the file
- Download the final excel file
- Runs completely in memory
