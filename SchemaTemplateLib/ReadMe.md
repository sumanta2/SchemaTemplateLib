# Creating and Using Local NuGet Packages

A complete guide to creating a NuGet package from your local .NET library and using it in your ASP.NET Web API project.

## Creating a NuGet Package

### 1. Prepare Your Class Library Project

First, ensure your library project (.csproj) has the necessary metadata. Open your `.csproj` file and add/update these properties:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <PackageId>MyCompany.MyLibrary</PackageId>
    <Version>1.0.0</Version>
    <Authors>Your Name</Authors>
    <Company>Your Company</Company>
    <Description>Description of your library</Description>
    <PackageLicenseExpression>MIT</PackageLicenseExpression>
  </PropertyGroup>
</Project>
```

### 2. Build the NuGet Package

Navigate to your library project folder in terminal/command prompt and run:

```bash
dotnet pack -c Release
```

This creates a `.nupkg` file in `bin/Release/` directory.

### 3. Set Up Local NuGet Source

Create a folder to store your local packages, for example: `C:\LocalNuGetPackages`

Copy your `.nupkg` file to this folder, then add it as a NuGet source:

Add it via Visual Studio:

- Tools → Options → NuGet Package Manager → Package Sources
- Click + → Add your folder path

## Using in ASP.NET Web API Project

### 1. Install the Package

In your Web API project directory, run:

```bash
dotnet add package MyCompany.MyLibrary --version 1.0.0
```

Or use Package Manager Console in Visual Studio:

```powershell
Install-Package MyCompany.MyLibrary -Version 1.0.0
```

### 2. Use the Library

Now you can use your library in your Web API:

## Updating the Package

When you make changes to your library:

1. Update the `<Version>` in your `.csproj` (e.g., to 1.0.1)
2. Run `dotnet pack -c Release` again
3. Copy the new `.nupkg` to your local folder
4. Update the package in your Web API: `dotnet add package MyCompany.MyLibrary --version 1.0.1`
