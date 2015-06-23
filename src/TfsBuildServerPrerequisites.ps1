#REQUIRES -Version 2.0

#------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#
# Licensed under the Microsoft Limited Public License (the "License");
# you may not use this file except in compliance with the License.
# A full copy of the license is provided in the root folder of the
# project directory.
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# Description:
#   This script collects and installs the files required to build SharePoint
#   projects on a TFS build server. It supports TFS 2012.
#
# Step 1. Run the script on a developer machine where VS 2012, SP 2013 and
#         Office developer tools from WebPI are installed.
#         Use the "-Collect" parameter to collect files required for
#         installation.
#
# Step 2. Copy the script and folder of collected files to a TFS build server.
#         Use the "-Install" parameter to install the files.
#------------------------------------------------------------------------------

param (
  [switch]$Collect,
  [switch]$Install,
  [switch]$Quiet
)

#==============================================================================
# Script variables
#==============================================================================

# Name of the script file
$MyScriptFileName = Split-Path $MyInvocation.MyCommand.Path -Leaf

# Directory from which script is run
$MyScriptDirectory = Split-Path $MyInvocation.MyCommand.Path

function InitializeScriptVariables() {
  # Program Files path is different on 32- and 64-bit machines.
  $Script:ProgramFilesPath = GetProgramFilesPath

  # SharePoint assemblies that can be referenced from SharePoint projects.
  # If you need more assemblies, add them here.
  $Script:SharePoint15ReferenceAssemblies = @(
    "Microsoft.SharePoint.dll",
    "Microsoft.SharePoint.Security.dll",
    "Microsoft.SharePoint.WorkflowActions.dll",
    "Microsoft.Office.Server.dll",
    "Microsoft.Office.Server.UserProfiles.dll",
    "Microsoft.SharePoint.Client.dll",
    "Microsoft.SharePoint.Client.Runtime.dll",
    "Microsoft.SharePoint.Client.ServerRuntime.dll",
    "Microsoft.SharePoint.Linq.dll",
    "Microsoft.SharePoint.Portal.dll",
    "Microsoft.SharePoint.Publishing.dll",
    "Microsoft.SharePoint.Taxonomy.dll",
    "Microsoft.SharePoint.WorkflowActions.dll",
    "Microsoft.Web.CommandUI.dll"
  )

  # Source path where the SharePoint15 Reference Assemblies are located (this is a better source than relying on Gac40 as multiple versions can exist in the GAC)
  $Script:SharePoint15ReferenceAssembliesPath = Join-Path $env:ProgramFiles "Common Files\microsoft shared\Web Server Extensions\15\ISAPI"

  # Destination path where SharePoint15 assemblies will be copied
  $Script:SharePoint15ReferenceAssemblyPath = Join-Path $ProgramFilesPath "Reference Assemblies\Microsoft\SharePoint15"

  # Folder where we collect assemblies and use as a source for installation
  $Script:FilesFolder = "Files"
  $Script:FilesPath = Join-Path $MyScriptDirectory $FilesFolder

  # Path to .NET Framework 4.0 GAC
  $Script:Gac40Path = Join-Path $Env:SystemRoot "Microsoft.NET\Assembly"

  # MSBuild extensions folders needed for SharePoint project.
  $Script:MSBuildSharePointDependencies = @(
    "SharePointTools",
    "WebApplications",
    "Web"
  )

  # MSBuild extensions path of SharePoint project dependencies.
  $Script:MSBuildSharePointDependenciesPath = Join-Path $ProgramFilesPath "MSBuild\Microsoft\VisualStudio\v12.0"

  # Workflow MSBuild extensions.
  $Script:MSBuildWorkflowManager = @(
    "Microsoft.WorkflowServiceBuildExtensions.targets",
    "Microsoft.Workflow.Service.Build.dll"
  )

  # Workflow Manager MSBuild extension path.
  $Script:MSBuildWorkflowManagerPath = Join-Path $ProgramFilesPath "MSBuild\Microsoft\Workflow Manager\1.0"

  # SharePoint project assemblies required to create WSP files.
  $Script:SharePointProjectAssemblies = @(
    "Microsoft.VisualStudio.SharePoint.dll",
    "Microsoft.VisualStudio.SharePoint.Designers.Models.dll",
    "Microsoft.VisualStudio.SharePoint.Designers.Models.Features.dll",
    "Microsoft.VisualStudio.SharePoint.Designers.Models.Packages.dll"
  )

  # Workflow tools for visual studio redist assemblies.
  $Script:WorkflowAssemblies = @(
    "Microsoft.Activities.Design.dll"
  )

  # Files needed to install assemblies to the GAC on TFS build machine.
  $Script:GacUtilFiles = @(
    "gacutil.exe",
    "gacutil.exe.config"
    "1033\gacutlrc.dll"
  )

  # Path where the GacUtils.exe can be found on the development machine.
  $Script:GacUtilFilePath = Join-Path $ProgramFilesPath "Microsoft SDKs\Windows\v8.0A\Bin\NETFX 4.0 Tools"

  # Registry path which indicates if SharePiont is installed or not
  $Script:SharePoint15InstallationRegistryKey = "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\15.0"
}

#==============================================================================
# Top-level script routines
#==============================================================================

# Install SharePoint project files on TFS build server
function InstallSharePointProjectFiles() {
  WriteText "Installing SharePoint project files for TFS Build."
  WriteText ""

  WriteText "Checking that the .Net Framework 4.0 is installed:"
  WriteText "Found .Net Framework version $(GetDotNet40Version)"
  WriteText ""

  if (TestSharePoint15) {
    WriteText "SharePoint 15 is installed on this machine."
    WriteText "Skipping installation of SharePoint 15 reference assemblies."
    WriteText ""
  } else {
    WriteText "Copying the SharePoint reference assemblies:"
    $SharePoint15ReferenceAssemblies | CopyFileToFolder $SharePoint15ReferenceAssemblyPath
    WriteText ""

    WriteText "Setting the SharePoint reference assemblies path in the registry:"
    AddSharePoint15ReferencePathToRegistry
    WriteText ""
  }

  WriteText "Copying the MSBuild extensions folders of SharePoint project dependencies:"
  $MSBuildSharePointDependencies | CopyFileToFolder $MSBuildSharePointDependenciesPath
  WriteText ""

  WriteText "Copying the Workflow Manager MSBuild extensions:"
  $MSBuildWorkflowManager | CopyFileToFolder $MSBuildWorkflowManagerPath
  WriteText ""

  WriteText "Installing the SharePoint project assemblies in the GAC:"
  $SharePointProjectAssemblies | InstallAssemblyToGac40
  WriteText ""

  WriteText "Installing workflow assemblies in the GAC:"
  $WorkflowAssemblies | InstallAssemblyToGac40
  WriteText ""

  WriteText "Installation is complete."
  WriteText ""
}

# Collect SharePoint project files on developement machine for TFS build server.
function CollectSharePointProjectFiles() {
  WriteText "Collecting SharePoint project files for TFS Build."
  WriteText ""

  WriteText "Checking that the .Net Framework 4.0 is installed:"
  WriteText "Found .Net Framework version: $(GetDotNet40Version)"
  WriteText ""

  WriteText "Creating the $FilesFolder directory:"
  $FilesPathInfo = [System.IO.Directory]::CreateDirectory($FilesPath)
  WriteText $FilesPathInfo.FullName
  WriteText ""

  WriteText "Collecting the SharePoint reference assemblies:"
  $SharePoint15ReferenceAssemblies | CopyFileFromFolder $SharePoint15ReferenceAssembliesPath
  WriteText ""

  WriteText "Collecting the MSBuild extensions dependencies of SharePoint project:"
  $MSBuildSharePointDependencies | CopyFileFromFolder $MSBuildSharePointDependenciesPath
  WriteText ""

  WriteText "Collecting the Workflow Manager MSBuild extensions:"
  $MSBuildWorkflowManager | CopyFileFromFolder $MSBuildWorkflowManagerPath
  WriteText ""

  WriteText "Collecting the SharePoint project assemblies:"
  $SharePointProjectAssemblies | CopyFileFromFolder $Gac40Path -Recurse -ProductVersion "12.0"
  WriteText ""

  WriteText "Collecting workflow assemblies:"
  $WorkflowAssemblies | CopyFileFromFolder $Gac40Path -Recurse
  WriteText ""

  WriteText "Collecting the gacutil.exe files:"
  $GacUtilFiles | CopyFileFromFolder $GacUtilFilePath
  WriteText ""

  WriteText "Colection is complete."
  WriteText ""
}

# Writes instructions how to use the script.
# If -Quiet flag is provided it throws an exception.
function WriteHelp() {
  if ($Quiet) {
    throw "Specify either the -Install or -Collect parameter."
  } else {
    Write-Host ""
    Write-Host "This script collects and installs the files to build SharePoint projects on a TFS build server 2012."
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host "  -Collect   - Collect files into the $FilesFolder folder on a development machine."
    Write-Host "  -Install   - Install files from the $FilesFolder folder on a TFS build server."
    Write-Host "  -Quiet     - Suppress messages. Errors will still be shown."
    Write-Host ""
    Write-Host "Usage:"
    Write-Host ""
    Write-Host "  .\$MyScriptFileName -Collect"
    Write-Host "     - Collects files into the $FilesFolder folder on a development machine."
    Write-Host ""
    Write-Host "  .\$MyScriptFileName -Install"
    Write-Host "     - Installs files from the $FilesFolder folder on a TFS build server."
    Write-Host ""
  }
}

#==============================================================================
# Utility functions
#==============================================================================

# Writes to Host if -Quiet flag is not provided
function WriteText($Value) {
  if (-Not $Quiet) {
    Write-Host $Value
  }
}

# Writes warning to Host if -Quiet flag is not provided
function WriteWarning($Value) {
  if (-Not $Quiet) {
    Write-Host $Value -ForegroundColor Yellow
  }
}

# Gets the Program Files (x86) path
function GetProgramFilesPath() {
  if (TestWindows64) {
    return ${Env:ProgramFiles(x86)}
  } else {
    return $Env:ProgramFiles
  }
}

# Installs assembly to the GAC
function InstallAssemblyToGac40([switch]$SkipIfExists) {
  begin {
    $GacUtilCommand = Join-Path $FilesPath "gacutil.exe"
  }
  process {
    $FileName = $_
    $FileFullPath = Join-Path $FilesPath $FileName

    if (-Not (Test-Path $FileFullPath)) {
      WriteWarning """$FileName"" was not found in $FilesPath"
      return
    }

    if ($SkipIfExists -and (TestAssemblyInGac40 $FileName)) {
      WriteText "Assembly $FileName is already in the GAC ($Gac40Path)."
      WriteText "Skipping installation in the GAC."
      return
    }

    & "$GacUtilCommand" /if "$FileFullPath" /nologo

    # Verify that assembly was installed to the GAC
    if (-not (TestAssemblyInGac40 $FileName)) {
      throw "Assembly $FileName was not installed in the GAC ($Gac40Path)"
    }

    WriteText "$FileName [is installed in ==>] GAC 4.0"
  }
}

# Test if the assembly file name is already in the GAC
function TestAssemblyInGac40($AssemblyFileName) {
  $FileInfo = Get-ChildItem $Gac40Path -Recurse | Where { $_.Name -eq $AssemblyFileName }
  return ($FileInfo -ne $null)
}

# Copies file from the specified path or its sub-directory to the $FilesPath
<#
#>
function CopyFileFromFolder([string]$Path, [string]$ProductVersion = '', [switch]$Recurse) {
  process {
    $FileName = $_
    if (!$Recurse) {
      $FileInfo = [System.IO.FileInfo]$(Join-Path $Path $FileName)
    } else {
      $FileInfo = Get-ChildItem $Path -Recurse | Where { $_.Name -eq $FileName }
    }

    if (!$FileInfo) {
      WriteWarning """$FileName"" was not found in $Path"
      return
    }

    if ($ProductVersion -ne '') {
      $FileInfo = $FileInfo | ?{ $_.VersionInfo.ProductVersion -match "^$ProductVersion" }
    }

    if ($FileInfo -is [array]) {
      throw "Multiple instances of $FileName were found in $Path"
    }

    Copy-Item $FileInfo.FullName $FilesPath -Recurse -Force
    WriteText """$($FileInfo.FullName)"" [was copied to ==>] ""$FilesPath"""
  }
}

# Copies file to the specified path from the $FilesPath
function CopyFileToFolder([string]$Path) {
  begin {
    # Ensure that the folder exists
    $null = [System.IO.Directory]::CreateDirectory($Path)
  }
  process {
    $FileName = $_
    $FileInfo = Join-Path $FilesPath $FileName

    if (-Not (Test-Path $FileInfo)) {
      WriteWarning """$FileName"" was not found in $FilesPath"
      return
    }

    Copy-Item $FileInfo $Path -Recurse -Force
    WriteText """$FileInfo"" [was copied to ==>] ""$Path"""
  }
}

# Adds folder with SharePoint reference libraries to Registry to be found by MSBuild
function AddSharePoint15ReferencePathToRegistry() {
  $RegistryValue = $SharePoint15ReferenceAssemblyPath + "\"
  if (TestWindows64) {
    $RegistryPath = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319\AssemblyFoldersEx\SharePoint15"
  } else {
    $RegistryPath = "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319\AssemblyFoldersEx\SharePoint15"
  }
  # Do not set value if it is already exists and has the same value
  if (-Not ((Test-Path $RegistryPath) -And ((Get-ItemProperty $RegistryPath)."(default)" -Eq $RegistryValue))) {
    $null = New-Item -ItemType String $RegistryPath -Value $RegistryValue -Force
  }
  WriteText "Registry key $RegistryPath [was set to ==>] $RegistryValue"
}

# Returns full path for the .Net 4.0 framework installation
function GetDotNet40Directory() {
  # Find a folder in Microsoft.NET that starts with 'v4.0'
  $DirectoryInfo = Get-ChildItem (Join-Path $Env:SystemRoot "Microsoft.NET\Framework") `
    | Where { $_ -is [System.IO.DirectoryInfo] } `
    | Where { $_.Name -like 'v4.0*' }

  if ($DirectoryInfo -eq $null) {
    throw ".Net 4.0 is not found on this machine."
  }

  return $DirectoryInfo
}

# Returns version of installed .Net 4.0 framework
function GetDotNet40Version() {
  return (GetDotNet40Directory).Name
}

# Returns $true if this is 64-bit version of Windows
function TestWindows64() {
  return (Get-WmiObject Win32_OperatingSystem | Select OSArchitecture).OSArchitecture.StartsWith("64")
}

# Returns $true if machine has SharePoint 15 installed
function TestSharePoint15() {
  if (Test-Path "$SharePoint15InstallationRegistryKey") {
    $node = Get-Item -LiteralPath $SharePoint15InstallationRegistryKey
    return $($node.GetValue("SharePoint") -eq "installed");
  }
  return $false
}

#==============================================================================
# Main script code
#==============================================================================

InitializeScriptVariables

if ($Install) {
  InstallSharePointProjectFiles
} elseif ($Collect) {
  CollectSharePointProjectFiles
 } else {
  WriteHelp
}
