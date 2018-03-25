# This module contains functions to read the appxmanifest from appx and appxbundle files.

# Add the Open Package Convention Types
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName WindowsBase
    
function Get-AppxManifestFromPackage
{
    <#
    .SYNOPSIS
    This function returns the xml contents of the AppxManifest of an appx package.
    .DESCRIPTION
    THis function can extract the contents of the AppxManifest of the package
    without extracting the appx package first. This is especially useful when packages are large.
    .EXAMPLE
    Get-AppxManifestFromPackage .\MyPackage.appx
    .PARAMETER AppxPackagePath
    The path to the appx file.
    #>
    # Use CmdletBinding, see https://blogs.technet.microsoft.com/heyscriptingguy/2012/07/07/weekend-scripter-cmdletbinding-attribute-simplifies-powershell-functions/
    [CmdletBinding()]
    param(

        # Path to the package
        [Parameter(Mandatory = $true)]
        [string]$AppxPackagePath
    )    
    
    try
    {        
        if (!(Test-Path -PathType Leaf -Path $AppxPackagePath))
        {
            throw "$AppxPackagePath does not exist."
        }
        
        $appxManifest = PSUsing([System.IO.Packaging.Package] $mainPackage = [System.IO.Packaging.Package]::Open(
                $AppxPackagePath,    
                [System.IO.FileMode]::Open, 
                [System.IO.FileAccess]::Read)) {
                
            # return $appxManifestXml
            $packageManifest = Get-ManifestFromPackage -Package $mainPackage
            
            Write-Output $packageManifest
        }
        
        Write-Output $appxManifest
    }                
    catch
    {
        $PSCmdlet.ThrowTerminatingError($PSItem)
    }    
}

function Get-AppxManifestFromBundle
{
    <#
    .SYNOPSIS
    This function returns the xml contents of the AppxManifest of the main package in the bundle.
    .DESCRIPTION
    THis function can extract the contents of the AppxManifest of the main package of the bundle
    without extracting the bundle first. This is especially useful when bundles are large.
    .EXAMPLE
    Get-AppxManifestFromBundle .\MyBundle.appxbundle
    .PARAMETER AppxBundlePath
    The path to the appxbundle file.
    #>
    # Use CmdletBinding, see https://blogs.technet.microsoft.com/heyscriptingguy/2012/07/07/weekend-scripter-cmdletbinding-attribute-simplifies-powershell-functions/
    [CmdletBinding()]
    param(

        # Path to the bundle
        [Parameter(Mandatory = $true)]
        [string]$AppxBundlePath
    )
    
    try
    {         
        if (!(Test-Path -PathType Leaf -Path $AppxBundlePath))
        {
            throw "$AppxBundlePath does not exist."
        }
        
        $appxManifest = PSUsing([System.IO.Packaging.Package] $package = [System.IO.Packaging.Package]::Open(
            $AppxBundlePath,    
            [System.IO.FileMode]::Open, 
            [System.IO.FileAccess]::Read)) {
            
            # Get all the parts in the package
            $packageParts = $package.GetParts()
            
            # First, find the main package part of the bundle.
            $bundleManifestPart = $packageParts | ? { $_.ContentType -eq "application/vnd.ms-appx.bundlemanifest+xml" } | Select-Object -first 1            
            $mainPackageInfoFileName = PSUsing([System.IO.StreamReader]$streamReader = New-Object System.IO.StreamReader($bundleManifestPart.GetStream())) {
            
                # The main package is the first element with type "Application" in the bundle manifest xml
                [xml]$bundleManifestXml = $streamReader.ReadToEnd()                
                $bundleManifestXmlMainPackage = $bundleManifestXml.Bundle.Packages.Package | ? { $_.Type -eq "Application" } | Select-Object -first 1
                
                Write-Output $bundleManifestXmlMainPackage.FileName
            }

            Write-Verbose "Main package: $mainPackageInfoFileName"
            
            # Once we have the main package, let's open it.
            $mainPackagePart = $packageParts | ? { $_.Uri -eq "/$mainPackageInfoFileName" } | Select-Object -first 1
            $mainPackageManifest = PSUsing([System.IO.Packaging.Package] $mainPackage = [System.IO.Packaging.Package]::Open(
                $mainPackagePart.GetStream(),    
                [System.IO.FileMode]::Open, 
                [System.IO.FileAccess]::Read)) {
                
                # return $appxManifestXml
                $mainPackageManifest = Get-ManifestFromPackage -Package $mainPackage
                
                Write-Output $mainPackageManifest
            }
            
            Write-Output $mainPackageManifest
        }
        
        Write-Output $appxManifest
    }    
    catch
    {
        $PSCmdlet.ThrowTerminatingError($PSItem)
    }
}

function Get-ManifestFromPackage
{
    <#
    .SYNOPSIS
    This function returns the xml contents of the AppxManifest from an open package.
    .DESCRIPTION
    THis function can extract the contents of the AppxManifest of the package by looking for
    a part with content type "application/vnd.ms-appx.manifest+xml" and returns the contents of that part.
    .PARAMETER Package
    The opened instance of System.IO.Packaging.Package.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Packaging.Package]$Package
    )
    
    $packageParts = $package.GetParts()            
                
    # Get the appxmanifest.
    $appxManifestPart = $packageParts | ? { $_.ContentType -eq "application/vnd.ms-appx.manifest+xml" } | Select-Object -first 1            
    $appxManifestXml = PSUsing([System.IO.StreamReader]$streamReader = New-Object System.IO.StreamReader($appxManifestPart.GetStream())) {
    
        # Cast contents to xml
        [xml]$appxManifestXml = $streamReader.ReadToEnd()        
        Write-Output $appxManifestXml
    }
    
    Write-Output $appxManifestXml
}

function PSUsing
{
    <#
    .SYNOPSIS
    Heklper function to ensure object that implement System.IDisposable get properly disposed.
    .DESCRIPTION
    This function is a powershell implementation of using(...) { } in C#.
    .PARAMETER InputObject
    The IDisposable instance that should be cleaned up after executing $ScriptBlock.
    .PARAMETER ScriptBlock
    The script block to execute while $InputObject is active.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [AllowNull()]
        [Object]
        $InputObject,
 
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock
    )
 
    try
    {
        . $ScriptBlock
    }
    finally
    {
        if ($null -ne $InputObject -and $InputObject -is [System.IDisposable])
        {
            $InputObject.Dispose()
        }
    }
}

Export-ModuleMember -Function @("Get-AppxManifestFromPackage", "Get-AppxManifestFromBundle")