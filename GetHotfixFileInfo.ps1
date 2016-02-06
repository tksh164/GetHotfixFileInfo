[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][ValidateScript({ Test-Path -LiteralPath $_ })]
    [string] $Path
)

##
## Invoke Commandline.
##
function Invoke-Commandline
{
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $CommandFileName,

        [Parameter(Mandatory = $true, Position = 1)][AllowEmptyString()]
        [string] $CommandArguments
    )

    $process = New-Object -TypeName 'System.Diagnostics.Process'

    try
    {
        # Set process information.
        $process.StartInfo.FileName = $CommandFileName
        $process.StartInfo.Arguments = $CommandArguments
        $process.StartInfo.CreateNoWindow = $true
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true

        # Register events for read stdout/stderr.
        $stringBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        $eventHandler = {
            if ($EventArgs.Data -ne $null)
            {
                [void] $Event.MessageData.AppendLine($EventArgs.Data)
            }
        }
        $outputEventJob = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action $eventHandler -MessageData $stringBuilder
        $errorEventJob = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action $eventHandler -MessageData $stringBuilder

        # Start process.
        [void] $process.Start()
        Write-Verbose ('[PID:{0}] "{1}" {2}' -f $process.Id, $CommandFileName, $CommandArguments)

        # Wait for process exit.
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        $process.WaitForExit()
        Write-Verbose ('[PID:{0}] ExitCode:{1}' -f $process.Id, $process.ExitCode)

        # Unregister events.
        Unregister-Event -SourceIdentifier $outputEventJob.Name
        Unregister-Event -SourceIdentifier $errorEventJob.Name

        return [pscustomobject] @{
            CommandFileName  = $CommandFileName
            CommandArguments = $CommandArguments
            Pid              = $process.Id
            ExitCode         = $process.ExitCode
            StartTime        = $process.StartTime
            ExitTime         = $process.ExitTime
            Output           = $stringBuilder.ToString().Trim()
        }
    }
    catch
    {
        throw $_
    }
}

##
## Get hotfix package information.
##
function Get-HotfixPackageInfo
{
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo] $HotfixFileInfo,

        [Parameter(Mandatory = $false)]
        [string] $WorkDirectory = (Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath (New-Guid).Guid)
    )

    # Create working directory.
    [void](New-Item -Path $WorkDirectory -ItemType 'Directory')

    # Extract package properties file from msu package.
    $params = @{
        CommandFileName  = 'C:\Windows\System32\expand.exe'
        CommandArguments = ('"{0}"' -f $HotfixFileInfo.FullName),'-F:*-pkgProperties.txt',("{0}" -f $WorkDirectory) -join ' '
    }
    $invokeResult = Invoke-Commandline @params

    # Retrieve FileInfo of extracted package properties file.
    $pkgPropFileInfo = (Get-ChildItem -LiteralPath $WorkDirectory -Filter '*-pkgProperties.txt')[0]
    
    # Get properties from file.
    $pkgProps = Get-Content -LiteralPath $pkgPropFileInfo.FullName | ConvertFrom-StringData

    # Remove work directory
    Remove-Item -LiteralPath $WorkDirectory -Recurse -Force

    # Create hashtable for return value.
    $ht = @{}
    foreach ($pkgProp in $pkgProps)
    {
        foreach ($key in $pkgProp.Keys)
        {
            switch ($key)
            {
                'ApplicabilityInfo' {
                    $ht.ApplicabilityInfo = $pkgProp[$key].Trim('"') -split ';'
                }
                'Applies to' {
                    $ht.AppliesTo = $pkgProp[$key].Trim('"')
                }
                'Build Date' {
                    $ht.BuildDate = [System.DateTime]::Parse($pkgProp[$key].Trim('"'))
                }
                'Company' {
                    $ht.Company = $pkgProp[$key].Trim('"')
                }
                'File Version' {
                    $ht.FileVersion = $pkgProp[$key].Trim('"')
                }
                'Installation Type' {
                    $ht.InstallationType = $pkgProp[$key].Trim('"')
                }
                'Installer Engine' {
                    $ht.InstallerEngine = $pkgProp[$key].Trim('"')
                }
                'Installer Version' {
                    $ht.InstallerVersion = $pkgProp[$key].Trim('"')
                }
                'KB Article Number' {
                    $ht.KBArticleNumber = $pkgProp[$key].Trim('"')
                }
                'Language' {
                    $ht.Language = $pkgProp[$key].Trim('"')
                }
                'Package Type' {
                    $ht.PackageType = $pkgProp[$key].Trim('"')
                }
                'Processor Architecture' {
                    $ht.ProcessorArchitecture = $pkgProp[$key].Trim('"')
                }
                'Product Name' {
                    $ht.ProductName = $pkgProp[$key].Trim('"')
                }
                'Support Link' {
                    $ht.SupportLink = $pkgProp[$key].Trim('"')
                }
            }
        }
    }
        
    return [pscustomobject] $ht
}


Get-ChildItem -LiteralPath $Path -Filter '*.msu' | ForEach-Object -Process {
    Get-HotfixPackageInfo -HotfixFileInfo $_
}
