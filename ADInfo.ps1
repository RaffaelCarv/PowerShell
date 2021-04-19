<#
.NOTES
	Name: ADInfo.ps1
	Original Author: Shane Moore
    Author: Shane Moore
    Contributor: Rafael Carvalho
	Requires: PowerShell and administrator rights on the target ADs
	server as well as the local machine.
   
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Check the integrity of Active Directory in the internal domain.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER HtmlReportFile
    Name of the HTML output file from the BuildHtmlServersReport. Default is ADllServersReport.html

.DESCRIPTION
	This script checks the integrity of Active Directory for various configuration recommendations.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "HealthChecker")]
param(
    [Parameter(Mandatory = $false, ParameterSetName = "HealthChecker")]
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport")]
    [string]$Server = ($env:COMPUTERNAME),
    [Parameter(Mandatory = $false)]
    [ValidateScript( { -not $_.ToString().EndsWith('\') })][string]$OutputFilePath = ".",
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport")]
    [switch]$MailboxReport,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [switch]$LoadBalancingReport,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [array]$CasServerList = $null,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [string]$SiteName = ([string]::Empty),
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly")]
    [ValidateScript( { -not $_.ToString().EndsWith('\') })][string]$XMLDirectoryPath = ".",
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [switch]$BuildHtmlServersReport,
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [string]$HtmlReportFile = "DCsServersReport.html",
    [Parameter(Mandatory = $false, ParameterSetName = "DCCoreReport")]
    [switch]$DCCoreRatio,
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly")]
    [switch]$AnalyzeDataOnly,
    [Parameter(Mandatory = $false)][switch]$SaveDebugLog
)
Function Set-ScriptLogFileLocation {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $false)][bool]$IncludeServerName = $false
    )
    $endName = "-{0}.txt" -f $dateTimeStringFormat

    if ($IncludeServerName) {
        $endName = "-{0}{1}" -f $Script:Server, $endName
    }

    $Script:OutputFullPath = "{0}\{1}{2}" -f $OutputFilePath, $FileName, $endName
    $Script:OutXmlFullPath = $Script:OutputFullPath.Replace(".txt", ".xml")

    if ($AnalyzeDataOnly -or
        $BuildHtmlServersReport) {
        return
    }

    Write-VerboseWriter("Script Executing on Server $env:COMPUTERNAME")
}
$Computers = (Get-ADComputer -Filter *).count
$Workstations = (Get-ADComputer -LDAPFilter "(&(objectClass=Computer)(!operatingSystem=*server*))" -Searchbase (Get-ADDomain).distinguishedName).count
$Servers = (Get-ADComputer -LDAPFilter "(&(objectClass=Computer)(operatingSystem=*server*))" -Searchbase (Get-ADDomain).distinguishedName).count
$Users = (get-aduser -filter *).count 
$domain = Get-ADDomain |FT Forest
$FSMO = netdom query FSMO
$ADForest = (Get-ADForest).ForestMode
$ADDomain = (Get-ADDomain).DomainMode
$DC = (Get-ADComputer -LDAPFilter "(&(objectClass=Computer)(primaryGroupID=516)(operatingSystem=*server*))" -Searchbase (Get-ADDomain).distinguishedName).count
$ADVer = Get-ADObject (Get-ADRootDSE).schemaNamingContext -property objectVersion | Select objectVersion
$ADNUM = $ADVer -replace "@{objectVersion=","" -replace "}",""

If ($ADNum -eq '88') {$srv = 'Windows Server 2019'}
ElseIf ($ADNum -eq '87') {$srv = 'Windows Server 2016'}
ElseIf ($ADNum -eq '69') {$srv = 'Windows Server 2012 R2'}
ElseIf ($ADNum -eq '56') {$srv = 'Windows Server 2012'}
ElseIf ($ADNum -eq '47') {$srv = 'Windows Server 2008 R2'}
ElseIf ($ADNum -eq '44') {$srv = 'Windows Server 2008'}
ElseIf ($ADNum -eq '31') {$srv = 'Windows Server 2003 R2'}
ElseIf ($ADNum -eq '30') {$srv = 'Windows Server 2003'}



Write-Host "For this Domain there are;" -ForegroundColor Yellow
Write-Host "Domain Controller = "$DC -ForegroundColor Cyan
Write-Host "Computers         = "$Computers -ForegroundColor Cyan
Write-Host "Workstions        = "$Workstations -ForegroundColor Cyan
Write-Host "Servers           = "$Servers -ForegroundColor Cyan
Write-Host "Users             = "$Users -ForegroundColor Cyan
Write-host ""
Write-host "Active Directory Info" -ForegroundColor Yellow
Write-Host "Active Directory Forest Mode =  "$ADForest -ForegroundColor Cyan
Write-Host "Active Directory Domain Mode =  "$ADDomain -ForegroundColor Cyan
Write-Host "Active Directory Schema Version is $ADNum which corresponds to $Srv" -ForegroundColor Cyan
Write-Host ""
Write-Host "FSMO Role Owners" -ForegroundColor Yellow
$FSMO

#####################################Get ALL DC Servers#################################
$getForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()

$DCServers = $getForest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 

$timeout = "60"

Write-Host "Domain Controllers List" -ForegroundColor Yellow
Write-Host "Domain Controllers = "$DCServers -ForegroundColor Cyan
Write-Host ""
Write-Host "Active Directory Health Check" -ForegroundColor Yellow
Write-Host ""

foreach ($DC in $DCServers){
$Identity = $DC
################Ping Test######
if ( Test-Connection -ComputerName $DC -Count 1 -ErrorAction SilentlyContinue ) {
Write-Host $DC `t $DC `t Ping Success -ForegroundColor Green
 ##############Netlogon Service Status################
		$serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name "Netlogon" -ErrorAction SilentlyContinue} -ArgumentList $DC
                wait-job $serviceStatus -timeout $timeout
                if($serviceStatus.state -like "Running")
                {
                 Write-Host $DC `t Netlogon Service TimeOut -ForegroundColor Yellow
                 stop-job $serviceStatus
                }
                else
                {
                $serviceStatus1 = Receive-job $serviceStatus
                 if ($serviceStatus1.status -eq "Running") {
 		   Write-Host $DC `t $serviceStatus1.name `t $serviceStatus1.status -ForegroundColor Green 
         	   $svcName = $serviceStatus1.name 
         	   $svcState = $serviceStatus1.status          
                  }
                 else 
                  { 
       		  Write-Host $DC `t $serviceStatus1.name `t $serviceStatus1.status -ForegroundColor Red 
         	  $svcName = $serviceStatus1.name 
         	  $svcState = $serviceStatus1.status          
                  } 
                }
               }
##############NTDS Service Status################
		$serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name "NTDS" -ErrorAction SilentlyContinue} -ArgumentList $DC
                wait-job $serviceStatus -timeout $timeout
                if($serviceStatus.state -like "Running")
                {
                 Write-Host $DC `t NTDS Service TimeOut -ForegroundColor Yellow
                 stop-job $serviceStatus
                }
                else
                {
                $serviceStatus1 = Receive-job $serviceStatus
                 if ($serviceStatus1.status -eq "Running") {
 		   Write-Host $DC `t $serviceStatus1.name `t $serviceStatus1.status -ForegroundColor Green 
         	   $svcName = $serviceStatus1.name 
         	   $svcState = $serviceStatus1.status          
                  }
                 else 
                  { 
       		  Write-Host $DC `t $serviceStatus1.name `t $serviceStatus1.status -ForegroundColor Red 
         	  $svcName = $serviceStatus1.name 
         	  $svcState = $serviceStatus1.status          
                  } 
                }
##############DNS Service Status################
		$serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name "DNS" -ErrorAction SilentlyContinue} -ArgumentList $DC
                wait-job $serviceStatus -timeout $timeout
                if($serviceStatus.state -like "Running")
                {
                 Write-Host $DC `t DNS Server Service TimeOut -ForegroundColor Yellow
                 stop-job $serviceStatus
                }
                else
                {
                $serviceStatus1 = Receive-job $serviceStatus
                 if ($serviceStatus1.status -eq "Running") {
 		   Write-Host $DC `t $serviceStatus1.name `t $serviceStatus1.status -ForegroundColor Green 
         	   $svcName = $serviceStatus1.name 
         	   $svcState = $serviceStatus1.status          
                  }
                 else 
                  { 
       		  Write-Host $DC `t $serviceStatus1.name `t $serviceStatus1.status -ForegroundColor Red 
         	  $svcName = $serviceStatus1.name 
         	  $svcState = $serviceStatus1.status          
                  } 
                }
####################Netlogons status##################
               add-type -AssemblyName microsoft.visualbasic 
               $cmp = "microsoft.visualbasic.strings" -as [type]
               $sysvol = start-job -scriptblock {dcdiag /test:netlogons /s:$($args[0])} -ArgumentList $DC
               wait-job $sysvol -timeout $timeout
               if($sysvol.state -like "Running")
               {
               Write-Host $DC `t Netlogons Test TimeOut -ForegroundColor Yellow
               stop-job $sysvol
               }
               else
               {
               $sysvol1 = Receive-job $sysvol
               if($cmp::instr($sysvol1, "passed test NetLogons"))
                  {
                  Write-Host $DC `t Netlogons Test passed -ForegroundColor Green
                  }
               else
                  {
                  Write-Host $DC `t Netlogons Test Failed -ForegroundColor Red
                  }
                }

####################Replications status#################
               add-type -AssemblyName microsoft.visualbasic 
               $cmp = "microsoft.visualbasic.strings" -as [type]
               $sysvol = start-job -scriptblock {dcdiag /test:Replications /s:$($args[0])} -ArgumentList $DC
               wait-job $sysvol -timeout $timeout
               if($sysvol.state -like "Running")
               {
               Write-Host $DC `t Replications Test TimeOut -ForegroundColor Yellow
               stop-job $sysvol
               }
               else
               {
               $sysvol1 = Receive-job $sysvol
               if($cmp::instr($sysvol1, "passed test Replications"))
                  {
                  Write-Host $DC `t Replications Test passed -ForegroundColor Green
                  }
               else
                  {
                  Write-Host $DC `t Replications Test Failed -ForegroundColor Red
                  }
                }
####################Services status#####################
               add-type -AssemblyName microsoft.visualbasic 
               $cmp = "microsoft.visualbasic.strings" -as [type]
               $sysvol = start-job -scriptblock {dcdiag /test:Services /s:$($args[0])} -ArgumentList $DC
               wait-job $sysvol -timeout $timeout
               if($sysvol.state -like "Running")
               {
               Write-Host $DC `t Services Test TimeOut -ForegroundColor Yellow
               stop-job $sysvol
               }
               else
               {
               $sysvol1 = Receive-job $sysvol
               if($cmp::instr($sysvol1, "passed test Services"))
                  {
                  Write-Host $DC `t Services Test passed -ForegroundColor Green
                  }
               else
                  {
                  Write-Host $DC `t Services Test Failed -ForegroundColor Red
                  }
                }
 ####################Advertising status##################
               add-type -AssemblyName microsoft.visualbasic 
               $cmp = "microsoft.visualbasic.strings" -as [type]
               $sysvol = start-job -scriptblock {dcdiag /test:Advertising /s:$($args[0])} -ArgumentList $DC
               wait-job $sysvol -timeout $timeout
               if($sysvol.state -like "Running")
               {
               Write-Host $DC `t Advertising Test TimeOut -ForegroundColor Yellow
               stop-job $sysvol
               }
               else
               {
               $sysvol1 = Receive-job $sysvol
               if($cmp::instr($sysvol1, "passed test Advertising"))
                  {
                  Write-Host $DC `t Advertising Test passed -ForegroundColor Green
                  }
               else
                  {
                  Write-Host $DC `t Advertising Test Failed -ForegroundColor Red
                  }
                }
####################FSMOCheck status##################
               add-type -AssemblyName microsoft.visualbasic 
               $cmp = "microsoft.visualbasic.strings" -as [type]
               $sysvol = start-job -scriptblock {dcdiag /test:FSMOCheck /s:$($args[0])} -ArgumentList $DC
               wait-job $sysvol -timeout $timeout
               if($sysvol.state -like "Running")
               {
               Write-Host $DC `t FSMOCheck Test TimeOut -ForegroundColor Yellow
               stop-job $sysvol
               }
               else
               {
               $sysvol1 = Receive-job $sysvol
               if($cmp::instr($sysvol1, "passed test FsmoCheck"))
                  {
                  Write-Host $DC `t FSMOCheck Test passed -ForegroundColor Green
                  }
               else
                  {
                  Write-Host $DC `t FSMOCheck Test Failed -ForegroundColor Red
                  }
                }
                Write-Host ""
             }
Function Get-ErrorsThatOccurred {

    if ($Error.Count -gt 0) {
        Write-Grey(" "); Write-Grey(" ")
        Function Write-Errors {
            Write-VerboseOutput("`r`n`r`nErrors that occurred that wasn't handled")

            $index = 0
            $Error |
                ForEach-Object {
                    $index++
                    $currentError = $_
                    $handledError = $Script:ErrorsExcluded |
                        Where-Object { $_.Equals($currentError) }

                        if ($null -eq $handledError) {
                            Write-VerboseOutput("Error Index: $index")
                            Write-VerboseOutput($currentError)

                            if ($null -ne $currentError.ScriptStackTrace) {
                                Write-VerboseOutput($currentError.ScriptStackTrace)
                            }
                            Write-VerboseOutput("-----------------------------------`r`n`r`n")
                        }
                    }

            Write-VerboseOutput("`r`n`r`nErrors that were handled")
            $index = 0
            $Error |
                ForEach-Object {
                    $index++
                    $currentError = $_
                    $handledError = $Script:ErrorsExcluded |
                        Where-Object { $_.Equals($currentError) }

                        if ($null -ne $handledError) {
                            Write-VerboseOutput("Error Index: $index")
                            Write-VerboseOutput($handledError)

                            if ($null -ne $handledError.ScriptStackTrace) {
                                Write-VerboseOutput($handledError.ScriptStackTrace)
                            }
                            Write-VerboseOutput("-----------------------------------`r`n`r`n")
                        }
                    }
        }

        if ($Error.Count -ne $Script:ErrorsExcludedCount) {
            Write-Red("There appears to have been some errors in the script. To assist with debugging of the script, please send the HealthChecker-Debug_*.txt, HealthChecker-Errors.json, and .xml file to ExToolsFeedback@microsoft.com.")
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
            #Need to convert Error to Json because running into odd issues with trying to export $Error out in my lab. Got StackOverflowException for one of the errors i always see there.
            try {
                $Error |
                    ConvertTo-Json |
                    Out-File ("$OutputFilePath\HealthChecker-Errors.json")
            } catch {
                Write-Red("Failed to export the HealthChecker-Errors.json")
                Invoke-CatchActions
            }
        } elseif ($Script:VerboseEnabled -or
            $SaveDebugLog) {
            Write-VerboseOutput("All errors that occurred were in try catch blocks and was handled correctly.")
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
        }
    } else {
        Write-VerboseOutput("No errors occurred in the script.")
    }
}
Function Get-HealthCheckFilesItemsFromLocation {
    $items = Get-ChildItem $XMLDirectoryPath | Where-Object { $_.Name -like "HealthCheck-*-*.xml" }

    if ($null -eq $items) {
        Write-Host("Doesn't appear to be any Health Check XML files here....stopping the script")
        exit
    }
    return $items
}
Function Main {

    if (-not (Confirm-Administrator) -and
        (-not $AnalyzeDataOnly -and
            -not $BuildHtmlServersReport)) {
        Write-Warning "The script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator."
        $Error.Clear()
        Start-Sleep -Seconds 2;
        exit
    }

    $Error.Clear() #Always clear out the errors
    $Script:ErrorsExcludedCount = 0 #this is a way to determine if the only errors occurred were in try catch blocks. If there is a combination of errors in and out, then i will just dump it all out to avoid complex issues.
    $Script:ErrorsExcluded = @()
    $Script:date = (Get-Date)
    $Script:dateTimeStringFormat = $date.ToString("yyyyMMddHHmmss")

    if ($BuildHtmlServersReport) {
        Set-ScriptLogFileLocation -FileName "HealthChecker-HTMLServerReport"
        $files = Get-HealthCheckFilesItemsFromLocation
        $fullPaths = Get-OnlyRecentUniqueServersXMLs $files
        $importData = Import-MyData -FilePaths $fullPaths
        New-HtmlServerReport -AnalyzedHtmlServerValues $importData.HtmlServerValues
        Start-Sleep 2;
        return
    }

    if ((Test-Path $OutputFilePath) -eq $false) {
        Write-Host "Invalid value specified for -OutputFilePath." -ForegroundColor Red
        return
    }

    if ($LoadBalancingReport) {
        Set-ScriptLogFileLocation -FileName "LoadBalancingReport"
        Write-HealthCheckerVersion
        Write-Green("Client Access Load Balancing Report on " + $date)
        Get-CASLoadBalancingReport
        Write-Grey("Output file written to " + $OutputFullPath)
        Write-Break
        Write-Break
        return
    }

    if ($DCCoreRatio) {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try {
            Get-ExchangeDCCoreRatio
            return
        } finally {
            $ErrorActionPreference = $oldErrorAction
        }
    }

    if ($MailboxReport) {
        Set-ScriptLogFileLocation -FileName "HealthCheck-MailboxReport" -IncludeServerName $true
        Get-MailboxDatabaseAndMailboxStatistics
        Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
        return
    }

    if ($AnalyzeDataOnly) {
        Set-ScriptLogFileLocation -FileName "HealthChecker-Analyzer"
        $files = Get-HealthCheckFilesItemsFromLocation
        $fullPaths = Get-OnlyRecentUniqueServersXMLs $files
        $importData = Import-MyData -FilePaths $fullPaths

        $analyzedResults = @()
        foreach ($serverData in $importData) {
            $analyzedServerResults = Start-AnalyzerEngine -HealthServerObject $serverData.HealthCheckerExchangeServer
            Write-ResultsToScreen -ResultsToWrite $analyzedServerResults.DisplayResults
            $analyzedResults += $analyzedServerResults
        }

        New-HtmlServerReport -AnalyzedHtmlServerValues $analyzedResults.HtmlServerValues
        return
    }

    Set-ScriptLogFileLocation -FileName "HealthCheck" -IncludeServerName $true
    Test-RequiresServerFqdn
    Write-HealthCheckerVersion
    [HealthChecker.HealthCheckerExchangeServer]$HealthObject = Get-HealthCheckerExchangeServer
    $analyzedResults = Start-AnalyzerEngine -HealthServerObject $HealthObject
    Write-ResultsToScreen -ResultsToWrite $analyzedResults.DisplayResults
    $currentErrors = $Error.Count

    try {
        $analyzedResults | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 6 -ErrorAction SilentlyContinue
    } catch {
        Write-VerboseOutput("Failed to Export-Clixml. Converting HealthCheckerExchangeServer to json")
        $jsonHealthChecker = $analyzedResults.HealthCheckerExchangeServer | ConvertTo-Json

        $testOuputxml = [PSCustomObject]@{
            HealthCheckerExchangeServer = $jsonHealthChecker | ConvertFrom-Json
            HtmlServerValues            = $analyzedResults.HtmlServerValues
            DisplayResults              = $analyzedResults.DisplayResults
        }

        $testOuputxml | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 6 -ErrorAction Stop
    } finally {
        if ($currentErrors -ne $Error.Count) {
            $index = 0
            while ($index -lt ($Error.Count - $currentErrors)) {
                Invoke-CatchActions $Error[$index]
                $index++
            }
        }

        Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
        Write-Grey("Exported Data Object Written to {0} " -f $Script:OutXmlFullPath)
    }
}

if ($scriptBuildDate -eq "Today") {
    Write-Error ("Script isn't built. Do not run source code directly.`r`nIf developer, follow build process.")
    Write-Host("`r`n`r`nDownload Built Script: https://aka.ms/ExHCDownload")
    exit
}