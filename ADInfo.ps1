<#
.NOTES
	Name: ADInfo.ps1
	Original Author: Shane Moore
    	Author: Shane Moore
    	Contributor: Rafael Carvalho - Live3 IT Solutions
	** Requires: PowerShell and administrator rights on the target ADs server as well as the local machine.**
   
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Check the integrity of Active Directory in the internal domain.
.DESCRIPTION
	This script checks the integrity of Active Directory for various configuration recommendations.
.EXAMPLE
	This cmdlet with run ADInfo Script by default and run against the local server.
	.\HealthChecker.ps1

#>

Import-Module ActiveDirectory

#---------------------------------------------------------------------
# Variables
#---------------------------------------------------------------------

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

Foreach ($domain in $domains)
{
    $domaindetails = Get-ADDomain $domain

    $domaininfo = New-Object PSObject
    
    $domaininfo | Add-Member NoteProperty -Name "Name" -Value $domaindetails.Name    
}

#---------------------------------------------------------------------
# Collect AD Forest information
#---------------------------------------------------------------------

Write-Host "*** For this Domain there are ***" -ForegroundColor Yellow
Write-Host "Domain Controller = "$DC
Write-Host "Computers         = "$Computers
Write-Host "Workstions        = "$Workstations
Write-Host "Servers           = "$Servers
Write-Host "Users             = "$Users
Write-Host ""
Write-host "*** Active Directory Info ***" -ForegroundColor Yellow
Write-Host "Active Directory Forest Mode =  "$ADForest
Write-Host "Active Directory Domain Mode =  "$ADDomain
Write-Host "Active Directory Schema Version is $ADNum which corresponds to $Srv"
Write-Host "Additional UPN Suffixes: $($forest.UPNSuffixes)"
Write-Host "NetBIOS Name: $($domaindetails.NetBIOSName)"
Write-Host "*** FSMO Role Owners ***" -ForegroundColor Yellow
$FSMO
Write-Host ""

#####################################Get ALL DC Servers#################################

$getForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()

$DCServers = $getForest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 

$timeout = "60"

Write-Host "*** Domain Controllers List ***" -ForegroundColor Yellow
Write-Host "Domain Controllers = "$DCServers
Write-Host ""
Write-Host "*** Active Directory Health Check ***" -ForegroundColor Yellow
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
