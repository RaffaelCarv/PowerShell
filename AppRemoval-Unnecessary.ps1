﻿<#
.SYNOPSIS
    Remove built-in apps (modern apps) from Windows 10.
.DESCRIPTION
    This script was removal apps unnecessary.
.EXAMPLE
    .\AppRemoval-Unnecessary.ps1
.NOTES
    FileName:    AppRemoval-Unnecessary.ps1
    Author:      Rafael Carvalho
    Created:     2021-04-26
    Version history:
    1.0.0 - (2021-04-26) Initial script updated with help section and a fix for randomly freezing
    
#>
$listOfApps = get-appxpackage
$appToRemove = $listOfApps | where-object {$_ -like "*Xbox*"}
Remove-AppxPackage -package $appToRemove.packagefullname

$listOfApps = get-appxpackage
$appToRemove = $listOfApps | where-object {$_ -like "*Solitaire*"}
Remove-AppxPackage -package $appToRemove.packagefullname