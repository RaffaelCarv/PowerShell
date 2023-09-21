$ADUser = (Get-AdUser -Filter *).Count
$ADGroup = (Get-ADGroup -Filter *).Count
$ADComputer = (Get-ADComputer -Filter *).Count
$ADObjects = $ADUser + $ADGroup + $ADComputer
$ADObjects