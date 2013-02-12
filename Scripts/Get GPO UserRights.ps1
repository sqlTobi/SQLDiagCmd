#Based on code from Bart Tukker
#http://www.powershellcommunity.org/Forums/tabid/54/aft/5949/Default.aspx

$a = @{}
$tf = [IO.Path]::GetTempFileName()
$privs = @('SeManageVolumePrivilege','SeLockMemoryPrivilege')

$table = New-Object System.Data.DataTable "UserRights"
$table.columns.add((New-Object System.Data.DataColumn user,([string])))
$table.columns.add((New-Object System.Data.DataColumn privilege,([string])))

$p=[diagnostics.process]::Start('secedit.exe', '/export /cfg ' + $tf)
$p.WaitForExit()
get-content $tf | ForEach-Object {
	foreach ($priv in $privs)
	{
		if ($_ -like $priv + '*') {
			
			$sids = @()
			$sids += ($_ -split {$_ -eq '=' -or $_ -eq ',' -or $_ -eq ' ' -or $_ -eq '*'} | where-object {$_ -notlike 'Se*' -and $_ -notlike ''})
			$i = 0 
			while ($i -le $sids.length -and $sids[$i] -notlike $null) {
				$row = $table.NewRow()
				$row.privilege = $priv
				if ($sids[$i] -notlike 'S-*' ){
					$row.user = $sids[$i]
				} else {
					$sid = New-Object System.Security.Principal.SecurityIdentifier($sids[$i])
					if (($sid.Value -notlike 'S-1-5-21-2000*') -and ($sid.Value -notlike 'S-1-5-32-*')) {
						$row.user = $SID.Value
					} else {
						$User = $sid.Translate([System.Security.Principal.NTAccount])
						$row.user = $User.Value
					}
				}
				
				$table.Rows.Add($row)
				$i++
			}
		}
	}
}
Remove-Item $tf

$ms = New-Object System.IO.MemoryStream
$table.WriteXml($ms, [System.Data.XmlWriteMode]::WriteSchema)
$enc = [System.Text.Encoding]::UTF8
$byte = $ms.ToArray()

$ms.Close | Out-Null
$ms.Dispose | Out-Null

$xml = [xml] $enc.GetString($byte)
$xml.OuterXml