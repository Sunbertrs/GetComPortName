# Get COM port and the bus reported device description script
# By Sunbertrs

$number = (Get-WMIObject Win32_PnPEntity | where {$_.name -match "com\d"}).Name
$data = (Get-WMIObject Win32_PnPEntity | where {$_.name -match "com\d"}).GetDeviceProperties("DEVPKEY_Device_BusReportedDeviceDesc").deviceProperties.Data
$i=0
foreach($num in $number){
	$num = $num -replace "[^com\d\d]", ""
	if($exist -eq 1){
		break
	}else {
		[PSCustomObject]@{ Number = $num; Name = $data[$i] }
		$exist=1
	}
	$i++
	$exist=0
}
