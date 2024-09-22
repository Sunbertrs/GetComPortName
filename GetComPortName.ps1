# Get COM port and the bus reported device description script

$number = (Get-WMIObject Win32_PnPEntity | where {$_.name -match "com\d"}).Name
try{
	$name = (Get-WMIObject Win32_PnPEntity | where {$_.name -match "com\d"}).GetDeviceProperties("DEVPKEY_Device_BusReportedDeviceDesc").deviceProperties.Data
}catch{
	echo "No COM ports are detected."
	exit
}

$i=0
foreach($num in $number){
	$num = $num -replace "[^com\d\d]", ""
	if($exist -eq 1){
		break
	}else{
		[PSCustomObject]@{ Number = $num; Name = $name[$i] }
		$exist=1
	}
	$i++
	$exist=0
}