param(
    [string]$EXEC_Code,
	[string]$EXEC_Value
    )
	
	import-module ActiveDirectory

#Umlaute Encoding
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

#FPITLBXXX -> Creates Listbox
#FPITSLXXX -> Creates Testbox
#FPITLSXXX -> Creates Testbox




if ($EXEC_Code -eq "FPITSL002")
{
	if ($EXEC_Value -eq "rueckgabe1")
	{
		return "zu_webjea_1";
	}
	elseif ($EXEC_Value -eq "rueckgabe3" )
	{
		return "zu_webjea_2";
	}
	else
	{
		return "zu_webjea_3";
	}
}
elseif($EXEC_Code -eq "FPITLB020")
{
	$blub = @()
	$blub = (get-adgroup -Filter * -Properties * | select Name).Name
	$Ausgabe = ""
	foreach($user in $blub)
	{
		$Ausgabe = $Ausgabe + $user + ";"
	}

	return $Ausgabe
}
else
{
	return "zurueck"
}
