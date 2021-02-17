# SharePoint to Share Drive
# https://github.com/tcsung/SharePoint-to-share-drive
# By Daniel Sung	Feb 2021


# ============================
# Define - basic structure
# ============================


$ErrorActionPreference		= 'SilentlyContinue'
$script:info			= @{}
$script:env_path		= @{}
$script:reg_path		= @{}
$script:vbmsg			= new-object -comobject wscript.shell

$reg_path.self			= 'Registry::HKCU\SOFTWARE\SP2SD'
$reg_path.addresses		= $reg_path.self + '\addresses'
$reg_path.shares		= $reg_path.self + '\shares'
$reg_path.zonemap		= 'Registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains'
$reg_path.zone			= 'Registry::HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2'
$env_path.progdir		= $env:programfiles
$env_path.ie			= $env_path.progdir + '\Internet Explorer\iexplore.exe'
$env_path.letters		= 'G:','H:','I:','J:','K:','L:','M:','N:','O:','P:','Q:','R:','S:','T:','U:','V:','W:','X:','Y:','Z:'
$info.producttype		= (get-itemproperty -path $reg_path.hklm_ntcurrentver).'ProductName'
$info.this_version		= '1.0'

# =================================
# FUNCTION - get_max 
# =================================
function get_max{
	param(
		[array]$array
	)

	$max = $null

	foreach($value in $array){
		if($value -eq '(default)'){
			$value = $null
		}else{
			$value = [int]$value
		}
		if($max -eq $null){ $max = $value }
		if($value -gt $max){ $max = $value }
	}
	return $max
}


# =================================
# FUNCTION - write_reg
# =================================
function write_reg{
	param(
		[parameter(mandatory=$true)][string]$reg_key,
		[string]$reg_value,
		[string]$reg_name,
		[string]$reg_type
	)

	$outcome = 0

	switch($reg_type){
		'REG_SZ'	{$private:type_name = 'string'}
		'REG_EXPAND_SZ'	{$private:type_name = 'ExpandString'}
		'REG_BINARY'	{$private:type_name = 'Binary'}
		'REG_DWORD'	{$private:type_name = 'DWord'}
		'REG_MUTI_SZ'	{$private:type_name = 'MultiString'}
		'REG_QWORD'	{$private:type_name = 'QWord'}
		default		{
				$private:type_name = 'string'
				$reg_type = 'REG_SZ'
				}	
	}

	if($reg_key -imatch '^Registry\:\:.*'){
		$reg_key = $reg_key -replace 'Registry\:\:',''
	}
	

	if(! $reg_name){$reg_name = '(Default)'}
	if(!(test-path Registry::$reg_key)){new-item -path ('Registry::' + $reg_key) -Force -Erroraction silentlyContinue}
	if(test-path Registry::$reg_key){
		if($reg_type -eq 'REG_BINARY'){
			$value = [byte[]][System.Text.Encoding]::UniCode.GetBytes($reg_value)
		}else{
			$value = $reg_value
		}
		try{
			set-itemproperty -path ('Registry::' + $reg_key) -name $reg_name -value $value -type $type_name | out-null
		}
		catch{
			write-warning 'ERROR : Fail to right value :' $_
			$outcome = -1
		}
	}else{
		write-warning 'ERROR : Fail create new registry value due to registry key path fail to create.'
		$outcome = -1
	}
	return $outcome
}


# =================================
# FUNCTION - run_ie
# =================================
function run_ie{
	param()

	$info.trustzone = $T_address.text
	if($info.trustzone -match '^(http.*)\:\/\/(.+)\.(.+\.com)$'){
		$http = $matches[1]
		$subdomain = $matches[2]
		$domain = $matches[3]
		if($http -eq 'https'){
			$http = 2
		}else{
			$http = 0
		}
		
		$cracked = $true

		$reg_trust_zone = $reg_path.zonemap + '\' + $domain + '\' + $subdomain
		$search = (get-itemproperty -path $reg_trust_zone).'https'
		if(! $search){
			$do = &write_reg -reg_key $reg_trust_zone -reg_name 'https' -reg_value $http -reg_type 'REG_DWORD'
			if($do -eq -1){
				$cracked = $false
			}
		}
		if($cracked -eq $false){
			$warn = $vbmsg.popup("Something went wrong, fail to prepare related configuration.",0,"SharePoint2ShareDrive Fail.",0)
			$L_result.text = 'Fail to prepare related configuration.'
			$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ff0000")
			$B_action2.Enabled = $false
			$B_action3.Enabled = $false
		}else{
			$L_result.text = 'Opening your Internet Explorer ...'
			$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#002bc9")
			start-process $env_path.ie -ArgumentList $info.trustzone -ErrorAction silentlyContinue
			$Listbox1.text = $T_address.text + '/'
			$search_list =@()
			$search_list += get-item -path $reg_path.shares | select-object -expandproperty property
			foreach($line in $search_list){
				if($line){
					$saved_share = (get-itemProperty -path $reg_path.shares).$line
					if (($saved_share -like "*$subdomain*" ) -and ($saved_share -like "*$domain*")){
						$Listbox1.Items.add($saved_share)
					}
				}
			}
		}
	}else{
		$warn = $vbmsg.popup("The web address you input seems incorrect.`nYou may omitted the 'https://' at the begin.",0,"SharePoint2ShareDrive Fail.",0)
		$L_result.text = 'The web address contain some mistake.'
		$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ff0000")
	}

	
}

# =================================
# FUNCTION - map_drive
# =================================
function map_drive{
	param()
	$share_drive = $Listbox2.text
	$go_map_drive = $false
	if(test-path -path ($share_drive + '\')){
		$c = 0
		$parameters = "use $share_drive /delete /y"
		$do = start-process net -argumentlist $parameters -ErrorAction silentlyContinue -PassThru -WindowStyle hidden
		do{
			start-sleep -Seconds 1
			$c++
		}until($do.HasExited -eq $true -or $c -ge 15)
		#(New-object -ComObject WScript.Network).RemoveNetworkDrive("$share_drive")
		#if($?){
		if($do.exitcode -eq 0){
			$go_map_drive = $true
		}else{
			$warn = $vbmsg.popup("Fail to remove the existing drive mapping.",0,"SharePoint2ShareDrive Fail.",0)
			$L_result.text = "Fail to remove the existing drive mapping."
			$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ff0000")
		}
	}else{
		$go_map_drive = $true
	}

	if($go_map_drive -eq $true){
		$path = $Listbox1.text
		if($info.producttype -match 'Windows 7'){
			net use $share_drive \\192.168.0.250\abc | out-null
			start-sleep -Seconds 2
		}
		(New-object -ComObject WScript.Network).MapNetworkDrive("$share_drive","$path",$true)
		if($?){
			$L_result.text = ""
			start-sleep -Milliseconds 100
			$L_result.text = "Share drive ready."
			$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#002bc9")


			$go_save = $true
			$search_list = @()
			$search_list += get-item -path "$reg_path.addresses" | select-object -expandproperty property
			foreach($line in $search_list){
				if($line){
					$search_address = (get-itemProperty -path $reg_path.addresses).$line
					if($search_address -eq $T_address.text){
						$go_save = $false
						break
					}
				}
			}
			if($go_save -eq $true){
				if($search_list.Count -eq 0){
					$name = 0
				}else{
					$name = (&get_max -array $search_list) + 1
				}
				$do = &write_reg -reg_key $reg_path.addresses -reg_name $name -reg_value $T_address.text -reg_type 'REG_SZ'
			}
			$do = &write_reg -reg_key $reg_path.addresses -reg_value $T_address.text -reg_type 'REG_SZ'

			$go_save = $true
			$search_list = @()
			$search_list += get-item -path $reg_path.shares | select-object -expandproperty property
			foreach($line in $search_list){
				if($line){
					$search_address = (get-itemProperty -path $reg_path.shares).$line
					if($search_address -eq $Listbox1.text){
						$go_save = $false
						break
					}
				}
			}
			if($go_save -eq $true){
				if($search_list.Count -eq 0){
					$name = 0
				}else{
					$name = (&get_max -array $search_list) + 1
				}
				$do = &write_reg -reg_key $reg_path.shares -reg_name $name -reg_value $Listbox1.text -reg_type 'REG_SZ'
			}

		}else{
			$warn = $vbmsg.popup("Something went wrong, please ensure you logged in your SharePoint account in IE.",0,"SharePoint2ShareDrive Fail.",0)
			$L_result.text = "Fail to Map drive."
			$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ff0000")
		}
	}
}


# ======================
#  Main Script
# ======================

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Form                     = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(720,520)
$Form.text                       = "SharePoint to Share Drive " + $info.this_version
$Form.TopMost                    = $false

$script:L_title                  = New-Object system.Windows.Forms.Label
$L_title.text                    = "SharePoint to Share Drive " + $info.this_version
$L_title.AutoSize                = $true
$L_title.width                   = 25
$L_title.height                  = 10
$L_title.location                = New-Object System.Drawing.Point(195,31)
$L_title.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',16,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

$script:L_description            = New-Object system.Windows.Forms.Label
$L_description.text              = "To connect your SharePoint like a share drive on your computer,`nplease follow the steps below to enable this feature."
$L_description.AutoSize          = $true
$L_description.width             = 25
$L_description.height            = 20
$L_description.location          = New-Object System.Drawing.Point(82,80)
$L_description.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$script:L_step1                  = New-Object system.Windows.Forms.Label
$L_step1.text                    = '1. Input your SharePoint address or choose from the list :'
$L_step1.AutoSize                = $true
$L_step1.width                   = 25
$L_step1.height                  = 10
$L_step1.location                = New-Object System.Drawing.Point(77,155)
$L_step1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$script:T_address                 = New-Object system.Windows.Forms.ComboBox
$T_address.text                   = (get-itemProperty -path $reg_path.addresses).'(default)'
$T_address.width                  = 330
$T_address.height                 = 10
$T_address.location               = New-Object System.Drawing.Point(97,185)
$T_address.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$get_addresses = @()
$get_addresses += get-item -path $reg_path.addresses | select-object -expandproperty property
foreach($line in $get_addresses){
	if($line -and ($line -ne '(default)')){
		$saved_address = (get-itemProperty -path $reg_path.addresses).$line
		$x = $T_address.Items.add($saved_address)
	}
}


$script:B_action2                = New-Object system.Windows.Forms.Button
$B_action2.text                  = 'Run Internet Explorer'
$B_action2.width                 = 171
$B_action2.height                = 30
$B_action2.location              = New-Object System.Drawing.Point(487,183)
$B_action2.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$B_action2.Add_Click({ $do = &run_ie })

$script:L_step3                  = New-Object system.Windows.Forms.Label
$L_step3.text                    = '2. Input the full path you need to map as drive or choose from the list :'
$L_step3.AutoSize                = $true
$L_step3.width                   = 25
$L_step3.height                  = 10
$L_step3.location                = New-Object System.Drawing.Point(77,245)
$L_step3.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$script:Listbox1                 = New-Object system.Windows.Forms.ComboBox
$Listbox1.text                   = ''
$Listbox1.width                  = 500
$Listbox1.height                 = 32
$Listbox1.location               = New-Object System.Drawing.Point(97,270)
$Listbox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',9.5)

$script:L_step3a                 = New-Object system.Windows.Forms.Label
$L_step3a.text                   = 'Select the Drive letter :'
$L_step3a.AutoSize               = $true
$L_step3a.width                  = 25
$L_step3a.height                 = 10
$L_step3a.location               = New-Object System.Drawing.Point(98,305)
$L_step3a.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$script:Listbox2                 = New-Object system.Windows.Forms.ComboBox
$Listbox2.text                   = 'Z:'
$Listbox2.width                  = 60
$Listbox2.height                 = 32
$Listbox2.location               = New-Object System.Drawing.Point(100,330)
$Listbox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
foreach($l in $env_path.letters){
	$x = $Listbox2.Items.add($l)
}

$script:B_action3                = New-Object system.Windows.Forms.Button
$B_action3.text                  = 'Create share drive'
$B_action3.width                 = 172
$B_action3.height                = 30
$B_action3.location              = New-Object System.Drawing.Point(487,328)
$B_action3.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$B_action3.Add_Click({
	if($Listbox1.text -eq $null){
		$warn = $vbmsg.popup("You are not yet choose / input the link for map the drive,`nplease try again.",0,"SharePoint2ShareDrive Fail.",0)
	}else{
		$L_result.text = 'Connecting .....'
		$L_result.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#002bc9")
		$do = &map_drive
	}
})

$script:G_stat                   = New-Object system.Windows.Forms.Groupbox
$G_stat.width                    = 580
$G_stat.height                   = 50
$G_stat.text                     = 'Status'
$G_stat.location                 = New-Object System.Drawing.Point(80,375)
$G_stat.Font			 = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$script:L_result                 = New-Object system.Windows.Forms.Label
$L_result.AutoSize               = $true
$L_result.width                  = 25
$L_result.height                 = 27
$L_result.location               = New-Object System.Drawing.Point(18,20)
$L_result.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',11)
$L_result.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#002bc9")

$script:B_close                  = New-Object system.Windows.Forms.Button
$B_close.text                    = 'Close'
$B_close.width                   = 115
$B_close.height                  = 34
$B_close.location                = New-Object System.Drawing.Point(540,450)
$B_close.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$B_close.Add_Click({ 
	$form.close()
})

$Form.controls.AddRange(@($L_title,$L_description,$L_step1,$T_address,$B_action2,$L_step3,$B_action3,$B_close,$G_stat,$Listbox1,$L_step3a,$Listbox2))
$G_stat.controls.AddRange(@($L_result))



# ---------------[Show form]-------------------
# Show the form
$result = $form.ShowDialog()

