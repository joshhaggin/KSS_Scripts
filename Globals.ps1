#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory
$source = "\\kiewitplaza\ktg\Active\KSS\KSS_Toolkit\KSS MultiTool"


$KSS_leads = Get-Content "$source\lib\KSS_SD_Lead_Associates.csv"

$lib = "$source\lib\scripts\"
$lib_KSS_Notes = "$source\lib\KSS_Notes"
$lib_nobackslash = "$source\lib\scripts"

$names_kss = get-childItem $lib -filter *.txt -Recurse
$global:filenames_tech = get-childItem $env:LOCALAPPDATA\KSS\clipboard_scripts\ -filter *.txt -Recurse

$text = ".txt"

$local = "\KSS\clipboard_scripts\"
$appdata_local_kss = $env:LOCALAPPDATA + $local

$global:TechScripts = 0
$global:KSSScripts = 0




$script_refresh = {
	if ($(test-path c:\Users\$env:USERNAME\AppData\Local\KSS\clipboard_scripts) -eq $false)
	{
		New-Item c:\Users\$env:USERNAME\AppData\Local\KSS\clipboard_scripts -Type Directory
	}
	
	if ($KSS_leads -contains $env:USERNAME)
	{
		$richtextbox_KSSScripts.readonly = $true
	}
	
	$global:filenames_kss = get-childItem $lib -filter *.txt -Recurse
	$global:filenames_tech = get-childItem $env:LOCALAPPDATA\KSS\clipboard_scripts\ -filter *.txt -Recurse
	
	$script_path = Test-Path c:\Users\$env:UserName\AppData\Local\kss\
	if ($script_path -ne $true)
	{
		New-Item c:\Users\$env:UserName\AppData\Local\KSS\ -Type Directory -Force
		$script_file = test-path c:\Users\$env:UserName\AppData\Local\kss\user_script.txt
		if ($script_file -ne $true)
		{
			New-Item c:\Users\$env:UserName\AppData\Local\kss\user_script.txt -Type file
		}
	}
	
	try
	{
		$richtextbox_TechScripts.Lines = Get-Content c:\Users\$env:UserName\AppData\Local\kss\user_script.txt
	}
	catch
	{
		$richtextbox_TechScripts.Text = "Could not read from source file or file empty."
	}
	
	try
	{
		$richtextbox_KSSScripts.Lines = Get-Content $lib_KSS_Notes\KSS_Notes.txt
	}
	catch
	{
		$richtextbox_KSSScripts.Text = "Could not read from source file or file empty."
	}
	
	#--------------------------------------------------------
	## get names of Tech scripts
	#--------------------------------------------------------
	$linklabel_Custom1.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[0].name)
	$linklabel_Custom2.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[1].name)
	$linklabel_Custom3.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[2].name)
	$linklabel_Custom4.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[3].name)
	$linklabel_Custom5.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[4].name)
	$linklabel_Custom6.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[5].name)
	$linklabel_Custom7.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[6].name)
	$linklabel_Custom8.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[7].name)
	$linklabel_Custom9.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[8].name)
	$linklabel_Custom10.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[9].name)
	$linklabel_Custom11.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[10].name)
	$linklabel_Custom12.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[11].name)
	$linklabel_Custom13.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[12].name)
	$linklabel_Custom14.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[13].name)
	$linklabel_Custom15.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[14].name)
	$linklabel_Custom16.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[15].name)
	$linklabel_Custom17.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[16].name)
	$linklabel_Custom18.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[17].name)
	$linklabel_Custom19.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[18].name)
	$linklabel_Custom20.Text = [io.path]::GetFileNameWithoutExtension($global:filenames_tech[19].name)
	
	#--------------------------------------------------------
	## make links visible if they have content
	#--------------------------------------------------------
	if ($linklabel_Custom1.Text -eq $null)
	{
		$linklabel_Custom1.Visible = $false
	}
	if ($linklabel_Custom2.Text -eq $null)
	{
		$linklabel_Custom2.Visible = $false
	}
	if ($linklabel_Custom3.Text -eq $null)
	{
		$linklabel_Custom3.Visible = $false
	}
	if ($linklabel_Custom4.Text -eq $null)
	{
		$linklabel_Custom4.Visible = $false
	}
	if ($linklabel_Custom5.Text -eq $null)
	{
		$linklabel_Custom5.Visible = $false
	}
	if ($linklabel_Custom6.Text -eq $null)
	{
		$linklabel_Custom6.Visible = $false
	}
	if ($linklabel_Custom7.Text -eq $null)
	{
		$linklabel_Custom7.Visible = $false
	}
	if ($linklabel_Custom8.Text -eq $null)
	{
		$linklabel_Custom8.Visible = $false
	}
	if ($linklabel_Custom9.Text -eq $null)
	{
		$linklabel_Custom9.Visible = $false
	}
	if ($linklabel_Custom10.Text -eq $null)
	{
		$linklabel_Custom10.Visible = $false
	}
	if ($linklabel_Custom11.Text -eq $null)
	{
		$linklabel_Custom11.Visible = $false
	}
	if ($linklabel_Custom12.Text -eq $null)
	{
		$linklabel_Custom12.Visible = $false
	}
	if ($linklabel_Custom13.Text -eq $null)
	{
		$linklabel_Custom13.Visible = $false
	}
	if ($linklabel_Custom14.Text -eq $null)
	{
		$linklabel_Custom14.Visible = $false
	}
	if ($linklabel_Custom15.Text -eq $null)
	{
		$linklabel_Custom15.Visible = $false
	}
	if ($linklabel_Custom16.Text -eq $null)
	{
		$linklabel_Custom16.Visible = $false
	}
	
	if ($linklabel_Custom17.Text -eq $null)
	{
		$linklabel_Custom17.Visible = $false
	}
	if ($linklabel_Custom18.Text -eq $null)
	{
		$linklabel_Custom18.Visible = $false
	}
	if ($linklabel_Custom19.Text -eq $null)
	{
		$linklabel_Custom19.Visible = $false
	}
	if ($linklabel_Custom20.Text -eq $null)
	{
		$linklabel_Custom20.Visible = $false
	}
	
	#--------------------------------------------------------
	## get names of KSS scripts
	#--------------------------------------------------------
	$linklabel_Content1.Text = [io.path]::GetFileNameWithoutExtension($names_kss[0].name)
	$linklabel_Content2.Text = [io.path]::GetFileNameWithoutExtension($names_kss[1].name)
	$linklabel_Content3.Text = [io.path]::GetFileNameWithoutExtension($names_kss[2].name)
	$linklabel_Content4.Text = [io.path]::GetFileNameWithoutExtension($names_kss[3].name)
	
	#--------------------------------------------------------
	## make links visible if they have content
	#--------------------------------------------------------
	if ($linklabel_Content1.Text -eq $null)
	{
		$linklabel_Content1.Visible = $false
	}
	if ($linklabel_Content2.Text -eq $null)
	{
		$linklabel_Content2.Visible = $false
	}
	if ($linklabel_Content3.Text -eq $null)
	{
		$linklabel_Content3.Visible = $false
	}
	if ($linklabel_Content4.Text -eq $null)
	{
		$linklabel_Content4.Visible = $false
	}
	
	
	#--------------------------------------------------------
	## Load tabs or not, based upon if they have text in the tabs
	#--------------------------------------------------------
	$script:savedpages = @{ }
	$tabcontrol1.TabPages | %{ $script:savedpages.Add($_.Name, $_) }
	
	$newtabs = $(Get-ChildItem $env:LOCALAPPDATA\kss\tabpage_new\).basename
	
	if (Test-Path $env:LOCALAPPDATA\kss\tabpage_new\$($newtabs[0]).txt)
	{
		try
		{
			$tabcontrol1.TabPages[2].Text = $newtabs[0]
			$richtextbox_New1.Lines = Get-Content $env:LOCALAPPDATA\kss\tabpage_new\$($newtabs[0]).txt
		}
		catch
		{
			$richtextbox_New1.Text = "Could not read from source file or file empty."
		}
	}
	else
	{
		$tabcontrol1.TabPages | %{ $tabcontrol1.TabPages.Remove($tabpage_New1) }
	}
	
	if (Test-Path $env:LOCALAPPDATA\kss\tabpage_new\$($newtabs[1]).txt)
	{
		try
		{
			$tabcontrol1.TabPages[3].Text = $newtabs[1]
			$richtextbox_New2.Lines = Get-Content $env:LOCALAPPDATA\kss\tabpage_new\$($newtabs[1]).txt
		}
		catch
		{
			$richtextbox_New2.Text = "Could not read from source file or file empty."
		}
	}
	else
	{
		$tabcontrol1.TabPages | %{ $tabcontrol1.TabPages.Remove($tabpage_New2) }
	}
	
	if (Test-Path $env:LOCALAPPDATA\kss\tabpage_new\$($newtabs[2]).txt)
	{
		try
		{
			$tabcontrol1.TabPages[4].Text = $newtabs[2]
			$richtextbox_New3.Lines = Get-Content $env:LOCALAPPDATA\kss\tabpage_new\$($newtabs[2]).txt
		}
		catch
		{
			$richtextbox_New3.Text = "Could not read from source file or file empty."
		}
	}
	else
	{
		$tabcontrol1.TabPages | %{ $tabcontrol1.TabPages.Remove($tabpage_New3) }
	}
	
	<#
	#--------------------------------------------------------
	## Load tabs in combo box
	#--------------------------------------------------------
	$fill_combobox = @()
	$fill_combobox = $(Get-ChildItem $env:LOCALAPPDATA\KSS\tabpage_new).basename
	Update-ToolStripComboBox $toolstripcombobox_tabs $fill_combobox
	#>
	
	#--------------------------------------------------------
	## sizing the tech scripts groupbox
	#--------------------------------------------------------
	if ($global:filenames_tech.count -eq "4")
	{
		$groupbox_size = New-Object drawing.size(352, 60)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "5" -or $global:filenames_tech.count -eq "6")
	{
		$groupbox_size = New-Object drawing.size(352, 82)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "7" -or $global:filenames_tech.count -eq "8")
	{
		$groupbox_size = New-Object drawing.size(352, 104)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "9" -or $global:filenames_tech.count -eq "10")
	{
		$groupbox_size = New-Object drawing.size(352, 126)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "11" -or $global:filenames_tech.count -eq "12")
	{
		$groupbox_size = New-Object drawing.size(352, 148)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "13" -or $global:filenames_tech.count -eq "14")
	{
		$groupbox_size = New-Object drawing.size(352, 170)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "15" -or $global:filenames_tech.count -eq "16")
	{
		$groupbox_size = New-Object drawing.size(352, 192)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "17" -or $global:filenames_tech.count -eq "18")
	{
		$groupbox_size = New-Object drawing.size(352, 214)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	elseif ($global:filenames_tech.count -eq "19" -or $global:filenames_tech.count -eq "20")
	{
		$groupbox_size = New-Object drawing.size(352, 236)
		$groupbox_TechScripts.Size = $groupbox_size
	}
	
	$splitcontainer2.Panel1minsize = $groupbox_TechScripts.Size.height
	
}



#region Control Helper Functions
function Update-ComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.
	
	.PARAMETER ComboBox
		The ComboBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ComboBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red" -Append
		Update-ComboBox $combobox1 "White" -Append
		Update-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Update-ComboBox $combobox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ComboBox]$ComboBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	$ComboBox.DisplayMember = $DisplayMember
}
#endregion

#region Control Helper Functions
function Update-ToolStripComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ToolStripComboBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ToolStripComboBox control.
	
	.PARAMETER ToolStripComboBox
		The ToolStripComboBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ToolStripComboBox's Items collection.
	
	.PARAMETER Append
		Adds the item(s) to the ToolStripComboBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ToolStripComboBox $toolStripComboBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ToolStripComboBox $toolStripComboBox1 "Red" -Append
		Update-ToolStripComboBox $toolStripComboBox1 "White" -Append
		Update-ToolStripComboBox $toolStripComboBox1 "Blue" -Append
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ToolStripComboBox]$ToolStripComboBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ToolStripComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ToolStripComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [Array])
	{
		$ToolStripComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ToolStripComboBox.Items.Add($obj)
		}
		$ToolStripComboBox.EndUpdate()
	}
	else
	{
		$ToolStripComboBox.Items.Add($Items)
	}
}
#endregion