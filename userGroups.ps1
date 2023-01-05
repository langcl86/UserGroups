
Add-Type -AssemblyName System.Windows.Forms
 
class userInput {
    [string]$username;
    [hashtable]$options;
    [string]$output; 
}

function newFLP ($width = 0, $height = 0) {   
    $flp = [System.Windows.Forms.FlowLayoutPanel]::new();
    #$flp.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;

    if (0 -lt $width)  { $flp.Width = $width;   }
    if (0 -lt $height) { $flp.Height = $height; }     
    
    return $flp;
}

function newGroup ($text) {
    $groupBox = [System.Windows.Forms.GroupBox]::new();
    $groupBox.Margin = 10;
    $groupBox.Text = $text;
    $groupBox.Padding = 10;
    return $groupBox;
}

function main {  
    $main = [System.Windows.Forms.Form]::new();
    $main.Width = 800;
    $main.Height = 600;

    $table = [System.Windows.Forms.TableLayoutPanel]::new();
    $table.ColumnCount = 2;
    $table.RowCount = 3;
    $table.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
    $table.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $table.AutoSize = $true;
    $table.Location = "10, 10";
    $table.Padding = 10;
    $table.Margin = 0;

    $ColA = newFLP ($main.Width * 0.35) ($main.Height);

    $usernameUIGroup = newGroup " Username: ";
    $usernameUIGroup.Height = $ColA.Height * 0.12;
    $usernameUIGroup.Width = $ColA.Width * 0.90;
    $usernameUIGroup.Margin = 10;
    
    $usernameFP = newFLP;
    $usernameFP.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $usernameFP.Padding = "10, 0, 5, 0";
   
    $usernameTB = [System.Windows.Forms.TextBox]::new();
    $usernameTB.Width = ($usernameUIGroup.Width * 0.60);
    #$usernameTB.Dock = [System.Windows.Forms.DockStyle]::Fill;
 
    $btnSubmit = [System.Windows.Forms.Button]::new();
    $btnSubmit.Width = ($usernameUIGroup.Width * 0.20);
    #$btnSubmit.Margin = "$($ColA.Width * 0.25), 0, $($ColA.Width * 0.25), 10";
    $btnSubmit.Text = "Submit";
    $btnSubmit.Add_Click({
        $ui = [userInput]::new();
        $ui.username = $usernameTB.Text;

        $ui.options = @{
        "AD"  = $ad_CB.Checked;
        "AAD" = $aad_CB.Checked;
        "OB"  = $ob_CB.Checked; 
        }

        foreach ($c in $optionFLP.Controls) {
            [System.Windows.Forms.CheckBox]$c = $c;
             
        }

        submit($ui);
        $resultsRTB.Text = $ui.output;
        $ui = $null;
    });

    $usernameFP.Controls.Add($usernameTB);
    $usernameFP.Controls.Add($btnSubmit);
    $usernameUIGroup.Controls.Add($usernameFP);
    $ColA.Controls.Add($usernameUIGroup);

    $deptuserGroup = newGroup " Department Users ";
    $deptuserGroup.Height = $ColA.Height * 0.70;
    $deptuserGroup.Width = $ColA.Width * 0.90;

    $deptUsersTree = [System.Windows.Forms.TreeView]::new();
    $deptUsersTree.Dock = [System.Windows.Forms.DockStyle]::Fill;    
    $deptUsersTree.ShowNodeToolTips = $true;

    $deptUsers = Get-ADUser -filter * -Properties Department,LockedOut | ? Enabled -EQ $true;
    $deptUsers.Department | Select -Unique | % {
    $dept = $_
        if ($dept -ne $null) {
            $deptNode = $deptUsersTree.Nodes.Add($_);
            $deptNode.ToolTipText = $dept;
            $deptUsers | ? Department -Match $dept | % {
                $user = $_.SAMAccountName;
                $userNode = $deptNode.Nodes.Add($user);
                $userNode.ToolTipText = $_.Name; 
                
                if ($_.LockedOut) {
                    $deptNode.Expand();
                    $userNode.ForeColor = [System.Drawing.Color]::Red;
                }
            };
        }
    };

    <#
    $deptUser_Group1 = $deptUserLST.Nodes.Add("Group 1");       
        $deptUser_Group1_Item1 = $deptUser_Group1.Nodes.Add("Item 1");
            $deptUser_Group1_Item1.IsSelected
        $deptUser_Group1_Item2 = $deptUser_Group1.Nodes.Add("Item 2");        
    $deptUser_Group2 = $deptUserLST.Nodes.Add("Group 2");

    $deptUserLST.Add_AfterSelect({
        $selItem = $deptUserLST.SelectedNode;
        if ($null -ne $($selItem.Parent)) {
            $usernameTB.Text = $selItem.Text;
        }
    });

    foreach ($node in $deptUserLST.Nodes) {
        [System.Windows.Forms.TreeNode]$node = $node;

        if ($node.Nodes.Count -lt 1) {
            $node.ForeColor = [System.Drawing.Color]::Red;
        }
        else {
            $node.Expand();
        }
    }
    #>

    $deptuserGroup.Controls.Add($deptUsersTree);
    $ColA.Controls.Add($deptuserGroup);

    $table.Controls.Add($ColA, 0, 0);
    $table.SetRowSpan($ColA, 2);

    $ColB1 = newFLP ($main.Width * 0.60) ($usernameUIGroup.Height + 20)

    $optionsGroup = newGroup " Groups ";
    $optionsGroup.Width = $ColB1.Width - 40;
    $optionsGroup.Height = $usernameUIGroup.Height;

    $ad_CB = [System.Windows.Forms.CheckBox]::new();
    $aad_CB = [System.Windows.Forms.CheckBox]::new();
    $ob_CB = [System.Windows.Forms.CheckBox]::new();

    $ad_CB.Text = "Active Directory";
    $ad_CB.Checked = $true;

    $aad_CB.Text = "Azure AD";
    $aad_CB.Checked = $true;

    $ob_CB.Text = "OnBase";
    $ob_CB.Checked = $true;

    $optionFLP = newFLP
    $optionFLP.BorderStyle = [System.Windows.Forms.BorderStyle]::None;
    $optionFLP.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $optionFLP.Padding = "15, 0, 0, 0";

    $optionFLP.Controls.Add($ad_CB);
    $optionFLP.Controls.Add($aad_CB);
    $optionFLP.Controls.Add($ob_CB);

    $optionsGroup.Controls.Add($optionFLP);

    $ColB1.Controls.Add($optionsGroup);

    $table.Controls.Add($ColB1, 1, 0);

    $ColB2 = newFLP $ColB1.Width ($deptuserGroup.Height + 10);
    $ColB2.Margin = "0, 0, 0, 0";

    $resultsGroup = newGroup " Assigned Groups: ";
    $resultsGroup.Width = $colB2.Width -40;
    $resultsGroup.Height = $deptuserGroup.Height;
    $resultsGroup.Padding = 10;

    $resultsRTB = [System.Windows.Forms.RichTextBox]::new();
    $resultsRTB.AutoSize = $true;
    $resultsRTB.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $resultsRTB.Padding = 5;
    $resultsRTB.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
    $resultsMenu = [System.Windows.Forms.ContextMenu]::new();
    $copyResults = $resultsMenu.MenuItems.Add("Copy Results");
    $copyResults.Add_Click({ Set-Clipboard ($resultsRTB.Text); });
    $resultsRTB.ContextMenu = $resultsMenu;
    $resultsGroup.Controls.Add($resultsRTB);

    $ColB2.Controls.Add($resultsGroup);

    $table.Controls.Add($ColB2, 1, 1);

    $main.Controls.Add($table);
    $main.ShowDialog() | Out-Null;
    
    ## Disconnect open AzureAD Connections
    try { 
        Get-AzureADTenantDetail | Out-Null;
        azureAD disconnect;
    }
    catch { <# not connected #> }
}

function submit ([userInput]$ui) {

    if ($ui.options.AD)  { getAD $ui;  }
    if ($ui.options.AAD) { getAAD $ui; }
    if ($ui.options.OB)  { getOB $ui;  }

<#
    $str = @();
    $str += "Write-Host `"Username: $($ui.username)`"";
    $str += "Write-Host `"Options: `"";
    $ui.options | % { $str += "Write-Host `"$_`""; } 

    $sf = Join-Path $env:TEMP "usrgroups.ps1";
    [System.IO.File]::WriteAllLines($sf, $str);
    $pi = [System.Diagnostics.ProcessStartInfo]::new();
    $pi.FileName = "powershell.exe";
    $pi.Arguments = "-NoLogo -WindowStyle Hidden -file $sf";
    $pi.CreateNoWindow = $true;
    $pi.RedirectStandardError = $true;
    $pi.RedirectStandardOutput = $true;
    $pi.UseShellExecute = $false;

    $p = [System.Diagnostics.Process]::new();
    $p.StartInfo = $pi;
    $p.Start();
    
    while (!$p.HasExited) {
        $out = $p.StandardOutput.ReadToEnd();
        $err = $p.StandardError.ReadToEnd();

        if ($out.Length -gt 0) { $resultsRTB.Text += $out.ToString(); }
        
        if ($err.Length -gt 0) { $resultsRTB.Text += $out.ToString(); }

        $resultsRTB.Text += [System.Environment]::NewLine;
    }
#>

}

function newHR {
    $hr = "--------------------------------------------`r`n";
    return $hr;
}

function azureAD ([string]$a) {
 switch ($a)
 {
  connect 
  {
	Write-Host "Connecting To Azure..." -NoNewLine;	  	
	
	$modExists = (Get-Module).Name | ForEach-Object { if ($_ -contains "AzureAD") { return $true; } }
	if ($modExists -ne $true) {
		try {
			Install-Module AzureAD -Force
		}
		catch {
			Write-Host "FAILED to add AzureAD Module `r`n" -ForegroundColor Red;
			Write-Host "Unable to copy O365 information`r`n";
			Write-Host $_.Exception.Message -ForegroundColor Yellow;
			return;
		}
	}

	try {
		Connect-AzureAD -ErrorAction SilentlyContinue | Out-Null		## Attempt connection to AzureAD
		Write-Host "DONE`r" -Foreground Green 
	}
	catch { 
		Write-Host "FAILED to connect to AzureAD`r`n" -ForegroundColor Red;
		Write-Host "Unable to copy O365 information`r`n";
		error $_.Exception.Message -ForegroundColor Yellow;
		return;
	 }
  }
  
  disconnect 
  {
	 Write-Host "Disconnecting from AzureAD..." -NoNewLine
	 Disconnect-AzureAd 
	 Write-Host "DONE`r" -Foreground Green
  }
}

}

function getAD {
	Param (
	[Parameter (Mandatory=$true,Position=0,HelpMessage="userInput Class object")][userInput]$ui 
	)

    $username = $ui.username;

	Write-Host "Searching Active Directory..." -NoNewLine;
		try {
			$aduser = Get-ADUser $username -Properties DistinguishedName,MemberOf
			$adOu = $aduser.DistinguishedName;											## Get OU Data
			$adGroup = $aduser.MemberOf -replace '(CN=)|(,.*)',''						## Get Membership Data	
			Write-Host "DONE" -Foreground Green
		}
		catch { error $_.Exception.Message }

        $string += newHR
        $string += "Active Directory`r`n";
        $string += newHR
        $string += $adGroup | ForEach-Object { $_ + ";`r`n" }
        $string += newHR
        $ui.output += $string;
}

function getOB {
	Param (
	[Parameter (Mandatory=$true,Position=0,HelpMessage="userInput Class object")][userInput]$ui 
	)

    $username = $ui.username;
	
	$server = "database_server_here";
	$database = "databse_name_here";
	$uid = "databse_username_here";
	$pw = "database_password_here";

	## Query User Groups 
	$q = Get-Content "$PSScriptRoot\ob-groups.sql"
	$q = $q.Replace("mpolo", $username);
	
	[string] $connectionString = "Server=$server;Database=$database;Integrated Security = False; User ID = $uid; Password = $pw;"
	
		try {
            Write-Host "Searching OnBase..." -NoNewline
			$conn = New-Object System.Data.SqlClient.SqlConnection($connectionString);
			$conn.Open();
			$command = $conn.CreateCommand()
			$command.CommandText = $q
			$result = $command.ExecuteReader()
			$table = New-Object "System.Data.DataTable";
			$table.Load($result)
			$conn.Close();
			$output = $table.OBGroup
			
			if($output.Count -lt 0 ) { throw; }
			Write-Host "DONE" -ForegroundColor Green

           $string += newHR;
           $string += "OnBase Groups`r`n";
           $string += newHR;
           $string += $output.Trim() | ForEach-Object { $_ + "`r`n"; } 
           $string += newHR;
           $ui.output += $string;
		}
	
		catch { Write-Output "FAILED`r`n"$_.Exception.Message"`r`n"; }
}

function getAAD {
	Param (
	[Parameter (Mandatory=$true,Position=0,HelpMessage="userInput Class object")][userInput]$ui 
	)

    $username = $ui.username;

    ## O365 License Name Lookup
    $names = @{
	    "O365_BUSINESS_ESSENTIALS"		     = "Office 365 Business Essentials"
	    "EXCHANGESTANDARD"				     = "Office 365 Exchange Online Only"
	    "STANDARDPACK"					     = "Enterprise Plan E1"
	    "ENTERPRISEPACK"					 = "Enterprise Plan E3"
	    "POWER_BI_STANDARD"				     = "Power-BI Standard"
	    "MCOMEETADV"						 = "PSTN conferencing"
	    "SHAREPOINTSTORAGE"				     = "SharePoint storage"
	    "ATP_ENTERPRISE"					 = "Exchange Online Advanced Threat Protection"
	    "AAD_PREMIUM"					     = "Azure Active Directory Premium"
	    "DYN365_FINANCIALS_TEAM_MEMBERS_SKU" = "Dynamics 365 for Team Members Business Edition"
	    "FLOW_FREE"						     = "Microsoft Flow Free"
	    "POWER_BI_PRO"					     = "Power BI Pro"
	    "DYN365_ENTERPRISE_SALES"		     = "Dynamics Office 365 Enterprise Sales"
	    "EXCHANGEENTERPRISE"				 = "Exchange Online Plan 2"
	    "WINDOWS_STORE"					     = "Windows Store for Business"
	    "MCOEV"							     = "Microsoft Phone System"
	    "MCOPSTN2"						     = "Domestic and International Calling Plan"
	    "EMSPREMIUM"                         = "ENTERPRISE MOBILITY + SECURITY E5"
	    "SPB"								 = "Microsoft 365 Business Premium"
	    "STREAM"							 = "Microsoft Stream Trial"	
	    "DYN365_AI_SERVICE_INSIGHTS"		 = "Dynamics 365 Customer Service Insights"
	    "POWERAPPS_VIRAL"					 = "Microsoft PowerApps Plan 2 Trial"
	    "MCOPSTNC"							 = "Skype for Business Communication Credits"
	    "TEAMS_EXPLORATORY"					 = "Teams Exploratory Trial"
    }
	    try { Get-AzureADTenantDetail | Out-Null }
	    catch { azureAD connect }

    Write-Host "Searching Azure..." -NoNewLine;
    try {
	    $user = $username + "@contoso.com";
	    $groups = Get-AzureADUSerMembership -ObjectID $user | Select-Object -ExpandProperty DisplayName		## Get Azure Memberships

	    ## Get Azure Licenses
	    $licensePlanList = Get-AzureADSubscribedSku
	    $userList = Get-AzureADUser -ObjectID $user | Select-Object -ExpandProperty AssignedLicenses | Select-Object SkuID 
	    $userList | ForEach-Object { $sku=$_.SkuId ;$licensePlanList | ForEach-Object { If ( $sku -eq $_.ObjectId.substring($_.ObjectId.length - 36, 36) ) { $licenses = $names[$_.SkuPartNumber];} } }
	    Write-Host "DONE" -Foreground Green
    }

    catch { error $_.Exception.Message }

    $string += newHR;
    $string += "Azure AD Groups`r`n";
    $string += newHR;
    $string += $groups  | ForEach-Object { $_ + "`r`n" }
    $string += newHR;
    $string += "`r`n";
    $string += newHR;
    $string += "Azure AD Licenses:`r`n";
    $string += newHR;
    $string += $licenses | ForEach-Object { $_  + "`r`n" }
    $string += newHR;
    $ui.output += $string;
}

main;