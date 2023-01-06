
Add-Type -AssemblyName System.Windows.Forms, Microsoft.VisualBasic
 
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
    checkReq;

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
    $deptUsersMenu = [System.Windows.Forms.ContextMenu]::new();
    
    $refreshDeptUsers = $deptUsersMenu.MenuItems.Add("Refresh");
    $refreshDeptUsers.Add_Click({getDeptUsers});
    
    $unlockUser = $deptUsersMenu.MenuItems.Add("Unlock User");
    $unlockUser.Enabled = $false;
    
    $findUser = $deptUsersMenu.MenuItems.Add("Find User");
    $findUser.Add_Click({
        $u = [Microsoft.VisualBasic.Interaction]::InputBox("Enter username to find department", "Find User Department");
        $d = $deptUsers | ? { $_.SAMAccountName -eq $u } | Select -ExpandProperty Department;

        if ($d.Count -eq 1) {
            foreach ($dept in $d) {
                $dNode = $deptUsersTree.Nodes | ? Text -EQ $dept;
                $dNode.Expand();
                $uNode = $dNode.Nodes | ? Text -eq $u;
                $deptUsersTree.SelectedNode = $uNode;
            }
        }

        elseif ($d.Count -eq 0 -and $u.Length -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("Username $u was not found", "User not found", "OK", "Information");
        }

    });

    function getDeptUsers {
    
        try {
            $deptUsers = Get-ADUser -filter * -Properties Department,LockedOut | ? Enabled -EQ $true;
        }
        catch {
            $msg = "Get-ADUser failed to retrieve AD users`r`n`r`n$($_.Exception.Message)"; 
            newError $msg;
        }

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

                }
            };
        };

        <#
        $deptUser_Group1 = $deptUserLST.Nodes.Add("Group 1");       
            $deptUser_Group1_Item1 = $deptUser_Group1.Nodes.Add("Item 1");
                $deptUser_Group1_Item1.IsSelected
            $deptUser_Group1_Item2 = $deptUser_Group1.Nodes.Add("Item 2");        
        $deptUser_Group2 = $deptUserLST.Nodes.Add("Group 2");

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
    }

    $unlockUser.Add_Click({
        try {
            $u = $deptUsersTree.SelectedNode.Text;
            Unlock-ADAccount $u;
            getDeptUsers;
        }
        catch {
            $msg = "Failed to unlock user account $($deptUsersTree.SelectedNode.Text)`r`n`r`n$($_.Exception.Message)";
            newError $msg;
        }
    });

    $deptUsersTree.Add_AfterSelect({
        $selItem = $deptUsersTree.SelectedNode;
            
        if ($null -ne $($selItem.Parent)) {
            $usernameTB.Text = $selItem.Text;
        }

        $isLocked = $deptUsers | ? { $_.SAMAccountName -eq $selItem.Text -and $_.LockedOut -eq $true };
        if ($isLocked) { $unlockUser.Enabled = $true; }
        else  { $unlockUser.Enabled = $false; }
    });

    getDeptUsers;
    $deptuserGroup.Controls.Add($deptUsersTree);
    $deptUsersTree.ContextMenu = $deptUsersMenu;
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

function submit {
    Param(
        [Parameter(Mandatory=$true, Position=0)][userInput]$ui
     )

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

function newError {
<#
    .SYNOPSIS
        Helper function displays Error MessageBox. 
#>
    Param(
        [Parameter(Mandatory=$true,Position=0)][string]$msg
    )

    $btns = [System.Windows.Forms.MessageBoxButtons]::OK;
    $icns = [System.Windows.Forms.MessageBoxIcon]::Error;
    [System.Windows.Forms.MessageBox]::Show($msg, "Error", $btns, $icns) | Out-Null;
}

function newHR {
<#
    .SYNOPSIS
        Returns a string used as a horizonal rule.
#>
    $hr = "--------------------------------------------`r`n";
    return $hr;
}

function checkReq {
<#
    .SYNOPSIS
        Check for/Install required PowerShell Modules.

    .DESCRIPTION
        This helper function checks to see if required Powershell modules are installed.
        If modules are not found, an attempt will be made to install the module.
        Failure to install required modules will result in an error and limited or no script functionality. 

    .REMARKS
        https://learn.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps
        https://learn.microsoft.com/en-us/powershell/azure/active-directory/overview?view=azureadps-2.0
#>
    $reqMods = @("AzureAD", "ActiveDirectory");

    foreach ($mod in $reqMods) {
        try {
            if ($null -eq (Get-Module $mod)) {
                Write-Host "Attempting to install PS module $mod..." -NoNewline
                Install-Module $mod -Force;
                Write-Host "DONE" -ForegroundColor Green;
            }
        }
        catch {
            Write-Host "FAILED" -ForegroundColor Red;
            $msg = "PS Module $mod Not Found`r`n`r`n$($_.Exception.Message)";
            newError $msg;
        }
    }

}

function azureAD {
<#
    .SYNOPSIS
        This function is used to establish or disconnect Azure AD connections. 

    .DESCRIPTION
        An AzureAD connection must be established before obtaining Azure AD/O365 group memberships and licenses. 
        Actor must authenticate with an administator AAD account. 
#>
    Param(
        [Parameter(Mandatory=$true,Position=0)][string]$a
    )

     switch ($a)
     {
      connect 
      {
	    Write-Host "Connecting To Azure..." -NoNewLine;	  	
        try {	
	            $modExists = (Get-Module).Name | ForEach-Object { if ($_ -contains "AzureAD") { return $true; } }
	            if ($modExists -ne $true) {
    	            Install-Module AzureAD -Force
	            }

	            Connect-AzureAD -ErrorAction SilentlyContinue | Out-Null		## Attempt connection to AzureAD
	            Write-Host "DONE`r" -Foreground Green 

        }
	    catch { 
            Write-Host "FAILED" -ForegroundColor Red;
		    $msg = "FAILED to connect to AzureAD`r`n`r`n$($_.Exception.Message)";
		    newError $msg;
		    throw $msg;
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
<#
    .SYNOPSIS
        Get ActiveDirectory Group memberships

    .DESCRIPTION
        Active Directory Group memberships are obtained using the ActiveDirectory Powershell module.

    .REMARKS
        https://learn.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps
#>
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
		catch {
            Write-Host "FAILED" -ForegroundColor Red;
            $msg = "Get-ADUser failed to get AD users`r`n`r`n$($_.Exception.Message)";
            newError $msg;
            return;
        }

        $string += newHR
        $string += "Active Directory`r`n";
        $string += newHR
        $string += $adGroup | ForEach-Object { $_ + ";`r`n" }
        $string += newHR
        $ui.output += $string;
}

function getOB {
<#
    .SYNOPSIS
        Get OnBase User Group Memberships

    .DESCRIPTION
        OnBase User Group memberships are obtained by querying the OnBase database.
        Database connection must be established with read privlidges. 
#>
	Param (
	[Parameter (Mandatory=$true,Position=0)][userInput]$ui 
	)

    $username = $ui.username;
	
	$server = "database_server_here";
	$database = "databse_name_here";
	$uid = "databse_username_here";
	$pw = "database_password_here";

	try {
        Write-Host "Searching OnBase..." -NoNewline

	    ## Query User Groups 
	    $q = Get-Content "$PSScriptRoot\ob-groups.sql" -ErrorAction SilentlyContinue;
	
        if ([System.String]::IsNullOrEmpty($q)) {
            throw "OnBase Usergroup SQL Query not found";
        }
	
	    [string] $connectionString = "Server=$server;Database=$database;Integrated Security = False; User ID = $uid; Password = $pw;"
	
		$conn = New-Object System.Data.SqlClient.SqlConnection($connectionString);
		$conn.Open();
		$command = $conn.CreateCommand();
        $q = $q.Replace("mpolo", $username);
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
	
	catch {
        Write-Host "FAILED" -ForegroundColor Red;
        $msg = "OnBase Databse Connection Failed`r`n`r`n$($_.Exception.Message)";
        newError $msg;
    }
}

function getAAD {
<#
    .SYNOPSIS
        Get Azure AD Group Memberships and Licenses

    .DESCRIPTION
        Azure AD groups and licenses are obtained with the Azure-AD Powershell module.

    .REMARKS
        https://learn.microsoft.com/en-us/powershell/azure/active-directory/overview?view=azureadps-2.0

#>
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
	    try { 
            Get-AzureADTenantDetail | Out-Null;
        }
        
	    catch {
            try { azureAD connect; }
            catch { return; }
        }

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

    catch {
     Write-Host "FAILED" -ForegroundColor Red;
     $msg = "Failed to retrieve Azure-AD groups.`r`n`r`n$($_.Exception.Message)";
     newError $msg;
     return;
    }

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
