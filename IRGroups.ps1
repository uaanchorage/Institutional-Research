<#
    Author:	John Zetterman
	Script:	IRGroups.ps1
	Date:	06/01/2018
	Contact: systems@uaa.alaska.edu
	
	Version 1.0
	
	Purpose:
        This script populates security groups in the UA domain from the table entries in the IRPROD database located on ANC-IR-SQL Always On cluster. 
        Runs under user context IRadmin 
		
#>
#Version History-
#1.0 Heavily modified original script written by Josh Ford.
#####

[string] $body = @()

Add-Type @"
	using System;
	public class secgroup {
		public string dn;
		public string query;
		public string friendlyName;
        public int numofmembers;
        public int failedCount;
        public string inputData;
	}
"@

# Get Start Time
$startDTM = (Get-Date)

$Lists = @(
    [secgroup] @{   
                    dn = "CN=UA_BANNER_FINANCE_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_BANNER_FINANCE_USER] ";
                    numofmembers = 0;
                    friendlyName = "UA_BANNER_FINANCE_USER";
                    failedCount = 0;
                    inputData = $null;
                }
    
    [secgroup] @{   
                    dn = "CN=UA_BANNER_HR_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_BANNER_HR_USER] ";
                    numofmembers = 0;
                    friendlyName = "UA_BANNER_HR_USER";
                    failedCount = 0;
                    inputData = $null;
                }
				
    [secgroup] @{   
                    dn = "CN=UA_EES_REGULAR,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_EES_REGULAR] ";
                    numofmembers = 0;
                    friendlyName = "UA_EES_REGULAR";
                    failedCount = 0;
                    inputData = $null;
                }
    
    [secgroup] @{   
                    dn = "CN=UA_SW_BANNER_FINANCE_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_SW_BANNER_FINANCE_USER] ";
                    numofmembers = 0;
                    friendlyName = "UA_SW_BANNER_FINANCE_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UA_SW_BANNER_HR_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_SW_BANNER_HR_USER] ";
                    numofmembers = 0;
                    friendlyName = "UA_SW_BANNER_HR_USER";
                    failedCount = 0;
                    inputData = $null;
                }
				
    [secgroup] @{   
                    dn = "CN=UA_SW_EES_REGULAR,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_SW_EES_REGULAR] ";
                    numofmembers = 0;
                    friendlyName = "UA_SW_EES_REGULAR";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAA_BANNER_FINANCE_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAA_BANNER_FINANCE_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAA_BANNER_FINANCE_USER";
                    failedCount = 0;
                    inputData = $null;
                }
                
    [secgroup] @{   
                    dn = "CN=UAA_BANNER_HR_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAA_BANNER_HR_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAA_BANNER_HR_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAA_EES_REGULAR,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAA_EES_REGULAR] ";
                    numofmembers = 0;
                    friendlyName = "UAA_EES_REGULAR";
                    failedCount = 0;
                    inputData = $null;
                }
    
    [secgroup] @{   
                    dn = "CN=UAF_BANNER_FINANCE_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAF_BANNER_FINANCE_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAF_BANNER_FINANCE_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAF_BANNER_HR_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAF_BANNER_HR_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAF_BANNER_HR_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAF_EES_REGULAR,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAF_EES_REGULAR] ";
                    numofmembers = 0;
                    friendlyName = "UAF_EES_REGULAR";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAS_BANNER_FINANCE_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAS_BANNER_FINANCE_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAS_BANNER_FINANCE_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAS_BANNER_HR_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAS_BANNER_HR_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAS_BANNER_HR_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAS_EES_REGULAR,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAS_EES_REGULAR] ";
                    numofmembers = 0;
                    friendlyName = "UAS_EES_REGULAR";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UA_BANNER_FERPA_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_BANNER_FERPA_USER] ";
                    numofmembers = 0;
                    friendlyName = "UA_BANNER_FERPA_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UA_SW_BANNER_FERPA_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UA_SW_BANNER_FERPA_USER] ";
                    numofmembers = 0;
                    friendlyName = "UA_SW_BANNER_FERPA_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAA_BANNER_FERPA_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAA_BANNER_FERPA_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAA_BANNER_FERPA_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAF_BANNER_FERPA_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAF_BANNER_FERPA_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAF_BANNER_FERPA_USER";
                    failedCount = 0;
                    inputData = $null;
                }

    [secgroup] @{   
                    dn = "CN=UAS_BANNER_FERPA_USER,OU=InstitutionalResearch,OU=InstitutionalEffectiveness,OU=Anc,OU=UAA,DC=ua,DC=ad,DC=alaska,DC=edu";
                    query = "SELECT ua_id FROM [IRPROD].[active_directory].[UAS_BANNER_FERPA_USER] ";
                    numofmembers = 0;
                    friendlyName = "UAS_BANNER_FERPA_USER";
                    failedCount = 0;
                    inputData = $null;
                }
				
)

$domaincontroller = "ua.ad.alaska.edu"
$Count1 = 0

function Get-Users ($sqlQuery) {

    $SQL_CONN = "Server=anc-ir-sql.apps.ad.alaska.edu;Database=IRPROD;Integrated Security=True;"
    # $sqlQuery = "SELECT * FROM [IRPROD].[active_directory].[UAA_EES_REGULAR] WHERE UA_ID = '31028884'"

	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = $SQL_CONN
    $SqlConnection.Open()

	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.Connection = $SqlConnection
	$SqlCmd.CommandText = $sqlQuery
    $SqlReader = $SqlCmd.ExecuteReader()

	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd

	$DataSet = New-Object System.Data.DataTable
	$DataSet.Load($SqlReader)
	$SqlConnection.Close()

    return $DataSet
    
}

# Emails report of runtime results
function Mail ($emailBody) {

    $date = Get-Date
    $emailFrom = "IRGroups@uaa.alaska.edu"
    #$emailTo = "wjohns73@uaa.alaska.edu"
    $emailTo = "systems@uaa.alaska.edu", "clgullett@alaska.edu"
    $subject = "Group Report $date"
    $smtpServer = "aspam-out.uaa.alaska.edu"
    Send-MailMessage -SmtpServer $smtpServer -From $emailFrom -To $emailTo -Subject $subject -Body $emailBody

}

ForEach ($list in $Lists) {

    $Count1++
    $Percent = "$([math]::Round($Count1/$Lists.Count*100))%"
    Write-Progress -Activity "Processing IRPROD Table: $($list.friendlyName)" -Status "Progress $Count1 of $($Lists.Count): $Percent" -PercentComplete (($Count1 / $Lists.Count) * 100)
    
    $groupDN = $list.dn
    $groupFN = $list.friendlyName
    $data = Get-Users $list.query

    $Count2 = 0
    
    ForEach ($row in $data) {

        $Count2++
        $Percent = "$([math]::Round($Count2/$data.Count*100))%"
        Write-Progress -Activity "Processing UAID: $($row.ua_id)" -Status "Progress $Count2 of $($data.Count): $Percent" -PercentComplete (($Count2 / $data.Count) * 100)

        Try {

            $uaidentifier = $row.ua_id
            # Write-Host $uaidentifier -ForegroundColor Cyan
            $fedUser = Get-ADUser -filter 'uaidentifier -eq $uaidentifier' -SearchBase 'OU=userAccounts,DC=UA,DC=AD,DC=ALASKA,DC=EDU' -Server $domaincontroller -ErrorAction Stop
            $samAccountName = $fedUser.samAccountName
            # Write-Host (Get-ADUser -Filter 'uaidentifier -eq $uaidentifier' -SearchBase 'OU=userAccounts,DC=UA,DC=AD,DC=ALASKA,DC=EDU') -ForegroundColor Green

        }

        Catch {

            Write-Host "User lookup failed, check source data for non-UA ID data." -ForegroundColor Red
            If ($uaidentifier -eq "") {
                Write-Host "Input from Database: NULL" -ForegroundColor Red
            } 
            Else {
                Write-Host "Input from Database: $uaidentifier" -ForegroundColor Red
            }

        }

        Try {

            Add-ADGroupMember -Identity $groupDN -Members $samAccountName -Server $domaincontroller
            # Write-Host "User successfully added" -ForegroundColor Green
            $list.numofmembers++

        }
        
        Catch {
            
            Write-Host "Adding user failed" -ForegroundColor Red
            Write-Host $uaidentifier -ForegroundColor Red
            Write-Host (Get-ADUser -Filter 'uaidentifier -eq $uaidentifier' -SearchBase 'OU=userAccounts,DC=UA,DC=AD,DC=ALASKA,DC=EDU' -Server $domaincontroller) -ForegroundColor Red
            $list.failedCount++
            If ($row.ua_id -eq $null) {
                $list.inputData += "NULL"
            }
            Else {
                $list.inputData += "$($row.ua_id)`r"
            }
            
        }
    }

    # Measure how long it took to Add
    $endDTM = (Get-Date)
    $runtime= $endDTM - $startDTM
    $scripttime= [math]::round($runtime.totalminutes , 2)

    $count = $list.numofmembers
    $body += "This Script ran in $scripttime minutes $count user were added to the group $groupFN. Failed to add $($list.failedCount) `n
                The following input data did not match any user in AD: `n$($list.inputData)`r" 
    
}
    
Mail $body