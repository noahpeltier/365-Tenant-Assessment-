Param(
    [parameter(Mandatory = $true)]
    $clientId,
    [parameter(Mandatory = $true)]
    $tenantId,
    [parameter(Mandatory = $true)]
    $certificateThumbprint,
    [parameter(Mandatory = $false)]
    [switch]$IncludeGroupMembership = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeMailboxPermissions = $false,
    [string]$ReportOutputPath = "C:\temp"
)

#region Helper functions
function Invoke-MgGraphRequestAllPages {
    param([string]$Uri)
    $allResults = @()
    do {
        $result = Invoke-MgGraphRequest -Method GET -URI $Uri -OutputType PSObject
        if ($result.value) {
            $allResults += $result.value
        }
        $Uri = $result.'@odata.nextLink'
    } while ($Uri)
    return $allResults
}

function Get-MGReport {
    param(
        $Report,
        [ValidateSet("D7", "D30", "D90", "D180")]
        $Period
    )
    $ReportRaw = Invoke-RestMethod (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/reports/$Report(period='$Period')" -OutputType HttpResponseMessage).RequestMessage.RequestUri.OriginalString
    $ReportRaw | ConvertFrom-Csv
}

function Update-Progress {
    param(
        $ProgressStatus = "Next Steps...",
        [switch]$AdvanceProgressTracker
    )
    if ($AdvanceProgressTracker) {
        $Script:ProgressTracker++
    }
    Write-Progress -Activity "Tenant Assessment in Progress" -Status "Processing Task $Script:ProgressTracker of $($Script:TotalProgressTasks): $ProgressStatus" -PercentComplete (($Script:ProgressTracker / $Script:TotalProgressTasks) * 100)
}

function Connect-ToGraph {
    Connect-Entra -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certificateThumbprint -NoWelcome
}

function Invoke-PrepareEnvironment {
    if (!(Test-Path -Path $ReportOutputPath)) {
        Try {
            New-Item -Path $ReportOutputPath -ItemType Directory -ErrorAction Stop
        }
        catch {
            Write-Host "Could not create folder at $ReportOutputPath - check permissions." -ForegroundColor Red
            return $false
        }
    }
    $Script:Filename = "TenantAssessment-$((Get-Date).ToString('yyyyMMddHHmmss')).xlsx"
    $Script:TemplatePresent = Test-Path "TenantAssessment-Template.xlsx"
    return $true
}

function Connect-Services {
    
}
#endregion



#region Data gathering functions
function Get-TenantUsers {
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,userprincipalname,mail,displayname,givenname,surname,licenseAssignmentStates,proxyaddresses,usagelocation,usertype,accountenabled,onPremisesSyncEnabled'
    $Users = Invoke-MgGraphRequestAllPages -Uri $uri

    return $users
}

function Get-TenantGroups {
    $uri = "https://graph.microsoft.com/v1.0/groups"
    $Groups = Invoke-MgGraphRequestAllPages -Uri $uri

    return $Groups
}

function Get-AllGroupMemberships {
    $GroupMembersObject = [System.Collections.Generic.List[Object]]::new()
    $i = 1
    foreach ($group in $Script:Groups) {
        Update-Progress "Processing group memberships on group $i of $($Script:Groups.count)..."
        $i++
        $apiuri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members"
        $Members = Invoke-MgGraphRequestAllPages -Uri $apiuri

        foreach ($member in $members) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.displayname
                MemberUserPrincipalName = $member.userprincipalname
                MemberType              = "Member"
                MemberObjectType        = $member.'@odata.type'.replace('#microsoft.graph.', '')

            }

            $GroupMembersObject.Add($memberEntry)

        }

        $apiuri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners"
        $Owners = Invoke-MgGraphRequestAllPages -Uri $apiuri 

        foreach ($member in $Owners) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.displayname
                MemberUserPrincipalName = $member.userprincipalname
                MemberType              = "Owner"
                MemberObjectType        = $member.'@odata.type'.replace('#microsoft.graph.', '')

            }

            $GroupMembersObject.Add($memberEntry)
        }
    }
    $GroupMembersObject
}

function Get-AllTeams {
    $uri = 'https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c eq ''Unified'') and resourceProvisioningOptions/any(x:x eq ''Team'')'
    $TeamGroups = Invoke-MgGraphRequestAllPages -Uri $uri
    $i = 1
    foreach ($teamgroup in $TeamGroups) {
        Update-Progress "Processing Team $i of $($Teamgroups.count)..."
        $i++
    
        $apiuri = "https://graph.microsoft.com/beta/teams/$($teamgroup.id)/allchannels"
        $Teamchannels = Invoke-MgGraphRequestAllPages -Uri $apiuri
        $standardchannels = ($teamchannels| Where-Object  { $_.membershipType -eq "standard" })
        $privatechannels = ($teamchannels| Where-Object  { $_.membershipType -eq "private" })
        $outgoingsharedchannels = ($teamchannels| Where-Object  { ($_.membershipType -eq "shared") -and (($_."@odata.id") -like "*$($teamgroup.id)*") })
        $incomingsharedchannels = ($teamchannels| Where-Object  { ($_.membershipType -eq "shared") -and ($_."@odata.id" -notlike "*$($teamgroup.id)*") })
        $teamgroup | Add-Member -MemberType NoteProperty -Name "StandardChannels" -Value $standardchannels.id.count -Force
        $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannels" -Value $privatechannels.id.count -Force
        $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannels" -Value $outgoingsharedchannels.id.count -Force
        $teamgroup | Add-Member -MemberType NoteProperty -Name "IncomingSharedChannels" -Value $incomingsharedchannels.id.count -Force
        $privatechannelSize = 0
        
        foreach ($Privatechannel in $privatechannels) {
            $PrivateChannelObject = $null
            $apiuri = "https://graph.microsoft.com/beta/teams/$($teamgroup.id)/channels/$($Privatechannel.id)/FilesFolder"
            Try {
                $PrivateChannelObject = Invoke-MgGraphRequest -Uri $apiUri -Method Get -OutputType PSObject
                $Privatechannelsize += $PrivateChannelObject.size
    
            }
            Catch {
                $Privatechannelsize += 0
            }
        }
    
        $sharedchannelSize = 0
        
        foreach ($sharedchannel in $outgoingsharedchannels) {
            $sharedChannelObject = $null
            $apiuri = "https://graph.microsoft.com/beta/teams/$($teamgroup.id)/channels/$($Sharedchannel.id)/FilesFolder"
            Try {
                $SharedChannelObject = Invoke-MgGraphRequest -Uri $apiUri -Method Get -OutputType PSObject
                $Sharedchannelsize += $SharedChannelObject.size
    
            }
            Catch {
                $Sharedchannelsize += 0
            }
        }
    
        $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannelsSize" -Value $privatechannelSize -Force
        $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannelsSize" -Value $sharedchannelSize -Force
        
    
        $TeamDetails = $null
        $apiuri = "https://graph.microsoft.com/v1.0/groups/$($teamgroup.id)/drive/"
        $TeamDetails = Invoke-MgGraphRequest -Uri $apiUri -Method Get -OutputType PSObject
    
        $teamgroup | Add-Member -MemberType NoteProperty -Name "DataSize" -Value $TeamDetails.quota.used -Force
        $teamgroup | Add-Member -MemberType NoteProperty -Name "URL" -Value $TeamDetails.webUrl.replace("/Shared%20Documents", "") -Force
    
    }
    return $TeamGroups
}

function Get-TenantLicenses {
    $uri = "https://graph.microsoft.com/v1.0/subscribedskus"
    $SKUs = Invoke-MgGraphRequestAllPages -Uri $uri

    return $SKUs
}

function Get-TenantOrgDetails {
    $uri = "https://graph.microsoft.com/v1.0/organization"
    $OrgDetails = Invoke-MgGraphRequestAllPages -Uri $uri

    return $OrgDetails
}

function Get-AADApps {
    $uri = "https://graph.microsoft.com/beta/servicePrincipals"
    $AADApps = Invoke-MgGraphRequestAllPages -Uri $uri

    return $AADApps
}

function Invoke-ProcessUserLicenseInfo {
    foreach ($user in $Script:Users) {
        $user | Add-Member -MemberType NoteProperty -Name "License SKUs" -Value ($user.licenseAssignmentStates.skuid -join ";") -Force
        $user | Add-Member -MemberType NoteProperty -Name "Group License Assignments" -Value ($user.licenseAssignmentStates.assignedByGroup -join ";") -Force
        $user | Add-Member -MemberType NoteProperty -Name "Disabled Plan IDs" -Value ($user.licenseAssignmentStates.disabledplans -join ";") -Force
    }
    foreach ($user in $Script:Users) {
        foreach ($group in $Script:Groups) {
            $user.'Group License Assignments' = $user.'Group License Assignments'.Replace($group.id, $group.displayname)
        }
        foreach ($SKU in $Script:SKUs) {
            $user.'License SKUs' = $user.'License SKUs'.Replace($SKU.skuid, $SKU.skuPartNumber)
        }
        foreach ($SKUplan in $Script:SKUs.servicePlans) {
            $user.'Disabled Plan IDs' = $user.'Disabled Plan IDs'.Replace($SKUplan.servicePlanId, $SKUplan.servicePlanName)
        }
    }
}

function Invoke-ProcessConditionalAccessPolicies {
    $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
    $policies = Invoke-MgGraphRequestAllPages -Uri $uri
    $policiesJSON = $policies | ConvertTo-Json -Depth 5
    if ($policiesJSON) {
        foreach ($user in $Script:Users) {
            $policiesJSON = $policiesJSON.Replace($user.id, "$($user.displayname) - $($user.userPrincipalName)")
        }
        foreach ($group in $Script:Groups) {
            $policiesJSON = $policiesJSON.Replace($group.id, "$($group.displayname) - $($group.id)")
        }
        foreach ($dirRole in (Invoke-MgGraphRequestAllPages -Uri "https://graph.microsoft.com/beta/directoryRoleTemplates")) {
            $policiesJSON = $policiesJSON.Replace($dirRole.Id, $dirRole.displayname)
        }
        foreach ($app in $Script:AADApps) {
            $policiesJSON = $policiesJSON.Replace($app.appid, $app.displayname)
            $policiesJSON = $policiesJSON.Replace($app.id, $app.displayname)
        }
        foreach ($loc in (Invoke-MgGraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations")) {
            $policiesJSON = $policiesJSON.Replace($loc.id, $loc.displayname)
        }
        $policies = $policiesJSON | ConvertFrom-Json
        $CAOutput = @()
        $CAHeadings = @(
            "displayName",
            "createdDateTime",
            "modifiedDateTime",
            "state",
            "Conditions.users.includeusers",
            "Conditions.users.excludeusers",
            "Conditions.users.includegroups",
            "Conditions.users.excludegroups",
            "Conditions.users.includeroles",
            "Conditions.users.excluderoles",
            "Conditions.clientApplications.includeServicePrincipals",
            "Conditions.clientApplications.excludeServicePrincipals",
            "Conditions.applications.includeApplications",
            "Conditions.applications.excludeApplications",
            "Conditions.applications.includeUserActions",
            "Conditions.applications.includeAuthenticationContextClassReferences",
            "Conditions.userRiskLevels",
            "Conditions.signInRiskLevels",
            "Conditions.platforms.includePlatforms",
            "Conditions.platforms.excludePlatforms",
            "Conditions.locations.includLocations",
            "Conditions.locations.excludeLocations",
            "Conditions.clientAppTypes",
            "Conditions.devices.deviceFilter.mode",
            "Conditions.devices.deviceFilter.rule",
            "GrantControls.operator",
            "grantcontrols.builtInControls",
            "grantcontrols.customAuthenticationFactors",
            "grantcontrols.termsOfUse",
            "SessionControls.disableResilienceDefaults",
            "SessionControls.applicationEnforcedRestrictions",
            "SessionControls.persistentBrowser",
            "SessionControls.cloudAppSecurity",
            "SessionControls.signInFrequency"
        )
        foreach ($heading in $CAHeadings) {
            $row = New-Object PSObject -Property @{ PolicyName = $heading }
            foreach ($policy in $policies) {
                $nestCount = $heading.Split('.').Count
                if ($nestCount -eq 1) {
                    $row | Add-Member -MemberType NoteProperty -Name $policy.displayname -Value $policy.$heading -Force
                }
                elseif ($nestCount -eq 2) {
                    $split = $heading.Split('.')
                    $value = $policy.$($split[0]).$($split[1])
                    $row | Add-Member -MemberType NoteProperty -Name $policy.displayname -Value ($value -join ';') -Force
                }
                elseif ($nestCount -eq 3) {
                    $split = $heading.Split('.')
                    $value = $policy.$($split[0]).$($split[1]).$($split[2])
                    $row | Add-Member -MemberType NoteProperty -Name $policy.displayname -Value ($value -join ';') -Force
                }
                elseif ($nestCount -eq 4) {
                    $split = $heading.Split('.')
                    $value = $policy.$($split[0]).$($split[1]).$($split[2]).$($split[3])
                    $row | Add-Member -MemberType NoteProperty -Name $policy.displayname -Value ($value -join ';') -Force
                }
            }
            $CAOutput += $row
        }
        return $CAOutput
    }
}


function Get-SPSite {
    param(
        $SiteID
    )
    try {
        invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$SiteID"
    }
    catch {

    }
}

function Invoke-ProcessSharePointTeamID {
    $i = 1
    foreach ($site in $Script:SharepointUsageReport) {
        Update-Progress -ProgressStatus "Resolving SharePoint site URLs $i of $($Script:SharepointUsageReport.count)..."
        $i++
        $team = $Script:TeamGroups| Where-Object { $_.URL -like "*$($site.'site url')*" }
        $site.TeamID = if ($team) { $team.id } else { "" }
    }
}

function Invoke-ProcessGroupMembership {
    $Script:GroupMembersObject = New-Object System.Collections.Generic.List[Object]
    $i = 1
    foreach ($group in $Script:Groups) {
        Update-Progress "Enumerating Group Membership: Processing Group $i of $($Script:Groups.count)"
        $i++
        $uri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members"
        $members = Invoke-MgGraphRequestAllPages -Uri $uri
        foreach ($member in $members) {
            $entry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.displayname
                MemberUserPrincipalName = $member.userprincipalname
                MemberType              = "Member"
                MemberObjectType        = $member.'@odata.type'.Replace('#microsoft.graph.', '')
            }
            $Script:GroupMembersObject.Add($entry)
        }
        $uri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners"
        $owners = Invoke-MgGraphRequestAllPages -Uri $uri
        foreach ($owner in $owners) {
            $entry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $owner.id
                MemberName              = $owner.displayname
                MemberUserPrincipalName = $owner.userprincipalname
                MemberType              = "Owner"
                MemberObjectType        = $owner.'@odata.type'.Replace('#microsoft.graph.', '')
            }
            $Script:GroupMembersObject.Add($entry)
        }
    }
}

function Get-DirectoryRoleTemplates {
    $apiURI = "https://graph.microsoft.com/beta/directoryRoleTemplates"
    Invoke-MgGraphRequestAllPages -Uri $apiuri 
}

function Get-RoomMailboxes{
    $i = 1
    $Roommailboxes = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited
    foreach ($room in $Roommailboxes) {
        Update-Progress -ProgressStatus "Getting room mailbox statistics $i of $($Roommailboxes.count)..."
        $i++
        $RoomStats = get-EXOmailboxstatistics $room.primarysmtpaddress
        $room | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $RoomStats.TotalItemSize -Force
        $room | Add-Member -MemberType NoteProperty -Name ItemCount -Value $RoomStats.ItemCount -Force

        ##Clean email addresses value
        $room.EmailAddresses = $room.EmailAddresses -join ';'

        $Roommailboxes
    }
}

function Get-EquipmentMailboxes {
    $EquipmentMailboxes = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited
    $i = 1
    foreach ($equipment in $EquipmentMailboxes) {
        Update-Progress -ProgressStatus "Getting Equipment mailbox statistics $i of $($EquipmentMailboxes.count)..."
        $i++
    
        $EquipmentStats = get-EXOmailboxstatistics $equipment.primarysmtpaddress
        $equipment | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $EquipmentStats.TotalItemSize -Force
        $equipment | Add-Member -MemberType NoteProperty -Name ItemCount -Value $EquipmentStats.ItemCount -Force
    
        ##Clean email addresses value
        $equipment.EmailAddresses = $equipment.EmailAddresses -join ';'
    }

    $EquipmentMailboxes
}

function Get-SharedMailboxes {
    $SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    $i =1
    foreach ($SharedMailbox in $SharedMailboxes) {
        Update-Progress -ProgressStatus "Getting shared mailbox statistics $i of $($SharedMailboxes.count)..."
        $i++
    
        $SharedStats = $null
        $SharedStats = get-EXOmailboxstatistics $SharedMailbox.primarysmtpaddress
        $SharedMailbox | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $SharedStats.TotalItemSize -Force
        $SharedMailbox | Add-Member -MemberType NoteProperty -Name ItemCount -Value $SharedStats.ItemCount -Force
        
        ##Clean email addresses value
        $SharedMailbox.EmailAddresses = $SharedMailbox.EmailAddresses -join ';'
    }

    $SharedMailboxes
}

function Get-UserMailboxStats {
    $MailboxStats = [system.Collections.Generic.List[Object]]::new()
    $i = 1
    foreach ($user in ($Script:Users| Where-Object { ($null -ne $_.mail ) -and ($_.userType -eq "Member") })) {
        Update-Progress -ProgressStatus "Getting user mailbox statistics $i of $($Script:Users.count)..."
        $i++
        $stats = $Script:MailboxStatsReport| Where-Object { $_.'User Principal Name' -eq $user.userprincipalname }
        $stats | Add-Member -MemberType NoteProperty -Name ObjectID -Value $user.id -Force
        $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $user.mail -Force
        $MailboxStats.Add($stats)
    }
    $MailboxStats
}

function Get-ArchiveMailboxStats {
    $ArchiveStats = [system.Collections.Generic.List[Object]]::new()
    foreach ($archive in $Script:ArchiveMailboxes) {
        Update-Progress -ProgressStatus "Getting archive mailbox statistics $i of $($ArchiveMailboxes.count)..."
        $i++

        $stats = get-EXOmailboxstatistics $archive.PrimarySmtpAddress -Archive #-erroraction SilentlyContinue
        $stats | Add-Member -MemberType NoteProperty -Name ObjectID -Value $archive.ExternalDirectoryObjectId -Force
        $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $archive.primarysmtpaddress -Force
        $ArchiveStats.Add($stats)
        
    }
    $ArchiveStats
}

function Get-ConditionAccessPolicies {
    $apiURI = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
    Invoke-MgGraphRequestAllPages -Uri $apiuri 
}

function Get-NamedLocations {
    $apiURI = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
    Invoke-MgGraphRequestAllPages -Uri $apiuri
}

function Get-Mailcontacts {
    $MailContacts = Get-MailContact -ResultSize unlimited | Select-Object displayname, alias, externalemailaddress, emailaddresses, HiddenFromAddressListsEnabled
    foreach ($mailcontact in $MailContacts) {
        $mailcontact.emailaddresses = $mailcontact.emailaddresses -join ';'
    }
    $MailContacts
}

function Get-AllMailboxPermissions {
    $PermissionOutput = [System.Collections.Generic.List[Object]]::new()
    ##Get all mailboxes
    $MailboxList = Get-EXOMailbox -ResultSize unlimited
    $PermissionProgress = 1
    foreach ($mailbox in $MailboxList) {
        Update-Progress -ProgressStatus "Getting mailbox permissions for ($PermissionProgress of $($MailboxList.count)) This takes a while..."
        $Permissions = Get-EXOMailboxPermission -UserPrincipalName $mailbox.UserPrincipalName | Where-Object { $_.User -ne "NT AUTHORITY\SELF" }
        foreach ($permission in $Permissions) {

            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.user

            }
            
            $PermissionOutput.Add($PermissionObject)
        }

        $RecipientPermissions = Get-EXORecipientPermission $mailbox.UserPrincipalName | Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" }
        foreach ($permission in $RecipientPermissions) {

            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.trustee

            }
            
            $PermissionOutput.Add($PermissionObject)
        }

        $PermissionProgress++
    }
}

function Get-InboundConnectors {
    $InboundConnectors = Get-InboundConnector | Select-Object enabled, name, connectortype, connectorsource, SenderIPAddresses, SenderDomains, RequireTLS, RestrictDomainsToIPAddresses, RestrictDomainsToCertificate, CloudServicesMailEnabled, TreatMessagesAsInternal, TlsSenderCertificateName, EFTestMode, Comment 
    foreach ($inboundconnector in $InboundConnectors) {
        $inboundconnector.senderipaddresses = $inboundconnector.senderipaddresses -join ';'
        $inboundconnector.senderdomains = $inboundconnector.senderdomains -join ';'
    }
    $InboundConnectors
}

function Get-OutboundConnectors {
    $OutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors:$true | Select-Object enabled, name, connectortype, connectorsource, TLSSettings, RecipientDomains, UseMXRecord, SmartHosts, Comment
    foreach ($OutboundConnector in $OutboundConnectors) {
        $OutboundConnector.RecipientDomains = $OutboundConnector.RecipientDomains -join ';'
        $OutboundConnector.SmartHosts = $OutboundConnector.SmartHosts -join ';'
    }
    $OutboundConnectors
}

function Get-MXRecords {
    foreach ($domain in $Script:OrgDetails.verifieddomains) {
        Try {
            Resolve-DnsName -Name $domain.name -Type mx -ErrorAction SilentlyContinue
        }
        catch {
            write-host "Error obtaining MX Record for $($domain.name)"
        }
    }
}
#endregion

# Setup Process tracking variables
$Script:ProgressTracker = 1
$Script:TotalProgressTasks = 25
if ($IncludeGroupMembership) {
    $Script:TotalProgressTasks++
}

if ($IncludeMailboxPermissions) {
    $Script:TotalProgressTasks++
}

# 1. Setup: Define Parameters, Functions, and Import Modules
Update-Progress -ProgressStatus "Initializing Tenant Assessment Script..."
Try {
    import-module ExchangeOnlineManagement #-RequiredVersion 3.4.0
    import-module importexcel
    Import-Module Microsoft.Entra
}
catch {
    Clear-Host
    write-host "Could not import modules, make sure you have the following modules availablenMSAL.PSnExchangeOnlineManagementnnThese modules can be installed by running the following commands:nInstall-Module MSAL.PSnnInstall-Module ExchangeOnlineManagement" -ForegroundColor red
    exit
}
$Filename = "TenantAssessment-$((get-date).tostring().replace('/','').replace(':','')).xlsx"
##File Location
$FilePath = $ReportOutputPath
Try {
    if (!(test-path -Path $FilePath)) {
        New-Item -Path $FilePath -ItemType Directory
    }
}
catch {
    write-host "Could not create folder at c:\temp - check you have appropriate permissions" -ForegroundColor red
    exit
}


# 2. Authenticate and Prepare Environment (Access Token, Output Directory, Template Check)
Update-Progress -ProgressStatus "Connecting to Microsoft Graph..."
Connect-Entra -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome


# 3. Graph API Collections: 
#    - Retrieve Users, Groups, Teams (with Channel & Drive details)
Update-Progress -ProgressStatus "Retrieving Users..." -AdvanceProgressTracker
$Script:Users = Get-TenantUsers

Update-Progress -ProgressStatus "Retrieving Groups..." -AdvanceProgressTracker
$script:Groups = Get-TenantGroups

Update-Progress -ProgressStatus "Retrieving Teams..." -AdvanceProgressTracker
$Script:TeamGroups = Get-AllTeams

#    - Retrieve License SKUs, Org Details, Azure AD Apps, and Conditional Access Policies
Update-Progress -ProgressStatus "Retrieving License SKUs..." -AdvanceProgressTracker
$Script:SKUs = Get-TenantLicenses

Update-Progress -ProgressStatus "Retrieving Organization Details..." -AdvanceProgressTracker
$Script:OrgDetails = Get-TenantOrgDetails

Update-Progress -ProgressStatus "Retrieving Azure AD Apps..." -AdvanceProgressTracker
$Script:aadApps = Get-AADApps

Update-Progress -ProgressStatus "Retrieving Conditional Access Policies..." -AdvanceProgressTracker
$Script:ConditionalAccessPolicies = Get-ConditionAccessPolicies

# 4. Graph Reports Collection: Get OneDrive, SharePoint, Mailbox, and M365 Apps Usage Reports
Update-Progress -ProgressStatus "Retrieving OneDrive Usage Report..." -AdvanceProgressTracker
$Script:OneDriveUsageReport = Get-MGReport -Report "getOneDriveUsageAccountDetail" -Period D30
$Script:OneDriveUsageReport | Add-Member -MemberType NoteProperty -Name "TeamID" -Value "" -force

Update-Progress -ProgressStatus "Retrieving SharePoint Usage Report..." -AdvanceProgressTracker
$Script:SharePointUsageReport = Get-MGReport -Report "getSharePointSiteUsageDetail" -Period D30
foreach ($site in $Script:SharePointUsageReport) {
    $SPSite = Get-SPSite $site.'Site Id'
    $site.'Site URL' = $SPSite.webURL
}
$Script:SharePointUsageReport | Add-Member -MemberType NoteProperty -Name "TeamID" -Value "" -force

Update-Progress -ProgressStatus "Retrieving Mailbox Usage Report..." -AdvanceProgressTracker
$Script:MailboxStatsReport = Get-MGReport -Report "getMailboxUsageDetail" -Period D30

Update-Progress -ProgressStatus "Retrieving M365 Apps Usage Report..." -AdvanceProgressTracker
$Script:M365AppsUsageReport = Get-MGReport -Report "getOffice365ServicesUserCounts" -Period D30

# 5. Optional Graph Collection: Enumerate Group Membership (Members & Owners)
if ($IncludeGroupMembership) {
    Update-Progress -ProgressStatus "Retrieving Group Memberships..." -AdvanceProgressTracker
    $Script:GroupMemberships = Get-AllGroupMemberships
}

# 6. Exchange Online Collections:
#    - Connect to Exchange Online
$orgDomain = ($Script:OrgDetails.VerifiedDomains| Where-Object { $_.IsDefault -eq $true }).Name
Connect-ExchangeOnline -CertificateThumbPrint $CertificateThumbprint -AppId $ClientId -Organization $orgDomain -ShowBanner:$false

#    - Retrieve Room, Equipment, and Shared Mailboxes and their Statistics (User & Archive)
Update-Progress -ProgressStatus "Retrieving Room Mailboxes..." -AdvanceProgressTracker
$Script:RoomMailboxes = Get-RoomMailboxes

Update-Progress -ProgressStatus "Retrieving Equipment Mailboxes..." -AdvanceProgressTracker
$Script:EquipmentMailboxes = Get-EquipmentMailboxes

Update-Progress -ProgressStatus "Retrieving Shared Mailboxes..." -AdvanceProgressTracker
$Script:SharedMailboxes = Get-SharedMailboxes
$Script:SharedMailboxes | Add-Member -MemberType NoteProperty -Name MailboxSize -Value "" -Force

Update-Progress -ProgressStatus "Retrieving archive Mailboxes..." -AdvanceProgressTracker
$Script:ArchiveMailboxes = get-EXOmailbox -Archive -ResultSize unlimited

Update-Progress -ProgressStatus "Retrieving User Mailbox Statistics..." -AdvanceProgressTracker
$Script:MailboxStats = Get-UserMailboxStats

Update-Progress -ProgressStatus "Retrieving Archive Mailbox Statistics..." -AdvanceProgressTracker
$Script:ArchiveMailboxStats = Get-ArchiveMailboxStats

#    - Retrieve Mail Contacts, Transport Rules, Mail Connectors, and MX Records
Update-Progress -ProgressStatus "Retrieving Mail Contacts..." -AdvanceProgressTracker
$Script:MailContacts = Get-Mailcontacts

Update-Progress -ProgressStatus "Retrieving TransportRules..." -AdvanceProgressTracker
$Script:TransportRules = Get-TransportRule -ResultSize unlimited | Select-Object name, state, mode, priority, description, comments

Update-Progress -ProgressStatus "Retrieving Inbound Connectors..." -AdvanceProgressTracker
$Script:InboundConnectors = Get-InboundConnectors

Update-Progress -ProgressStatus "Retrieving Outbound Connectors..." -AdvanceProgressTracker
$Script:OutboundConnectors = Get-OutboundConnectors

Update-Progress -ProgressStatus "Retrieving MX Records..." -AdvanceProgressTracker
$Script:MXRecords = Get-MXRecords

if ($IncludeMailboxPermissions) {
    Update-Progress -ProgressStatus "Retrieving Mailbox Permissions..." -AdvanceProgressTracker
    $Script:MailboxPermissions = Get-AllMailboxPermissions
}

Update-Progress -ProgressStatus "Finalizing report data..." -AdvanceProgressTracker
# 7. Data Enrichment: Update and Format User Data (License Translation, Calculations)
##Update users tab with Values
$Script:users | Add-Member -MemberType NoteProperty -Name MailboxSizeGB -Value "" -Force
$Script:users | Add-Member -MemberType NoteProperty -Name MailboxItemCount -Value "" -Force
$Script:users | Add-Member -MemberType NoteProperty -Name OneDriveSizeGB -Value "" -Force
$Script:users | Add-Member -MemberType NoteProperty -Name OneDriveFileCount -Value "" -Force
$Script:users | Add-Member -MemberType NoteProperty -Name ArchiveSizeGB -Value "" -Force
$Script:users | Add-Member -MemberType NoteProperty -Name Mailboxtype -Value "" -Force
$Script:users | Add-Member -MemberType NoteProperty -Name ArchiveItemCount -Value "" -Force

##Tidy up user Proxyaddresses
foreach ($user in $Script:users) {
    $user.proxyaddresses = $user.proxyaddresses -join ';'
}
##Tidy up group Proxyaddresses
foreach ($group in $script:Groups) {
    $group.proxyaddresses = $group.proxyaddresses -join ';'
}

##Translate License SKUs and groups
foreach ($user in $Script:Users) {
    $user | Add-Member -MemberType NoteProperty -Name "License SKUs" -Value ($user.licenseAssignmentStates.skuid -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Group License Assignments" -Value ($user.licenseAssignmentStates.assignedByGroup -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Disabled Plan IDs" -Value ($user.licenseAssignmentStates.disabledplans -join ";") -Force
}
foreach ($user in $Script:Users) {

    foreach ($Group in $script:Groups) {
        $user.'Group License Assignments' = $user.'Group License Assignments'.replace($group.id, $group.displayName) 
    }
    foreach ($SKU in $SKUs) {
        $user.'License SKUs' = $user.'License SKUs'.replace($SKU.skuid, $SKU.skuPartNumber)
    }
    foreach ($SKUplan in $Script:SKUs.servicePlans) {
        $user.'Disabled Plan IDs' = $user.'Disabled Plan IDs'.replace($SKUplan.servicePlanId, $SKUplan.servicePlanName)
    }

}
foreach ($user in ($Script:users| Where-Object  { $_.usertype -ne "Guest" })) {
    ##Set Mailbox Type
    if ($Script:roommailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Room"
    }    
    elseif ($Script:EquipmentMailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Equipment"
    }
    elseif ($Script:sharedmailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Shared"
    }
    else {
        $user.Mailboxtype = "User"
    }

    ##Set Mailbox Size and count
    If ($Script:MailboxStats| Where-Object  { $_.objectID -eq $user.id }) {
        $user.MailboxSizeGB = (((($Script:MailboxStats| Where-Object  { $_.objectID -eq $user.id }).'Storage Used (Byte)' / 1024) / 1024) / 1024) 
        $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
        $user.MailboxItemCount = ($Script:MailboxStats| Where-Object  { $_.objectID -eq $user.id }).'item count'
    }

    ##Set Shared Mailbox size and count
    If ($Script:SharedMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($Script:SharedMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($Script:SharedMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($Script:SharedMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }

    ##Set Equipment Mailbox size and count
    If ($Script:EquipmentMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($Script:EquipmentMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($Script:EquipmentMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($Script:EquipmentMailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }


    ##Set Room Mailbox size and count
    If ($Script:Roommailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($Script:Roommailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($Script:Roommailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($Script:Roommailboxes| Where-Object  { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }

    ##Set archive size and count
    If ($Script:ArchiveStats| Where-Object  { $_.objectID -eq $user.id }) {
        $user.ArchiveSizeGB = (((($Script:ArchiveStats| Where-Object  { $_.objectID -eq $user.id }).totalitemsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
        $user.ArchiveSizeGB = [math]::Round($user.ArchiveSizeGB, 2)
        $user.ArchiveItemCount = ($Script:ArchiveStats| Where-Object  { $_.objectID -eq $user.id }).ItemCount
    }

    ##Set OneDrive Size and count
    if ($Script:OneDriveUsageReport| Where-Object  { $_.'Owner Principal Name' -eq $user.userPrincipalName }) {
        if (($Script:OneDriveUsageReport| Where-Object  { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'Storage Used (Byte)') {
            If (!$user.OneDriveSizeGB ) {
                $user.OneDriveSizeGB = (((($Script:OneDriveUsageReport| Where-Object  { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'Storage Used (Byte)' / 1024) / 1024) / 1024)
                $user.OneDriveSizeGB = [math]::Round($user.OneDriveSizeGB, 2)
                $user.OneDriveFileCount = ($Script:OneDriveUsageReport| Where-Object  { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'file count'
            }
        }
    }
}

# 8. Export: Generate the Excel Report (Multiple Worksheets)
Update-Progress -ProgressStatus "Generating Excel Report..." -AdvanceProgressTracker
try {
    if (Test-Path ".\TenantAssessment-Template.xlsx") {
        ##Add cover sheet
        Copy-ExcelWorksheet -SourceObject TenantAssessment-Template.xlsx -SourceWorksheet "High-Level" -DestinationWorkbook "$FilePath\$Filename" -DestinationWorksheet "High-Level"
        
    }
    ##Export Data File##
    ##Export User Accounts tab
    $Script:Users | Where-Object { ($_.usertype -ne "Guest") -and ($_.mailboxtype -eq "User") } | Select-Object Migrate, id, displayName, givenName, surname, userPrincipalName, mail, onPremisesSyncEnabled, accountenabled, targetobjectID, targetUPN, TargetMail, MailboxItemCount, MailboxSizeGB, OneDriveSizeGB, OneDriveFileCount, MailboxType, ArchiveSizeGB, ArchiveItemCount, proxyaddresses, 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "User Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export Shared Mailboxes tab
    $Script:Users | Where-Object { ($_.usertype -ne "Guest") -and ($_.mailboxtype -eq "shared") } | Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, proxyaddresses, 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Shared Mailboxes" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export Resource Accounts tab
    $Script:Users | Where-Object { ($_.usertype -ne "Guest") -and (($_.mailboxtype -eq "Room") -or ($_.mailboxtype -eq "Equipment")) } | Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, proxyaddresses, 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Resource Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export SharePoint Tab
    $Script:SharePointUsageReport | Where-Object { ([string]::IsNullOrWhiteSpace($_.teamid)) -and ($_.'Root Web Template' -ne "Team Channel") } | Select-Object 'Site ID', 'Site URL', 'Owner Display Name', 'Is Deleted', 'Last Activity Date', 'File Count', 'Active File Count', 'Page View Count', 'Storage Used (Byte)', 'Root Web Template', 'Owner Principal Name' | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "SharePoint Sites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Teams Tab
    $Script:TeamGroups | Select-Object id, displayname, standardchannels, privatechannels, SharedChannels, Datasize, PrivateChannelsSize, SharedChannelsSize, IncomingSharedChannels, mail, URL, description, createdDateTime, mailEnabled, securityenabled, mailNickname, proxyAddresses, visibility | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Teams"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Guest Accounts tab
    $Script:Users | Where-Object { $_.usertype -eq "Guest" } | Select-Object id, accountenabled, userPrincipalName, mail, displayName, givenName, surname, proxyaddresses, 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usertype | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Guest Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export AAD Apps Tab
    $Script:AADApps | Where-Object { $_.publishername -notlike "Microsoft*" } | Select-Object createddatetime, displayname, publisherName, signinaudience | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "AAD Apps" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Conditional Access Tab
    $Script:ConditionalAccessPolicies | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Conditional Access" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export M365 Apps Usage
    $Script:M365AppsUsageReport | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "M365 Apps Usage" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Unified Groups tab
    $Script:Groups | Where-Object { ($_.grouptypes -Contains "unified") -and ($_.resourceProvisioningOptions -notcontains "Team") } | Select-Object id, displayname, mail, description, createdDateTime, mailEnabled, securityenabled, mailNickname, proxyAddresses, visibility, membershipRule | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Unified Groups"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Standard Groups tab
    $Script:Groups | Where-Object { $_.grouptypes -notContains "unified" } | Select-Object id, displayname, mail, description, createdDateTime, mailEnabled, securityenabled, mailNickname, proxyAddresses, visibility, membershipRule | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Standard Groups"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Mail Contacts tab
    $Script:MailContacts | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "MailContacts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export MX Records tab
    $Script:MXRecords | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "MX Records"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Verified Domains tab
    $Script:OrgDetails.verifieddomains | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Verified Domains"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Transport Rules tab
    $Script:TransportRules | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Transport Rules" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Receive Connectors Tab
    $Script:InboundConnectors  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Receive Connectors" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Send Connectors Tab
    $Script:OutboundConnectors  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Send Connectors" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export OneDrive Tab
    $Script:OneDriveUsageReport  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "OneDrive Sites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    If ($IncludeMailboxPermissions) {
        ##Export Mailbox Permissions Tab
        $Script:MailboxPermissions | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Mailbox Permissions" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    If ($IncludeGroupMembership) {
        ##Export Group Membership Tab
        $Script:GroupMemberships | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Group Membership" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
}
catch {
    write-host "Error exporting report, check permissions and make sure the file is not open! $_"
    pause

}
# 9. Finalize: Complete and Wrap Up the Report
