<#
.SYNOPSIS
Zoekt alle Microsoft Teams Call Queues waarin een gebruiker agent is (direct of via groepen).

.DESCRIPTION
Dit script resolveert een opgegeven User Principal Name (UPN) via Microsoft Graph,
haalt transitive groepslidmaatschappen op en controleert vervolgens alle Teams Call Queues
(Get-CsCallQueue) op directe en indirecte membership.

Het script ondersteunt paginering voor call queues met -First 100 en -Skip,
omdat Get-CsCallQueue standaard/maximaal 100 resultaten per call geeft.

.PARAMETER UserPrincipalName
UPN van de gebruiker die je wilt controleren, bijvoorbeeld kees@bedrijf.nl.

.PARAMETER CsvPath
Optioneel pad voor CSV-export van de resultaten.

.PARAMETER InstallMissingModules
Installeert ontbrekende modules automatisch (CurrentUser scope) als deze switch is opgegeven.

.PARAMETER DebugProperties
Toont per call queue potentieel relevante properties voor directe agents en groepen,
zodat wijzigingen in Teams PowerShell output makkelijker te debuggen zijn.

.EXAMPLE
.\Find-TeamsCallQueuesForUser.ps1 -UserPrincipalName kees@bedrijf.nl

.EXAMPLE
.\Find-TeamsCallQueuesForUser.ps1 -UserPrincipalName kees@bedrijf.nl -CsvPath .\queues-kees.csv

.NOTES
Benodigde modules:
- MicrosoftTeams
- Microsoft.Graph.Users
- Microsoft.Graph.Groups

Benodigde Graph scopes (minimaal):
- User.Read.All
- GroupMember.Read.All
(Alternatief: Directory.Read.All)

Vereist voldoende rechten in Teams en Graph tenantcontext.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [string]$CsvPath,

    [switch]$InstallMissingModules,

    [switch]$DebugProperties,

    [int]$LoginTimeoutSeconds = 600
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,

        [switch]$InstallIfMissing
    )

    $loaded = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue
    if (-not $loaded) {
        if ($InstallIfMissing) {
            Write-Verbose "Module '$ModuleName' ontbreekt. Installatie gestart..."
            try {
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            }
            catch {
                throw "Kon module '$ModuleName' niet installeren. Fout: $($_.Exception.Message)"
            }
        }
        else {
            throw "Vereiste module '$ModuleName' ontbreekt. Installeer handmatig of gebruik -InstallMissingModules."
        }
    }

    try {
        Import-Module -Name $ModuleName -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Kon module '$ModuleName' niet importeren. Fout: $($_.Exception.Message)"
    }
}

function Connect-RequiredServices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds
    )

    $requiredScopes = @('User.Read.All', 'GroupMember.Read.All', 'Directory.Read.All')

    while ($true) {
        try {
            Write-Host 'Stap 1/2: aanmelden bij Microsoft Teams'
            Write-Host 'Open de browser met de device-code prompt en rond de login volledig af.'
            Connect-MicrosoftTeams -UseDeviceAuthentication -ErrorAction Stop | Out-Null
            [void](Get-CsCallQueue -First 1 -WarningAction SilentlyContinue -ErrorAction Stop)
        }
        catch {
            Write-Error "Microsoft Teams login/validatie mislukt. Fout: $($_.Exception.Message)"
            Read-Host 'Druk op Enter om opnieuw te proberen'
            continue
        }

        try {
            Write-Host 'Stap 2/2: aanmelden bij Microsoft Graph'
            Write-Host 'Open de browser met de device-code prompt en rond de login volledig af.'

            $ctx = $null
            try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch { $ctx = $null }

            $needGraphConnect = $true
            if ($ctx -and $ctx.Account) {
                $grantedScopes = @($ctx.Scopes)
                $missing = $requiredScopes | Where-Object { $_ -notin $grantedScopes }
                if (-not $missing) {
                    $needGraphConnect = $false
                }
            }

            if ($needGraphConnect) {
                Connect-MgGraph -Scopes $requiredScopes -UseDeviceAuthentication -ContextScope Process -ClientTimeout $TimeoutSeconds -NoWelcome -ErrorAction Stop | Out-Null
            }

            $validatedCtx = Get-MgContext -ErrorAction Stop
            if (-not $validatedCtx -or -not $validatedCtx.Account) {
                throw 'Graph context is niet actief na aanmelden.'
            }

            [void](Get-MgUser -Top 1 -Property Id -ErrorAction Stop)
        }
        catch {
            Write-Error "Microsoft Graph login/validatie mislukt of timeout bereikt. Fout: $($_.Exception.Message)"
            Read-Host 'Druk op Enter om opnieuw te proberen'
            continue
        }

        Write-Host 'Alle verbindingen zijn succesvol. Script gaat nu verder.'
        return
    }
}

function ConvertTo-NormalizedGuidString {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [AllowNull()]
        $Value
    )

    process {
        if ($null -eq $Value) { return $null }

        try {
            $text = [string]$Value
        }
        catch {
            return $null
        }
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }

        $text = $text.Trim().Trim('{}').Trim()

        $guidValue = [Guid]::Empty
        if ([Guid]::TryParse($text, [ref]$guidValue)) {
            return $guidValue.ToString().ToLowerInvariant()
        }

        return $null
    }
}

function Get-PossibleIdValuesFromObject {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [AllowNull()]
        $InputObject
    )

    process {
        $results = New-Object System.Collections.Generic.List[string]

        if ($null -eq $InputObject) {
            return @()
        }

        if ($InputObject -is [string] -or $InputObject -is [Guid]) {
            try { $results.Add([string]$InputObject) } catch {}
            return @($results)
        }

        $propCandidates = @(
            'Id', 'ID', 'ObjectId', 'ObjectID', 'Identity',
            'UserPrincipalName', 'UPN', 'SipAddress', 'Mail',
            'GroupId', 'TeamId', 'ChannelId', 'DistributionList',
            'DistributionListId', 'Guid'
        )

        foreach ($propName in $propCandidates) {
            try {
                $p = $InputObject.PSObject.Properties[$propName]
                if ($p -and $null -ne $p.Value -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
                    $results.Add([string]$p.Value)
                }
            }
            catch {
                continue
            }
        }

        if ($results.Count -eq 0) {
            try {
                foreach ($p in $InputObject.PSObject.Properties) {
                    if ($p.Name -match '(?i)(id|guid|principalname|upn|sip|mail)') {
                        if ($null -ne $p.Value -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
                            $results.Add([string]$p.Value)
                        }
                    }
                }
            }
            catch {
                # negeer parsingfouten; functie blijft best-effort
            }
        }

        return @($results)
    }
}

function Get-UserTransitiveGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    $groupMap = @{}
    $groupIdSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    try {
        $memberships = Get-MgUserTransitiveMemberOf -UserId $UserId -All -ErrorAction Stop
    }
    catch {
        throw "Kon transitive memberships niet ophalen voor userId '$UserId'. Controleer Graph-rechten. Fout: $($_.Exception.Message)"
    }

    foreach ($m in $memberships) {
        $odataType = $null
        if ($m.PSObject.Properties['AdditionalProperties']) {
            $ap = $m.AdditionalProperties
            if ($ap -is [hashtable] -and $ap.ContainsKey('@odata.type')) {
                $odataType = [string]$ap['@odata.type']
            }
        }

        if ($odataType -and $odataType -notmatch '(?i)microsoft\.graph\.group') {
            continue
        }

        $gid = ConvertTo-NormalizedGuidString $m.Id
        if (-not $gid) { continue }

        [void]$groupIdSet.Add($gid)

        $displayName = $null
        $mail = $null

        if ($m.PSObject.Properties['DisplayName']) { $displayName = [string]$m.DisplayName }
        if ($m.PSObject.Properties['Mail']) { $mail = [string]$m.Mail }

        if ([string]::IsNullOrWhiteSpace($displayName) -and $m.PSObject.Properties['AdditionalProperties']) {
            $ap = $m.AdditionalProperties
            if ($ap -is [hashtable]) {
                if ($ap.ContainsKey('displayName')) { $displayName = [string]$ap['displayName'] }
                if ($ap.ContainsKey('mail')) { $mail = [string]$ap['mail'] }
            }
        }

        if (-not $groupMap.ContainsKey($gid)) {
            $groupMap[$gid] = [PSCustomObject]@{
                Id          = $gid
                DisplayName = $displayName
                Mail        = $mail
            }
        }
    }

    return [PSCustomObject]@{
        GroupIdSet = $groupIdSet
        GroupMap   = $groupMap
    }
}

function Get-AllTeamsCallQueues {
    [CmdletBinding()]
    param()

    $all = New-Object System.Collections.Generic.List[object]
    $skip = 0
    $pageSize = 100

    while ($true) {
        Write-Verbose "Ophalen call queues: -First $pageSize -Skip $skip"
        try {
            $page = @(Get-CsCallQueue -First $pageSize -Skip $skip -WarningAction SilentlyContinue -ErrorAction Stop)
        }
        catch {
            throw "Kon call queues niet ophalen met Get-CsCallQueue. Fout: $($_.Exception.Message)"
        }

        if (-not $page -or $page.Count -eq 0) {
            break
        }

        foreach ($q in $page) { [void]$all.Add($q) }

        if ($page.Count -lt $pageSize) {
            break
        }

        $skip += $pageSize
    }

    return @($all)
}

function Test-QueueDirectMembership {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Queue,
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$UserUpn
    )

    $matchIds = New-Object System.Collections.Generic.List[string]

    $candidateProps = $Queue.PSObject.Properties | Where-Object {
        $_.Name -match '(?i)^(users?|agents?)$' -or $_.Name -match '(?i)(user|agent)'
    }

    foreach ($p in $candidateProps) {
        $val = $p.Value
        if ($null -eq $val) { continue }

        $items = @()
        if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
            $items = @($val)
        }
        else {
            $items = @($val)
        }

        foreach ($item in $items) {
            $ids = @(Get-PossibleIdValuesFromObject -InputObject $item)
            foreach ($idVal in $ids) {
                if ([string]::IsNullOrWhiteSpace($idVal)) { continue }

                $normalized = ConvertTo-NormalizedGuidString $idVal
                if ($normalized -and $normalized -eq $UserId) {
                    $matchIds.Add($idVal)
                    continue
                }

                if ([string]$idVal -ieq $UserUpn) {
                    $matchIds.Add($idVal)
                    continue
                }
            }
        }
    }

    return [PSCustomObject]@{
        IsMatch    = ($matchIds.Count -gt 0)
        MatchedIds = @($matchIds | Select-Object -Unique)
    }
}

function Test-QueueGroupMembership {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Queue,
        [Parameter(Mandatory = $true)]$GroupData
    )

    $matched = New-Object System.Collections.Generic.List[object]

    $candidateProps = $Queue.PSObject.Properties | Where-Object {
        $_.Name -match '(?i)(distributionlist|group|teamid|channelid)'
    }

    foreach ($p in $candidateProps) {
        $val = $p.Value
        if ($null -eq $val) { continue }

        $items = @()
        if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
            $items = @($val)
        }
        else {
            $items = @($val)
        }

        foreach ($item in $items) {
            $ids = @(Get-PossibleIdValuesFromObject -InputObject $item)
            foreach ($idVal in $ids) {
                $gid = ConvertTo-NormalizedGuidString $idVal
                if (-not $gid) { continue }

                if ($GroupData.GroupIdSet.Contains($gid)) {
                    if ($GroupData.GroupMap.ContainsKey($gid)) {
                        $matched.Add($GroupData.GroupMap[$gid])
                    }
                    else {
                        $matched.Add([PSCustomObject]@{ Id = $gid; DisplayName = $null; Mail = $null })
                    }
                }
            }
        }
    }

    $unique = @($matched | Sort-Object Id -Unique)
    return [PSCustomObject]@{
        IsMatch        = ($unique.Count -gt 0)
        MatchedGroups  = $unique
    }
}

try {
    Assert-Module -ModuleName 'MicrosoftTeams' -InstallIfMissing:$InstallMissingModules
    Assert-Module -ModuleName 'Microsoft.Graph.Authentication' -InstallIfMissing:$InstallMissingModules
    Assert-Module -ModuleName 'Microsoft.Graph.Users' -InstallIfMissing:$InstallMissingModules
    Assert-Module -ModuleName 'Microsoft.Graph.Groups' -InstallIfMissing:$InstallMissingModules

    Connect-RequiredServices -TimeoutSeconds $LoginTimeoutSeconds

    Write-Verbose "Resolven gebruiker '$UserPrincipalName' via Graph..."
    try {
        $user = Get-MgUser -UserId $UserPrincipalName -Property Id,UserPrincipalName,DisplayName,Mail -ErrorAction Stop
    }
    catch {
        throw "Gebruiker '$UserPrincipalName' niet gevonden of geen rechten om deze te lezen. Fout: $($_.Exception.Message)"
    }

    if (-not $user -or -not $user.Id) {
        throw "Gebruiker '$UserPrincipalName' bestaat niet of kon niet worden geresolved."
    }

    $userId = ConvertTo-NormalizedGuidString $user.Id
    if (-not $userId) {
        throw "Gebruiker '$UserPrincipalName' heeft geen valide GUID Id in Graph-response."
    }

    $groupData = Get-UserTransitiveGroups -UserId $user.Id
    if ($groupData.GroupIdSet.Count -eq 0) {
        Write-Warning "Geen transitive groepslidmaatschappen gevonden voor '$($user.UserPrincipalName)'. Alleen directe membership wordt gecontroleerd."
    }

    $queues = @(Get-AllTeamsCallQueues)
    if (-not $queues -or $queues.Count -eq 0) {
        throw 'Er zijn geen call queues gevonden in de tenant of je hebt onvoldoende rechten.'
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($q in $queues) {
        if ($DebugProperties) {
            $dbgProps = $q.PSObject.Properties.Name | Where-Object {
                $_ -match '(?i)(user|agent|group|distributionlist|teamid|channelid)'
            }
            Write-Host ("[DebugProperties] Queue '{0}': {1}" -f ($q.Name), (($dbgProps -join ', ')))
        }

        $direct = Test-QueueDirectMembership -Queue $q -UserId $userId -UserUpn $user.UserPrincipalName
        $viaGroup = Test-QueueGroupMembership -Queue $q -GroupData $groupData

        if (-not $direct.IsMatch -and -not $viaGroup.IsMatch) { continue }

        $membershipType = if ($direct.IsMatch -and $viaGroup.IsMatch) { 'Both' } elseif ($direct.IsMatch) { 'Direct' } else { 'Group' }

        $matchedGroupNames = @($viaGroup.MatchedGroups | ForEach-Object { $_.DisplayName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $matchedGroupMail = @($viaGroup.MatchedGroups | ForEach-Object { $_.Mail } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $matchedGroupIds = @($viaGroup.MatchedGroups | ForEach-Object { $_.Id } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        $results.Add([PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            UserDisplayName   = $user.DisplayName
            CallQueueName     = $q.Name
            CallQueueIdentity = $q.Identity
            MembershipType    = $membershipType
            DirectMatchedIds  = ($direct.MatchedIds -join '; ')
            MatchedGroupNames = ($matchedGroupNames -join '; ')
            MatchedGroupMail  = ($matchedGroupMail -join '; ')
            MatchedGroupIds   = ($matchedGroupIds -join '; ')
            RoutingMethod     = $q.RoutingMethod
            AllowOptOut       = $q.AllowOptOut
            ConferenceMode    = $q.ConferenceMode
            ChannelId         = $q.ChannelId
            Notes             = if ($DebugProperties) { 'DebugProperties enabled' } else { $null }
        }) | Out-Null
    }

    if ($results.Count -eq 0) {
        Write-Warning "Geen call queue memberships gevonden voor gebruiker '$($user.UserPrincipalName)'."
    }

    $results |
        Sort-Object CallQueueName |
        Format-Table UserPrincipalName, UserDisplayName, CallQueueName, MembershipType, MatchedGroupNames, RoutingMethod -AutoSize

    Write-Host "Totaal gevonden call queues: $($results.Count)"

    if ($CsvPath) {
        try {
            $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            Write-Host "CSV geëxporteerd naar: $CsvPath"
        }
        catch {
            throw "CSV export mislukt naar '$CsvPath'. Fout: $($_.Exception.Message)"
        }
    }

    $results
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
