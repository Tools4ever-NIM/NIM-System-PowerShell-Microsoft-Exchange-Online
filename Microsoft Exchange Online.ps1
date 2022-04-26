#
# Microsoft Exchange Online.ps1 - IDM System PowerShell Script for Microsoft Exchange Online Services.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#


#
# https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/exchange-online-powershell-v2/exchange-online-powershell-v2?view=exchange-ps
#

$EXOManagementMinVersion = @{ Major = 2; Minor = 0; Build = 5 }

if (!(Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (!(Get-PackageProvider -ListAvailable | Where-Object { $_.Name -eq 'NuGet' }) -or (Get-PackageProvider -Name 'NuGet').Version -lt '2.8.5.201') {
        Install-PackageProvider -Name 'NuGet' -MinimumVersion '2.8.5.201' -Scope 'CurrentUser' -Force
    }

    Install-Module -Name 'ExchangeOnlineManagement' -Scope 'CurrentUser' -Force
}


$Log_MaskableKeys = @(
    'password'
)


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'UseModernAuthentication'
                type = 'checkbox'
                label = 'Use modern authentication'
                value = $true
            }
            @{
                name = 'ConnectionUri'
                type = 'textbox'
                label = 'Connection URI'
                label_indent = $true
                value = ''
                hidden = 'UseModernAuthentication'
            }
            @{
                name = 'AzureADAuthorizationEndpointUri'
                type = 'textbox'
                label = 'Azure AD authorization endpoint URI'
                label_indent = $true
                value = ''
                hidden = 'UseModernAuthentication'
            }
            @{
                name = 'DelegatedOrganization'
                type = 'textbox'
                label = 'Delegated organization name'
                label_indent = $true
                value = ''
                hidden = 'UseModernAuthentication'
            }
            @{
                name = 'ExchangeEnvironmentName'
                type = 'combo'
                label = 'Exchange environment name'
                label_indent = $true
                table = @{
                    rows = @(
                        @{ id = 'O365China';        display_text = 'O365China' }
                        @{ id = 'O365Default';      display_text = 'O365Default' }
                        @{ id = 'O365GermanyCloud'; display_text = 'O365GermanyCloud' }
                        @{ id = 'O365USGovDoD';     display_text = 'O365USGovDoD' }
                        @{ id = 'O365USGovGCCHigh'; display_text = 'O365USGovGCCHigh' }
                    )
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = 'O365Default'
                hidden = 'UseModernAuthentication'
            }
            @{
                name = 'Username'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                value = ''
                hidden = 'UseModernAuthentication'
            }
            @{
                name = 'Password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                value = ''
                hidden = 'UseModernAuthentication'
            }
            @{
                name = 'AppId'
                type = 'textbox'
                label = 'Application ID'
                label_indent = $true
                value = ''
                hidden = '!UseModernAuthentication'
            }
            @{
                name = 'Organization'
                type = 'textbox'
                label = 'Organization'
                label_indent = $true
                value = ''
                hidden = '!UseModernAuthentication'
            }
            @{
                name = 'certificate'
                type = 'textbox'
                label = 'Certificate name'
                label_indent = $true
                value = ''
                hidden = '!UseModernAuthentication'
            }
            @{
                name = 'PageSize'
                type = 'textbox'
                label = 'Page size'
                value = '1000'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                value = 10
            }
        )
    }

    if ($TestConnection) {
        Open-MsExchangeSession (ConvertFrom-Json2 $ConnectionParams)
    }

    if ($Configuration) {
        Open-MsExchangeSession (ConvertFrom-Json2 $ConnectionParams)

        @(
            @{
                name = 'organizational_unit'
                type = 'combo'
                label = 'Organizational unit'
                table = @{
                    rows = @( @{ display = '*'; value = '*' } ) + @( Get-MsExchangeOrganizationalUnit | Sort-Object -Property 'canonicalName' | ForEach-Object { @{ display = $_.canonicalName; value = $_.distinguishedName } } )
                    settings_combo = @{
                        display_column = 'display'
                        value_column = 'value'
                    }
                }
                value = '*'
            }
        )
    }

    Log info "Done"
}


function Idm-OnUnload {
    Close-MsExchangeSession
}


#
# CRUD functions
#

$Properties = @{
    CASMailbox = @(
        @{ name = 'ActiveSyncEnabled';             default = $true; set = $true; }
        @{ name = 'ECPEnabled';                    default = $true; set = $true; }
        @{ name = 'Guid';                          default = $true; key = $true; }
        @{ name = 'Id';                            default = $true;              }
        @{ name = 'Identity';                      default = $true;              }
        @{ name = 'ImapEnabled';                   default = $true; set = $true; }
        @{ name = 'IsValid';                       default = $true;              }
        @{ name = 'LinkedMasterAccount';           default = $true;              }
        @{ name = 'OWAEnabled';                    default = $true; set = $true; }
        @{ name = 'PopEnabled';                    default = $true; set = $true; }
    )

    Mailbox = @(
        @{ name = 'Alias';                         default = $true; enable = $true; set = $true; }
        @{ name = 'ArchiveName';                   default = $true; enable = $true; set = $true; }
        @{ name = 'DeliverToMailboxAndForward';    set = $true;                                  }
        @{ name = 'DisplayName';                   default = $true; enable = $true; set = $true; }
        @{ name = 'DistinguishedName';                                                           }
        @{ name = 'EmailAddresses';                default = $true; set = $true;                 }
        @{ name = 'Guid';                          default = $true; key = $true;                 }
        @{ name = 'HiddenFromAddressListsEnabled'; set = $true;                                  }
        @{ name = 'Id';                            default = $true;                              }
    )
}


# Default properties and IDM properties are the same
foreach ($key in $Properties.Keys) {
    for ($i = 0; $i -lt $Properties.$key.Count; $i++) {
        if ($Properties.$key[$i].default) {
            $Properties.$key[$i].idm = $true
        }
    }
}


function Idm-CASMailboxesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'CASMailbox'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'Unlimited'
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.CASMailbox | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.CASMailbox | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/client-access/get-casmailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Get-MsExchangeCASMailbox" -In @call_params
            Get-MsExchangeCASMailbox @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-CASMailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.CASMailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.CASMailbox | Where-Object { !$_.key -and !$_.set } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.CASMailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/client-access/set-casmailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Set-MsExchangeCASMailbox" -In @call_params
                $rv = Set-MsExchangeCASMailbox @call_params
            LogIO info "Set-MsExchangeCASMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxEnable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.enable } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/enable-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Enable-MsExchangeMailbox" -In @call_params
                $rv = Enable-MsExchangeMailbox @call_params
            LogIO info "Enable-MsExchangeMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'Mailbox'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'Unlimited'
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.Mailbox | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Get-MsExchangeMailbox" -In @call_params
            Get-MsExchangeMailbox @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.set } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Set-MsExchangeMailbox" -In @call_params
                $rv = Set-MsExchangeMailbox @call_params
            LogIO info "Set-MsExchangeMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxDisable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.disable } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
            Confirm  = $false   # Be non-interactive
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/disable-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Disable-MsExchangeMailbox" -In @call_params
                $rv = Disable-MsExchangeMailbox @call_params
            LogIO info "Disable-MsExchangeMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxPermissionAdd {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.permissionAdd } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Add-MsExchangeMailboxPermission" -In @call_params
                $rv = Add-MsExchangeMailboxPermission @call_params
            LogIO info "Add-MsExchangeMailboxPermission" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxPermissionRemove {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.permissionRemove } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
            Confirm  = $false   # Be non-interactive
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/remove-mailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Remove-MsExchangeMailboxPermission" -In @call_params
                $rv = Remove-MsExchangeMailboxPermission @call_params
            LogIO info "Remove-MsExchangeMailboxPermission" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


#
# Helper functions
#

function Open-MsExchangeSession {
    param (
        [hashtable] $SystemParams
    )

    # Use connection related parameters only
    $connection_params = if ($SystemParams.UseModernAuthentication) {
        [ordered]@{
            AppId        = $SystemParams.AppId
            Organization = $SystemParams.Organization
            Certificate  = $SystemParams.certificate
            PageSize     = $SystemParams.PageSize
        }
    }
    else {
        [ordered]@{
            ConnectionUri                   = $SystemParams.ConnectionUri
            AzureADAuthorizationEndpointUri = $SystemParams.AzureADAuthorizationEndpointUri
            DelegatedOrganization           = $SystemParams.DelegatedOrganization
            ExchangeEnvironmentName         = $SystemParams.ExchangeEnvironmentName
            Username                        = $SystemParams.Username
            Password                        = $SystemParams.Password
            PageSize                        = $SystemParams.PageSize
        }
    }

    $connection_string = ConvertTo-Json $connection_params -Compress -Depth 32

    if ($Global:MsExchangePSSession -and $connection_string -ne $Global:MsExchangeConnectionString) {
        Log info "MsExchangePSSession connection parameters changed"
        Close-MsExchangeSession
    }

    if ($Global:MsExchangePSSession -and $Global:MsExchangePSSession.State -ne 'Opened') {
        Log warn "MsExchangePSSession State is '$($Global:MsExchangePSSession.State)'"
        Close-MsExchangeSession
    }

    if ($Global:MsExchangePSSession) {
        #Log debug "Reusing MsExchangePSSession"
    }
    else {
        Log info "Opening MsExchangePSSession '$connection_string'"

        $params = Copy-Object $connection_params

        if ($SystemParams.UseModernAuthentication) {
            $v_act = (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement').Version

            if ($v_act.Major -lt $EXOManagementMinVersion.Major -or $v_act.Major -eq $EXOManagementMinVersion.Major -and ($v_act.Minor -lt $EXOManagementMinVersion.Minor -or $v_act.Minor -eq $EXOManagementMinVersion.Minor -and $v_act.Build -lt $EXOManagementMinVersion.Build)) {
                throw "ExchangeOnlineManagement PowerShell Module version older than $($EXOManagementMinVersion.Major).$($EXOManagementMinVersion.Minor).$($EXOManagementMinVersion.Build)"
            }

            $params.Certificate = Nim-GetCertificate $connection_params.certificate
        }
        else {
            $params.Credential = New-Object System.Management.Automation.PSCredential($connection_params.Username, (ConvertTo-SecureString $connection_params.Password -AsPlainText -Force))
            $params.Remove('Username')
            $params.Remove('Password')
        }

        try {
            Connect-ExchangeOnline @params -Prefix 'MsExchange' -ShowBanner:$false

            $Global:MsExchangePSSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}
            $Global:MsExchangeConnectionString = $connection_string
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }

        Log info "Done"
    }
}


function Close-MsExchangeSession {
    if ($Global:MsExchangePSSession) {
        Log info "Closing MsExchangePSSession"

        try {
            Remove-PSSession -Session $Global:MsExchangePSSession -ErrorAction SilentlyContinue
            $Global:MsExchangePSSession = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log info "Done"
    }
}


function Get-ClassMetaData {
    param (
        [string] $SystemParams,
        [string] $Class
    )

    @(
        @{
            name = 'properties'
            type = 'grid'
            label = 'Properties'
            table = @{
                rows = @( $Global:Properties.$Class | ForEach-Object {
                    @{
                        name = $_.name
                        usage_hint = @( @(
                            foreach ($key in $_.Keys) {
                                if ($key -notin @('default', 'idm', 'key')) { continue }

                                if ($key -eq 'idm') {
                                    $key.Toupper()
                                }
                                else {
                                    $key.Substring(0,1).Toupper() + $key.Substring(1)
                                }
                            }
                        ) | Sort-Object) -join ' | '
                    }
                })
                settings_grid = @{
                    selection = 'multiple'
                    key_column = 'name'
                    checkbox = $true
                    filter = $true
                    columns = @(
                        @{
                            name = 'name'
                            display_name = 'Name'
                        }
                        @{
                            name = 'usage_hint'
                            display_name = 'Usage hint'
                        }
                    )
                }
            }
            value = ($Global:Properties.$Class | Where-Object { $_.default }).name
        }
    )
}
