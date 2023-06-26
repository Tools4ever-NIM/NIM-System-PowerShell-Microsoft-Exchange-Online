#
# Microsoft Exchange Online.ps1 - IDM System PowerShell Script for Microsoft Exchange Online Services.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#


#
# https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/exchange-online-powershell-v2/exchange-online-powershell-v2?view=exchange-ps
#

#
# Properly configure API permissions in Microsoft Azure to work with this connector.
#
# - Go to portal.azure.com
# - Find the used App registration
# - Select 'API permissions'
# - Assure 'Office 365 Exchange Online / Exchange.ManageAsApp' Application permissions are granted
#
# If not, do the following:
#
# - Click 'Add a permission'
# - Select 'APIs my organization uses'
# - Select 'Office 365 Exchange Online'
# - Choose 'Application permissions'
# - Unfold 'Exchange'
# - Select 'Exchange.ManageAsApp'
# - Click 'Add permission'
# - Click 'Grand admin consent for ...'
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
        @{ name = 'ActiveSyncAllowedDeviceIDs';                          options = @('set')                      }
        @{ name = 'ActiveSyncBlockedDeviceIDs';                          options = @('set')                      }
        @{ name = 'ActiveSyncDebugLogging';                              options = @('set')                      }
        @{ name = 'ActiveSyncEnabled';                                   options = @('default', 'set')           }
        @{ name = 'ActiveSyncMailboxPolicy';                             options = @('set')                      }
        @{ name = 'ActiveSyncMailboxPolicyIsDefaulted';                                                          }
        @{ name = 'ActiveSyncSuppressReadReceipt';                       options = @('set')                      }
        @{ name = 'DistinguishedName';                                                                           }
        @{ name = 'EwsAllowEntourage';                                   options = @('set')                      }
        @{ name = 'EwsAllowList';                                        options = @('set')                      }
        @{ name = 'EwsAllowMacOutlook';                                  options = @('set')                      }
        @{ name = 'EwsAllowOutlook';                                     options = @('set')                      }
        @{ name = 'EwsApplicationAccessPolicy';                          options = @('set')                      }
        @{ name = 'EwsBlockList';                                        options = @('set')                      }
        @{ name = 'EwsEnabled';                                          options = @('set')                      }
        @{ name = 'ExchangeVersion';                                                                             }
        @{ name = 'ExternalImapSettings';                                                                        }
        @{ name = 'ExternalPopSettings';                                                                         }
        @{ name = 'ExternalSmtpSettings';                                                                        }
        @{ name = 'Guid';                                                options = @('default', 'key')           }
        @{ name = 'Id';                                                  options = @('default')                  }
        @{ name = 'Identity';                                            options = @('default')                  }
        @{ name = 'ImapEnabled';                                         options = @('default', 'set')           }
        @{ name = 'ImapForceICalForCalendarRetrievalOption';             options = @('set')                      }
        @{ name = 'ImapMessagesRetrievalMimeFormat';                     options = @('set')                      }
        @{ name = 'ImapSuppressReadReceipt';                             options = @('set')                      }
        @{ name = 'ImapUseProtocolDefaults';                             options = @('set')                      }
        @{ name = 'InternalImapSettings';                                                                        }
        @{ name = 'InternalPopSettings';                                                                         }
        @{ name = 'InternalSmtpSettings';                                                                        }
        @{ name = 'IsOptimizedForAccessibility';                         options = @('set')                      }
        @{ name = 'IsValid';                                             options = @('default')                  }
        @{ name = 'LegacyExchangeDN';                                                                            }
        @{ name = 'LinkedMasterAccount';                                 options = @('default')                  }
        @{ name = 'MacOutlookEnabled';                                   options = @('set')                      }
        @{ name = 'MapiHttpEnabled';                                     options = @('set')                      }
        @{ name = 'ObjectCategory';                                                                              }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'OneWinNativeOutlookEnabled';                          options = @('set')                      }
        @{ name = 'OrganizationId';                                                                              }
        @{ name = 'OriginatingServer';                                                                           }
        @{ name = 'OutlookMobileEnabled';                                options = @('set')                      }
        @{ name = 'OWAEnabled';                                          options = @('default', 'set')           }
        @{ name = 'OWAforDevicesEnabled';                                options = @('set')                      }
        @{ name = 'OwaMailboxPolicy';                                    options = @('set')                      }
        @{ name = 'PopEnabled';                                          options = @('default', 'set')           }
        @{ name = 'PopForceICalForCalendarRetrievalOption';              options = @('set')                      }
        @{ name = 'PopMessageDeleteEnabled';                                                                     }
        @{ name = 'PopMessagesRetrievalMimeFormat';                      options = @('set')                      }
        @{ name = 'PopSuppressReadReceipt';                              options = @('set')                      }
        @{ name = 'PopUseProtocolDefaults';                              options = @('set')                      }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'PublicFolderClientAccess';                            options = @('set')                      }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'ServerLegacyDN';                                                                              }
        @{ name = 'ServerName';                                                                                  }
        @{ name = 'ShowGalAsDefaultView';                                options = @('set')                      }
        @{ name = 'SmtpClientAuthenticationDisabled';                    options = @('set')                      }
        @{ name = 'UniversalOutlookEnabled';                             options = @('set')                      }
        @{ name = 'WhenChanged';                                                                                 }
        @{ name = 'WhenChangedUTC';                                                                              }
        @{ name = 'WhenCreated';                                                                                 }
        @{ name = 'WhenCreatedUTC';                                                                              }
    )

    Mailbox = @(
        @{ name = 'AcceptMessagesOnlyFrom';                              options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';                     options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';              options = @('set')                      }
        @{ name = 'AccountDisabled';                                     options = @('set')                      }
        @{ name = 'AddressBookPolicy';                                   options = @('enable', 'set')            }
        @{ name = 'AddressListMembership';                                                                       }
        @{ name = 'AdminDisplayVersion';                                                                         }
        @{ name = 'AdministrativeUnits';                                                                         }
        @{ name = 'AggregatedMailboxGuids';                                                                      }
        @{ name = 'Alias';                                               options = @('default', 'enable', 'set') }
        @{ name = 'ArchiveGuid';                                         options = @('enable')                   }
        @{ name = 'ArchiveName';                                         options = @('default', 'enable', 'set') }
        @{ name = 'ArchiveRelease';                                                                              }
        @{ name = 'ArchiveState';                                        options = @('set')                      }
        @{ name = 'AuditAdmin';                                          options = @('set')                      }
        @{ name = 'AuditDelegate';                                       options = @('set')                      }
        @{ name = 'AuditEnabled';                                        options = @('set')                      }
        @{ name = 'AuditLogAgeLimit';                                    options = @('set')                      }
        @{ name = 'AuditOwner';                                          options = @('set')                      }
        @{ name = 'AutoExpandingArchiveEnabled';                                                                 }
        @{ name = 'BypassModerationFromSendersOrMembers';                options = @('set')                      }
        @{ name = 'CalendarRepairDisabled';                              options = @('set')                      }
        @{ name = 'CalendarVersionStoreDisabled';                        options = @('set')                      }
        @{ name = 'ComplianceTagHoldApplied';                                                                    }
        @{ name = 'CustomAttribute1';                                    options = @('set')                      }
        @{ name = 'CustomAttribute2';                                    options = @('set')                      }
        @{ name = 'CustomAttribute3';                                    options = @('set')                      }
        @{ name = 'CustomAttribute4';                                    options = @('set')                      }
        @{ name = 'CustomAttribute5';                                    options = @('set')                      }
        @{ name = 'CustomAttribute6';                                    options = @('set')                      }
        @{ name = 'CustomAttribute7';                                    options = @('set')                      }
        @{ name = 'CustomAttribute8';                                    options = @('set')                      }
        @{ name = 'CustomAttribute9';                                    options = @('set')                      }
        @{ name = 'CustomAttribute10';                                   options = @('set')                      }
        @{ name = 'CustomAttribute11';                                   options = @('set')                      }
        @{ name = 'CustomAttribute12';                                   options = @('set')                      }
        @{ name = 'CustomAttribute13';                                   options = @('set')                      }
        @{ name = 'CustomAttribute14';                                   options = @('set')                      }
        @{ name = 'CustomAttribute15';                                   options = @('set')                      }
        @{ name = 'DataEncryptionPolicy';                                options = @('set')                      }
        @{ name = 'DefaultAuditSet';                                     options = @('set')                      }
        @{ name = 'DefaultPublicFolderMailbox';                          options = @('set')                      }
        @{ name = 'DelayHoldApplied';                                                                            }
        @{ name = 'DeliverToMailboxAndForward';                          options = @('set')                      }
        @{ name = 'DisabledArchiveDatabase';                                                                     }
        @{ name = 'DisabledArchiveGuid';                                                                         }
        @{ name = 'DisabledMailboxLocations';                                                                    }
        @{ name = 'DisplayName';                                         options = @('default', 'enable', 'set') }
        @{ name = 'DistinguishedName';                                                                           }
        @{ name = 'EffectivePublicFolderMailbox';                                                                }
        @{ name = 'ElcProcessingDisabled';                               options = @('set')                      }
        @{ name = 'EmailAddresses';                                      options = @('default', 'set')           }
        @{ name = 'EndDateForRetentionHold';                             options = @('set')                      }
        @{ name = 'EnforcedTimestamps';                                  options = @('set')                      }
        @{ name = 'ExchangeGuid';                                                                                }
        @{ name = 'ExchangeUserAccountControl';                                                                  }
        @{ name = 'ExchangeSecurityDescriptor';                                                                  }
        @{ name = 'ExchangeVersion';                                                                             }
        @{ name = 'ExtensionCustomAttribute1';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute2';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute3';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute4';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute5';                           options = @('set')                      }
        @{ name = 'Extensions';                                                                                  }
        @{ name = 'ExternalDirectoryObjectId';                                                                   }
        @{ name = 'ExternalOofOptions';                                  options = @('set')                      }
        @{ name = 'ForwardingAddress';                                   options = @('set')                      }
        @{ name = 'ForwardingSmtpAddress';                               options = @('set')                      }
        @{ name = 'GeneratedOfflineAddressBooks';                                                                }
        @{ name = 'GrantSendOnBehalfTo';                                 options = @('set')                      }
        @{ name = 'GroupMailbox';                                        options = @('set')                      }
        @{ name = 'Guid';                                                options = @('default', 'key')           }
        @{ name = 'HasPicture';                                                                                  }
        @{ name = 'HasSnackyAppData';                                                                            }
        @{ name = 'HasSpokenName';                                                                               }
        @{ name = 'HiddenFromAddressListsEnabled';                       options = @('set')                      }
        @{ name = 'Id';                                                  options = @('default')                  }
        @{ name = 'Identity';                                                                                    }
        @{ name = 'ImmutableId';                                         options = @('set')                      }
        @{ name = 'InactiveMailbox';                                     options = @('set')                      }
        @{ name = 'InactiveMailboxRetireTime';                                                                   }
        @{ name = 'IncludeInGarbageCollection';                                                                  }
        @{ name = 'InPlaceHolds';                                                                                }
        @{ name = 'IsDirSynced';                                                                                 }
        @{ name = 'IsExcludedFromServingHierarchy';                      options = @('set')                      }
        @{ name = 'IsInactiveMailbox';                                                                           }
        @{ name = 'IsLinked';                                                                                    }
        @{ name = 'IsMachineToPersonTextMessagingEnabled';                                                       }
        @{ name = 'IsMailboxEnabled';                                                                            }
        @{ name = 'IsPersonToPersonTextMessagingEnabled';                                                        }
        @{ name = 'IsResource';                                                                                  }
        @{ name = 'IsRootPublicFolderMailbox';                                                                   }
        @{ name = 'IsShared';                                                                                    }
        @{ name = 'IsSoftDeletedByDisable';                                                                      }
        @{ name = 'IsSoftDeletedByRemove';                                                                       }
        @{ name = 'IssueWarningQuota';                                   options = @('set')                      }
        @{ name = 'IsValid';                                                                                     }
        @{ name = 'JournalArchiveAddress';                               options = @('set')                      }
        @{ name = 'Languages';                                           options = @('set')                      }
        @{ name = 'LastExchangeChangedTime';                                                                     }
        @{ name = 'LegacyExchangeDN';                                                                            }
        @{ name = 'LitigationHoldDate';                                  options = @('set')                      }
        @{ name = 'LitigationHoldDuration';                              options = @('set')                      }
        @{ name = 'LitigationHoldEnabled';                               options = @('set')                      }
        @{ name = 'LitigationHoldOwner';                                 options = @('set')                      }
        @{ name = 'MailboxContainerGuid';                                                                        }
        @{ name = 'MailboxLocations';                                                                            }
        @{ name = 'MailboxMoveBatchName';                                                                        }
        @{ name = 'MailboxMoveFlags';                                                                            }
        @{ name = 'MailboxMoveRemoteHostName';                                                                   }
        @{ name = 'MailboxMoveSourceMDB';                                                                        }
        @{ name = 'MailboxMoveStatus';                                                                           }
        @{ name = 'MailboxMoveTargetMDB';                                                                        }
        @{ name = 'MailboxPlan';                                                                                 }
        @{ name = 'MailboxProvisioningConstraint';                                                               }
        @{ name = 'MailboxProvisioningPreferences';                                                              }
        @{ name = 'MailboxRegion';                                       options = @('set')                      }
        @{ name = 'MailboxRegionLastUpdateTime';                                                                 }
        @{ name = 'MailboxRelease';                                                                              }
        @{ name = 'MailTip';                                             options = @('set')                      }
        @{ name = 'MailTipTranslations';                                 options = @('set')                      }
        @{ name = 'ManagedFolderMailboxPolicy';                          options = @('enable', 'set')            }
        @{ name = 'MaxReceiveSize';                                      options = @('set')                      }
        @{ name = 'MaxSendSize';                                         options = @('set')                      }
        @{ name = 'MessageCopyForSendOnBehalfEnabled';                   options = @('set')                      }
        @{ name = 'MessageCopyForSentAsEnabled';                         options = @('set')                      }
        @{ name = 'MessageCopyForSMTPClientSubmissionEnabled';           options = @('set')                      }
        @{ name = 'MessageTrackingReadStatusEnabled';                                                            }
        @{ name = 'MicrosoftOnlineServicesID';                           options = @('set')                      }
        @{ name = 'ModeratedBy';                                         options = @('set')                      }
        @{ name = 'ModerationEnabled';                                   options = @('set')                      }
        @{ name = 'Name';                                                options = @('set')                      }
        @{ name = 'NetID';                                                                                       }
        @{ name = 'NonCompliantDevices';                                 options = @('set')                      }
        @{ name = 'ObjectCategory';                                                                              }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'Office';                                              options = @('set')                      }
        @{ name = 'OrganizationalUnit';                                                                          }
        @{ name = 'OrganizationId';                                                                              }
        @{ name = 'OriginatingServer';                                                                           }
        @{ name = 'OrphanSoftDeleteTrackingTime';                                                                }
        @{ name = 'PersistedCapabilities';                                                                       }
        @{ name = 'PitrCopyIntervalInSeconds';                           options = @('set')                      }
        @{ name = 'PitrEnabled';                                         options = @('set')                      }
        @{ name = 'PoliciesExcluded';                                                                            }
        @{ name = 'PoliciesIncluded';                                                                            }
        @{ name = 'ProhibitSendQuota';                                   options = @('set')                      }
        @{ name = 'ProhibitSendReceiveQuota';                            options = @('set')                      }
        @{ name = 'ProtocolSettings';                                                                            }
        @{ name = 'ProvisionedForOfficeGraph';                           options = @('set')                      }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'QueryBaseDNRestrictionEnabled';                                                               }
        @{ name = 'RecipientLimits';                                     options = @('set')                      }
        @{ name = 'RecipientType';                                                                               }
        @{ name = 'RecipientTypeDetails';                                                                        }
        @{ name = 'RecalculateInactiveMailbox';                          options = @('set')                      }
        @{ name = 'ReconciliationId';                                                                            }
        @{ name = 'RejectMessagesFrom';                                  options = @('set')                      }
        @{ name = 'RejectMessagesFromDLMembers';                         options = @('set')                      }
        @{ name = 'RejectMessagesFromSendersOrMembers';                  options = @('set')                      }
        @{ name = 'RemoteAccountPolicy';                                                                         }
        @{ name = 'RemoveDelayHoldApplied';                              options = @('set')                      }
        @{ name = 'RemoveDelayReleaseHoldApplied';                       options = @('set')                      }
        @{ name = 'RemoveDisabledArchive';                               options = @('set')                      }
        @{ name = 'RemoveMailboxProvisioningConstraint';                 options = @('set')                      }
        @{ name = 'RemoveOrphanedHolds';                                 options = @('set')                      }
        @{ name = 'RequireSenderAuthenticationEnabled';                  options = @('set')                      }
        @{ name = 'ResourceCapacity';                                    options = @('set')                      }
        @{ name = 'ResourceCustom';                                      options = @('set')                      }
        @{ name = 'ResourceType';                                                                                }
        @{ name = 'RetainDeletedItemsFor';                               options = @('set')                      }
        @{ name = 'RetentionComment';                                    options = @('set')                      }
        @{ name = 'RetentionHoldEnabled';                                options = @('set')                      }
        @{ name = 'RetentionPolicy';                                     options = @('enable', 'set')            }
        @{ name = 'RetentionUrl';                                        options = @('set')                      }
        @{ name = 'RoleAssignmentPolicy';                                options = @('enable', 'set')            }
        @{ name = 'RoomMailboxAccountEnabled';                                                                   }
        @{ name = 'RulesQuota';                                          options = @('set')                      }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'SchedulerAssistant';                                  options = @('set')                      }
        @{ name = 'SendModerationNotifications';                         options = @('set')                      }
        @{ name = 'ServerLegacyDN';                                                                              }
        @{ name = 'ServerName';                                                                                  }
        @{ name = 'SharingPolicy';                                       options = @('set')                      }
        @{ name = 'SiloName';                                                                                    }
        @{ name = 'SimpleDisplayName';                                   options = @('set')                      }
        @{ name = 'SingleItemRecoveryEnabled';                           options = @('set')                      }
        @{ name = 'SKUAssigned';                                                                                 }
        @{ name = 'SourceAnchor';                                                                                }
        @{ name = 'StartDateForRetentionHold';                           options = @('set')                      }
        @{ name = 'StsRefreshTokensValidFrom';                           options = @('set')                      }
        @{ name = 'UMEnabled';                                                                                   }
        @{ name = 'UpdateEnforcedTimestamp';                             options = @('set')                      }
        @{ name = 'UnifiedMailbox';                                                                              }
        @{ name = 'UsageLocation';                                                                               }
        @{ name = 'UseDatabaseQuotaDefaults';                            options = @('set')                      }
        @{ name = 'UseDatabaseRetentionDefaults';                        options = @('set')                      }
        @{ name = 'UserCertificate';                                     options = @('set')                      }
        @{ name = 'UserSMimeCertificate';                                options = @('set')                      }
        @{ name = 'WasInactiveMailbox';                                                                          }
        @{ name = 'WhenChanged';                                                                                 }
        @{ name = 'WhenChangedUTC';                                                                              }
        @{ name = 'WhenCreated';                                                                                 }
        @{ name = 'WhenCreatedUTC';                                                                              }
        @{ name = 'WhenMailboxCreated';                                                                          }
        @{ name = 'WhenSoftDeleted';                                                                             }
        @{ name = 'WindowsEmailAddress';                                 options = @('set')                      }
        @{ name = 'WindowsLiveID';                                                                               }
    )
}


# Default properties and IDM properties are the same
foreach ($class in $Properties.Keys) {
    foreach ($e in $Properties.$class) {
        if (!$e.options) { $e.options = @() }
        if ($e.options.Contains('default')) { $e.options += 'idm' }
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
            $properties = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('key') }).name
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
                @{ name = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.CASMailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
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

        $key = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('key') }).name

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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('enable') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
            $properties = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name
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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('disable') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('permissionAdd') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('permissionRemove') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
                            foreach ($opt in $_.options) {
                                if ($opt -notin @('default', 'idm', 'key')) { continue }

                                if ($opt -eq 'idm') {
                                    $opt.Toupper()
                                }
                                else {
                                    $opt.Substring(0,1).Toupper() + $opt.Substring(1)
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
            value = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
    )
}
