#
# Microsoft Exchange Online.ps1 - IDM System PowerShell Script for Microsoft Exchange Online Services.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#

# Resolve any potential TLS issues
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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


# RPS (New-PSSession) is no longer supported in ExchangeOnline and RPS will be deprecated through the second half of 2023
# ExchangeManagement v3 is now required to support certificate authentication using Connect-ExchangeOnline cmdlet
# https://techcommunity.microsoft.com/t5/exchange-team-blog/deprecation-of-remote-powershell-in-exchange-online-re-enabling/ba-p/3779692

$Global:EXOManagementMinVersion = '3.0.0.0'
$Global:Mailboxes = [System.Collections.ArrayList]@()
$Global:DistributionGroups = [System.Collections.ArrayList]@()
$Global:ModuleStatus = '<b><div class="alert alert-danger" role="alert">Exchange Online PowerShell Module is not installed.</div></b>'

$Global:Module = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement'
if (!$Global:Module) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $nugetProvider = Get-PackageProvider | Where-Object { $_.Name -eq 'NuGet' }

    if (-not $nugetProvider -or $nugetProvider.Version -lt [version]'2.8.5.201') {
        Install-PackageProvider -Name 'NuGet' -MinimumVersion '2.8.5.201' -Scope 'CurrentUser' -Force 
    }

    Install-Module -Name 'ExchangeOnlineManagement' -Scope 'CurrentUser' -Force -AllowClobber
    $Global:Module = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement'
} 

# Check Module after install
if ($Global:Module) {
    $Global:ModuleStatus = "<b><div class=`"alert alert-success`" role=`"alert`">Exchange Online PowerShell Module $($Global:Module[0].Version) is installed.</div></b>"
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

    Log verbose "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'ModuleStatus'
                type = 'text'
                label = 'Module Status'
                text = $Global:ModuleStatus
            }
            @{
                name = 'AppId'
                type = 'textbox'
                label = 'Application ID'
                value = ''
            }
            @{
                name = 'Organization'
                type = 'textbox'
                label = 'Organization'
                value = ''
            }
            @{
                name = 'certificate'
                type = 'textbox'
                label = 'Certificate name'
                value = ''
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
                value = 1
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

    Log verbose "Done"
}


function Idm-OnUnload {
    Close-MsExchangeSession
}


#
# CRUD functions
#

$Properties = @{
    CASMailbox = @(
        @{ name = 'ActiveSyncAllowedDeviceIDs';                 options = @('set')              }
        @{ name = 'ActiveSyncBlockedDeviceIDs';                 options = @('set')              }
        @{ name = 'ActiveSyncDebugLogging';                     options = @('set')              }
        @{ name = 'ActiveSyncEnabled';                          options = @('default','set')    }
        @{ name = 'ActiveSyncMailboxPolicy';                    options = @('set')              }
        @{ name = 'ActiveSyncMailboxPolicyIsDefaulted';                                         }
        @{ name = 'ActiveSyncSuppressReadReceipt';              options = @('set')              }
        @{ name = 'DisplayName';                                options = @('default','set')    }
        @{ name = 'DistinguishedName';                                                          }
        @{ name = 'ECPEnabled';                                                                 }
        @{ name = 'EmailAddresses';                             options = @('default','set')    }
        @{ name = 'EwsAllowEntourage';                          options = @('set')              }
        @{ name = 'EwsAllowList';                               options = @('set')              }
        @{ name = 'EwsAllowMacOutlook';                         options = @('set')              }
        @{ name = 'EwsAllowOutlook';                            options = @('set')              }
        @{ name = 'EwsApplicationAccessPolicy';                 options = @('set')              }
        @{ name = 'EwsBlockList';                               options = @('set')              }
        @{ name = 'EwsEnabled';                                 options = @('set')              }
        @{ name = 'ExchangeObjectId';                                                           }
        @{ name = 'ExchangeVersion';                                                            }
        @{ name = 'ExternalDirectoryObjectId';                                                  }
        @{ name = 'Guid';                                       options = @('default','key')    }
        @{ name = 'HasActiveSyncDevicePartnership';                                             }
        @{ name = 'Identity';                                   options = @('default')          }
        @{ name = 'ImapEnabled';                                options = @('default','set')    }
        @{ name = 'ImapEnableExactRFC822Size';                                                  }
        @{ name = 'ImapForceICalForCalendarRetrievalOption';    options = @('set')              }
        @{ name = 'ImapMessagesRetrievalMimeFormat';            options = @('set')              }
        @{ name = 'ImapSuppressReadReceipt';                    options = @('set')              }
        @{ name = 'ImapUseProtocolDefaults';                    options = @('set')              }
        @{ name = 'IsOptimizedForAccessibility';                options = @('set')              }
        @{ name = 'LegacyExchangeDN';                                                           }
        @{ name = 'LinkedMasterAccount';                        options = @('default')          }
        @{ name = 'MacOutlookEnabled';                          options = @('set')              }
        @{ name = 'MAPIBlockOutlookExternalConnectivity';                                       }
        @{ name = 'MAPIBlockOutlookNonCachedMode';                                              }
        @{ name = 'MAPIBlockOutlookRpcHttp';                                                    }
        @{ name = 'MAPIBlockOutlookVersions';                                                   }
        @{ name = 'MAPIEnabled';                                options = @('set')              }
        @{ name = 'MapiHttpEnabled';                            options = @('set')              }
        @{ name = 'Name';                                       options = @('default')          }
        @{ name = 'ObjectCategory';                                                             }
        @{ name = 'ObjectClass';                                                                }
        @{ name = 'OrganizationId';                                                             }
        @{ name = 'OutlookMobileEnabled';                       options = @('set')              }
        @{ name = 'OWAEnabled';                                 options = @('default','set')    }
        @{ name = 'OWAforDevicesEnabled';                       options = @('set')              }
        @{ name = 'OwaMailboxPolicy';                           options = @('set')              }
        @{ name = 'PopEnabled';                                 options = @('default','set')    }
        @{ name = 'PopEnableExactRFC822Size';                                                   }
        @{ name = 'PopForceICalForCalendarRetrievalOption';     options = @('set')              }
        @{ name = 'PopMessageDeleteEnabled';                                                    }
        @{ name = 'PopMessagesRetrievalMimeFormat';             options = @('set')              }
        @{ name = 'PopSuppressReadReceipt';                     options = @('set')              }
        @{ name = 'PopUseProtocolDefaults';                     options = @('set')              }
        @{ name = 'PrimarySmtpAddress';                         options = @('default')          }
        @{ name = 'PublicFolderClientAccess';                   options = @('set')              }
        @{ name = 'SamAccountName';                             options = @('default')          }
        @{ name = 'ServerLegacyDN';                                                             }
        @{ name = 'ShowGalAsDefaultView';                       options = @('set')              }
        @{ name = 'SmtpClientAuthenticationDisabled';           options = @('set')              }
        @{ name = 'UniversalOutlookEnabled';                    options = @('set')              }
        @{ name = 'WhenChanged';                                                                }
        @{ name = 'WhenChangedUTC';                                                             }
        @{ name = 'WhenCreated';                                                                }
        @{ name = 'WhenCreatedUTC';                                                             }
    )
    DistributionGroup = @(
        @{ name = 'AcceptMessagesOnlyFrom';                     options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';            options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';     options = @('set')                      }
        @{ name = 'AddressListMembership';                                                              }
        @{ name = 'AdministrativeUnits';                                                                }
        @{ name = 'Alias';                                      options = @('default','create','set')   }
        @{ name = 'ArbitrationMailbox';                         options = @('create','set')             }
        @{ name = 'BccBlocked';                                 options = @('create','set')             }
        @{ name = 'BypassModerationFromSendersOrMembers';       options = @('create','set')                     }
        @{ name = 'BypassNestedModerationEnabled';              options = @('set')                      }
        @{ name = 'CustomAttribute1';                           options = @('set')                      }
        @{ name = 'CustomAttribute10';                          options = @('set')                      }
        @{ name = 'CustomAttribute11';                          options = @('set')                      }
        @{ name = 'CustomAttribute12';                          options = @('set')                      }
        @{ name = 'CustomAttribute13';                          options = @('set')                      }
        @{ name = 'CustomAttribute14';                          options = @('set')                      }
        @{ name = 'CustomAttribute15';                          options = @('set')                      }
        @{ name = 'CustomAttribute2';                           options = @('set')                      }
        @{ name = 'CustomAttribute3';                           options = @('set')                      }
        @{ name = 'CustomAttribute4';                           options = @('set')                      }
        @{ name = 'CustomAttribute5';                           options = @('set')                      }
        @{ name = 'CustomAttribute6';                           options = @('set')                      }
        @{ name = 'CustomAttribute7';                           options = @('set')                      }
        @{ name = 'CustomAttribute8';                           options = @('set')                      }
        @{ name = 'CustomAttribute9';                           options = @('set')                      }
        @{ name = 'Description';                                options = @('create','set')             }
        @{ name = 'DisplayName';                                options = @('default','create','set')   }
        @{ name = 'DistinguishedName';                                                                  }
        @{ name = 'EmailAddresses';                             options = @('set')                      }
        @{ name = 'EmailAddressPolicyEnabled';                                                          }
        @{ name = 'ExchangeObjectId';                                                                   }
        @{ name = 'ExchangeVersion';                                                                    }
        @{ name = 'ExtensionCustomAttribute1';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute2';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute3';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute4';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute5';                  options = @('set')                      }
        @{ name = 'GrantSendOnBehalfTo';                        options = @('set')                      }
        @{ name = 'GroupType';                                                                          }
        @{ name = 'Guid';                                       options = @('default','key')            }       
        @{ name = 'HiddenFromAddressListsEnabled';              options = @('set')                      }
        @{ name = 'HiddenGroupMembershipEnabled';               options = @('create','set')             }
        @{ name = 'Id';                                         options = @('default')                  }
        @{ name = 'Identity';                                                                           }
        @{ name = 'IsDirSynced';                                                                        }
        @{ name = 'IsValid';                                                                            }
        @{ name = 'LastExchangeChangedTime';                                                            }
        @{ name = 'LegacyExchangeDN';                                                                   }
        @{ name = 'MailTip';                                    options = @('set')                      }
        @{ name = 'MailTipTranslations';                        options = @('set')                      }
        @{ name = 'ManagedBy';                                  options = @('create','set')             }
        @{ name = 'MaxReceiveSize';                             options = @('set')                      }
        @{ name = 'MaxSendSize';                                options = @('set')                      }
        @{ name = 'MemberDepartRestriction';                    options = @('create','set')             }
        @{ name = 'MemberJoinRestriction';                      options = @('create','set')             }
        @{ name = 'MigrationToUnifiedGroupInProgress';                                                  }
        @{ name = 'ModeratedBy';                                options = @('create','set')             }
        @{ name = 'ModerationEnabled';                          options = @('create','set')             }
        @{ name = 'Name';                                       options = @('create','set')             }
        @{ name = 'ObjectCategory';                                                                     }
        @{ name = 'ObjectClass';                                                                        }
        @{ name = 'OrganizationalUnit';                                                                 }
        @{ name = 'OrganizationalUnitRoot';                                                             }
        @{ name = 'OrganizationId';                                                                     }
        @{ name = 'OriginatingServer';                                                                  }
        @{ name = 'PoliciesExcluded';                                                                   }
        @{ name = 'PoliciesIncluded';                                                                   }
        @{ name = 'PrimarySmtpAddress';                         options = @('default','create','set')   }
        @{ name = 'RecipientType';                                                                      }
        @{ name = 'RecipientTypeDetails';                                                               }
        @{ name = 'RejectMessagesFrom';                         options = @('set')                      }
        @{ name = 'RejectMessagesFromDLMembers';                options = @('set')                      }
        @{ name = 'RejectMessagesFromSendersOrMembers';         options = @('set')                      }
        @{ name = 'ReportToManagerEnabled';                                                             }
        @{ name = 'ReportToOriginatorEnabled';                                                          }
        @{ name = 'SamAccountName';                             options = @('create','set')             }
        @{ name = 'SendModerationNotifications';                options = @('create','set')             }
        @{ name = 'SendOofMessageToOriginatorEnabled';          options = @('set')                      }
        @{ name = 'Type';                                       options = @('create')                   }
        @{ name = 'UMDtmfMap';                                                                          }
        @{ name = 'WhenChanged';                                                                        }
        @{ name = 'WhenChangedUTC';                                                                     }
        @{ name = 'WhenCreated';                                                                        }
        @{ name = 'WhenCreatedUTC';                                                                     }
        @{ name = 'WindowsEmailAddress';                        options = @('set')                      }
    )
    DistributionGroupMember = @(
        @{ name = 'GroupGuid';                      options = @('default','set')        }
        @{ name = 'Guid';           options = @('default','set')                        }
        @{ name = 'RecipientType';                                                      }
    )
    Mailbox = @(
        @{ name = 'AcceptMessagesOnlyFrom';                     options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';            options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';     options = @('set')                      }
        @{ name = 'AccountDisabled';                            options = @('default','set')            }
        @{ name = 'AddressBookPolicy';                          options = @('enable','set')             }
        @{ name = 'AddressListMembership';                                                              }
        @{ name = 'AdministrativeUnits';                                                                }
        @{ name = 'AggregatedMailboxGuids';                                                             }
        @{ name = 'Alias';                                      options = @('default','enable','set')   }
        @{ name = 'AntispamBypassEnabled';                                                              }
        @{ name = 'ArbitrationMailbox';                                                                 }
        @{ name = 'ArchiveDatabase';                            options = @('enable')                   }
        @{ name = 'ArchiveDomain';                              options = @('enable')                   }
        @{ name = 'ArchiveGuid';                                options = @('enable')                   }
        @{ name = 'ArchiveName';                                options = @('set')                      }
        @{ name = 'ArchiveQuota';                                                                       }
        @{ name = 'ArchiveRelease';                                                                     }
        @{ name = 'ArchiveState';                               options = @('set')                      }
        @{ name = 'ArchiveStatus';                                                                      }
        @{ name = 'ArchiveWarningQuota';                                                                }
        @{ name = 'AuditAdmin';                                 options = @('set')                      }
        @{ name = 'AuditDelegate';                              options = @('set')                      }
        @{ name = 'AuditEnabled';                               options = @('set')                      }
        @{ name = 'AuditLogAgeLimit';                           options = @('set')                      }
        @{ name = 'AuditOwner';                                 options = @('set')                      }
        @{ name = 'AutoExpandingArchiveEnabled';                                                        }
        @{ name = 'BypassModerationFromSendersOrMembers';       options = @('set')                      }
        @{ name = 'CalendarLoggingQuota';                                                               }
        @{ name = 'CalendarRepairDisabled';                     options = @('set')                      }
        @{ name = 'CalendarVersionStoreDisabled';               options = @('set')                      }
        @{ name = 'ComplianceTagHoldApplied';                                                           }
        @{ name = 'CustomAttribute1';                           options = @('set')                      }
        @{ name = 'CustomAttribute10';                          options = @('set')                      }
        @{ name = 'CustomAttribute11';                          options = @('set')                      }
        @{ name = 'CustomAttribute12';                          options = @('set')                      }
        @{ name = 'CustomAttribute13';                          options = @('set')                      }
        @{ name = 'CustomAttribute14';                          options = @('set')                      }
        @{ name = 'CustomAttribute15';                          options = @('set')                      }
        @{ name = 'CustomAttribute2';                           options = @('set')                      }
        @{ name = 'CustomAttribute3';                           options = @('set')                      }
        @{ name = 'CustomAttribute4';                           options = @('set')                      }
        @{ name = 'CustomAttribute5';                           options = @('set')                      }
        @{ name = 'CustomAttribute6';                           options = @('set')                      }
        @{ name = 'CustomAttribute7';                           options = @('set')                      }
        @{ name = 'CustomAttribute8';                           options = @('set')                      }
        @{ name = 'CustomAttribute9';                           options = @('set')                      }
        @{ name = 'Database';                                                                           }
        @{ name = 'DataEncryptionPolicy';                                                               }
        @{ name = 'DefaultAuditSet';                            options = @('set')                      }
        @{ name = 'DefaultPublicFolderMailbox';                 options = @('set')                      }
        @{ name = 'DelayHoldApplied';                                                                   }
        @{ name = 'DeliverToMailboxAndForward';                 options = @('set')                      }
        @{ name = 'DisabledArchiveDatabase';                                                            }
        @{ name = 'DisabledArchiveGuid';                                                                }
        @{ name = 'DisabledMailboxLocations';                                                           }
        @{ name = 'DisplayName';                                options = @('default','enable','set')   }
        @{ name = 'DistinguishedName';                                                                  }
        @{ name = 'DowngradeHighPriorityMessagesEnabled';                                               }
        @{ name = 'EffectivePublicFolderMailbox';                                                       }
        @{ name = 'ElcProcessingDisabled';                      options = @('set')                      }
        @{ name = 'EmailAddresses';                             options = @('default','set')            }
        @{ name = 'EmailAddressPolicyEnabled';                                                          }
        @{ name = 'EndDateForRetentionHold';                    options = @('set')                      }
        @{ name = 'ExchangeGuid';                                                                       }
        @{ name = 'ExchangeObjectId';                                                                   }
        @{ name = 'ExchangeSecurityDescriptor';                                                         }
        @{ name = 'ExchangeUserAccountControl';                                                         }
        @{ name = 'ExchangeVersion';                                                                    }
        @{ name = 'ExtensionCustomAttribute1';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute2';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute3';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute4';                  options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute5';                  options = @('set')                      }
        @{ name = 'Extensions';                                                                         }
        @{ name = 'ExternalDirectoryObjectId';                                                          }
        @{ name = 'ExternalOofOptions';                         options = @('set')                      }
        @{ name = 'ForwardingAddress';                          options = @('set')                      }
        @{ name = 'ForwardingSmtpAddress';                      options = @('set')                      }
        @{ name = 'GeneratedOfflineAddressBooks';                                                       }
        @{ name = 'GrantSendOnBehalfTo';                        options = @('set')                      }
        @{ name = 'Guid';                                       options = @('default')                  }
        @{ name = 'HasPicture';                                                                         }
        @{ name = 'HasSnackyAppData';                                                                   }
        @{ name = 'HasSpokenName';                                                                      }
        @{ name = 'HiddenFromAddressListsEnabled';              options = @('set')                      }
        @{ name = 'Id';                                         options = @('default')                  }
        @{ name = 'Identity';                                   options = @('default','key')            }
        @{ name = 'ImListMigrationCompleted';                                                           }
        @{ name = 'ImmutableId';                                options = @('set')                      }
        @{ name = 'InactiveMailboxRetireTime';                                                          }
        @{ name = 'IncludeInGarbageCollection';                                                         }
        @{ name = 'InPlaceHolds';                                                                       }
        @{ name = 'IsDirSynced';                                                                        }
        @{ name = 'IsExcludedFromServingHierarchy';             options = @('set')                      }
        @{ name = 'IsHierarchyReady';                                                                   }
        @{ name = 'IsHierarchySyncEnabled';                                                             }
        @{ name = 'IsInactiveMailbox';                          options = @('set')                      }
        @{ name = 'IsLinked';                                                                           }
        @{ name = 'IsMachineToPersonTextMessagingEnabled';                                              }
        @{ name = 'IsMailboxEnabled';                                                                   }
        @{ name = 'IsMonitoringMailbox';                                                                }
        @{ name = 'IsPersonToPersonTextMessagingEnabled';                                               }
        @{ name = 'IsResource';                                                                         }
        @{ name = 'IsRootPublicFolderMailbox';                                                          }
        @{ name = 'IsShared';                                                                           }
        @{ name = 'IsSoftDeletedByDisable';                                                             }
        @{ name = 'IsSoftDeletedByRemove';                                                              }
        @{ name = 'IssueWarningQuota';                          options = @('set')                      }
        @{ name = 'JournalArchiveAddress';                      options = @('set')                      }
        @{ name = 'Languages';                                  options = @('set')                      }
        @{ name = 'LastExchangeChangedTime';                                                            }
        @{ name = 'LegacyExchangeDN';                                                                   }
        @{ name = 'LinkedMasterAccount';                                                                }
        @{ name = 'LitigationHoldDate';                         options = @('set')                      }
        @{ name = 'LitigationHoldDuration';                     options = @('set')                      }
        @{ name = 'LitigationHoldEnabled';                      options = @('set')                      }
        @{ name = 'LitigationHoldOwner';                        options = @('set')                      }
        @{ name = 'MailboxContainerGuid';                                                               }
        @{ name = 'MailboxLocations';                                                                   }
        @{ name = 'MailboxMoveBatchName';                                                               }
        @{ name = 'MailboxMoveFlags';                                                                   }
        @{ name = 'MailboxMoveRemoteHostName';                                                          }
        @{ name = 'MailboxMoveSourceMDB';                                                               }
        @{ name = 'MailboxMoveStatus';                                                                  }
        @{ name = 'MailboxMoveTargetMDB';                                                               }
        @{ name = 'MailboxPlan';                                                                        }
        @{ name = 'MailboxProvisioningConstraint';                                                      }
        @{ name = 'MailboxProvisioningPreferences';                                                     }
        @{ name = 'MailboxRegion';                              options = @('set')                      }
        @{ name = 'MailboxRegionLastUpdateTime';                                                        }
        @{ name = 'MailboxRelease';                                                                     }
        @{ name = 'MailTip';                                    options = @('set')                      }
        @{ name = 'MailTipTranslations';                        options = @('set')                      }
        @{ name = 'ManagedFolderMailboxPolicy';                 options = @('enable','set')             }
        @{ name = 'MaxBlockedSenders';                                                                  }
        @{ name = 'MaxReceiveSize';                             options = @('set')                      }
        @{ name = 'MaxSafeSenders';                                                                     }
        @{ name = 'MaxSendSize';                                options = @('set')                      }
        @{ name = 'MessageCopyForSendOnBehalfEnabled';          options = @('set')                      }
        @{ name = 'MessageCopyForSentAsEnabled';                options = @('set')                      }
        @{ name = 'MessageCopyForSMTPClientSubmissionEnabled';  options = @('set')                      }
        @{ name = 'MessageRecallProcessingEnabled';                                                     }
        @{ name = 'MessageTrackingReadStatusEnabled';                                                   }
        @{ name = 'MicrosoftOnlineServicesID';                  options = @('set')                      }
        @{ name = 'ModeratedBy';                                options = @('set')                      }
        @{ name = 'ModerationEnabled';                          options = @('set')                      }
        @{ name = 'Name';                                       options = @('set')                      }
        @{ name = 'NetID';                                                                              }
        @{ name = 'NonCompliantDevices';                        options = @('set')                      }
        @{ name = 'ObjectCategory';                                                                     }
        @{ name = 'ObjectClass';                                                                        }
        @{ name = 'Office';                                     options = @('set')                      }
        @{ name = 'OfflineAddressBook';                                                                 }
        @{ name = 'OrganizationalUnit';                                                                 }
        @{ name = 'OrganizationId';                                                                     }
        @{ name = 'OrphanSoftDeleteTrackingTime';                                                       }
        @{ name = 'PersistedCapabilities';                                                              }
        @{ name = 'PoliciesExcluded';                                                                   }
        @{ name = 'PoliciesIncluded';                                                                   }
        @{ name = 'PrimarySmtpAddress';                         options = @('default','set')            }
        @{ name = 'ProhibitSendQuota';                          options = @('set')                      }
        @{ name = 'ProhibitSendReceiveQuota';                   options = @('set')                      }
        @{ name = 'ProtocolSettings';                                                                   }
        @{ name = 'QueryBaseDN';                                                                        }
        @{ name = 'QueryBaseDNRestrictionEnabled';                                                      }
        @{ name = 'RecipientLimits';                            options = @('set')                      }
        @{ name = 'RecipientType';                                                                      }
        @{ name = 'RecipientTypeDetails';                                                               }
        @{ name = 'ReconciliationId';                                                                   }
        @{ name = 'RecoverableItemsQuota';                                                              }
        @{ name = 'RecoverableItemsWarningQuota';                                                       }
        @{ name = 'RejectMessagesFrom';                         options = @('set')                      }
        @{ name = 'RejectMessagesFromDLMembers';                options = @('set')                      }
        @{ name = 'RejectMessagesFromSendersOrMembers';         options = @('set')                      }
        @{ name = 'RemoteAccountPolicy';                                                                }
        @{ name = 'RemoteRecipientType';                                                                }
        @{ name = 'RequireSenderAuthenticationEnabled';         options = @('set')                      }
        @{ name = 'ResetPasswordOnNextLogon';                   options = @('set')                      }
        @{ name = 'ResourceCapacity';                           options = @('set')                      }
        @{ name = 'ResourceCustom';                             options = @('set')                      }
        @{ name = 'ResourceType';                                                                       }
        @{ name = 'RetainDeletedItemsFor';                      options = @('set')                      }
        @{ name = 'RetainDeletedItemsUntilBackup';                                                      }
        @{ name = 'RetentionComment';                           options = @('set')                      }
        @{ name = 'RetentionHoldEnabled';                       options = @('set')                      }
        @{ name = 'RetentionPolicy';                            options = @('set')                      }
        @{ name = 'RetentionUrl';                               options = @('set')                      }
        @{ name = 'RoleAssignmentPolicy';                       options = @('enable','set')             }
        @{ name = 'RoomMailboxAccountEnabled';                                                          }
        @{ name = 'RulesQuota';                                 options = @('set')                      }
        @{ name = 'SamAccountName';                             options = @('enable','set')             }
        @{ name = 'SCLDeleteEnabled';                                                                   }
        @{ name = 'SCLDeleteThreshold';                                                                 }
        @{ name = 'SCLJunkEnabled';                                                                     }
        @{ name = 'SCLJunkThreshold';                                                                   }
        @{ name = 'SCLQuarantineEnabled';                                                               }
        @{ name = 'SCLQuarantineThreshold';                                                             }
        @{ name = 'SCLRejectEnabled';                                                                   }
        @{ name = 'SCLRejectThreshold';                                                                 }
        @{ name = 'SendModerationNotifications';                options = @('set')                      }
        @{ name = 'ServerLegacyDN';                                                                     }
        @{ name = 'SharingPolicy';                              options = @('set')                      }
        @{ name = 'SiloName';                                                                           }
        @{ name = 'SimpleDisplayName';                          options = @('set')                      }
        @{ name = 'SingleItemRecoveryEnabled';                  options = @('set')                      }
        @{ name = 'SKUAssigned';                                                                        }
        @{ name = 'SourceAnchor';                                                                       }
        @{ name = 'StartDateForRetentionHold';                  options = @('set')                      }
        @{ name = 'StsRefreshTokensValidFrom';                  options = @('set')                      }
        @{ name = 'ThrottlingPolicy';                                                                   }
        @{ name = 'Type';                                       options = @('set')                      }
        @{ name = 'UMDtmfMap';                                                                          }
        @{ name = 'UMEnabled';                                                                          }
        @{ name = 'UnifiedMailbox';                                                                     }
        @{ name = 'UsageLocation';                                                                      }
        @{ name = 'UseDatabaseQuotaDefaults';                   options = @('set')                      }
        @{ name = 'UseDatabaseRetentionDefaults';               options = @('set')                      }
        @{ name = 'UserPrincipalName';                          options = @('default','enable','set')   }
        @{ name = 'WasInactiveMailbox';                                                                 }
        @{ name = 'WhenChanged';                                                                        }
        @{ name = 'WhenChangedUTC';                                                                     }
        @{ name = 'WhenCreated';                                                                        }
        @{ name = 'WhenCreatedUTC';                                                                     }
        @{ name = 'WhenMailboxCreated';                                                                 }
        @{ name = 'WhenSoftDeleted';                                                                    }
        @{ name = 'WindowsEmailAddress';                        options = @('set')                      }
        @{ name = 'WindowsLiveID';                                                                      }
    )
    MailboxAutoReplyConfiguration = @(
        @{ name = 'AutoDeclineFutureRequestsWhenOOF';   options = @('default','set')        }
        @{ name = 'AutoReplyState';                     options = @('default','set')        }
        @{ name = 'CreateOOFEvent';                     options = @('default','set')        }
        @{ name = 'DeclineAllEventsForScheduledOOF';    options = @('default','set')        }
        @{ name = 'DeclineEventsForScheduledOOF';       options = @('default','set')        }
        @{ name = 'DeclineMeetingMessage';              options = @('default','set')        }
        @{ name = 'DomainController';                   options = @('set')                  }
        @{ name = 'EventsToDeleteIDs';                                                      }
        @{ name = 'EndTime';                            options = @('default','set')        }
        @{ name = 'ExternalAudience';                   options = @('default','set')        }
        @{ name = 'ExternalMessage';                    options = @('default','set')        }
        @{ name = 'InternalMessage';                    options = @('default','set')        }
        @{ name = 'OOFEventSubject';                    options = @('set')                  }
        @{ name = 'StartTime';                          options = @('default','set')        }
        @{ name = 'Recipients';                                                             }
        @{ name = 'ReminderMinutesBeforeStart';                                             }
        @{ name = 'ReminderMessage';                                                        }
        @{ name = 'MailboxOwnerId';                     options = @('default')              }
        @{ name = 'Identity';                           options = @('default','key')        }
        @{ name = 'IsValid';                                                                }
        @{ name = 'ObjectState';                                                            }
        @{ name = 'ID';                                                                     }
        
    )
    MailboxPermission = @(
        @{ name = 'AccessRights';       options = @('default','add','remove')   }
        @{ name = 'Deny';               options = @('default')                  }
        @{ name = 'DomainController';   options = @('add')                      }
        @{ name = 'Identity';           options = @('default','add','remove')   }
        @{ name = 'InheritanceType';    options = @('default','add')            }
        @{ name = 'IsInherited';        options = @('default')                  }
        @{ name = 'User';               options = @('default','add','remove')   }
        @{ name = 'Owner';              options = @('add')                      }
        @{ name = 'ID';           options = @('default','key')            }
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

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'CASMailbox'

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class -CanFilter $true
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'unlimited'
            PropertySets = @('All')
        }
        
        if($function_params.filter.length -gt 0) {
            $call_params.Filter = $function_params.filter
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-exocasmailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud

            LogIO info "Get-EXOCasMailbox" -In @call_params
            
            # EXO cmdlets cannot be prefixed because "EXO" is effectively a prefix already
            Get-EXOCasMailbox @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}


function Idm-CASMailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'CASMailbox'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

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

    Log verbose "Done"
}

function Idm-DistributionGroupsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'DistributionGroup'

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class -CanFilter $true
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'unlimited'
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud
            
            LogIO info "Get-MsExchangeDistributionGroup" -In @call_params
            
            if($Global:DistributionGroups.count -gt 0) {
                Log verbose "Using cached distribution groups"
                $groups = $Global:DistributionGroups
            } else {
                # EXO cmdlets cannot be prefixed because "EXO" is effectively a prefix already
                $groups = Get-MsExchangeDistributionGroup @call_params

                # Push group GUIDs into a global collection
                $Global:DistributionGroups.Clear()
                foreach($grp in $groups) {
                    [void]$Global:DistributionGroups.Add( $grp )
                }
            }
            # Return Data
            $groups | Select-Object $properties

        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}

function Idm-DistributionGroupCreate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'DistributionGroup'
    
    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                $Global:Properties.$Class | Where-Object { $_.name -in @('Name','Type') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'mandatory' }
                }    
                $Global:Properties.$Class | Where-Object { $_.options.Contains('create') -and !$_.options.Contains('key') -and $_.name -notin @('Name','Type') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'optional' }
                }
                $Global:Properties.$Class | Where-Object { $_.options.Contains('key') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }
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

        $call_params = $function_params

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/new-distributiongroup?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "New-MsExchangeDistributionGroup" -In @call_params
                $rv = New-MsExchangeDistributionGroup @call_params
            LogIO info "New-MsExchangeDistributionGroup" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}

function Idm-DistributionGroupUpdate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'DistributionGroup'
    
    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }
                $Global:Properties.$Class | Where-Object { $_.options.Contains('set') -and !$_.options.Contains('key') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'optional' }
                }
                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }
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
        
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/set-distributiongroup?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Set-MsExchangeDistributionGroup" -In @call_params
                $rv = Set-MsExchangeDistributionGroup @call_params
            LogIO info "Set-MsExchangeDistributionGroup" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}

function Idm-DistributionGroupDelete {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'DistributionGroup'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('disable') } | ForEach-Object {
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

        $call_params = @{
            Identity = $function_params.$key
            Confirm  = $false   # Be non-interactive
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/remove-distributiongroup?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Remove-MsExchangeDistributionGroup" -In @call_params
                $rv = Remove-MsExchangeDistributionGroup @call_params
            LogIO info "Disable-MsExchangeDistributionGroup" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}

function Idm-DistributionGroupMembersRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )
    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    $Class = 'DistributionGroupMember'

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        # Check Cache State
        EvaluateCacheState -Type 'DistributionGroups'

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroupmember?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud           
            $i = $Global:DistributionGroups.count

            foreach($grp in $Global:DistributionGroups) {
                $call_params = @{
                    Identity = $grp.Guid
                    ResultSize = 'unlimited'
                }
                
                LogIO info "Get-MsExchangeDistributionGroupMember" -In @call_params
                $result = Get-MsExchangeDistributionGroupMember @call_params
                
                foreach($member in $result) {
                    [PSCustomObject]@{
                        GroupGuid = $grp.Guid
                        Guid = $member.Guid
                        RecipientType = $member.RecipientType
                    }
                }
                
                if(($i -= 1) % 100 -eq 0) {
                    Log debug ("[Progress][$($Class)] $($i) remaining distribution groups to search")
                }
            }
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}

function Idm-DistributionGroupMemberCreate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = "GroupGuid";  allowance = 'mandatory'  }
                @{ name = "Guid"; allowance = 'mandatory'  }
                @{ name = '*';      allowance = 'prohibited' }
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

        $call_params = @{
            Identity = $function_params["GroupGuid"]
            Member = $function_params["Guid"]
        }
        LogIO info "Add-MsExchangeDistributionGroupMember" -In @call_params
               Add-MsExchangeDistributionGroupMember @call_params -Confirm:$false >$null 2>&1
        
        $rv = [PSCustomObject]@{
            GroupGuid = $function_params["GroupGuid"]
            Guid = $function_params["Guid"]
        }
        LogIO info "Add-MsExchangeDistributionGroupMember" -Out $rv
        
        $rv
    }

    Log verbose "Done"
}

function Idm-DistributionGroupMemberDelete {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'delete'
            parameters = @(
                @{ name = "GroupGuid";  allowance = 'mandatory'  }
                @{ name = "Guid"; allowance = 'mandatory'  }
                @{ name = '*';      allowance = 'prohibited' }
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

        $call_params= @{
            Identity = $function_params["GroupGuid"]
            Member = $function_params["Guid"]
        }

        LogIO info "Remove-MsExchangeDistributionGroupMember" -In @call_params
            Remove-MsExchangeDistributionGroupMember @call_params -Confirm:$false >$null 2>&1
        LogIO info "Remove-MsExchangeDistributionGroupMember"
    }

    Log verbose "Done"
}

function Idm-MailboxesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'Mailbox'

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class -CanFilter $true
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'unlimited'
            PropertySets = @('All')
        }
        
        if($function_params.filter.length -gt 0) {
            $call_params.Filter = $function_params.filter
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        # Skip retrieval if already available, else return current dataset
        if($Global:Mailboxes.count -lt 1) {
            try {
                # https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailbox?view=exchange-ps
                #
                # Cmdlet availability:
                # v Cloud
                
                LogIO info "Get-EXOMailbox" -In @call_params
                
                # EXO cmdlets cannot be prefixed because "EXO" is effectively a prefix already
                $mailboxes = Get-EXOMailbox @call_params 
                $mailboxes | Select-Object $properties
                
                # Push mailbox GUIDs into a global collection
                $Global:Mailboxes.Clear()
                foreach($mb in $mailboxes) {
                    [void]$Global:Mailboxes.Add($mb)
                }
            }
            catch {
                Log error "Failed: $_"
                Write-Error $_
            }
        } else { 
            Log verbose "Using cached mailboxes"
            foreach($mbx in $Global:Mailboxes) {
                $mbx | Select-Object $properties
            }
        }
    }

    Log verbose "Done"
}


function Idm-MailboxEnable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'Mailbox'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('enable') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

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

    Log verbose "Done"
}


function Idm-MailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'Mailbox'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }
                @{ name = 'Type'; allowance = 'optional' }
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

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

            $rv = Get-EXOMailbox -Identity $function_params.$key
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}


function Idm-MailboxDisable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'Mailbox'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('disable') } | ForEach-Object {
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

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

    Log verbose "Done"
}


function Idm-MailboxAutoReplyConfigurationsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    
    $Class = 'MailboxAutoReplyConfiguration'

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'unlimited'
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        # Check Cache State
        EvaluateCacheState -Type 'Mailboxes'

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailboxautoreplyconfiguration?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud
            
            LogIO info "Get-MsExchangeMailboxAutoReplyConfiguration" -In @call_params
            
            $data = $Global:Mailboxes.GUID.GUID | Get-MsExchangeMailboxAutoReplyConfiguration @call_params
            
            $data | Select-Object $properties
            
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}


function Idm-MailboxAutoReplyConfigurationSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    
    $Class = 'MailboxAutoReplyConfiguration'
    
    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        
        $call_params = @{
            Identity = $function_params.$key
        }
        
        $function_params.Remove($key)
        
        $call_params += $function_params
        
        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailboxautoreplyconfiguration?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Set-MsExchangeMailboxAutoReplyConfiguration" -In @call_params
                $rv = Set-MsExchangeMailboxAutoReplyConfiguration @call_params
            LogIO info "Set-MsExchangeMailboxAutoReplyConfiguration" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}


function Idm-MailboxPermissionsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    
    $Class = 'MailboxPermission'

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class $Class
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'unlimited'
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        # Check Cache State
        EvaluateCacheState -Type 'Mailboxes'

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud
            
            LogIO info "Get-EXOMailboxPermission" -In @call_params

            $data = $Global:Mailboxes.GUID.GUID | Get-EXOMailboxPermission @call_params

            foreach ($item in $data) {
                # Convert the selected fields to JSON
                $json = $item | ConvertTo-Json -Depth 10 -Compress

                # Hash the JSON string
                $sha256 = [System.Security.Cryptography.SHA256]::Create()
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                $hash = $sha256.ComputeHash($bytes)
                $key = [BitConverter]::ToString($hash) -replace '-', ''

                # Create new object with hashed ID
                $newItem = [PSCustomObject]@{
                    ID = $key
                }

                # Add selected properties to the new object
                foreach ($prop in $properties) {
                    try {
                        $value = $item.$prop
                        $newItem | Add-Member -MemberType NoteProperty -Name $prop -Value $value
                    }
                    catch {
                        Log warn "Property '$prop' not found on item or failed to add: $_"
                    }
                }

                # Output the new object
                $newItem
            }

        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}


function Idm-MailboxPermissionAdd {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'MailboxPermission'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = 'Identity'; allowance = 'mandatory' }
            
                $Global:Properties.$Class | Where-Object { $_.options.Contains('key') -or !$_.options.Contains('add') } | ForEach-Object {
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

        $call_params += $function_params
        
        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Add-MsExchangeMailboxPermission" -In @call_params
                $rv = Add-MsExchangeMailboxPermission @call_params  
                $json = $rv | ConvertTo-Json -Depth 10 -Compress
                $sha256 = [System.Security.Cryptography.SHA256]::Create()
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                $hash = $sha256.ComputeHash($bytes)
                $key = [BitConverter]::ToString($hash) -replace '-', ''
                
                
                $returnObj = @{}
                foreach($prop in $rv.PSObject.Properties) {
                    $returnObj[$prop.Name] = $prop.Value
                }
                
                $returnObj['User'] = $function_params.User
                $returnObj["ID"] = $key
            LogIO info "Add-MsExchangeMailboxPermission" -Out $returnObj

            $returnObj
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log verbose "Done"
}


function Idm-MailboxPermissionRemove {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = 'MailboxPermission'
    
    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('remove') } | ForEach-Object {
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

        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name

        $call_params = @{
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

    Log verbose "Done"
}

#
# Helper functions
#

function Open-MsExchangeSession {
    param (
        [hashtable] $SystemParams
    )

    if ((Get-Module -ListAvailable -Name 'ExchangeOnlineManagement').Version -lt $Global:EXOManagementMinVersion) {
        throw "ExchangeOnlineManagement PowerShell Module version older than $($Global:EXOManagementMinVersion)"
    }

    # Use connection related parameters only
    $connection_params = [ordered]@{
        AppId        = $SystemParams.AppId
        Organization = $SystemParams.Organization
        Certificate  = $SystemParams.certificate
        PageSize     = $SystemParams.PageSize
    }

    $connection_string = ConvertTo-Json $connection_params -Compress -Depth 32

    if (Get-ConnectionInformation | ? { $_.State -eq 'Connected' }) {
        #Log debug "Reusing MsExchangePSSession"
    }
    else {
        Log verbose "Opening ExchangeOnline session '$connection_string'"

        $params = Copy-Object $connection_params
        $params.Certificate = Nim-GetCertificate $connection_params.certificate

        try {
            Connect-ExchangeOnline @params -Prefix 'MsExchange' -ShowBanner:$false
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }

        Log verbose "Done"
    }
}


function Close-MsExchangeSession {
    if (Get-ConnectionInformation | ? { $_.State -eq 'Connected' }) {
        Log verbose "Closing ExchangeOnline session"
        Disconnect-ExchangeOnline -Confirm:$false
        Log verbose "Done"
    }
}


function Get-ClassMetaData {
    param (
        [string] $SystemParams,
        [string] $Class,
        [switch] $CanFilter
    )

    @(
        if($CanFilter) {
            @{
                name = 'filter'
                type = 'textbox'
                label = 'Filter'
                description = '-Filter expression'
                value = ''
            }
        }
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

function EvaluateCacheState {
    param (
        [string] $Type
    )

    if( ($Type -eq 'Mailboxes' -or $Type -eq '*') -and $Global:Mailboxes.count -lt 1) {
        Log verbose "Refreshing Mailboxes Cache"
        Idm-MailboxesRead | Out-Null
    }

    if(($Type -eq 'DistributionGroups' -or $Type -eq '*') -and $Global:DistributionGroups.count -lt 1) {
        Log verbose "Refreshing Distribution Groups Cache"
        Idm-DistributionGroupsRead | Out-Null
    }
}

$configScenarios = @'
[{"name":"Default","description":"Default Configuration","version":"1.0","createTime":1738209486625,
"modifyTime":1744130247208,"name_values":[{"name":"AppId","value":null},{"name":"Organization","valu
e":null},{"name":"PageSize","value":null},{"name":"available_certificates","value":null},{"name":"ce
rtificate","value":null},{"name":"collections","value":["CASMailboxes","DistributionGroupMembers","D
istributionGroups","MailboxAutoReplyConfigurations","Mailboxes","MailboxPermissions"]},{"name":"nr_o
f_sessions","value":null},{"name":"organizational_unit","value":"*"},{"name":"sessions_idle_timeout"
,"value":null}],"collections":[{"col_name":"DistributionGroups","fields":[{"field_name":"Guid","fiel
d_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[]
,"reference":false,"ref_col_fields":[]},{"field_name":"Alias","field_type":"string","include":true,"
field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_field
s":[]},{"field_name":"DisplayName","field_type":"string","include":true,"field_format":"","field_sou
rce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Id","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"PrimarySmtpAddress","field_type":"string"
,"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fal
se,"ref_col_fields":[]},{"field_name":"HiddenFromAddressListsEnabled","field_type":"string","include
":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_c
ol_fields":[]},{"field_name":"HiddenGroupMembershipEnabled","field_type":"string","include":true,"fi
eld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields"
:[]},{"field_name":"IsDirSynced","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IsReadO
nly","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","r
ef_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IsFixedSize","field_type":"string",
"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fals
e,"ref_col_fields":[]},{"field_name":"IsSynchronized","field_type":"string","include":true,"field_fo
rmat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{
"field_name":"Keys","field_type":"string","include":true,"field_format":"","field_source":"data","ja
vascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Values","field_type"
:"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refer
ence":false,"ref_col_fields":[]},{"field_name":"SyncRoot","field_type":"string","include":true,"fiel
d_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[
]},{"field_name":"Count","field_type":"string","include":true,"field_format":"","field_source":"data
","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}],"key":"Guid","display":"Alias
","name_values":[{"name":"filter","value":""},{"name":"properties","value":["Alias","DisplayName","G
uid","HiddenFromAddressListsEnabled","HiddenGroupMembershipEnabled","Id","IsDirSynced","PrimarySmtpA
ddress"]}],"sys_nn":[],"source":"data"},{"col_name":"DistributionGroupMembers","fields":[{"field_nam
e":"GroupGuid","field_type":"string","include":true,"field_format":"","field_source":"data","javascr
ipt":"","ref_col":["DistributionGroups"],"reference":false,"ref_col_fields":[]},{"field_name":"Guid"
,"field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_c
ol":["DistributionGroups","Mailboxes"],"reference":false,"ref_col_fields":[]},{"field_name":"Recipie
ntType","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":""
,"ref_col":[],"reference":false,"ref_col_fields":[]}],"key":"","display":"GroupGuid","name_values":[
],"sys_nn":[{"field_a":"GroupGuid","col_a":"DistributionGroups","field_b":"Guid","col_b":"Mailboxes"
},{"field_a":"GroupGuid","col_a":"DistributionGroups","field_b":"Guid","col_b":"DistributionGroups"}
],"container":"GroupGuid","source":"data"},{"col_name":"Mailboxes","fields":[{"field_name":"Guid","f
ield_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col"
:[],"reference":false,"ref_col_fields":[]},{"field_name":"AccountDisabled","field_type":"string","in
clude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"
ref_col_fields":[]},{"field_name":"Alias","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"DisplayName","field_type":"string","include":true,"field_format":"","field_source":"data","javascr
ipt":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"EmailAddresses","field_ty
pe":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"re
ference":false,"ref_col_fields":[]},{"field_name":"Id","field_type":"string","include":true,"field_f
ormat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},
{"field_name":"PrimarySmtpAddress","field_type":"string","include":true,"field_format":"","field_sou
rce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"UserP
rincipalName","field_type":"string","include":true,"field_format":"","field_source":"data","javascri
pt":"","ref_col":[],"reference":true,"ref_col_fields":[]},{"field_name":"LastLoggedOnUserAccount","f
ield_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col"
:[],"reference":false,"ref_col_fields":[]},{"field_name":"LastLogoffTime","field_type":"string","inc
lude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"r
ef_col_fields":[]},{"field_name":"LastLogonTime","field_type":"string","include":true,"field_format"
:"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fiel
d_name":"AcceptMessagesOnlyFrom","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AcceptM
essagesOnlyFromDLMembers","field_type":"string","include":true,"field_format":"","field_source":"dat
a","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AcceptMessages
OnlyFromSendersOrMembers","field_type":"string","include":true,"field_format":"","field_source":"dat
a","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AddressBookPol
icy","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","r
ef_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ArchiveDatabase","field_type":"stri
ng","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":
false,"ref_col_fields":[]},{"field_name":"ArchiveDomain","field_type":"string","include":true,"field
_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]
},{"field_name":"ArchiveGuid","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ArchiveNam
e","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref
_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ArchiveRelease","field_type":"string"
,"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fal
se,"ref_col_fields":[]},{"field_name":"ArchiveState","field_type":"string","include":true,"field_for
mat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"
field_name":"ArchiveStatus","field_type":"string","include":true,"field_format":"","field_source":"d
ata","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ArchiveWarni
ngQuota","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"
","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AuditAdmin","field_type":"strin
g","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":f
alse,"ref_col_fields":[]},{"field_name":"AuditDelegate","field_type":"string","include":true,"field_
format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}
,{"field_name":"AuditEnabled","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AuditLogAg
eLimit","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":""
,"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AuditOwner","field_type":"string
","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fa
lse,"ref_col_fields":[]},{"field_name":"AutoExpandingArchiveEnabled","field_type":"string","include"
:true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_co
l_fields":[]},{"field_name":"BypassModerationFromSendersOrMembers","field_type":"string","include":t
rue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_
fields":[]},{"field_name":"CustomAttribute1","field_type":"string","include":true,"field_format":"",
"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_na
me":"CustomAttribute10","field_type":"string","include":true,"field_format":"","field_source":"data"
,"javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CustomAttribute1
1","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref
_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CustomAttribute12","field_type":"stri
ng","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":
false,"ref_col_fields":[]},{"field_name":"CustomAttribute13","field_type":"string","include":true,"f
ield_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields
":[]},{"field_name":"CustomAttribute14","field_type":"string","include":true,"field_format":"","fiel
d_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"
CustomAttribute15","field_type":"string","include":true,"field_format":"","field_source":"data","jav
ascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CustomAttribute2","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"CustomAttribute3","field_type":"string","in
clude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"
ref_col_fields":[]},{"field_name":"CustomAttribute4","field_type":"string","include":true,"field_for
mat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"
field_name":"CustomAttribute5","field_type":"string","include":true,"field_format":"","field_source"
:"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CustomAtt
ribute6","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"
","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CustomAttribute7","field_type":
"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refere
nce":false,"ref_col_fields":[]},{"field_name":"CustomAttribute8","field_type":"string","include":tru
e,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fi
elds":[]},{"field_name":"CustomAttribute9","field_type":"string","include":true,"field_format":"","f
ield_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name
":"DeliverToMailboxAndForward","field_type":"string","include":true,"field_format":"","field_source"
:"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"DisabledA
rchiveDatabase","field_type":"string","include":true,"field_format":"","field_source":"data","javasc
ript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"DisabledArchiveGuid","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"EmailAddressPolicyEnabled","field_type":"st
ring","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference
":false,"ref_col_fields":[]},{"field_name":"EndDateForRetentionHold","field_type":"string","include"
:true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_co
l_fields":[]},{"field_name":"ExchangeObjectId","field_type":"string","include":true,"field_format":"
","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_
name":"ExtensionCustomAttribute1","field_type":"string","include":true,"field_format":"","field_sour
ce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Extens
ionCustomAttribute2","field_type":"string","include":true,"field_format":"","field_source":"data","j
avascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ExtensionCustomAttr
ibute3","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":""
,"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ExtensionCustomAttribute4","fiel
d_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[]
,"reference":false,"ref_col_fields":[]},{"field_name":"ExtensionCustomAttribute5","field_type":"stri
ng","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":
false,"ref_col_fields":[]},{"field_name":"ForwardingAddress","field_type":"string","include":true,"f
ield_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields
":[]},{"field_name":"ForwardingSmtpAddress","field_type":"string","include":true,"field_format":"","
field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_nam
e":"ImmutableId","field_type":"string","include":true,"field_format":"","field_source":"data","javas
cript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IsDirSynced","field_typ
e":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"ref
erence":false,"ref_col_fields":[]},{"field_name":"IsInactiveMailbox","field_type":"string","include"
:true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_co
l_fields":[]},{"field_name":"IsMailboxEnabled","field_type":"string","include":true,"field_format":"
","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_
name":"IsMonitoringMailbox","field_type":"string","include":true,"field_format":"","field_source":"d
ata","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IsShared","f
ield_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col"
:[],"reference":false,"ref_col_fields":[]},{"field_name":"LitigationHoldDate","field_type":"string",
"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fals
e,"ref_col_fields":[]},{"field_name":"LitigationHoldDuration","field_type":"string","include":true,"
field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_field
s":[]},{"field_name":"LitigationHoldEnabled","field_type":"string","include":true,"field_format":"",
"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_na
me":"LitigationHoldOwner","field_type":"string","include":true,"field_format":"","field_source":"dat
a","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"MailboxPlan","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"MailTip","field_type":"string","include":
true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col
_fields":[]},{"field_name":"MaxBlockedSenders","field_type":"string","include":true,"field_format":"
","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_
name":"MaxReceiveSize","field_type":"string","include":true,"field_format":"","field_source":"data",
"javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"MaxSafeSenders","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"MaxSendSize","field_type":"string","inclu
de":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref
_col_fields":[]},{"field_name":"MessageCopyForSendOnBehalfEnabled","field_type":"string","include":t
rue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_
fields":[]},{"field_name":"MessageCopyForSentAsEnabled","field_type":"string","include":true,"field_
format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}
,{"field_name":"MessageCopyForSMTPClientSubmissionEnabled","field_type":"string","include":true,"fie
ld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":
[]},{"field_name":"MessageRecallProcessingEnabled","field_type":"string","include":true,"field_forma
t":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fi
eld_name":"MessageTrackingReadStatusEnabled","field_type":"string","include":true,"field_format":"",
"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_na
me":"RecipientLimits","field_type":"string","include":true,"field_format":"","field_source":"data","
javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"RecipientType","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"RecipientTypeDetails","field_type":"string"
,"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fal
se,"ref_col_fields":[]},{"field_name":"RejectMessagesFrom","field_type":"string","include":true,"fie
ld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":
[]},{"field_name":"RejectMessagesFromDLMembers","field_type":"string","include":true,"field_format":
"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field
_name":"RejectMessagesFromSendersOrMembers","field_type":"string","include":true,"field_format":"","
field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_nam
e":"RetentionHoldEnabled","field_type":"string","include":true,"field_format":"","field_source":"dat
a","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"RetentionPolic
y","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref
_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"SCLDeleteEnabled","field_type":"strin
g","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":f
alse,"ref_col_fields":[]},{"field_name":"SCLDeleteThreshold","field_type":"string","include":true,"f
ield_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields
":[]},{"field_name":"SCLJunkEnabled","field_type":"string","include":true,"field_format":"","field_s
ource":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"SCL
JunkThreshold","field_type":"string","include":true,"field_format":"","field_source":"data","javascr
ipt":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"SCLQuarantineEnabled","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"SCLQuarantineThreshold","field_type":"strin
g","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":f
alse,"ref_col_fields":[]},{"field_name":"SCLRejectEnabled","field_type":"string","include":true,"fie
ld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":
[]},{"field_name":"SCLRejectThreshold","field_type":"string","include":true,"field_format":"","field
_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"S
endModerationNotifications","field_type":"string","include":true,"field_format":"","field_source":"d
ata","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"UsageLocatio
n","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref
_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"WhenCreated","field_type":"string","i
nclude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,
"ref_col_fields":[]},{"field_name":"WhenMailboxCreated","field_type":"string","include":true,"field_
format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}
,{"field_name":"WhenSoftDeleted","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Identit
y","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref
_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"AddressListMembership","field_type":"
string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"referen
ce":false,"ref_col_fields":[]},{"field_name":"AggregatedMailboxGuids","field_type":"string","include
":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_c
ol_fields":[]},{"field_name":"AntispamBypassEnabled","field_type":"string","include":true,"field_for
mat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"
field_name":"ArbitrationMailbox","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Archive
Quota","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"",
"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CalendarLoggingQuota","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refe
rence":false,"ref_col_fields":[]},{"field_name":"CalendarRepairDisabled","field_type":"string","incl
ude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"re
f_col_fields":[]},{"field_name":"CalendarVersionStoreDisabled","field_type":"string","include":true,
"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fiel
ds":[]},{"field_name":"ComplianceTagHoldApplied","field_type":"string","include":true,"field_format"
:"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fiel
d_name":"Database","field_type":"string","include":true,"field_format":"","field_source":"data","jav
ascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"DataEncryptionPolicy"
,"field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_c
ol":[],"reference":false,"ref_col_fields":[]},{"field_name":"DefaultAuditSet","field_type":"string",
"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fals
e,"ref_col_fields":[]},{"field_name":"DefaultPublicFolderMailbox","field_type":"string","include":tr
ue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_f
ields":[]},{"field_name":"DelayHoldApplied","field_type":"string","include":true,"field_format":"","
field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_nam
e":"DisabledMailboxLocations","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Distinguis
hedName","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"
","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"DowngradeHighPriorityMessagesEn
abled","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"",
"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"EffectivePublicFolderMailbox","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"ElcProcessingDisabled","field_type":"string
","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fa
lse,"ref_col_fields":[]},{"field_name":"ExchangeGuid","field_type":"string","include":true,"field_fo
rmat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{
"field_name":"ExchangeSecurityDescriptor","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"ExchangeUserAccountControl","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ExchangeVe
rsion","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"",
"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Extensions","field_type":"string"
,"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fal
se,"ref_col_fields":[]},{"field_name":"ExternalDirectoryObjectId","field_type":"string","include":tr
ue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_f
ields":[]},{"field_name":"ExternalOofOptions","field_type":"string","include":true,"field_format":""
,"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_n
ame":"GeneratedOfflineAddressBooks","field_type":"string","include":true,"field_format":"","field_so
urce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Gran
tSendOnBehalfTo","field_type":"string","include":true,"field_format":"","field_source":"data","javas
cript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"HasPicture","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refe
rence":false,"ref_col_fields":[]},{"field_name":"HasSnackyAppData","field_type":"string","include":t
rue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_
fields":[]},{"field_name":"HasSpokenName","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"HiddenFromAddressListsEnabled","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ImListM
igrationCompleted","field_type":"string","include":true,"field_format":"","field_source":"data","jav
ascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"InactiveMailboxRetire
Time","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","
ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IncludeInGarbageCollection","field
_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],
"reference":false,"ref_col_fields":[]},{"field_name":"InPlaceHolds","field_type":"string","include":
true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col
_fields":[]},{"field_name":"IsExcludedFromServingHierarchy","field_type":"string","include":true,"fi
eld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields"
:[]},{"field_name":"IsHierarchyReady","field_type":"string","include":true,"field_format":"","field_
source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Is
HierarchySyncEnabled","field_type":"string","include":true,"field_format":"","field_source":"data","
javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IsLinked","field_t
ype":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"r
eference":false,"ref_col_fields":[]},{"field_name":"IsMachineToPersonTextMessagingEnabled","field_ty
pe":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"re
ference":false,"ref_col_fields":[]},{"field_name":"IsPersonToPersonTextMessagingEnabled","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refe
rence":false,"ref_col_fields":[]},{"field_name":"IsResource","field_type":"string","include":true,"f
ield_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields
":[]},{"field_name":"IsRootPublicFolderMailbox","field_type":"string","include":true,"field_format":
"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field
_name":"IsSoftDeletedByDisable","field_type":"string","include":true,"field_format":"","field_source
":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IsSoftDe
letedByRemove","field_type":"string","include":true,"field_format":"","field_source":"data","javascr
ipt":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"IssueWarningQuota","field
_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],
"reference":false,"ref_col_fields":[]},{"field_name":"JournalArchiveAddress","field_type":"string","
include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false
,"ref_col_fields":[]},{"field_name":"Languages","field_type":"string","include":true,"field_format":
"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field
_name":"LastExchangeChangedTime","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"LegacyE
xchangeDN","field_type":"string","include":true,"field_format":"","field_source":"data","javascript"
:"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"LinkedMasterAccount","field_t
ype":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"r
eference":false,"ref_col_fields":[]},{"field_name":"MailboxContainerGuid","field_type":"string","inc
lude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"r
ef_col_fields":[]},{"field_name":"MailboxLocations","field_type":"string","include":true,"field_form
at":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"f
ield_name":"MailboxMoveBatchName","field_type":"string","include":true,"field_format":"","field_sour
ce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Mailbo
xMoveFlags","field_type":"string","include":true,"field_format":"","field_source":"data","javascript
":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"MailboxMoveRemoteHostName","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"MailboxMoveSourceMDB","field_type":"strin
g","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":f
alse,"ref_col_fields":[]},{"field_name":"MailboxMoveStatus","field_type":"string","include":true,"fi
eld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields"
:[]},{"field_name":"MailboxMoveTargetMDB","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"MailboxProvisioningConstraint","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Mailbox
ProvisioningPreferences","field_type":"string","include":true,"field_format":"","field_source":"data
","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"MailboxRegion",
"field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_co
l":[],"reference":false,"ref_col_fields":[]},{"field_name":"MailboxRegionLastUpdateTime","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refe
rence":false,"ref_col_fields":[]},{"field_name":"MailboxRelease","field_type":"string","include":tru
e,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fi
elds":[]},{"field_name":"MailTipTranslations","field_type":"string","include":true,"field_format":""
,"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_n
ame":"ManagedFolderMailboxPolicy","field_type":"string","include":true,"field_format":"","field_sour
ce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Micros
oftOnlineServicesID","field_type":"string","include":true,"field_format":"","field_source":"data","j
avascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ModeratedBy","field
_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],
"reference":false,"ref_col_fields":[]},{"field_name":"ModerationEnabled","field_type":"string","incl
ude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"re
f_col_fields":[]},{"field_name":"Name","field_type":"string","include":true,"field_format":"","field
_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"N
etID","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","
ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"NonCompliantDevices","field_type":
"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refere
nce":false,"ref_col_fields":[]},{"field_name":"ObjectCategory","field_type":"string","include":true,
"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fiel
ds":[]},{"field_name":"ObjectClass","field_type":"string","include":true,"field_format":"","field_so
urce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Offi
ce","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","re
f_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"OfflineAddressBook","field_type":"st
ring","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference
":false,"ref_col_fields":[]},{"field_name":"OrganizationalUnit","field_type":"string","include":true
,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fie
lds":[]},{"field_name":"OrganizationId","field_type":"string","include":true,"field_format":"","fiel
d_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"
OrphanSoftDeleteTrackingTime","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"PersistedC
apabilities","field_type":"string","include":true,"field_format":"","field_source":"data","javascrip
t":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"PoliciesExcluded","field_ty
pe":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"re
ference":false,"ref_col_fields":[]},{"field_name":"PoliciesIncluded","field_type":"string","include"
:true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_co
l_fields":[]},{"field_name":"ProhibitSendQuota","field_type":"string","include":true,"field_format":
"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field
_name":"ProhibitSendReceiveQuota","field_type":"string","include":true,"field_format":"","field_sour
ce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Protoc
olSettings","field_type":"string","include":true,"field_format":"","field_source":"data","javascript
":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"QueryBaseDN","field_type":"s
tring","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"referenc
e":false,"ref_col_fields":[]},{"field_name":"QueryBaseDNRestrictionEnabled","field_type":"string","i
nclude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,
"ref_col_fields":[]},{"field_name":"ReconciliationId","field_type":"string","include":true,"field_fo
rmat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{
"field_name":"RecoverableItemsQuota","field_type":"string","include":true,"field_format":"","field_s
ource":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Rec
overableItemsWarningQuota","field_type":"string","include":true,"field_format":"","field_source":"da
ta","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"RemoteAccount
Policy","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":""
,"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"RemoteRecipientType","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refe
rence":false,"ref_col_fields":[]},{"field_name":"RequireSenderAuthenticationEnabled","field_type":"s
tring","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"referenc
e":false,"ref_col_fields":[]},{"field_name":"ResetPasswordOnNextLogon","field_type":"string","includ
e":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_
col_fields":[]},{"field_name":"ResourceCapacity","field_type":"string","include":true,"field_format"
:"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fiel
d_name":"ResourceCustom","field_type":"string","include":true,"field_format":"","field_source":"data
","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ResourceType","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"RetainDeletedItemsFor","field_type":"stri
ng","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":
false,"ref_col_fields":[]},{"field_name":"RetainDeletedItemsUntilBackup","field_type":"string","incl
ude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"re
f_col_fields":[]},{"field_name":"RetentionComment","field_type":"string","include":true,"field_forma
t":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fi
eld_name":"RetentionUrl","field_type":"string","include":true,"field_format":"","field_source":"data
","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"RoleAssignmentP
olicy","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"",
"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"RoomMailboxAccountEnabled","field
_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],
"reference":false,"ref_col_fields":[]},{"field_name":"RulesQuota","field_type":"string","include":tr
ue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_f
ields":[]},{"field_name":"SamAccountName","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"ServerLegacyDN","field_type":"string","include":true,"field_format":"","field_source":"data","java
script":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"SharingPolicy","field_
type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"
reference":false,"ref_col_fields":[]},{"field_name":"SiloName","field_type":"string","include":true,
"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fiel
ds":[]},{"field_name":"SimpleDisplayName","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"SingleItemRecoveryEnabled","field_type":"string","include":true,"field_format":"","field_source":"
data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"SKUAssigned
","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_
col":[],"reference":false,"ref_col_fields":[]},{"field_name":"SourceAnchor","field_type":"string","i
nclude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,
"ref_col_fields":[]},{"field_name":"StartDateForRetentionHold","field_type":"string","include":true,
"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fiel
ds":[]},{"field_name":"StsRefreshTokensValidFrom","field_type":"string","include":true,"field_format
":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fie
ld_name":"ThrottlingPolicy","field_type":"string","include":true,"field_format":"","field_source":"d
ata","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"UMDtmfMap","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"UMEnabled","field_type":"string","include
":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_c
ol_fields":[]},{"field_name":"UnifiedMailbox","field_type":"string","include":true,"field_format":""
,"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_n
ame":"UseDatabaseQuotaDefaults","field_type":"string","include":true,"field_format":"","field_source
":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"UseDatab
aseRetentionDefaults","field_type":"string","include":true,"field_format":"","field_source":"data","
javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"WasInactiveMailbox
","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_
col":[],"reference":false,"ref_col_fields":[]},{"field_name":"WhenChanged","field_type":"string","in
clude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"
ref_col_fields":[]},{"field_name":"WhenChangedUTC","field_type":"string","include":true,"field_forma
t":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fi
eld_name":"WhenCreatedUTC","field_type":"string","include":true,"field_format":"","field_source":"da
ta","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"WindowsEmailA
ddress","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":""
,"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"WindowsLiveID","field_type":"str
ing","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference"
:false,"ref_col_fields":[]}],"key":"Guid","display":"Alias","name_values":[{"name":"filter","value":
""},{"name":"properties","value":["AcceptMessagesOnlyFrom","AcceptMessagesOnlyFromDLMembers","Accept
MessagesOnlyFromSendersOrMembers","AccountDisabled","AddressBookPolicy","Alias","ArchiveDatabase","A
rchiveDomain","ArchiveGuid","ArchiveName","ArchiveRelease","ArchiveState","ArchiveStatus","ArchiveWa
rningQuota","AuditAdmin","AuditDelegate","AuditEnabled","AuditLogAgeLimit","AuditOwner","AutoExpandi
ngArchiveEnabled","BypassModerationFromSendersOrMembers","CustomAttribute1","CustomAttribute10","Cus
tomAttribute11","CustomAttribute12","CustomAttribute13","CustomAttribute14","CustomAttribute15","Cus
tomAttribute2","CustomAttribute3","CustomAttribute4","CustomAttribute5","CustomAttribute6","CustomAt
tribute7","CustomAttribute8","CustomAttribute9","DeliverToMailboxAndForward","DisabledArchiveDatabas
e","DisabledArchiveGuid","DisplayName","EmailAddresses","EmailAddressPolicyEnabled","EndDateForReten
tionHold","ExchangeObjectId","ExtensionCustomAttribute1","ExtensionCustomAttribute2","ExtensionCusto
mAttribute3","ExtensionCustomAttribute4","ExtensionCustomAttribute5","ForwardingAddress","Forwarding
SmtpAddress","Guid","Id","ImmutableId","IsDirSynced","IsInactiveMailbox","IsMailboxEnabled","IsMonit
oringMailbox","IsShared","LitigationHoldDate","LitigationHoldDuration","LitigationHoldEnabled","Liti
gationHoldOwner","MailboxPlan","MailTip","MaxBlockedSenders","MaxReceiveSize","MaxSafeSenders","MaxS
endSize","MessageCopyForSendOnBehalfEnabled","MessageCopyForSentAsEnabled","MessageCopyForSMTPClient
SubmissionEnabled","MessageRecallProcessingEnabled","MessageTrackingReadStatusEnabled","PrimarySmtpA
ddress","RecipientLimits","RecipientType","RecipientTypeDetails","RejectMessagesFrom","RejectMessage
sFromDLMembers","RejectMessagesFromSendersOrMembers","RetentionHoldEnabled","RetentionPolicy","SCLDe
leteEnabled","SCLDeleteThreshold","SCLJunkEnabled","SCLJunkThreshold","SCLQuarantineEnabled","SCLQua
rantineThreshold","SCLRejectEnabled","SCLRejectThreshold","SendModerationNotifications","UsageLocati
on","UserPrincipalName","WhenCreated","WhenMailboxCreated","WhenSoftDeleted"]}],"sys_nn":[],"source"
:"data"},{"col_name":"MailboxAutoReplyConfigurations","fields":[{"field_name":"Identity","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":["Mailbo
xes"],"reference":false,"ref_col_fields":[]},{"field_name":"AutoDeclineFutureRequestsWhenOOF","field
_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],
"reference":false,"ref_col_fields":[]},{"field_name":"AutoReplyState","field_type":"string","include
":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_c
ol_fields":[]},{"field_name":"CreateOOFEvent","field_type":"string","include":true,"field_format":""
,"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_n
ame":"DeclineAllEventsForScheduledOOF","field_type":"string","include":true,"field_format":"","field
_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"D
eclineEventsForScheduledOOF","field_type":"string","include":true,"field_format":"","field_source":"
data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"DeclineMeet
ingMessage","field_type":"string","include":true,"field_format":"","field_source":"data","javascript
":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"EndTime","field_type":"strin
g","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":f
alse,"ref_col_fields":[]},{"field_name":"ExternalAudience","field_type":"string","include":true,"fie
ld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":
[]},{"field_name":"ExternalMessage","field_type":"string","include":true,"field_format":"","field_so
urce":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Inte
rnalMessage","field_type":"string","include":true,"field_format":"","field_source":"data","javascrip
t":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"StartTime","field_type":"st
ring","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference
":false,"ref_col_fields":[]},{"field_name":"MailboxOwnerId","field_type":"string","include":true,"fi
eld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields"
:[]}],"key":"Identity","display":"","name_values":[],"sys_nn":[],"source":"data"},{"col_name":"CASMa
ilboxes","fields":[{"field_name":"Guid","field_type":"string","include":true,"field_format":"","fiel
d_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"
ActiveSyncEnabled","field_type":"string","include":true,"field_format":"","field_source":"data","jav
ascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"DisplayName","field_t
ype":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"r
eference":false,"ref_col_fields":[]},{"field_name":"EmailAddresses","field_type":"string","include":
true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col
_fields":[]},{"field_name":"Identity","field_type":"string","include":true,"field_format":"","field_
source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Im
apEnabled","field_type":"string","include":true,"field_format":"","field_source":"data","javascript"
:"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"LinkedMasterAccount","field_t
ype":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"r
eference":false,"ref_col_fields":[]},{"field_name":"Name","field_type":"string","include":true,"fiel
d_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[
]},{"field_name":"OWAEnabled","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"PopEnabled
","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_
col":[],"reference":false,"ref_col_fields":[]},{"field_name":"PrimarySmtpAddress","field_type":"stri
ng","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":
false,"ref_col_fields":[]},{"field_name":"SamAccountName","field_type":"string","include":true,"fiel
d_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[
]}],"key":"Guid","display":"Name","name_values":[],"sys_nn":[],"source":"data"},{"col_name":"Mailbox
Permissions","fields":[{"field_name":"Identity","field_type":"string","include":true,"field_format":
"","field_source":"data","javascript":"","ref_col":["Mailboxes"],"reference":false,"ref_col_fields":
[]},{"field_name":"AccessRights","field_type":"string","include":true,"field_format":"","field_sourc
e":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"Deny","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"InheritanceType","field_type":"string","i
nclude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,
"ref_col_fields":[]},{"field_name":"IsInherited","field_type":"string","include":true,"field_format"
:"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fiel
d_name":"User","field_type":"string","include":true,"field_format":"","field_source":"data","javascr
ipt":"","ref_col":[],"reference":false,"ref_col_fields":[{"col":"Mailboxes","field":"UserPrincipalNa
me"}]}],"key":"Identity","display":"User","name_values":[],"sys_nn":[],"source":"data"}]}]
'@
