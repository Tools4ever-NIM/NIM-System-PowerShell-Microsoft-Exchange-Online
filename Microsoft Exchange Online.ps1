#
# Microsoft Exchange Online.ps1 - IDM System PowerShell Script for Microsoft Exchange Online Services.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#

# Resolve any potential TLS issues
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12

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
                value = 5
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
		@{ name = 'ActiveSyncAllowedDeviceIDs';					options = @('set')				}
		@{ name = 'ActiveSyncBlockedDeviceIDs';					options = @('set')				}
		@{ name = 'ActiveSyncDebugLogging';						options = @('set')				}
		@{ name = 'ActiveSyncEnabled';							options = @('default','set')	}
		@{ name = 'ActiveSyncMailboxPolicy';					options = @('set')				}
		@{ name = 'ActiveSyncMailboxPolicyIsDefaulted';		 									}
		@{ name = 'ActiveSyncSuppressReadReceipt';				options = @('set')				}
		@{ name = 'DisplayName';								options = @('default','set')	}
		@{ name = 'DistinguishedName';		 													}
		@{ name = 'ECPEnabled';		 															}
		@{ name = 'EmailAddresses';								options = @('default','set')	}
		@{ name = 'EwsAllowEntourage';							options = @('set')				}
		@{ name = 'EwsAllowList';								options = @('set')				}
		@{ name = 'EwsAllowMacOutlook';							options = @('set')				}
		@{ name = 'EwsAllowOutlook';							options = @('set')				}
		@{ name = 'EwsApplicationAccessPolicy';					options = @('set')				}
		@{ name = 'EwsBlockList';								options = @('set')				}
		@{ name = 'EwsEnabled';									options = @('set')				}
		@{ name = 'ExchangeObjectId';		 													}
		@{ name = 'ExchangeVersion';		 													}
		@{ name = 'ExternalDirectoryObjectId';		 											}
		@{ name = 'Guid';										options = @('default','key')	}
		@{ name = 'HasActiveSyncDevicePartnership';		 										}
		@{ name = 'Identity';									options = @('default')			}
		@{ name = 'ImapEnabled';								options = @('default','set')	}
		@{ name = 'ImapEnableExactRFC822Size';		 											}
		@{ name = 'ImapForceICalForCalendarRetrievalOption';	options = @('set')				}
		@{ name = 'ImapMessagesRetrievalMimeFormat';			options = @('set')				}
		@{ name = 'ImapSuppressReadReceipt';					options = @('set')				}
		@{ name = 'ImapUseProtocolDefaults';					options = @('set')				}
		@{ name = 'IsOptimizedForAccessibility';				options = @('set')				}
		@{ name = 'LegacyExchangeDN';		 													}
		@{ name = 'LinkedMasterAccount';						options = @('default')			}
		@{ name = 'MacOutlookEnabled';							options = @('set')				}
		@{ name = 'MAPIBlockOutlookExternalConnectivity';		 								}
		@{ name = 'MAPIBlockOutlookNonCachedMode';		 										}
		@{ name = 'MAPIBlockOutlookRpcHttp';		 											}
		@{ name = 'MAPIBlockOutlookVersions';		 											}
		@{ name = 'MAPIEnabled';								options = @('set')				}
		@{ name = 'MapiHttpEnabled';							options = @('set')				}
		@{ name = 'Name';										options = @('default')			}
		@{ name = 'ObjectCategory';		 														}
		@{ name = 'ObjectClass';		 														}
		@{ name = 'OrganizationId';		 														}
		@{ name = 'OutlookMobileEnabled';						options = @('set')				}
		@{ name = 'OWAEnabled';									options = @('default','set')	}
		@{ name = 'OWAforDevicesEnabled';						options = @('set')				}
		@{ name = 'OwaMailboxPolicy';							options = @('set')				}
		@{ name = 'PopEnabled';									options = @('default','set')	}
		@{ name = 'PopEnableExactRFC822Size';		 											}
		@{ name = 'PopForceICalForCalendarRetrievalOption';		options = @('set')				}
		@{ name = 'PopMessageDeleteEnabled';		 											}
		@{ name = 'PopMessagesRetrievalMimeFormat';				options = @('set')				}
		@{ name = 'PopSuppressReadReceipt';						options = @('set')				}
		@{ name = 'PopUseProtocolDefaults';						options = @('set')				}
		@{ name = 'PrimarySmtpAddress';							options = @('default')			}
		@{ name = 'PublicFolderClientAccess';					options = @('set')				}
		@{ name = 'SamAccountName';								options = @('default')			}
		@{ name = 'ServerLegacyDN';		 														}
		@{ name = 'ShowGalAsDefaultView';						options = @('set')				}
		@{ name = 'SmtpClientAuthenticationDisabled';			options = @('set')				}
		@{ name = 'UniversalOutlookEnabled';					options = @('set')				}
		@{ name = 'WhenChanged';		 														}
		@{ name = 'WhenChangedUTC';		 														}
		@{ name = 'WhenCreated';		 														}
		@{ name = 'WhenCreatedUTC';		 														}
	)
    DistributionGroup = @(
		@{ name = 'AcceptMessagesOnlyFrom';						options = @('set')						}
		@{ name = 'AcceptMessagesOnlyFromDLMembers';			options = @('set')						}
		@{ name = 'AcceptMessagesOnlyFromSendersOrMembers';		options = @('set')						}
		@{ name = 'AddressListMembership';					 											}
		@{ name = 'AdministrativeUnits';					 											}
		@{ name = 'Alias';										options = @('default','create','set')	}
		@{ name = 'ArbitrationMailbox';					 		options = @('create','set')	            }
		@{ name = 'BccBlocked';							        options = @('create','set')	            }
		@{ name = 'BypassModerationFromSendersOrMembers';		options = @('create','set')						}
		@{ name = 'BypassNestedModerationEnabled';				options = @('set')						}
		@{ name = 'CustomAttribute1';							options = @('set')						}
		@{ name = 'CustomAttribute10';							options = @('set')						}
		@{ name = 'CustomAttribute11';							options = @('set')						}
		@{ name = 'CustomAttribute12';							options = @('set')						}
		@{ name = 'CustomAttribute13';							options = @('set')						}
		@{ name = 'CustomAttribute14';							options = @('set')						}
		@{ name = 'CustomAttribute15';							options = @('set')						}
		@{ name = 'CustomAttribute2';							options = @('set')						}
		@{ name = 'CustomAttribute3';							options = @('set')						}
		@{ name = 'CustomAttribute4';							options = @('set')						}
		@{ name = 'CustomAttribute5';							options = @('set')						}
		@{ name = 'CustomAttribute6';							options = @('set')						}
		@{ name = 'CustomAttribute7';							options = @('set')						}
		@{ name = 'CustomAttribute8';							options = @('set')						}
		@{ name = 'CustomAttribute9';							options = @('set')						}
		@{ name = 'Description';				 				options = @('create','set')	            }
		@{ name = 'DisplayName';								options = @('default','create','set')	}
		@{ name = 'DistinguishedName';					 												}
		@{ name = 'EmailAddresses';								options = @('set')          			}
		@{ name = 'EmailAddressPolicyEnabled';															}
		@{ name = 'ExchangeObjectId';					 												}
		@{ name = 'ExchangeVersion';					 												}
		@{ name = 'ExtensionCustomAttribute1';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute2';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute3';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute4';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute5';					options = @('set')						}
		@{ name = 'GrantSendOnBehalfTo';						options = @('set')						}
		@{ name = 'GroupType';					 														}
        @{ name = 'Guid';										options = @('default','key')			}		
		@{ name = 'HiddenFromAddressListsEnabled';				options = @('set')						}
        @{ name = 'HiddenGroupMembershipEnabled';				options = @('create','set')			    }
		@{ name = 'Id';											options = @('default')					}
		@{ name = 'Identity';					 														}
		@{ name = 'IsDirSynced';																		}
		@{ name = 'IsValid';				            	 											}
		@{ name = 'LastExchangeChangedTime';															}
		@{ name = 'LegacyExchangeDN';					 												}
		@{ name = 'MailTip';									options = @('set')						}
		@{ name = 'MailTipTranslations';						options = @('set')						}
		@{ name = 'ManagedBy';									options = @('create','set')			    }
		@{ name = 'MaxReceiveSize';								options = @('set')						}
		@{ name = 'MaxSendSize';								options = @('set')						}
		@{ name = 'MemberDepartRestriction';    				options = @('create','set')			    }
		@{ name = 'MemberJoinRestriction';          			options = @('create','set')			    }
		@{ name = 'MigrationToUnifiedGroupInProgress';			                                        }
		@{ name = 'ModeratedBy';								options = @('create','set')			    }
		@{ name = 'ModerationEnabled';							options = @('create','set')			    }
		@{ name = 'Name';										options = @('create','set')     		}
		@{ name = 'ObjectCategory';					 													}
		@{ name = 'ObjectClass';					 													}
		@{ name = 'OrganizationalUnit';					 												}
        @{ name = 'OrganizationalUnitRoot';					 											}
		@{ name = 'OrganizationId';					 													}
		@{ name = 'OriginatingServer';          					 									}
		@{ name = 'PoliciesExcluded';					 												}
		@{ name = 'PoliciesIncluded';					 												}
		@{ name = 'PrimarySmtpAddress';							options = @('default','create','set')   }
		@{ name = 'RecipientType';					 													}
		@{ name = 'RecipientTypeDetails';					 											}
		@{ name = 'RejectMessagesFrom';							options = @('set')						}
		@{ name = 'RejectMessagesFromDLMembers';				options = @('set')						}
		@{ name = 'RejectMessagesFromSendersOrMembers';			options = @('set')						}
		@{ name = 'ReportToManagerEnabled';					 											}
		@{ name = 'ReportToOriginatorEnabled';  			 											}
		@{ name = 'SamAccountName';								options = @('create','set')				}
		@{ name = 'SendModerationNotifications';				options = @('create','set')			    }
        @{ name = 'SendOofMessageToOriginatorEnabled';			options = @('set')						}
        @{ name = 'Type';			                            options = @('create')   				}
		@{ name = 'UMDtmfMap';					 														}
		@{ name = 'WhenChanged';					 													}
		@{ name = 'WhenChangedUTC';					 													}
		@{ name = 'WhenCreated';					 													}
		@{ name = 'WhenCreatedUTC';					 													}
		@{ name = 'WindowsEmailAddress';						options = @('set')						}
	)
    DistributionGroupMember = @(
		@{ name = 'GroupGuid';						options = @('default','set')		}
		@{ name = 'Guid';			options = @('default','set')						}
		@{ name = 'RecipientType';		                                                }
	)
	Mailbox = @(
		@{ name = 'AcceptMessagesOnlyFrom';						options = @('set')						}
		@{ name = 'AcceptMessagesOnlyFromDLMembers';			options = @('set')						}
		@{ name = 'AcceptMessagesOnlyFromSendersOrMembers';		options = @('set')						}
		@{ name = 'AccountDisabled';							options = @('default','set')			}
		@{ name = 'AddressBookPolicy';							options = @('enable','set')				}
		@{ name = 'AddressListMembership';					 											}
		@{ name = 'AdministrativeUnits';					 											}
		@{ name = 'AggregatedMailboxGuids';					 											}
		@{ name = 'Alias';										options = @('default','enable','set')	}
		@{ name = 'AntispamBypassEnabled';																}
		@{ name = 'ArbitrationMailbox';					 												}
		@{ name = 'ArchiveDatabase';							options = @('enable')					}
		@{ name = 'ArchiveDomain';								options = @('enable')					}
		@{ name = 'ArchiveGuid';								options = @('enable')					}
		@{ name = 'ArchiveName';								options = @('set')						}
		@{ name = 'ArchiveQuota';					 													}
		@{ name = 'ArchiveRelease';					 													}
		@{ name = 'ArchiveState';								options = @('set')						}
		@{ name = 'ArchiveStatus';					 													}
		@{ name = 'ArchiveWarningQuota';																}
		@{ name = 'AuditAdmin';									options = @('set')						}
		@{ name = 'AuditDelegate';								options = @('set')						}
		@{ name = 'AuditEnabled';								options = @('set')						}
		@{ name = 'AuditLogAgeLimit';							options = @('set')						}
		@{ name = 'AuditOwner';									options = @('set')						}
		@{ name = 'AutoExpandingArchiveEnabled';					 									}
		@{ name = 'BypassModerationFromSendersOrMembers';		options = @('set')						}
		@{ name = 'CalendarLoggingQuota';					 											}
		@{ name = 'CalendarRepairDisabled';						options = @('set')						}
		@{ name = 'CalendarVersionStoreDisabled';				options = @('set')						}
		@{ name = 'ComplianceTagHoldApplied';					 										}
		@{ name = 'CustomAttribute1';							options = @('set')						}
		@{ name = 'CustomAttribute10';							options = @('set')						}
		@{ name = 'CustomAttribute11';							options = @('set')						}
		@{ name = 'CustomAttribute12';							options = @('set')						}
		@{ name = 'CustomAttribute13';							options = @('set')						}
		@{ name = 'CustomAttribute14';							options = @('set')						}
		@{ name = 'CustomAttribute15';							options = @('set')						}
		@{ name = 'CustomAttribute2';							options = @('set')						}
		@{ name = 'CustomAttribute3';							options = @('set')						}
		@{ name = 'CustomAttribute4';							options = @('set')						}
		@{ name = 'CustomAttribute5';							options = @('set')						}
		@{ name = 'CustomAttribute6';							options = @('set')						}
		@{ name = 'CustomAttribute7';							options = @('set')						}
		@{ name = 'CustomAttribute8';							options = @('set')						}
		@{ name = 'CustomAttribute9';							options = @('set')						}
		@{ name = 'Database';					 														}
		@{ name = 'DataEncryptionPolicy';					 											}
		@{ name = 'DefaultAuditSet';							options = @('set')						}
		@{ name = 'DefaultPublicFolderMailbox';					options = @('set')						}
		@{ name = 'DelayHoldApplied';					 												}
		@{ name = 'DeliverToMailboxAndForward';					options = @('set')						}
		@{ name = 'DisabledArchiveDatabase';															}
		@{ name = 'DisabledArchiveGuid';					 											}
		@{ name = 'DisabledMailboxLocations';															}
		@{ name = 'DisplayName';								options = @('default','enable','set')	}
		@{ name = 'DistinguishedName';					 												}
		@{ name = 'DowngradeHighPriorityMessagesEnabled';					 							}
		@{ name = 'EffectivePublicFolderMailbox';					 									}
		@{ name = 'ElcProcessingDisabled';						options = @('set')						}
		@{ name = 'EmailAddresses';								options = @('default','set')			}
		@{ name = 'EmailAddressPolicyEnabled';															}
		@{ name = 'EndDateForRetentionHold';					options = @('set')						}
		@{ name = 'ExchangeGuid';					 													}
		@{ name = 'ExchangeObjectId';					 												}
		@{ name = 'ExchangeSecurityDescriptor';															}
		@{ name = 'ExchangeUserAccountControl';															}
		@{ name = 'ExchangeVersion';					 												}
		@{ name = 'ExtensionCustomAttribute1';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute2';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute3';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute4';					options = @('set')						}
		@{ name = 'ExtensionCustomAttribute5';					options = @('set')						}
		@{ name = 'Extensions';					 														}
		@{ name = 'ExternalDirectoryObjectId';					 										}
		@{ name = 'ExternalOofOptions';							options = @('set')						}
		@{ name = 'ForwardingAddress';							options = @('set')						}
		@{ name = 'ForwardingSmtpAddress';						options = @('set')						}
		@{ name = 'GeneratedOfflineAddressBooks';					 									}
		@{ name = 'GrantSendOnBehalfTo';						options = @('set')						}
		@{ name = 'Guid';										options = @('default','key')			}
		@{ name = 'HasPicture';					 														}
		@{ name = 'HasSnackyAppData';																	}
		@{ name = 'HasSpokenName';					 													}
		@{ name = 'HiddenFromAddressListsEnabled';				options = @('set')						}
		@{ name = 'Id';											options = @('default')					}
		@{ name = 'Identity';					 														}
		@{ name = 'ImListMigrationCompleted';															}
		@{ name = 'ImmutableId';								options = @('set')						}
		@{ name = 'InactiveMailboxRetireTime';															}
		@{ name = 'IncludeInGarbageCollection';															}
		@{ name = 'InPlaceHolds';																		}
		@{ name = 'IsDirSynced';																		}
		@{ name = 'IsExcludedFromServingHierarchy';				options = @('set')						}
		@{ name = 'IsHierarchyReady';																	}
		@{ name = 'IsHierarchySyncEnabled';																}
		@{ name = 'IsInactiveMailbox';							options = @('set')						}
		@{ name = 'IsLinked';					 														}
		@{ name = 'IsMachineToPersonTextMessagingEnabled';					 							}
		@{ name = 'IsMailboxEnabled';					 												}
		@{ name = 'IsMonitoringMailbox';					 											}
		@{ name = 'IsPersonToPersonTextMessagingEnabled';					 							}
		@{ name = 'IsResource';					 														}
		@{ name = 'IsRootPublicFolderMailbox';					 										}
		@{ name = 'IsShared';					 														}
		@{ name = 'IsSoftDeletedByDisable';					 											}
		@{ name = 'IsSoftDeletedByRemove';					 											}
		@{ name = 'IssueWarningQuota';							options = @('set')						}
		@{ name = 'JournalArchiveAddress';						options = @('set')						}
		@{ name = 'Languages';									options = @('set')						}
		@{ name = 'LastExchangeChangedTime';															}
		@{ name = 'LegacyExchangeDN';					 												}
		@{ name = 'LinkedMasterAccount';					 											}
		@{ name = 'LitigationHoldDate';							options = @('set')						}
		@{ name = 'LitigationHoldDuration';						options = @('set')						}
		@{ name = 'LitigationHoldEnabled';						options = @('set')						}
		@{ name = 'LitigationHoldOwner';						options = @('set')						}
		@{ name = 'MailboxContainerGuid';																}
		@{ name = 'MailboxLocations';					 												}
		@{ name = 'MailboxMoveBatchName';																}
		@{ name = 'MailboxMoveFlags';																	}
		@{ name = 'MailboxMoveRemoteHostName';															}
		@{ name = 'MailboxMoveSourceMDB';																}
		@{ name = 'MailboxMoveStatus';																	}
		@{ name = 'MailboxMoveTargetMDB';																}
		@{ name = 'MailboxPlan';																		}
		@{ name = 'MailboxProvisioningConstraint';														}
		@{ name = 'MailboxProvisioningPreferences';														}
		@{ name = 'MailboxRegion';								options = @('set')						}
		@{ name = 'MailboxRegionLastUpdateTime';														}
		@{ name = 'MailboxRelease';																		}
		@{ name = 'MailTip';									options = @('set')						}
		@{ name = 'MailTipTranslations';						options = @('set')						}
		@{ name = 'ManagedFolderMailboxPolicy';					options = @('enable','set')				}
		@{ name = 'MaxBlockedSenders';																	}
		@{ name = 'MaxReceiveSize';								options = @('set')						}
		@{ name = 'MaxSafeSenders';																		}
		@{ name = 'MaxSendSize';								options = @('set')						}
		@{ name = 'MessageCopyForSendOnBehalfEnabled';			options = @('set')						}
		@{ name = 'MessageCopyForSentAsEnabled';				options = @('set')						}
		@{ name = 'MessageCopyForSMTPClientSubmissionEnabled';	options = @('set')						}
		@{ name = 'MessageRecallProcessingEnabled';														}
		@{ name = 'MessageTrackingReadStatusEnabled';													}
		@{ name = 'MicrosoftOnlineServicesID';					options = @('set')						}
		@{ name = 'ModeratedBy';								options = @('set')						}
		@{ name = 'ModerationEnabled';							options = @('set')						}
		@{ name = 'Name';										options = @('set')						}
		@{ name = 'NetID';																				}
		@{ name = 'NonCompliantDevices';						options = @('set')						}
		@{ name = 'ObjectCategory';					 													}
		@{ name = 'ObjectClass';					 													}
		@{ name = 'Office';										options = @('set')						}
		@{ name = 'OfflineAddressBook';					 												}
		@{ name = 'OrganizationalUnit';					 												}
		@{ name = 'OrganizationId';					 													}
		@{ name = 'OrphanSoftDeleteTrackingTime';					 									}
		@{ name = 'PersistedCapabilities';					 											}
		@{ name = 'PoliciesExcluded';					 												}
		@{ name = 'PoliciesIncluded';					 												}
		@{ name = 'PrimarySmtpAddress';							options = @('default','set')			}
		@{ name = 'ProhibitSendQuota';							options = @('set')						}
		@{ name = 'ProhibitSendReceiveQuota';					options = @('set')						}
		@{ name = 'ProtocolSettings';					 												}
		@{ name = 'QueryBaseDN';					 													}
		@{ name = 'QueryBaseDNRestrictionEnabled';					 									}
		@{ name = 'RecipientLimits';							options = @('set')						}
		@{ name = 'RecipientType';					 													}
		@{ name = 'RecipientTypeDetails';					 											}
		@{ name = 'ReconciliationId';					 												}
		@{ name = 'RecoverableItemsQuota';					 											}
		@{ name = 'RecoverableItemsWarningQuota';					 									}
		@{ name = 'RejectMessagesFrom';							options = @('set')						}
		@{ name = 'RejectMessagesFromDLMembers';				options = @('set')						}
		@{ name = 'RejectMessagesFromSendersOrMembers';			options = @('set')						}
		@{ name = 'RemoteAccountPolicy';					 											}
		@{ name = 'RemoteRecipientType';					 											}
		@{ name = 'RequireSenderAuthenticationEnabled';			options = @('set')						}
		@{ name = 'ResetPasswordOnNextLogon';					options = @('set')						}
		@{ name = 'ResourceCapacity';							options = @('set')						}
		@{ name = 'ResourceCustom';								options = @('set')						}
		@{ name = 'ResourceType';					 													}
		@{ name = 'RetainDeletedItemsFor';						options = @('set')						}
		@{ name = 'RetainDeletedItemsUntilBackup';					 									}
		@{ name = 'RetentionComment';							options = @('set')						}
		@{ name = 'RetentionHoldEnabled';						options = @('set')						}
		@{ name = 'RetentionPolicy';							options = @('set')						}
		@{ name = 'RetentionUrl';								options = @('set')						}
		@{ name = 'RoleAssignmentPolicy';						options = @('enable','set')				}
		@{ name = 'RoomMailboxAccountEnabled';					 										}
		@{ name = 'RulesQuota';									options = @('set')						}
		@{ name = 'SamAccountName';								options = @('enable','set')				}
		@{ name = 'SCLDeleteEnabled';					 												}
		@{ name = 'SCLDeleteThreshold';					 												}
		@{ name = 'SCLJunkEnabled';					 													}
		@{ name = 'SCLJunkThreshold';					 												}
		@{ name = 'SCLQuarantineEnabled';					 											}
		@{ name = 'SCLQuarantineThreshold';					 											}
		@{ name = 'SCLRejectEnabled';					 												}
		@{ name = 'SCLRejectThreshold';					 												}
		@{ name = 'SendModerationNotifications';				options = @('set')						}
		@{ name = 'ServerLegacyDN';					 													}
		@{ name = 'SharingPolicy';								options = @('set')						}
		@{ name = 'SiloName';					 														}
		@{ name = 'SimpleDisplayName';							options = @('set')						}
		@{ name = 'SingleItemRecoveryEnabled';					options = @('set')						}
		@{ name = 'SKUAssigned';					 													}
		@{ name = 'SourceAnchor';					 													}
		@{ name = 'StartDateForRetentionHold';					options = @('set')						}
		@{ name = 'StsRefreshTokensValidFrom';					options = @('set')						}
		@{ name = 'ThrottlingPolicy';					 												}
		@{ name = 'UMDtmfMap';					 														}
		@{ name = 'UMEnabled';					 														}
		@{ name = 'UnifiedMailbox';					 													}
		@{ name = 'UsageLocation';					 													}
		@{ name = 'UseDatabaseQuotaDefaults';					options = @('set')						}
		@{ name = 'UseDatabaseRetentionDefaults';				options = @('set')						}
		@{ name = 'UserPrincipalName';							options = @('default','enable','set')	}
		@{ name = 'WasInactiveMailbox';					 												}
		@{ name = 'WhenChanged';					 													}
		@{ name = 'WhenChangedUTC';					 													}
		@{ name = 'WhenCreated';					 													}
		@{ name = 'WhenCreatedUTC';					 													}
		@{ name = 'WhenMailboxCreated';					 												}
		@{ name = 'WhenSoftDeleted';					 												}
		@{ name = 'WindowsEmailAddress';						options = @('set')						}
		@{ name = 'WindowsLiveID';					 													}
	)
	MailboxAutoReplyConfiguration = @(
		@{ name = 'AutoDeclineFutureRequestsWhenOOF';	options = @('default','set')		}
		@{ name = 'AutoReplyState';						options = @('default','set')		}
		@{ name = 'CreateOOFEvent';						options = @('default','set')		}
		@{ name = 'DeclineAllEventsForScheduledOOF';	options = @('default','set')		}
		@{ name = 'DeclineEventsForScheduledOOF';		options = @('default','set')		}
		@{ name = 'DeclineMeetingMessage';				options = @('default','set')		}
		@{ name = 'DomainController';					options = @('set')					}
		@{ name = 'EventsToDeleteIDs';														}
		@{ name = 'EndTime';							options = @('default','set')		}
		@{ name = 'ExternalAudience';					options = @('default','set')		}
		@{ name = 'ExternalMessage';					options = @('default','set')		}
		@{ name = 'InternalMessage';					options = @('default','set')		}
		@{ name = 'OOFEventSubject';					options = @('set')					}
		@{ name = 'StartTime';							options = @('default','set')		}
		@{ name = 'Recipients';																}
		@{ name = 'ReminderMinutesBeforeStart';												}
		@{ name = 'ReminderMessage';														}
		@{ name = 'MailboxOwnerId';						options = @('default')				}
		@{ name = 'Identity';							options = @('default','key')		}
		@{ name = 'IsValid';																}
		@{ name = 'ObjectState';															}
		@{ name = 'ID';																		}
		
	)
	MailboxPermission = @(
		@{ name = 'AccessRights';    	options = @('default','add','remove')	}
		@{ name = 'Deny';            	options = @('default')					}
		@{ name = 'DomainController';	options = @('add')						}
		@{ name = 'Identity';        	options = @('default','key')			}
		@{ name = 'InheritanceType'; 	options = @('default','add')			}
		@{ name = 'IsInherited';     	options = @('default')					}
		@{ name = 'User';            	options = @('default','add','remove')	}
		@{ name = 'Owner'; 				options = @('add')						}
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

    Log info "Done"
}

function Idm-DistributionGroupsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
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
                foreach($group in $Global:DistributionGroups) {
                    $group
                }
            } else {
			    # EXO cmdlets cannot be prefixed because "EXO" is effectively a prefix already
                $groups = Get-MsExchangeDistributionGroup @call_params | Select-Object $properties
			    $groups

			    # Push group GUIDs into a global collection
			    $Global:DistributionGroups.Clear()
			    foreach($grp in $groups) {
				    [void]$Global:DistributionGroups.Add( $grp )
			    }
            }
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}

function Idm-DistributionGroupCreate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
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

    Log info "Done"
}

function Idm-DistributionGroupUpdate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
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

    Log info "Done"
}

function Idm-DistributionGroupDelete {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
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

    Log info "Done"
}

function Idm-DistributionGroupMembersRead {
	param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
	
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

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroupmember?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud			
            if($i -lt 1) {
                Log info "Retrieving Groups"
                Idm-DistributionGroupsRead > $null
            }

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

    Log info "Done"
}

function Idm-DistributionGroupMemberCreate {
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

    Log info "Done"
}

function Idm-DistributionGroupMemberDelete {
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

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud
			
            LogIO info "Get-EXOMailbox" -In @call_params
			
			# EXO cmdlets cannot be prefixed because "EXO" is effectively a prefix already
            $mailboxes = Get-EXOMailbox @call_params | Select-Object $properties
			$mailboxes
			
			# Push mailbox GUIDs into a global collection
			$Global:Mailboxes.Clear()
			foreach($mb in $mailboxes) {
				[void]$Global:Mailboxes.Add( @{ Identity = $mb.$key } )
			}
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

    Log info "Done"
}


function Idm-MailboxAutoReplyConfigurationsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
	
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

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailboxautoreplyconfiguration?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud
			
            LogIO info "Get-MsExchangeMailboxAutoReplyConfiguration" -In @call_params
			
			$i = $Global:Mailboxes.count
			$data = @()
			foreach($mb in $Global:Mailboxes) {
				$data += (Get-MsExchangeMailboxAutoReplyConfiguration -Identity $mb.Identity @call_params | Select-Object $properties)
				
				if(($i -= 1) % 100 -eq 0) {
					Log debug ("[Progress][$($Class)] $($i) remaining mailboxes to search")
				}
			}
			
			$data
			
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxAutoReplyConfigurationSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
	
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

    Log info "Done"
}


function Idm-MailboxPermissionsRead {
	param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
	
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

        try {
            # https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v Cloud
			
            LogIO info "Get-EXOMailboxPermission" -In @call_params
			
	    $data = $Global:Mailboxes.Identity | Get-EXOMailboxPermission @call_params
	    $data | Select-Object $Properties
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
	$Class = 'MailboxPermission'

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('add') } | ForEach-Object {
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
		Log info "Opening ExchangeOnline session '$connection_string'"

		$params = Copy-Object $connection_params
		$params.Certificate = Nim-GetCertificate $connection_params.certificate

		try {
			Connect-ExchangeOnline @params -Prefix 'MsExchange' -ShowBanner:$false
		}
		catch {
			Log error "Failed: $_"
			Write-Error $_
		}

		Log info "Done"
    }
}


function Close-MsExchangeSession {
	if (Get-ConnectionInformation | ? { $_.State -eq 'Connected' }) {
		Log info "Closing ExchangeOnline session"
        Disconnect-ExchangeOnline -Confirm:$false
		Log info "Done"
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
