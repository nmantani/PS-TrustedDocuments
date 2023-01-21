<#
Get-TrustedDocuments.ps1: PowerShell script to display information on trusted
documents for Microsoft Office stored in the Windows registry

Copyright (c) 2022, Nobutaka Mantani
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
   this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright
   notice, this list of conditions and the following disclaimer in the
   documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS
IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>

<#
.SYNOPSIS
Displays information on trusted documents for Microsoft Office stored in the Windows registry.

.DESCRIPTION
Retrieves information on trusted documents for Microsoft Office stored under the "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\[version]\[document type]\Security\Trusted Documents" registry keys. This script displays the status of trusted documents (editing enabled or content (macro) enabled), file creation timestamp, and status change timestamp (the last edit of the document or execution of the macro, the time resolution of the status change timestamp is minutes). This script is useful for incident response relating to malicious Microsoft Office documents by checking when and in which Microsoft Office document file a malicious macro was executed.

.PARAMETER DocumentType
Specifies the document type such as Word, Excel, and PowerPoint. Only information on the specified document type will be displayed. This parameter is case insensitive.

.PARAMETER EditingEnabledOnly
If this parameter is specified, this script will display only information on documents that editing is enabled.

.PARAMETER ContentEnabledOnly
If this parameter is specified, this script will display only information on documents that content (macro) is enabled.

.PARAMETER User
Specifies the user. If both this parameter and the HiveFilePath parameter are not specified, this script will display information for the current user. Administrator privilege is required to use this parameter to display information for another user.

.PARAMETER HiveFilePath
Specifies the path of an offline registry hive file (NTUSER.DAT file extracted from another computer) to display information. If both this parameter and the User parameter are not specified, this script uses the HKEY_CURRENT_USER registry hive of the current user. Administrator privilege is required to use this parameter because this script temporarily loads the offline registry hive file into the "HKEY_USERS\PS-TrustedDocuments" registry key.

.EXAMPLE
powershell -ExecutionPolicy Bypass .\Get-TrustedDocuments.ps1

Description
---------------------------------------------------------------
Displaying information on all trusted documents for the current user.

.EXAMPLE
powershell -ExecutionPolicy Bypass .\Get-TrustedDocuments.ps1 -DocumentType word

Description
---------------------------------------------------------------
Displaying information on trusted Word documents for the current user.

.EXAMPLE
powershell -ExecutionPolicy Bypass .\Get-TrustedDocuments.ps1 -EditingEnabledOnly

Description
---------------------------------------------------------------
Displaying information on trusted documents that editing is enabled.

.EXAMPLE
powershell -ExecutionPolicy Bypass .\Get-TrustedDocuments.ps1 -ContentEnabledOnly

Description
---------------------------------------------------------------
Displaying information on trusted documents that content (macro) is enabled.

.EXAMPLE
powershell -ExecutionPolicy Bypass .\Get-TrustedDocuments.ps1 -User exampleuser

Description
---------------------------------------------------------------
Displaying information on trusted documents for the specified user.
Administrator privilege is required to display information for another user.

.EXAMPLE
powershell -ExecutionPolicy Bypass .\Get-TrustedDocuments.ps1 -HiveFilePath .\extracted\NTUSER.DAT

Description
---------------------------------------------------------------
Displaying information on trusted documents from an offline registry hive file.
Administrator privilege is required.
#>

Param(
    [String]$DocumentType,
    [Switch]$EditingEnabledOnly,
    [Switch]$ContentEnabledOnly,
    [String]$User,
    [String]$HiveFilePath
)

function check_mutually_exclusive_options($name_a, $value_a, $name_b, $value_b) {
    if ($value_a -and $value_b) {
        Write-Output "Error: $name_a and $name_b cannot be used at the same time."
        exit
    }
}

function check_privilege {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $is_admin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (!$is_admin) {
        Write-Output "Error: administrator privilege is required."
        exit
    }
}

function load_hivefile($path) {
    if (Test-Path $path) {
        $ret = reg load HKU\PS-TrustedDocuments $path
        if ($null -eq $ret) {
            Write-Output "Error: failed to load $path."
            exit
        }
    } else {
        Write-Output "Error: $path is not found."
        exit
    }
}

function unload_hivefile {
    [gc]::collect()
    $null = reg unload HKU\PS-TrustedDocuments
}

function set_regpath_prefix($path, $user) {
    if ($path -ne "") {
        check_privilege
        load_hivefile $path
        $regpath_prefix = "Registry::HKU\PS-TrustedDocuments\SOFTWARE\Microsoft\Office"
    } elseif ($user -ne "") {
        if ($user -ne $env:USERNAME) {
            check_privilege
        }

        if(Get-Command 'Get-CimInstance' -ErrorAction SilentlyContinue) {
            $sid = Get-CimInstance win32_useraccount | Where-Object {$_.name -eq $user} | Select-Object sid -ExpandProperty sid
        } else {
            $sid = Get-WmiObject win32_useraccount | Where-Object {$_.name -eq $user} | Select-Object sid -ExpandProperty sid
        }

        if ($null -eq $sid) {
            Write-Output "Error: user $user not found."
            exit
        } else {
            $regpath_prefix = "Registry::HKU\$sid\SOFTWARE\Microsoft\Office"
        }
    } else {
        $regpath_prefix = "Registry::HKCU\SOFTWARE\Microsoft\Office"
    }

    return $regpath_prefix
}

function show_no_information_message($path, $user) {
    if ($path -eq "") {
        if ($user -ne "") {
            Write-Output "There is no information on trusted documents for user $user."
        } else {
            Write-Output "There is no information on trusted documents for user $env:USERNAME."
        }
    } else {
        Write-Output "There is no information on trusted documents in $path."
        unload_hivefile
    }
}

function bytes_to_status_change_filetime($bytes) {
    [UInt64]$timestamp_unix_min = [System.BitConverter]::ToUint32($bytes,16) # offset 16-19
    $timestamp_unix = $timestamp_unix_min * 60
    $timestamp_filetime = ($timestamp_unix * 10000000) + 116444736000000000 # conversion from UNIXTIME to FILETIME

    return $timestamp_filetime
}

check_mutually_exclusive_options "-EditingEnabledOnly" $EditingEnabledOnly "-ContentEnabledOnly" $ContentEnabledOnly
check_mutually_exclusive_options "-User" $User "-HiveFilePath" $HiveFilePath

$regpath_prefix = set_regpath_prefix $HiveFilePath $User

if (!(Test-Path $regpath_prefix)) {
    show_no_information_message $HiveFilePath $User
    exit
}

if ($DocumentType -ne "") {
    $keys = Resolve-Path "$regpath_prefix\*\$DocumentType\Security\Trusted Documents"
} else {
    $keys = Resolve-Path "$regpath_prefix\*\*\Security\Trusted Documents"
}

if ($null -eq $keys) {
    show_no_information_message $HiveFilePath $User
    exit
}

# This is necessary to use [System.Web.HttpUtility]::UrlDecode() with older version of PowerShell
Add-Type -AssemblyName System.Web

$count = 0
$not_found = $true
foreach ($k in $keys) {
    $item = Get-ChildItem $k

    if ($null -eq $item) {
        continue
    } else {
        $not_found = $false
    }

    foreach ($p in $item.Property) {
        $path = [System.Web.HttpUtility]::UrlDecode($p)
        $bytes = $item.GetValue($p) # Byte[]
        $timestamp = [System.BitConverter]::ToUint64($bytes,0) # offset 0-7
        $creation_timestamp_utc = [datetime]::FromFileTimeUtc($timestamp).toString("yyyy/MM/dd HH:mm:ss")

        $flag = [System.BitConverter]::ToInt32($bytes, 20) # offset 20-23
        $editing_enabled = $false
        $content_enabled = $false
        if ($flag -eq 1) { # \x01\x00\x00\x00 (little endian)
            $editing_enabled = $true
        } elseif ($flag -eq 2147483647) {# \xff\xff\xff\x7f (little endian)
            $content_enabled = $true
        }

        if ($editing_enabled -and -not $ContentEnabledOnly) {
            if ($count -gt 0) {
                Write-Output ""
            }
            Write-Output "File path: $path"
            Write-Output "Status: editing enabled"
        } elseif ($content_enabled -and -not $EditingEnabledOnly) {
            if ($count -gt 0) {
                Write-Output ""
            }
            Write-Output "File path: $path"
            Write-Output "Status: content (macro) enabled"
        }

        $timezone_offset = [System.BitConverter]::Toint64($bytes,8) # offset 8-15
        $timezone_offset_hour = $timezone_offset / 10000000 / 60 / 60
        $creation_timestamp_local = [datetime]::FromFileTimeUtc($timestamp + $timezone_offset).toString("yyyy/MM/dd HH:mm:ss")

        if (($editing_enabled -and -not $ContentEnabledOnly) `
            -or ($content_enabled -and -not $EditingEnabledOnly)) {
            Write-Output "File creation timestamp (UTC): $creation_timestamp_utc"
            if ($timezone_offset_hour -ne 0) {
                if ($timezone_offset_hour -gt 0) {
                    Write-Output "File creation timestamp (localtime, UTC+${timezone_offset_hour}): $creation_timestamp_local"
                } else {
                    Write-Output "File creation timestamp (localtime, UTC${timezone_offset_hour}): $creation_timestamp_local"
                }
            }
        }

        $status_timestamp_filetime = bytes_to_status_change_filetime $bytes
        $status_timestamp_utc = [datetime]::FromFileTimeUtc($status_timestamp_filetime).toString("yyyy/MM/dd HH:mm")
        $status_timestamp_local = [datetime]::FromFileTimeUtc($status_timestamp_filetime + $timezone_offset).toString("yyyy/MM/dd HH:mm")

        if (($editing_enabled -and -not $ContentEnabledOnly) `
            -or ($content_enabled -and -not $EditingEnabledOnly)) {
            Write-Output "Status change timestamp (UTC): $status_timestamp_utc"

            if ($timezone_offset_hour -ne 0) {
                if ($timezone_offset_hour -gt 0) {
                    Write-Output "Status change timestamp (localtime, UTC+${timezone_offset_hour}): $status_timestamp_local"
                } else {
                    Write-Output "Status change timestamp (localtime, UTC${timezone_offset_hour}): $status_timestamp_local"
                }
            }

            $count += 1
        }
    }

    # This is necessary to unload the offline registry hive file
    if ($HiveFilePath -ne "" -and $null -ne $item) {
        if ([Version] $PSVersionTable.PSVersion -gt "2.0") {
            $item.dispose()
        }
        $item.close()
    }
}

if ($not_found) {
    show_no_information_message $HiveFilePath $User
}

if ($HiveFilePath -ne "") {
    unload_hivefile
}
