# Citrix Admin Audit Utility

Read-only troubleshooting utility for Citrix administrators.  
The main entry point is a batch file that uses a PowerShell helper in the background.

## Files

- `Citrix_Admin_Audit_Utility.bat`
- `Citrix_Admin_Audit_Helper.ps1`

Keep both files in the same folder.

## Intended use

This utility is designed primarily for:

- Delivery Controllers
- Studio servers with the Citrix SDK installed
- VDA servers for local runtime checks

For separate Studio installations, you can set an **AdminAddress** target in the menu and point the script to a Delivery Controller.

## Main checks

- Site overview and licensing indicators
- Controllers and Citrix SDK service status
- Machine catalogs and delivery groups
- Machine registration, maintenance mode, and power state
- Current and disconnected sessions
- Access, entitlement, and assignment policy rules
- MCS provisioning schemes and provisioning tasks
- Hypervisor connections and hosting units
- Local Citrix software, VDA registry, and App Layering presence indicators
- Local Citrix-related Windows services and event logs
- Full report export to the desktop

## Notes

- The utility is intentionally **read-only**.
- It uses `Get-BrokerMachine` instead of the deprecated `Get-BrokerDesktop`.
- App Layering checks are **presence-based** only. They do not validate ELM health or connector communication.
- Provisioning task output can be limited for catalogs created through Web Studio.

## Requirements

- Windows PowerShell
- Citrix SDK / Studio / Controller-side snap-ins for site-wide Citrix cmdlets
- Administrative rights recommended
- Network access to the target Delivery Controller if `AdminAddress` is used


