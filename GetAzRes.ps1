This script collects information about Azure services and generates an Excel report. The following services are currently supported:
- Storage Accounts
- VMs
- Disks
- Public IPs

Installation.
You need PowerShell v7 with and the following modules installed:
- Az.Accounts
- Az.Billing
- Az.Compute
- Az.Network
- Az.Resources

You can check if the modules are installed with "Get-Module" command. Pelase run "Install-Module <MODULENAME>" and "Import-Module <MODULENAME>" to install missed modules. You need Reader access to the subscriptions you want to process. 

Execution.
You need to place the files GetAzRes.ps1 and xlsx.ps1 in the same directory. Don't forget to run the "Connect-AzAccount" command before executing the script.
.\GetAzRes.ps1 -Path <FILENAME> -Tenant <TENANT GUID> -Subscriptions <SUBSCRIPTION GUID> -Disks -VMs -IPs -Storages -Cost -Usage -Top <MAX NUMBER> -Continue

Below please find description of the options.
- -Path             Mandatory. Filename for the Excel report file.
- -Tenant           Mandatory. Id of the AAD tenant.
- -Subscriptions    Optional. List of the subscriptions that will be processed. All the subscriptions will be processed if the argument is not specified.
- -Disks            Optional. Collect information about disks.
- -VMs              Optional. Collect information about Virtual Machines.
- -IPs              Optional. Collect information about public IP addresses.
- -Storages         Optional. Collect information about Storage Accounts.
- -Cost             Optional. Calculate cost for each instance (i.e. each VM, Disk, IP, Storage Account) for the last 30 days. This option is time consuming since it invokes the requests that may take tens of seconds per instance.
- -Usage            Optional. Calculate resource usage for each VM for the last 30 days. This option is time consuming since it invokes the requests that may take tens of seconds per instance. The following usage characteristics are provided:
    - CPU usage - Percentage of time a VM consumes between 0% and 10% of CPU power. Percentage of time a VM consumes between 80% and 100% of CPU power.
    - RAM usage - Percentage of time a VM consumes between 0% and 20% of memory. Percentage of time a VM consumes between 80% and 100% of memory.
- -Top              Optional. Specify maximum number of instances in each category you want to process. The option is helpful if you want rough estimation of the script's execution time.
- -Continue         Optional. Sometimes authorization tocken is expired or something the script hangs executing an Az command. You can interrupt it with Ctrl+C and re-run again with this option. It will resume starting with where it was interrupted.
