param (
[Parameter(Mandatory=$true)] [string]$Tenant,
[Parameter(Mandatory=$true)] [string]$Path,
[string[]]$Subscriptions=@(),
[ValidateSet(5, 10, 15)] [int32]$Interval,
[switch]$Disks, 
[switch]$IPs, 
[switch]$Storages,
[switch]$VMs, 
[switch]$SQL,
[switch]$Usage, 
[switch]$Cost,
[switch]$Continue,
[int]$Top
)

$global:top = 0
$global:exPkg = $Null
$global:Interval = [TimeSpan]"00:15:00"
$titleDisks = @("Subscription","Disk Name", "Location", "Size", "Tier")
$titleIPs = @("Subscription","Name", "Type", "Tier", "Interface")
$titleStorages = @("Subscription", "Name", "Location", "Size", "Tier")
$titleVMs = @("Subscription", "Name", "Size", "Location", "OS", "Auto Shutdown")
$titleVMsUsage = @("CPU 80%-100%", "CPU 0%-20%", "RAM 80%-100%", "RAM 0%-20%")
$titleCost = "Cost (last 30 days)"
$logFile = ".\GetAzRes.tmp"
$listModules = @("Az.Billing", "Az.Accounts", "Az.Compute", "Az.Network", "Az.Resources")

. .\xlsx.ps1

# Suppress warnings
$WarningPreference = "SilentlyContinue"

# Check parameters are consistent
function CheckParams() {
# For VMs
    if (-not $VMs -and $SQL) {
        Write-Host 'The "-SQL" option makes sense only if "-VMs" option is specified'
        Exit
    }
}

# Log file for being able to continue if interrupted
$global:log = $null
function UpdateLog([string]$resId) {
    if ($global:log -eq $null) {
        $global:log = @(Get-Content -Path $logFile)
    }

    if (($global:log -match $resId).Count -le 0) {
        Add-Content $resId -Path $logFile
        $global:log += $resId
    }
}

function CheckLog([string]$resId) {
    if ($Continue) {
        if ($global:log -eq $null) {
            $global:log = @(Get-Content -Path $logFile)
        }

        return (($global:log -match $resId).Count -gt 0)
    }
    else {
        return $false
    }
}

$global:listCost=@()
function UpdateCost($listProps, $cost) {
    if ($listProps -eq $null) {
        $global:listCost=@()
    }
    else {
        $global:listCost += [pscustomobject]@{ Props=$listProps; Cost=$cost }
    }
}

function GetCost($listProps) {
    if ($listProps.Count -gt 0 -and $listProps -ne $null) {
        foreach ($cost in $global:listCost) {
            for ($i = 0; $i -lt $listProps.Count; $i++) {
                if ($listProps[$i] -eq $cost.Props[$i] -and $i -ge ($cost.Props.Count - 1)) {
                    return $cost.Cost
                }
            }
        }
    }

    return $null
}

function TitleAdd([ref][string[]]$title, [string[]]$titleAdd) {
    $title.Value += $titleAdd
}

function GetPercentageBetween($arr, $min, $max, $property) {
    $n = 0
    $invalid = 0

    for ($i = 0; $i -lt $arr.Count; $i++)  {
        $obj = $arr[$i]
        if ($obj.$property -ne $Null) {
            if ($obj.$property -ge $min -and $obj.$property -le $max) { 
                $n += 1 }
        }
        else {
            $invalid += 1
        }
    }

    if ($n -gt 0) {
        return [math]::Round(($n / ($arr.Count - $invalid)) * 100, 2)
    }

    return 0
}

function GetMinUsage($arr, $property) {
    $min = [int64]::MaxValue
    foreach ($obj in $arr) {
        if ($obj.$property -ne $Null -and $obj.$property -ne 0) {
            if ($obj.$property -lt $min) {
                $min = $obj.$property
            }
        }
    }

    return $min
}

# Get location name to display 
function GetLocation($location) {
    $loc = $global:listLocations | where {$_.Location -eq $location}
    if ($loc -eq $null) {
        $loc = ""
    }
    return $loc.DisplayName
}

function AddRawToExcel($ws, $out) {
    ExcelAdd-WorkSheetRaw $ws $out

    if ($Top -gt 0) {
        $global:top -= 1
        if ($global:top -le 0) {
            $global:top = $Top
            return $false
        }
    }

    return $true
}

# Check if all required modules are installed
function CheckModulesInstalled() {
    $ok = $true

    foreach ($module in $listModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host " Module '$module' is missing, please install it with PS command 'Install-Module $module -Force'"
            $ok = $false
        }
    }

    if (-not $ok) {
        Exit
    }
}

# Check if logged in into Azure
function CheckLogin() {
    if ((Get-AzContext).Account.Id -eq $null) {
        Write-Host " Please login in Azure with the 'Connect-AzAccount' command."
        Exit
    }
}

# Initialization
function Init() {
    Write-Host "Initializing..."
    $global:listLocations = Get-AzLocation
    $global:top = $Top

    CheckModulesInstalled

    if ($Interval -ne $null)  {
        $global:Interval = [TimeSpan]"00:00:00"
        $null = $global:Interval.Add([TimeSpan]"00:${Interval}:00")
    }

    if ($Tenant -eq $null) {
    	Write-host "Specify tenant ID!"
    	Exit
    }

    # Get allsubscriptions list
    Write-Host "Loading subscriptions list..."
    $global:listSubscriptions = Get-AzSubscription -Tenant $Tenant
    if ($global:listSubscriptions.Count -eq 0) {
        Write-host "No Subscriptions available for tenant" $Tenant
        Exit
    }

    if ($Subscriptions -gt 0) {
        $list = $global:listSubscriptions | Where-Object {$Subscriptions -contains $_}
        if ($list.Count -gt 0) {
            $global:listSubscriptions = $list
        }
        else {
            Write-host "Specified subscriptions are not found" $Subscriptions
            Exit
        }
    }

    # Convert file name to full path, if needed
    try {
        $xlFile=Get-Item -Path $Path -ErrorAction Stop
        $ErrorActionPreference = "Continue"
    }
    catch {
        $xlFile = [regex]::Match($_, "\'(.*?)\'")[0].Value.Trim("'")
    }

    # Check if continue after interruption
    if (-not $Continue) {
        "$(Get-Date)" | Set-Content -Path $logFile
        if (Test-Path $xlFile) {
            Remove-Item $xlFile -ErrorAction "SilentlyContinue"
            if (-not $?) {
                Write-Host "Cannot open output file '$xlFile'. Check if it is already open by Excel."
                Exit
            }
        }
    }

    # Excel initialization
    $global:exPkg = ExcelInit $xlFile -ErrorAction SilentlyContinue
    $global:workBook = ExcelNew-WorkBook $global:exPkg
}

# Unattached disks
function UnattachedDisks() {
    $costSum = $null
    UpdateCost

    $workSheet = ExcelAdd-WorkSheet $global:exPkg $global:workBook "UnattachedDisks"
    if (-not $Continue) {
        if ($Cost) { 
            TitleAdd ([ref]$titleDisks) @($titleCost) 
        }
        ExcelAdd-WorkSheetRaw $workSheet $titleDisks
    }

    Write-Host "`n`nProcessing unattached disks..." -NoNewLine

    foreach ($subs in $global:listSubscriptions) {
        Write-Host "`n  $($subs.Name) - reading..." -NoNewLine
        $Null = Set-AzContext -Subscription $subs

        $listDisks = (Get-AzDisk | where DiskState -eq "Unattached")
        for ($i = 0; $i -lt $listDisks.Count; $i++) {
            $disk = $listDisks[$i]

            Write-Host "`r  $($subs.Name) Disk ($($i+1))/$($listDisks.Count) $($disk.Name)                            " -NoNewLine

            # Check if this is ASR replica disk
            if ($disk.Name.EndsWith("-ASRReplica") -or $disk.Name.StartsWith("asr")) {
                continue
            }

            # Check if continue after interruption
            if (CheckLog $disk.Id) {
                continue
            }

            $out = $subs.Name, $disk.Name, (GetLocation $disk.Location), $disk.DiskSizeGB, $disk.Tier

    	    if ($disk.ManagedBy -eq $null) {
                if ($Cost) {
                    $costSum = GetCost $disk.Location, $disk.DiskSizeGB, $disk.Tier
                    if ($costSum -eq $null) {
                        try {
                            $costSum = (Get-AzConsumptionUsageDetail -InstanceId $disk.Id -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
                                       | Measure-Object PretaxCost -Sum).Sum
                            $costSum = [math]::Round($costSum, 2)
                            UpdateCost $disk.Location, $disk.DiskSizeGB, $disk.Tier $costSum
                        }
                        catch {
                            $costSum = 0
                        }
                    }
                    $out += $costSum
                }

                if (-not (AddRawToExcel $workSheet $out)) {
                    return
                }
                UpdateLog $disk.Id
    	    }
        }
    }
}

# Storage Accounts
function StorageAccounts() {
    $workSheet = ExcelAdd-WorkSheet $global:exPkg $global:workBook "Storages"
    if ($Cost)  { TitleAdd ([ref]$titleStorages) $titleCost }

    if (-not $Continue) {
        ExcelAdd-WorkSheetRaw $workSheet $titleStorages
    }

    Write-Host "`n`nProcessing Storage accounts..." -NoNewLine

    foreach ($subs in $global:listSubscriptions) {
        Write-Host "`n  $($subs.Name) - reading..." -NoNewLine
        $Null = Set-AzContext -Subscription $subs
    
        $listStorages = Get-AzStorageAccount
        for ($i = 0; $i -lt $listStorages.Count; $i++) {
            $storage = $listStorages[$i]

            Write-Host "`r  $($subs.Name) Storage Account ($($i+1))/$($listStorages.Count) $($storage.Name)                            " -NoNewLine

            # Check if continue after interruption
            if (CheckLog $storage.Id) {
                continue
            }

            $size = (Get-AzMetric -ResourceId $storage.id -MetricName "UsedCapacity" -WarningAction:SilentlyContinue).Data.Average
            $sizeMb = [math]::round($size/1Mb, 3)
            $sizeGb = [math]::round($size/1Gb, 3)
            if ($sizeGb -lt 1) {
                $size = $sizeMb.ToString() + " Mb"
            }
            else {
                $size = $sizeGb.ToString() + " Gb"
            }

            $out = $subs.Name, $storage.StorageAccountName, (GetLocation $storage.PrimaryLocation), $size, $storage.sku.Name

            if ($Cost) {
                $costSum = (Get-AzConsumptionUsageDetail -InstanceId $storage.Id -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
                           | Measure-Object PretaxCost -Sum).Sum
                $costSum = [math]::Round($costSum, 2)

                $out += "$costSum"
            }

            if (-not (AddRawToExcel $workSheet $out)) {
                return 
            }
            UpdateLog $storage.Id
        }
    }
}

# VMs
function VMs() {
    $ram80_100 = $null
    $ram0_20 = $null
    $cpu80_100 = $null
    $cpu0_10 = $null
    $costSum = $null

    $workSheet = ExcelAdd-WorkSheet $global:exPkg $global:workBook "VMs"
    if (-not $Continue) {
        if ($SQL)   { TitleAdd ([ref]$titleVMs) "SQL Agent" }
        if ($Usage) { TitleAdd ([ref]$titleVMs) $titleVMsUsage }
        if ($Cost)  { TitleAdd ([ref]$titleVMs) $titleCost }
        ExcelAdd-WorkSheetRaw $workSheet $titleVMs
    }

    Write-Host "`n`nProcessing VMs ..." -NoNewLine

    foreach ($subs in $global:listSubscriptions) {
        Write-Host "`n  $($subs.Name) - reading..." -NoNewLine
        $Null = Set-AzContext -Subscription $subs

        $listVMs = Get-AzVM
        $n = $listVMs.Count
        for ($i = 0; $i -lt $n; $i++) {
            $vm = $listVMs[$i]
            Write-Host "`r  $subs - VM ($($i+1)/$n) $($vm.Name)                           " -NoNewLine

            # Check if continue after interruption
            if (CheckLog $vm.Id) {
                continue
            }

            $size = $vm.HardwareProfile.VmSize
            $ram = (Get-AzVMSize -VMName $vm.Name -ResourceGroupName $vm.ResourceGroupName | where{$_.Name -eq $size}).MemoryInMB * 1Mb

            $autostop = ""
            $schedule = Get-AzResource -ResourceType microsoft.devtestlab/schedules -Name ("shutdown-computevm-"+$vm.Name) -ResourceGroupName $vm.ResourceGroupName -ExpandProperties -ErrorAction SilentlyContinue
            if ($schedule -ne $null) {
                $autostop = $schedule.Properties.dailyRecurrence.time.Insert(2, ":")
            }

            $out = $subs.Name, $vm.Name, $size, (GetLocation $vm.Location), $vm.StorageProfile.OsDisk.OsType, $autostop

            # Check if SQL agent is installed
            $isSQL = "No" 
            if ($SQL) {
                if ($vm.StorageProfile.OsDisk.OsType -eq "Windows") {
                    try  {
                        if ((Get-AzVMExtension -ResourceId $vm.Id).Name -contains "SqlIaasExtension") { 
                            $isSQL = "Yes" 
                        }
                    }
                    catch { }
                }
                else {
                    $isSQL = ""
                }

                $out += $isSQL
            }

            # Get RAM usage (avg / max), %
            if ($Usage) {
                try  {
                    $listRAM = Get-AzMetric -ResourceId $vm.Id -MetricName "Available Memory Bytes" -AggregationType Minimum `
                              -TimeGrain $global:Interval -StartTime (Get-Date).AddDays(-30) -EndTime (Get-Date)

                    $ram20 = $ram * 0.2
                    $ram80 = $ram * 0.8
                    # The array contains the available RAM so percentile 0 - 20% consumed RAM is between 80 and 100% of available RAM
                    $ram80_100 = GetPercentageBetween $listRAM.Data 0 $ram20 Minimum
                    $ram0_20 = GetPercentageBetween $listRAM.Data $ram80 $ram Minimum

                    # Get CPU usage, %
                    $listCPU = Get-AzMetric -ResourceId $vm.Id -MetricName "Percentage CPU" -AggregationType Maximum `
                              -TimeGrain $global:Interval  -StartTime (Get-Date).AddDays(-30) -EndTime (Get-Date)
                    $cpu80_100 = GetPercentageBetween $listCPU.Data 80 100 Maximum
                    $cpu0_10 = GetPercentageBetween $listCPU.Data 0 20 Maximum
                }
                catch {
                    $ram80_100 = 0
                    $ram0_20 = 0
                    $cpu80_100 = 0
                    $cpu0_10 = 0
                }

                $out += $cpu80_100.ToString(), $cpu0_10.ToString(), $ram80_100.ToString(), $ram0_20.ToString()
            }

            if ($Cost) {
                $costSum = (Get-AzConsumptionUsageDetail -InstanceId $vm.Id -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
                           | Measure-Object PretaxCost -Sum).Sum
                $costSum = [math]::Round($costSum, 2)

                $out += "$costSum"
            }

            if (-not (AddRawToExcel $workSheet $out)) { 
                return 
            }
            UpdateLog $vm.Id
        }
    }
}

# Unattached IPs
function UnattachedIPs() {
    $workSheet = ExcelAdd-WorkSheet $global:exPkg $global:workBook "IPs"
    if (-not $Continue) {
        if ($Cost) { TitleAdd ([ref]$titleIPs) @($titleCost) }
        ExcelAdd-WorkSheetRaw $workSheet $titleIPs
    }

    Write-Host "`n`nProcessing Unattached IP addresses..." -NoNewLine

    foreach ($subs in $global:listSubscriptions) {
        Write-Host "`n  $($subs.Name) - reading..." -NoNewLine
        $Null = Set-AzContext -Subscription $subs

        # Find all IPs without IP configuration associated with it
    	$listIPs = Get-AzPublicIpAddress | where {$_.IpConfiguration -eq $null}
        for ($i = 0; $i -lt $listIPs.Count; $i++) {
            $ip = $listIPs[$i]

            Write-Host "`r  $($subs.Name) IP ($($i+1))/$($listIPs.Count) $($ip.Name)                            " -NoNewLine

            # Check if continue after interruption
            if (CheckLog $ip.Id) {
                continue
            }

            $out = $subs.Name, $ip.Name, $ip.PublicIpAllocationMethod, $ip.Sku.Tier, ""

            if ($Cost) {
                $costSum = (Get-AzConsumptionUsageDetail -InstanceId $ip.Id -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
                           | Measure-Object PretaxCost -Sum).Sum
                $costSum = [math]::Round($costSum, 2)

                $out += "$costSum"
            }

            if (-not (AddRawToExcel $workSheet $out)) { 
                return 
            }
            UpdateLog $ip.Id
        }

        # Find all IPs assigned to unattached Network Interface
        $listNets = Get-AzNetworkInterface | where { $_.VirtualMachine -eq $null -and $_.PublicIpAddress -ne $null }
        for ($i = 0; $i -lt $listNets.Count; $i++) {
            $net = $listNets[$i]

            Write-Host "`rUnattached Network Interfaces remaining: $($listNets.Count - $i) " $net.Name "                            " -NoNewLine

            # Check if continue after interruption
            if (CheckLog $net.Id) {
                continue
            }

            $listIpConfigs = $net.IpConfigurations | where { $_.PublicIpAddress -ne $Null }
            foreach ($ipconfig in $listIpConfigs) {
                $ip = Get-AzPublicIpAddress -Name $ipconfig.PublicIpAddress.Id
                $out = $subs.Name, $ip.Name, $ip.PublicIpAllocationMethod, $ip.Sku.Tier, $net.Name

                if ($Cost) {
                    $costSum = (Get-AzConsumptionUsageDetail -InstanceId $ip.Id -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
                               | Measure-Object PretaxCost -Sum).Sum
                    $costSum = [math]::Round($costSum, 2)

                    $out += "$costSum"
                }

                if (-not (AddRawToExcel $workSheet $out)) { 
                    return 
                }
            }
            UpdateLog $net.Id
        }
    }
}

Init

try {
    if ($Disks)     { UnattachedDisks }
    if ($Storages)  { StorageAccounts }
    if ($IPs)       { UnattachedIPs }
    if ($VMs)       { VMs }
}
finally { 
# Close Excel file if Ctrl+C entered
    ExcelClose $global:exPkg 
}

# Close Excel file
ExcelClose $global:exPkg
