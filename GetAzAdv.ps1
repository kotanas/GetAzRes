param (
[Parameter(Mandatory=$true)] [string]$Tenant,
[Parameter(Mandatory=$true)] [string]$Path,
[string[]]$Subscriptions=@(),
[switch]$Continue,
[int]$Top
)

$global:top = 0
$global:exPkg = $Null
$titleAdv = @("Subscription","Advisory", "Impacted Resources", "Potential yearly savings")
$logFile = ".\GetAzRes.tmp"

. .\xlsx.ps1

# Suppress warnings
$WarningPreference = "SilentlyContinue"

# Check parameters are consistent
function CheckParams() {
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

function TitleAdd([ref][string[]]$title, [string[]]$titleAdd) {
    $title.Value += $titleAdd
}

function AddRawToExcel($ws, $out) {
    if ($Top -gt 0) {
        if ($global:top -gt 0) {
            ExcelAdd-WorkSheetRaw $ws $out
            $global:top -= 1
        }
        else {
            $global:top = $Top
            return $false
        }
    }
    else {
        ExcelAdd-WorkSheetRaw $ws $out
    }

    return $true
}

# Initialization
function Init() {
    Write-Host "Initializing..."
    $global:top = $Top

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
    }
    catch {
        $xlFile = [regex]::Match($_, "\'(.*?)\'")[0].Value.Trim("'")
    }

    # Check if continue after interruption
    if (-not $Continue) {
        "$(Get-Date)" | Set-Content -Path $logFile
        if (Test-Path $xlFile) {
            Remove-Item $xlFile
        }
    }

    # Excel initialization
    $global:exPkg = ExcelInit $xlFile
    $global:workBook = ExcelNew-WorkBook $global:exPkg
}

function Subscriptions_subscriptions($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        if (($rec.ExtendedProperty.AdditionalProperties["term"] -eq "P3Y") -and ($rec.ExtendedProperty.AdditionalProperties["lookbackPeriod"] -eq 30)) {
            $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
            $resNum += 1
        }
    }

    return $annualSum, $resNum
}

function ReservedInstances_reservedInstances($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        if ($rec.ExtendedProperty.AdditionalProperties["term"] -eq "P3Y") {
            $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
            $resNum += 1
        }
    }

    return $annualSum, $resNum
}

function Compute_virtualMachineScaleSets() {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
        $resNum += 1
    }

    return $annualSum, $resNum
}

function Compute_virtualMachines($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
        $resNum += 1
    }

    return $annualSum, $resNum}

function Documentdb_databaseaccounts($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
        $resNum += 1
    }

    return $annualSum, $resNum
}


function Compute_disks($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
        $resNum += 1
    }

    return $annualSum, $resNum
}

function Kusto_Clusters($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
        $resNum += 1
    }

    return $annualSum, $resNum
}

function Compute_snapshots($listRecommendations) {
    $annualSum = 0
    $resNum = 0

    foreach ($rec in $listRecommendations) {
        $annualSum += $rec.ExtendedProperty.AdditionalProperties["annualSavingsAmount"]
        $resNum += 1
    }

    return $annualSum, $resNum
}

# Get Cost Recommendations
function CostRecommendations() {
    Write-Host "Adding cost recommendations"

    $workSheet = ExcelAdd-WorkSheet $global:exPkg $global:workBook "Recommendations"
    if (-not $Continue) {
        ExcelAdd-WorkSheetRaw $workSheet $titleAdv
    }

    foreach ($subs in $global:listSubscriptions) {
        Write-Host "  Subscription $($subs.Name) ..."

        # Check if continue after interruption
        if (CheckLog $subs.Id) {
            continue
        }

        $listAdv = (Get-AzAdvisorRecommendation -SubscriptionId $subs.Id) | where { $_.Category -eq "Cost" }

        # Calculate total amount for each resource type
        $listSolutions = $listAdv.ShortDescriptionSolution | Select-Object -Unique
        $totalAnnualSum = 0
        foreach ($solution in $listSolutions) {
            $annualSum = 0
            $resNum = 0
            $listRecommendations = $listAdv | where { $_.ShortDescriptionSolution -eq $solution }
            if ($listRecommendations.Count -gt 0) {
                switch ($listRecommendations[0].ImpactedField) {
                    "Microsoft.ReservedInstances/reservedInstances" { $annualSum, $resNum = (ReservedInstances_reservedInstances $listRecommendations) }
                    "microsoft.documentdb/databaseaccounts" { $annualSum, $resNum = (Documentdb_databaseaccounts $listRecommendations) }
                    "Microsoft.Subscriptions/subscriptions" { $annualSum, $resNum = (Subscriptions_subscriptions $listRecommendations) }
                    "Microsoft.Compute/virtualMachineScaleSets" { $annualSum, $resNum = (Compute_virtualMachineScaleSets $listRecommendations) }
                    "Microsoft.Compute/virtualMachines"  { $annualSum, $resNum = (Compute_virtualMachines $listRecommendations) }
                    "Microsoft.Compute/disks" { $annualSum, $resNum = (Compute_disks $listRecommendations) }
                    "Microsoft.Kusto/Clusters" { $annualSum, $resNum = (Kusto_Clusters $listRecommendations) }
                    "Microsoft.Compute/snapshots" { $annualSum, $resNum = (Compute_snapshots $listRecommendations) }
                }
            }

            if ($annualSum -gt 0) {
                $out = $subs.Name, $solution, $resNum, $annualSum
                ExcelAdd-WorkSheetRaw $workSheet $out
            }

            $totalAnnualSum += $annualSum
        }

        if ($totalAnnualSum -gt 0) {
            ExcelAdd-WorkSheetRaw $workSheet "Total:", "", "", $totalAnnualSum
            ExcelAdd-WorkSheetRaw $workSheet ""
        }
    }
}

Init

try {
    CostRecommendations
}
finally { 
# Close Excel file if Ctrl+C entered
    ExcelClose $global:exPkg 
}

# Close Excel file
ExcelClose $global:exPkg
