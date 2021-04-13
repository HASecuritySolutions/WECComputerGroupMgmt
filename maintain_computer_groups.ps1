# Specify the computer groups to manage. These are the groups that should be utilized to subscribe
# computers to Windows Event Collectors. The main goal of this script is to load balance which
# computers get added to the below computer groups.
$computer_groups = @("MachineGroup1","MachineGroup2","MachineGroup3","MachineGroup5")

# Should the groups be created if they do not exist. A value of 1 will cause this script to
# create any missing groups
$create_missing_groups = 0

# Set the default search base distinguished name. This will limit the initial grab of all computers to this DN
$base_ou = "DC=tyrellcorp,DC=us"

# Explicit ignore list - Any computer name in this list will be excluding from being added to the computer groups
# Example - @("computer1","server1","DC03")
$excluded_computers = @("TestComp1","TestComp2")

# Any computer account found in the below OUs will be classified as a server - Set to @() to not classify by OU
$server_ous = @("OU=Servers,OU=Test,DC=tyrellcorp,DC=us","OU=Servers,OU=Home Office,DC=tyrellcorp,DC=us")
# Any computer account found in the below OUs will be classified as a workstation - Set to @() to not classify by OU
$workstation_ous = @("OU=Workstations,OU=Test,DC=tyrellcorp,DC=us", "OU=IT,OU=Home Office,DC=tyrellcorp,DC=us", "OU=HR,OU=Home Office,DC=tyrellcorp,DC=us")

# Anything matching the regex below will be classified as a server - This check only applies to things not identified
# with a server OU. Server OU processes first. This processes against computer not found by an OU check
$server_regex = @('^(A|B)DFS','^DC[0-9]+')
# Anything matching the regex below will be classified as a workstation - This check only applies to things not identified
# with a workstation OU. Workstation OU processes first. This processes against computer not found by an OU check
$workstation_regex = @('^TestComp[0-9]$','^HOIT[0-9]+','^HOACCT[0-9]+')

# Remaining check - How should remaining computers be treated? Set to workstation, server, or ignore
# If set to server, all remaining computers will be treated as if they are a server
# If set to workstation, all remaining computers will be treated as if they are a workstation
# If set to ignore, all remaining computers will be ignored and not added to the computer groups
$remaining_assets = "ignore"

# What is the name of the field you are storing point totals in - Recommendation is Description or 
$server_points = 4
$workstation_points = 1

# Do not change below this line

# Create computer accounts for testing - You can use this by manually running this function
function New-ComputerAccounts($number_of_accounts) {
    1..$number_of_accounts | ForEach-Object { New-ADComputer -Name "TestComp$_" -SAMAccountName "TestComp$_" }
}

# Start by checking if all the groups exist - Only runs if $create_missing_groups is 1
if ($create_missing_groups -eq 1){
    $computer_groups | ForEach-Object {
        $computer = $_
        try {
            Get-ADGroup -Identity $_ | Out-Null
        } 
        catch {
            Write-Host "$computer does not exist"
            New-ADGroup -Name $computer -GroupScope Global -GroupCategory Security
        }
    }
}
# Next, grab all computers and store them in $all_computer_assets
$all_computer_assets = Get-ADComputer -Filter * -SearchBase $base_ou


# Function takes an object list of assets and checks if they are workstations
# or servers - Property and point value added to object
function Get-AssetType($assets){
    foreach ($asset in $assets) {
        # BEGIN OU CHECKS
        if ($asset.DistinguishedName.IndexOf('OU=',[System.StringComparison]::CurrentCultureIgnoreCase) -ge 0){
            $ou = $asset.DistinguishedName.Substring($asset.DistinguishedName.IndexOf('OU=',[System.StringComparison]::CurrentCultureIgnoreCase))
            # First, check if asset OU is in server OUs
            if ($ou -in $server_ous){
                Add-Member -InputObject $asset -Name "AssetType" -Value "Server" -MemberType NoteProperty -Force
                Add-Member -InputObject $asset -Name "Points" -Value $server_points -MemberType NoteProperty -Force
                Add-Member -InputObject $asset -Name "MatchOn" -Value "ServerOU" -MemberType NoteProperty -Force
                continue
            }
            # Next, check if asset OU is in the workstation OUs
            if ($ou -in $workstation_ous){
                Add-Member -InputObject $asset -Name "AssetType" -Value "Workstation" -MemberType NoteProperty -Force
                Add-Member -InputObject $asset -Name "Points" -Value $workstation_points -MemberType NoteProperty -Force
                Add-Member -InputObject $asset -Name "MatchOn" -Value "WorkstationOU" -MemberType NoteProperty -Force
                continue
            }
        }
        # BEGIN REGEX CHECKS
        $found = 0
        # First, check if asset name matches server regex patterns
        if ($server_regex.Count -ne 0){
            foreach ($regex_pattern in $server_regex) {
                if ($asset.Name | Select-String -Pattern $regex_pattern){
                    Add-Member -InputObject $asset -Name "AssetType" -Value "Server" -MemberType NoteProperty -Force
                    Add-Member -InputObject $asset -Name "Points" -Value $server_points -MemberType NoteProperty -Force
                    Add-Member -InputObject $asset -Name "MatchOn" -Value "ServerRegex" -MemberType NoteProperty -Force
                    $found = 1
                    break
                }
            }
            if ($found -eq 1){
                continue
            }
        }
        # Next, check if asset name matches workstation regex patterns
        $found = 0
        if ($workstation_regex.Count -ne 0){
            foreach ($regex_pattern in $workstation_regex) {
                if ($asset.Name | Select-String -Pattern $regex_pattern){
                    Add-Member -InputObject $asset -Name "AssetType" -Value "Workstation" -MemberType NoteProperty -Force
                    Add-Member -InputObject $asset -Name "Points" -Value $workstation_points -MemberType NoteProperty -Force
                    Add-Member -InputObject $asset -Name "MatchOn" -Value "WorkstationRegex" -MemberType NoteProperty -Force
                    $found = 1
                    break
                }
            }
            if ($found -eq 1){
                continue
            }
        }
        
        # BEGIN FALLBACK CHECK
        # Apply remaining classification filter
        if ($remaining_assets -eq "server") {
            Add-Member -InputObject $asset -Name "AssetType" -Value "Server" -MemberType NoteProperty -Force
            Add-Member -InputObject $asset -Name "Points" -Value $server_points -MemberType NoteProperty -Force
            Add-Member -InputObject $asset -Name "MatchOn" -Value "ServerFallback" -MemberType NoteProperty -Force
            continue
        }
        if ($remaining_assets -eq "workstation") {
            Add-Member -InputObject $asset -Name "AssetType" -Value "Workstation" -MemberType NoteProperty -Force
            Add-Member -InputObject $asset -Name "Points" -Value $workstation_points -MemberType NoteProperty -Force
            Add-Member -InputObject $asset -Name "MatchOn" -Value "WorkstationFallback" -MemberType NoteProperty -Force
            continue
        }
        if ($remaining_assets -eq "ignore") {
            $asset | Add-Member -NotePropertyName "AssetType" -NotePropertyValue "Ignore" -Force -PassThru
            Add-Member -InputObject $asset -Name "Points" -Value 0 -MemberType NoteProperty -Force
            Add-Member -InputObject $asset -Name "MatchOn" -Value "IgnoreFallback" -MemberType NoteProperty -Force
            continue
        }
    }
    return $assets
}
$assets = Get-AssetType $all_computer_assets

# Get current membership and asset type of computer groups
$groups = @{}
$count = 0
foreach ($group in $computer_groups){
    $members = Get-ADGroupMember -Identity $group
    $groups[$group] = @{}
    $groups[$group]['members'] = [array](Get-AssetType $members)
    if (($groups[$group]['members']).Count -eq 0){
        $sum = 0
    } else {
        $sum = [int]($groups[$group]['members'] | Measure-Object -Property Points -Sum | Select-Object -ExpandProperty Sum)
    }
    $groups[$group]['points'] = $sum
    $count += 1
}
# Find all computers that are not assigned to the computer groups
$available_computers = $assets
$groups.Keys | Foreach-Object {
    if(($groups[$_]['members']).Count -ge 1){
        foreach($member in $groups[$_]['members']) {
            if($member.Name -in $available_computers.Name){
                $available_computers = $available_computers | Where-Object { $_.Name -ne $member.Name}
            }
        }
    }
}
$total_points = $assets | Measure-Object -Property Points -Sum | Select-Object -ExpandProperty Sum
$group_count = $computer_groups.Count
$points_per_server = $total_points / $group_count

# Check for overallocation, this should not happen unless someone manually modifies the groups
$count = 0
$new_total_points = $total_points
$groups.Keys | Foreach-Object {
    $consumable_points = $points_per_server - $groups[$_]['points']
    if($consumable_points -lt 0){
        $consumable_points
        $new_total_points = $new_total_points - $groups[$_]['points']
        $count += 1
    }
}
$new_count = $group_count - $count
$new_points_per_server = $new_total_points / $new_count

Write-Host "INFO - Total possible points is $total_points which should be spread across $group_count WEC instances"
Write-Host "INFO - Recommended load should be $points_per_server per WEC instance"
if ($new_total_points -gt $total_points + $server_points){
    Write-Host "INFO - Adjusted Total possible points is $new_total_points which should be spread across $new_count WEC instances"
    Write-Host "Recommended load should be $new_points_per_server per WEC instance"
}

# Allocate computers to the groups with weight distribution
$allocated_computers = @()
$margin = $points_per_server + $server_points
$groups.Keys | Foreach-Object {
    $points_consumed = $groups[$_]['points']
    Write-Host "STATUS $_ currently has $points_consumed points consumed"
    if($points_consumed -lt $new_points_per_server){
        foreach ($computer in $available_computers) {
            if($points_consumed -ge $new_points_per_server){
                Write-Host "ALLOCATION - $_ ending with $points_consumed points - Recommended level is $margin"
                break
            } else {
                Add-ADGroupMember -Identity $_ -Members $computer
                $consumable_points = $consumable_points - $computer.Points
                $allocated_computers += $computer
                $points_consumed = $points_consumed + $computer.Points
            }
        }
        foreach($allocation in $allocated_computers){
            if($allocation.Name -in $available_computers.Name){
                $available_computers = $available_computers | Where-Object { $_.Name -ne $allocation.Name}
            }
        }
    } else {
        if($points_consumed -gt $total_points + $server_points){
            Write-Host "NO ALLOCATION - $_ exceeds recommended maximum point allocation of $points_per_server. Current points is $points_consumed"
        } else {
            Write-Host "NO ALLOCATION - $_ within acceptable range at $points_consumed and recommended upper value of $margin"
        }
    }
}
