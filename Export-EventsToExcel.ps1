param ( 
    [string]$path = 'C:\Windows\System32\winevt\Logs', 
    [string]$output = 'C:\Users\mdoug016\OneDrive - Ascension\Desktop', 
    [switch]$show = $True 
) 
 
$logs = dir $path\Security.evtx 
 
$events = @() 
 
$events += $logs | sort length -Descending | Start-RSJob -Name {$_.BaseName} -ScriptBlock { 
        Get-WinEvent -Path $_.FullName | select @{Name='UTC';Expression={Get-Date($_.TimeCreated.ToUniversalTime()) -f o}},MachineName,ProviderName,LogName,Id,LevelDisplayName,Message         
    } | Wait-RSJob -ShowProgress | Receive-RSJob 
 
$events | sort 'UTC' -desc | Export-Excel $output\events.xlsx -Show:$show 
 
Get-RSJob | Remove-RSJob