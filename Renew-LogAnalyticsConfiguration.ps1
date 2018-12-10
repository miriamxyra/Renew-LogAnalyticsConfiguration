<#
.SYNOPSIS
	A brief description of the Renew-LogAnalyticsConfiguration function.
	
.DESCRIPTION
	A detailed description of the Renew-LogAnalyticsConfiguration function.
	
.PARAMETER NewWorkspaceId
        Enter the Workspace Id of the new workspace.
        
.PARAMETER NewWorkspaceKey
        Enter the Workspace Key of the new workspace.
        
.PARAMETER OldWorkspaceId
        Enter the the OldWorkspaceId of the Workspace you want to delete.
	
.PARAMETER Computer
	If this parameter is left empty, the script will be run on the local computer. Accepts a String or a CSV. Computername must be resolved by DNS.
	
.EXAMPLE
	PS C:\> Renew-LogAnalyticsConfiguration -OldWorkspaceId 'value1' -NewWorkspaceId 'value2' -NewWorkspaceKey 'value3' -Computer "pc1, pc2, pc3"
		
.NOTES
	Additional information about the function.
#>

param (
    
    [Parameter(Mandatory=$true)]
    [string]$NewWorkspaceId,

    [Parameter(Mandatory=$true)]
    [string]$NewWorkspaceKey,

    [Parameter(Mandatory=$true)]
    [string]$OldWorkspaceId,

    [string]$Computer,

    [pscredential]$Credential
    
)

function ConnectToPCAndCleanUp ($ComputerName) {

    $sessionParameters = @{
        ComputerName = $ComputerName
        ErrorAction  = 'Stop'
        Name         = 'MMASession'
    }

    if ($Credential) {
        $sessionParameters.Add('Credential', $Credential)
    }

    
    try {
        $session = New-PSSession @sessionParameters
    }
    catch {
        write-verbose ('Error establishing connection to {0}. Error message was {1}' -f $ComputerName, $_.Exception.Message) 
        Write-Error -Message ('Error establishing connection to {0}. Error message was {1}' -f $ComputerName, $_.Exception.Message) -Exception $_.Exception -TargetObject $ComputerName
        return $null
    }

    try {
        Invoke-Command -Session $session -ScriptBlock {$mma = New-Object -ComObject 'AgentConfigManager.MgmtSvcCfg'}
        Invoke-Command -Session $session -ScriptBlock {$mma.RemoveCloudWorkspace($Using:OldWorkspaceId)}
        Invoke-Command -Session $session -ScriptBlock {$mma.AddCloudWorkspace($Using:NewWorkspaceId, $Using:NewWorkspaceKey)}
        Invoke-Command -Session $session -ScriptBlock {$mma.ReloadConfiguration()}
    }
    catch {
        write-verbose ('Error executing command on {0}. Assuming issue with the connection. Error was {1}' -f $ComputerName, $_.Exception.Message)
        Write-Error -Message ('Error executing command on {0}. Assuming issue with the connection. Error was {1}' -f $ComputerName, $_.Exception.Message)
    }

}

if (!$Computer) {
    $Computer = "localhost"
}

If ($Computer -match ",") {
    $computers =$computer.Split(",")
    foreach ($name in $Computers) {
        ConnectToPCAndCleanUp $name
    }
}
else {
    ConnectToPCAndCleanUp $Computer
}


