Param(
    # Application Settings
    [String][ValidateNotNullOrEmpty()]$Subnet="129.120.55.0" #Target Subnet
   
) #end param

#Parameters for potentially unused registrations
$CHECK_LASTLOGIN = (Get-Date).adddays(-60); #Check if computer was last logged in older than X days ago
$DNSRegistration = @()

#Hold objects we believe are not using registration anymore
$UnusedRegistration = @()

if (($Subnet -Match "^([0-9]{1,3}\.){3}") -eq $True) {
    #Grab every DNS Record for the subnet
    $ScriptBlock = {
        param($Name)
        Resolve-DnsName -Name $Name
    }
    for ($i=1; $i -lt 255; $i++) {
        #We throttle because computers suck
        While ((Get-job -State "Running").Count -gt 30) { Start-Sleep 2 }
        
        $CurrentIP = ($Matches[0] + $i)
        Start-Job $ScriptBlock -ArgumentList $CurrentIP
    }
}
# Wait for all to complete
While (Get-Job -State "Running") { Start-Sleep 2 }

$DNSRegistration += (Get-Job | Receive-Job)
Remove-Job *

if ($DNSRegistration.Length -eq 0) {
    Throw "No DNS records were returned, User Input error?"
}

#Progress bars are cool
$prog = 0
#Iterate through all the objects
foreach ($Obj in $DNSRegistration) {
    
    Write-Progress -Activity "DNS Report" -Status $Obj.Name -PercentComplete (($prog/$DNSRegistration.length)*100)
    $prog++

    $ScriptBlock = {
        param ($Obj)

        $ADComputer = $null
        $result = $null
        $AITComputer = $null
        $Matches = $null
        $ReturnObject = $null #holds single object we believe isnt used

        #Match all hostnames denoting computers using RegEx (per our personal naming convention)
        $result = $Obj.NameHost -Match "^([^.]*)" #Grab everything before the first period
        if ($result) {
            
            #Can we ping it?
            $TestPing = Test-Connection $Obj.NameHost -Count 1 -Quiet

            #need to stringify before using get-adcomputer otherwise it throws a fit
            $CompName = $Matches[0].ToString()

            #Check if it's an AIT computer, if it is grab the service tag only
            $AITComputer = $Obj.NameHost -Match "^(([a-z]){3}-){2}(.{1,7})(?=\.)"
            if ($AITComputer) {
                $CompName = "*" + $CompName.Substring($CompName.length-7,7)
            }
            try {
                $ADComputer = Get-ADComputer -Filter {(Name -like $CompName)} -Properties Name,Enabled,operatingSystem,lastLogonTimeStamp -searchBase "OU=UNT,DC=unt,DC=ad,DC=unt,DC=edu" -Verbose
            } catch {
                $ReturnObject = New-Object PSObject -Property @{
                    Machine      =    $null
                    FQDN         =    $Obj.NameHost
                    IP           =    $Obj.Name
                    Enabled      =    $false
                    OS           =    $null
                    Comments     =    $_
                }
                return $ReturnObject
            }
            
            #Check for returned object and run tests
            if ($ADComputer -ne $NULL) {
                
                $newest = $null #Hold variable to check which one is newer

                #If we get back an array we need to do some evaluations first
                if ($ADComputer.Length -ne $NULL) {
                    foreach ($i in $ADComputer) {
                        if ($i.lastLogonTimeStamp -gt $newest.lastLogonTimeStamp) {
                            $newest = $i
                        }
                    }

                    $ADComputer = $newest
                }

                $ReturnObject = New-Object PSObject -Property @{
                        Machine      =    $ADComputer.Name
                        FQDN         =    $Obj.NameHost
                        IP           =    $Obj.Name
                        Enabled      =    $ADComputer.Enabled
                        OS           =    $ADComputer.OperatingSystem
                        Pingable     =    $null
                        Comments     =    $null
                }
                
                #If it's not been logged into in X days put it in the report
                if ([datetime]::FromFileTime($ADComputer.lastLogonTimeStamp) -le $CHECK_LASTLOGIN) {
                    $ReturnObject.Comments += "Hasn't logged in since " + ([datetime]::FromFileTime($ADComputer.lastLogonTimeStamp)) + ";"
                }
                #If it's disabled in AD
                if ($ReturnObject.Enabled -eq $false) {
                    $ReturnObject.Comments += "Disabled in AD;"
                }
                #If we got multiple back
                if ($newest -ne $NULL) {
                    $ReturnObject.Comments += "Multiple objects found in AD. Newest Selected;"
                }
            }

            #If there's no object put it in the report
            if (($ADComputer -eq $null)) {
                $ReturnObject = New-Object PSObject -Property @{
                        Machine      =    $null
                        FQDN         =    $Obj.NameHost
                        IP           =    $Obj.Name
                        Enabled      =    $false
                        OS           =    $null
                        Pingable     =    $null
                        Comments     =    "Not found in Active Directory"
                }
            }
            $ReturnObject.Pingable = $TestPing
            return $ReturnObject
        }
    }

    #There's some sort of rate limit for AD module usage that I can't figure out and I am losing my damn mind
    While ((Get-job -State "Running").Count -ge 5) { Start-Sleep 2 }
    Start-Job $ScriptBlock -ArgumentList $Obj
}

# Wait for all to complete
While (Get-Job -State "Running") { Start-Sleep 2 }

$UnusedRegistration += (Get-Job | Receive-Job)
Remove-Job *

$UnusedRegistration | Export-Csv -NoTypeInformation -Path ("Report.csv")