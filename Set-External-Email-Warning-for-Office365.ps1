param (
    [Parameter(
        Mandatory=$true,
        Position=1
        #HelpMessage="The name of the rule. If an rule already exist with this name, the script will edit it."
    )]
    [ValidateNotNullOrEmpty()]
    [string]$TransportRuleName,

    [Parameter(
        Mandatory=$true,
        Position=2
        #HelpMessage="Header of the disclaimer"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$DisclaimerCaution,

    [Parameter(
        Mandatory=$true,
        Position=3
        #HelpMessage="The message in the disclaimer"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$DisclaimerCautionMessage,

    [Parameter(
        Mandatory=$false,
        Position=4
        #HelpMessage="An array @() that contain the group to be exclude from this rule. Will editing a rule, if this parameter is empty, il will remove the present group."
    )]
    [ValidateNotNullOrEmpty()]
    [array]$ExcludeGroupMembers= @()
)
try {
    #Connect & Login to ExchangeOnline (MFA)
    $getsessions = Get-PSSession | Select-Object -Property State, Name
    $isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
    If ($isconnected -ne "True") 
    {
    Connect-ExchangeOnline
    }  
}
catch {
    Exit-PSSession
}
finally {
    
}
$Disclaimer="<p><div style='background-color:#FFEB9C; width:100%; border-style: solid; border-color:#9C6500; border-width:1pt; padding:2pt; font-size:10pt; line-height:12pt; font-family:Calibri; color:Black; text-align: left;'><span style='color:#9C6500'; font-weight:bold;>$DisclaimerCaution</span>$DisclaimerCautionMessage</div><br></p>"
$Time =Get-Date -Format "HH:mm:ss"
$ValidExcludeGroups = New-Object System.Collections.Generic.List[System.Object]
$ValidExcludeGroupsGUID = New-Object System.Collections.Generic.List[System.Object]
#Check for exclude groups
if ($PSBoundParameters.ContainsKey('ExcludeGroupMembers')) 
{
    [Int64]$GroupTotal = $ExcludeGroupMembers.Count
    [Int64]$GroupIndex = 1
    Write-Verbose "[$Time] $GroupTotal "
    Write-Verbose "[$Time] There is a total of $GroupTotal groups to verify."
    ##========================================================================================
    ##                                                                                      ##
    ##                Verify if the Groups in $ExcludeGroupMembers are valid.               ##
    ##                                                                                      ##
    ##========================================================================================
    foreach ($Group in $ExcludeGroupMembers) 
    {
        Write-Verbose "[$Time] Verifying if $Group is a distribution groups or mail-enabled security groups or and Unified groups."
        $VerifyDistributionGroup = Get-DistributionGroup -Identity $Group -ErrorAction SilentlyContinue #Use the Get-DistributionGroup cmdlet to view existing distribution groups or mail-enabled security groups.
        $VerifyUnifiedGroup = Get-UnifiedGroup -Identity $Group -ErrorAction silentlycontinue #Use the Get-UnifiedGroup cmdlet to view Microsoft 365 Groups in your cloud-based organization.
        ##========================================================================================
        ##                                                                                      ##
        ##                                   DistributionGroup                                  ##
        ##                                                                                      ##
        ##========================================================================================
        if ($null -ne $VerifyDistributionGroup)

        {
            $VerifyDistributionGroupName = $VerifyDistributionGroup.Name
            $ValidExcludeGroups.add($VerifyDistributionGroup)
            $DistributionGroupFound = $true 
            Write-Verbose "[$Time][$GroupIndex / $GroupTotal][VALIDATED] $VerifyDistributionGroupName was added to the valid groups to exclude."
        }
        else 
        {
             $DistributionGroupFound = $false    
        }

        ##========================================================================================
        ##                                                                                      ##
        ##                                     UnifiedGroup                                     ##
        ##                                                                                      ##
        ##========================================================================================
        if ($null -ne $VerifyUnifiedGroup)
        {
            $VerifyUnifiedGroupName = $VerifyUnifiedGroup.Name
            $ValidExcludeGroups.add($VerifyUnifiedGroup.GUID)
            $UnifiedGroupFound = $true
            Write-Verbose "[$Time][$GroupIndex / $GroupTotal][VALIDATED] $VerifyUnifiedGroupName was added to the valid groups to exclude."
        }
        else 
        {
             $UnifiedGroupFound = $false    
        }

        ##========================================================================================
        ##                                                                                      ##
        ##                                      None Found                                      ##
        ##                                                                                      ##
        ##========================================================================================

        if (($true -ne $DistributionGroupFound) -and ($true -ne $UnifiedGroupFound)) 
        {
            Write-Verbose "[$Time][$GroupIndex / $GroupTotal][Error] $Group was nowhere to be found."
        }

        $GroupIndex ++
    }

    $ValidExcludeGroupsCount = $ValidExcludeGroups.Count
    if ($ValidExcludeGroups.Count -eq $GroupTotal) # Proceed only if the group are all validated
    {
        Write-Verbose "[$Time][SUCCESS][$ValidExcludeGroupsCount / $GroupTotal] All the groups have been verified."
        foreach ($ValidExcludeGroup in $ValidExcludeGroups) 
        {
            $ValidExcludeGroupGUID =$ValidExcludeGroup.GUID
            $ValidExcludeGroupsGUID.add($ValidExcludeGroupGUID)
        }  
    }
    else 
    {
        $GroupError = $GroupTotal - $ValidExcludeGroupsCount 
        Write-Verbose "[$Time][ERROR][$ValidExcludeGroupsCount / $GroupTotal] There was $GroupError group that couldn't be verified over $GroupTotal groups."
        $ValidExcludeGroups =$null
        Write-Error "[$Time][ERROR] There was $GroupError group that couldn't be verified over $GroupTotal groups."
        Write-Verbose "[$Time][EXIT] The Transport Rule name $TransportRuleName will not be change. Exiting the script."
        Exit    
    }

}

#Verfiy if the TransportRule Exist
$TransportRule = Get-TransportRule $TransportRuleName -ErrorAction SilentlyContinue # Use the Get-TransportRule cmdlet to view transport rules (mail flow rules) in your organization.
if ($null -ne $TransportRule) 
{
##========================================================================================
##                                                                                      ##
##                                         FOUND                                        ##
##                                                                                      ##
##========================================================================================
    try 
    {
        Write-Verbose "[$Time][FOUND] The Transport Rule name $TransportRuleName was found."
        $Choice1Option1 = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Create a new rule name $TransportRuleName"
        $Choice1Option2 = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "The script will end."
        $Choice1Options = [System.Management.Automation.Host.ChoiceDescription[]]($Choice1Option1, $Choice1Option2)
        $Choice1 = $Host.ui.PromptForChoice("The Transport Rule name $TransportRuleName was found.", "Edit the rule name $TransportRuleName ?", $Choice1Options, 1) 
        switch ($Choice1)
        {
            0
            {
                Write-Verbose "[$Time][Yes] The Transport Rule name $TransportRuleName will be change."
                if ($null -ne $ValidExcludeGroups) 
                {
                    Set-TransportRule $TransportRuleName -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ExceptIfSentToMemberOf $ValidExcludeGroupsGUID -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap 
                }
                else 
                {
                    Set-TransportRule $TransportRuleName -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap  
                }
                Write-Verbose "[$Time][CHANGED] The Transport Rule name $TransportRuleName was change."
            }    
            1 
            {
                Write-Verbose "[$Time][No] The Transport Rule name $TransportRuleName will not be change. Exiting the script."
                Exit
            }
        }
    }
    catch 
    {
        
    } 
}
else 
{
##========================================================================================
##                                                                                      ##
##                                        MISSING                                       ##
##                                                                                      ##
##========================================================================================
    try {
        Write-Verbose "[$Time][Missing] The Transport Rule name $TransportRuleName was nowhere to be found."
        $Choice2Option1 = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Create a new rule name $TransportRuleName"
        $Choice2Option2 = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "The script will end."
        $Choice2Options = [System.Management.Automation.Host.ChoiceDescription[]]($Choice2Option1, $Choice2Option2)
        $Choice2 = $Host.ui.PromptForChoice("The Transport Rule name $TransportRuleName was nowhere to be found.", "Create a new rule name $TransportRuleName?", $Choice2Options, 1) 
        switch ($Choice2)
        {
            0
            {
                if ($null -ne $ValidExcludeGroups) 
                {
                    Write-Verbose "[$Time][Yes] The Transport Rule name $TransportRuleName will be create."
                    New-TransportRule $TransportRuleName -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ExceptIfSentToMemberOf $ValidExcludeGroupsGUID -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap
                }
                else 
                {
                    New-TransportRule $TransportRuleName -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap
                }
                Write-Verbose "[$Time][CREATED] The Transport Rule name $TransportRuleName was Create."        
            }    
            1 
            {
                Write-Verbose "[$Time][No] The Transport Rule name $TransportRuleName will not be create. Exiting the script."
                Exit
            }
        }
    }
    catch 
    {
        Write-Verbose "[$Time][No] The Transport Rule name $TransportRuleName will not be create."
    }
}

#New-TransportRule $TransportRuleName -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ExceptIfSentToMemberOf $ExcludeGroups -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap
#The ExceptIfSentTo parameter specifies an exception that looks for RECIPIENTS in messages. Name, Alias, Distinguished name (DN), Canonical DN, Email address, GUID. In on-premises Exchange, this exception is only available on Mailbox servers. 
#The ExceptIfSentToMemberOf parameter specifies an exception that looks for messages sent to members of GROUPS. Name, Distinguished name (DN), Email address, GUID 

