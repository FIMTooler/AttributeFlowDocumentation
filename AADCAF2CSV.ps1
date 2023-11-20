#Evaluates inputs and determines if defaults are used
Function CheckInputs([ref]$FilePath,[ref]$OutputPath,[ref]$ObjectFilter)
{
    Try
    {
        If ($Inputs.Count -eq 0)
        {
            #No args - output message with example to console and proceed with defaults
            "`r`nThere are no arguments.  Exiting" | Write-Host
            "`r`n`r`nInputs`r`n------" | Write-Host
            "Folder Path (Required)               Path to exported AADC Configuration `r`n" | Write-Host
            "Output File Path (Required)         Path for output CSV file`r`n" | Write-Host
            "Object filter (Optional)               Filter used to select which objects are drawn" | Write-Host
            "     Default - *`r`n" | Write-Host
            "`r`n`r`nExamples`r`n--------" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`" `"c:\out.csv`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`" `"c:\out.csv`" `"Person`"`r`n" | Write-Host
            Exit
        }
        ElseIf ($Inputs.Count -eq 1)
        {
            "`r`nThere's only 1 argument.  Exiting" | Write-Host
            "`r`n`r`nInputs`r`n------" | Write-Host
            "Folder Path (Required)               Path to exported AADC Configuration `r`n" | Write-Host
            "Output File Path (Required)         Path for output CSV file`r`n" | Write-Host
            "Object filter (Optional)               Filter used to select which objects are drawn" | Write-Host
            "     Default - *`r`n" | Write-Host
            "`r`n`r`nExamples`r`n--------" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`" `"c:\out.csv`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`" `"c:\out.csv`" `"Person`"`r`n" | Write-Host
            Exit
        }
        ElseIf ($Inputs.Count -eq 2)
        {
            #2 inputs provided
            $FilePath.Value = $Inputs[0]
            $OutputPath.Value = $Inputs[1]
            "`r`nFile Path - $($FilePath.Value)" | Write-Host
            "`r`nOutput File Path - $($OutputPath.Value)" | Write-Host
            "Object Filter - *" | Write-Host
        }
        ElseIf ($Inputs.Count -eq 3)
        {
            #3 inputs provided
            $FilePath.Value = $Inputs[0]
            $OutputPath.Value = $Inputs[1]
            $ObjectFilter.Value = $Inputs[2].Split(',')
            "`r`nFile Path - $($FilePath.Value)" | Write-Host
            "`r`nOutput File Path - $($OutputPath.Value)" | Write-Host
            "`r`nObject Filter - $($ObjectFilter.Value)" | Write-Host
        }
        Else
        {
            #More than 3 inputs detected
            "ERROR - more than three arguments were passed" | Write-Host
            Exit
        }
    }
    Catch
    {
        "Error within CheckInputs function`r`n" | Write-Host
        "`r`n`r`nInputs`r`n------" | Write-Host
        $Inputs | Write-Host
        "`r`nStack Trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}
<#
   	.SYNOPSIS 
   	Gets the Attribute Flow Rules from Sync Server Configuration

   	.DESCRIPTION
   	Reads the server configuration from the XML files, and outputs the Attribute Flow rules as PSObjects

   	.OUTPUTS
   	PSObjects containing the synchronization server attribute flow rules
   
   	.EXAMPLE
   	Get-AttributeFlows -ServerConfigurationFolder "E:\ServerConfiguration" | out-gridview
#>
Function Get-AttributeFlows
{
   Param
   (        
        [parameter(Mandatory=$false)]
		[String]
		[ValidateScript({Test-Path $_})]
		$ServerConfigurationFolder,
        [String[]]
        $ObjFilter
   ) 
   End
   {   	
        ### This is where the rules will be aggregated before we output them
        $rules = @()
        ###
        ### Loop through the management agent XMLs to get the Name:GUID mapping
        ###
        $maList = @{}
        $maFiles = (get-item (join-path $ServerConfigurationFolder "\Connectors\*.xml"))
        foreach ($maFile in $maFiles)
        {
	        ### Skip the file if it does NOT contain an ma-data node
	        if (select-xml $maFile -XPath "//ma-data" -ErrorAction 0)
	        {
		        ### Get the MA Name and MA ID
		        $maName = (select-xml $maFile -XPath "//ma-data/name").Node.InnerText
		        $maID = (select-xml $maFile -XPath "//ma-data/id").Node.InnerText  
			   
		        $maList.Add($maID,$maName)
	        }
        }

        $syncRuleFiles = get-item (join-path $ServerConfigurationFolder "\SynchronizationRules\*.xml")
        foreach ($srFile in $syncRuleFiles)
        {
            [xml]$SRxml = (Get-Content $srFile)
            $SR = $SRxml.synchronizationRule
            $name = $SR.name
            $id = $SR.id
            $description = $SR.description
            $flowDirection = $SR.direction
            $version = $SR.Version
            $disabled = $SR.disabled
            $connectorID = $SR.connector
	        $maName = $maList[$connectorID]
            $linkType = $SR.linkType
            [int]$precedence = $SR.precedence
            $enablePasswordSync = $SR.EnablePasswordSync
            switch ($flowDirection)
            {
                "Inbound"
                {
                    #targetObjecType because this in inbound rule
		            $mvObjectType = $SR.targetObjectType
                    $cdObjectType = $SR.sourceObjectType
                }
                "Outbound"
                {
                    #sourceObjecType because this in outbound rule
		            $mvObjectType = $SR.sourceObjectType
                    $cdObjectType = $SR.targetObjectType
                }
                default
                {
                    throw "Flow direction not Inbound or Outbound - $($flowDirection)"
                }
            
            }

            if ($ObjFilter -eq "*" -or $ObjFilter -contains $mvObjectType)
            {
		        foreach($importFlow in $SR.'attribute-mappings'.ChildNodes)
		        {
                    if ($flowDirection -eq "Inbound" -and $null -ne $importFlow.src)
                    {
                        #direct
                        $ruleType = "DIRECT"
                        $srcAttribute = $importFlow.src.attr
                        #dest because this in inbound rule
		                $mvAttribute = $importFlow.dest
                    }
                    elseif ($flowDirection -eq "Inbound" -and $null -ne $importFlow.expression)
                    {
                        #expression
                        $ruleType = "EXPRESSION"
                        $expression = $importFlow.expression
                        #get source attributes from expression
                        $srcAttribute = @()
                        $expressionSplit = @($expression.Split("[").Split("]"))
                        for ($i=0;$i -lt $expressionSplit.Count;$i++)
                        {
                            if (($i % 2) -eq 1 -and $srcAttribute -notcontains $expressionSplit[$i])
                            {
                                $srcAttribute += $expressionSplit[$i]
                            }
                            #handle ImportedValue which is parsed different
                            if ($expressionSplit[$i].contains("ImportedValue"))
                            {
                               $temp = $expressionSplit[$i].Substring($expressionSplit[$i].IndexOf("ImportedValue("))
                               $temp = $temp.Substring($temp.IndexOf('"')+1)
                               $srcAttribute += $temp.Remove($temp.IndexOf('"'))
                            }
                        }
                        #dest because this in inbound rule
		                $mvAttribute = $importFlow.dest
                    }
                    elseif ($flowDirection -eq "Outbound" -and $null -ne $importFlow.src)
                    {
                        #direct
                        $ruleType = "DIRECT"
                        $srcAttribute = $importFlow.dest
                        #dest because this in outbound rule
		                $mvAttribute = $importFlow.src.attr
                    }
                    elseif ($flowDirection -eq "Outbound" -and $null -ne $importFlow.expression)
                    {
                        #expression
                        $ruleType = "EXPRESSION"
                        $expression = $importFlow.expression
                        #get source attributes from expression
                        $mvAttribute = @()
                        $expressionSplit = @($expression.Split("[").Split("]"))
                        for ($i=1;$i -lt $expressionSplit.Count;$i+=2)
                        {
                            if ($mvAttribute -notcontains $expressionSplit[$i])
                            {
                                $mvAttribute += $expressionSplit[$i]
                            }
                        }
                        #dest because this in outbound rule
		                $srcAttribute = $importFlow.dest
                    }
                    #ERROR
                    else
                    {
                        throw "Unable to determine flow and attributes"
                    }
                    $valueMergeType = $importFlow.valueMergeType

		            $rule = New-Object PSObject
                    #RuleType,SourceMA,DestinationMA,CDObjectType,SourceMA,DestinationMA,CDObjectType,CDAttribute,expression,MVObjectType,MVAttribute,SR-Name,description,flowDirection,version,disabled,connectorID,linkType,precedence,enablePasswordSync,valueMergeType
		            $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value $ruleType
                    switch ($flowDirection)
                    {
                        "Inbound"
                        {
		                    $rule | Add-Member -MemberType noteproperty -name 'SourceMA' -value $maName
                        }
                        "Outbound"
                        {
		                    $rule | Add-Member -MemberType noteproperty -name 'DestinationMA' -value $maName
                        }
                        default
                        {
                            throw "Flow direction not Inbound or Outbound - $($flowDirection)"
                        }
            
                    }
		            $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		            $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $srcAttribute
		            $rule | Add-Member -MemberType noteproperty -name 'expression' -value $expression
		            $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		            $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $mvAttribute
		            $rule | Add-Member -MemberType noteproperty -name 'SR-Name' -value $name
		            $rule | Add-Member -MemberType noteproperty -name 'SR-ID' -value $id
		            $rule | Add-Member -MemberType noteproperty -name 'description' -value $description
		            $rule | Add-Member -MemberType noteproperty -name 'flowDirection' -value $flowDirection
		            $rule | Add-Member -MemberType noteproperty -name 'version' -value $version
		            $rule | Add-Member -MemberType noteproperty -name 'disabled' -value $disabled
		            $rule | Add-Member -MemberType noteproperty -name 'connectorID' -value $connectorID
		            $rule | Add-Member -MemberType noteproperty -name 'linkType' -value $linkType
		            $rule | Add-Member -MemberType noteproperty -name 'precedence' -value $precedence
		            $rule | Add-Member -MemberType noteproperty -name 'valueMergeType' -value $valueMergeType
		            $rule | Add-Member -MemberType noteproperty -name 'enablePasswordSync' -value $enablePasswordSync
		            $rules += $rule
                    Clear-Variable ruleType,srcAttribute,mvAttribute,expression,valueMergeType -ErrorAction SilentlyContinue
		        }
	        }
        }
		Write-Output $rules
    }
}


Function Get-ImportToExportAttributeFlow
{
   	Param
   	(        
        [parameter(Mandatory=$true)]
		[String]
		[ValidateScript({Test-Path $_})]
		$ServerConfigurationFolder,
        [parameter(Mandatory=$true)]
        [String[]]
        $Filter
   	) 
	End
	{
		### Get the Attribute Flow Rules
		$rules = Get-AttributeFlows -ServerConfigurationFolder $ServerConfigurationFolder $Filter

        #separate inbound and outbound also Sort by precedence
        $IAFs = $rules | where { $_.flowDirection -eq "Inbound" } | sort -Property precedence
        $EAFs = $rules | where { $_.flowDirection -eq "Outbound" } | sort -Property precedence

        ### Add extra property for matching IAFs from different MAs
        ### Used to prevent duplicating IAFs that were previously matched
        $IAFs | % { $_ | Add-Member -MemberType "NoteProperty" -Name "Matched" -Value $false }
        $EAFs | % { $_ | Add-Member -MemberType "NoteProperty" -Name "Matched" -Value $false }


        #RuleType,SourceMA,DestinationMA,CDObjectType,SourceMA,DestinationMA,CDObjectType,CDAttribute,expression,MVObjectType,MVAttribute,SR-Name,description,flowDirection,version,disabled,connectorID,linkType,precedence,enablePasswordSync,valueMergeType,Matched

        ### Array holding PSObjects with arrays of IAFs and EAFs
        $e2eFlowRules = @()

        ### Loops through IAFs looking for IAF and EAF matches until all IAFs are matched
        foreach ($IAF in $IAFs)
        {
            ### empty PS object to hold IAF and EAF arrays
            $e2eAF = New-Object PSObject
            $iafArray = @()
            $eafArray = @()

            ### if not already matched look for matches
            ### if already matched move on to next IAF
            if ($IAF.Matched -eq $false)
            {
                ### IAF isn't matched add it to array and set Matched true
                $IAF.Matched = $true
                $e2eAF | Add-Member -MemberType "NoteProperty" -Name "MVObjectType" -Value $IAF.MVObjectType
                $e2eAF | Add-Member -MemberType "NoteProperty" -Name "MVAttribute" -Value $IAF.MVAttribute
                $iafArray += ($IAF | select * -ExcludeProperty MVObjectType,MVAttribute,Matched)
                
                ### get matching IAFs
                $iafMatches = @($IAFs | where {$_.'MVAttribute' -contains $IAF.'MVAttribute' -and $_.'MVObjectType' -eq $IAF.'MVObjectType' -and $_.'Matched' -eq $false})
                if ($iafMatches.Count -gt 0)
                {
                    ### Loops through matching IAFs adding them to array and setting matched true
                    Foreach ($match in $iafMatches)
                    {
                        ### Set matched to true and add matches to array
                        $match.Matched = $true
                        $iafArray += ($match | select * -ExcludeProperty MVObjectType,MVAttribute,Matched)
                    }
                }

		        ### Look for a matchinging EAF rule    
		        $eafMatches = @($EAFs | where {$_.'MVAttribute' -contains $IAF.'MVAttribute' -and $_.'MVObjectType' -eq $IAF.'MVObjectType'})
		        if ($eafMatches.count -gt 0)
		        {
                    ### Loops through matching EAFs adding them to array and setting matched true
		            foreach($match in $eafMatches)
		            {
                        ### Set matched to true and add matches to array
                        $match.Matched = $true
                        ### use MVAttribute with EAF only if there's more than one attribute used.
                        if ($match.MVAttribute.Count -gt 1)
                        {
		                    $eafArray += ($match | select * -ExcludeProperty MVObjectType,Matched)
                        }
                        else
                        {
                            $eafArray += ($match | select * -ExcludeProperty MVObjectType,MVAttribute,Matched)
                        }
		            }
		        }
			    
                ### if arrays have objects add them to PSObject
                if ($iafArray.Length -gt 0)
                {
                    $e2eAF | Add-Member -MemberType "NoteProperty" -Name "IAFs" -Value $iafArray

                    if ($eafArray.Length -gt 0)
                    {
                        $e2eAF | Add-Member -MemberType "NoteProperty" -Name "EAFs" -Value $eafArray
                    }
                    ### Add PSObject to array for output
                    $e2eFlowRules += $e2eAF
                }
            }
        }

        ### Add unmatched EAFs to array
        foreach ($EAF in @($EAFs | where {$_.'Matched' -eq $false}))
        {
            ### Create empty array for EAF
            $eafArray = @()
            
            ### Create new PSObject for eafArray
            $e2eAF = New-Object PSObject
            
            ### Update Matched and add EAF to array
            $EAF.Matched = $true
            $e2eAF | Add-Member -MemberType "NoteProperty" -Name "MVObjectType" -Value $EAF.MVObjectType
            $e2eAF | Add-Member -MemberType "NoteProperty" -Name "MVAttribute" -Value $EAF.MVAttribute

            ### use MVAttribute with EAF only if there's more than one attribute used
            if ($EAF.MVAttribute.Count -gt 1)
            {
		        $eafArray += ($EAF | select * -ExcludeProperty MVObjectType,Matched)
            }
            else
            {
                $eafArray += ($EAF | select * -ExcludeProperty MVObjectType,MVAttribute,Matched)
            }
            
            ### Add array to PSObject
            $e2eAF | Add-Member -MemberType "NoteProperty" -Name "EAFs" -Value $eafArray
            
            ### Add PSObject to array for output
            $e2eFlowRules += $e2eAF
        }

        ######################################################
        ################## Output formatting #################
        ######################################################

        ### Uncomment following line and comment out the following Foreach
        ### to output array of objects to work with

        $e2eFlowRules | select MVObjectType,MVAttribute,IAFs,EAFs | sort MVObjectType,MVAttribute
    }
}

#Set default parameters
$FilePath = ""
$OutputPath = ""
$ObjectFilter = @()
$ObjectFilter += "*"

#puts script arguments into array so they can be used
$Inputs = @()
Foreach ($Arg in $args)
{
    $Inputs = $Inputs + $Arg
}

CheckInputs ([ref]$FilePath) ([ref]$OutputPath) ([ref]$ObjectFilter)
Get-Date

$AFs = Get-ImportToExportAttributeFlow $FilePath $ObjectFilter

"AFs count is " + $AFs.Count | Write-Host
#header
#out-file requires UTF8 encoding for Excel to respect multi-line cells
"Inbound SyncRule Name,Inbound MA Name,Inbound ObjectType,Inbound Attributes,Expression,Inbound Flow Type,Inbound Precedence,Disabled,MV Attribute,MV ObjectType,Disabled,Outbound Precedence,Outbound SyncRule Name,Outbound MA Name,Outbound ObjectType,Outbound Attribute,Expression,Outbound Flow Type,All MV Attributes used" | Out-File $OutputPath -Append -Encoding utf8
Foreach ($AF in $AFs)
{
    #Get MVAttributeName
    $i = 0
    #repeat while Inbound or Outbound flows exist
    While ($AF.IAFs.Count -gt $i -or $AF.EAFs.Count -gt $i)
    {
        $outLine = ""
        #region handle Inbound flows
        if ($AF.IAFs.Count -gt $i)
        {
            $outLine += $AF.IAFs[$i].'SR-Name'
            $outLine += "," + $AF.IAFs[$i].SourceMA
            $outLine += "," + $AF.IAFs[$i].CDObjectType
            $Attributes = ""
            if ($AF.IAFs[$i].CDAttribute.Count -gt 1)
            {
                $Attributes += "`""
                foreach ($str in $AF.IAFs[$i].CDAttribute)
                {
                    $Attributes += $str + "`n"
                }
                $Attributes = $Attributes.TrimEnd("`n")
                $Attributes += "`""
            }
            else
            {
                $Attributes += $AF.IAFs[$i].CDAttribute
            }
            $outLine += "," + $Attributes
            if ($null -ne $AF.IAFs[$i].Expression)
            {
                $outLine += ",`"" + $AF.IAFs[$i].Expression.Replace('"','""') + "`""
            }
            else
            {
                $outLine += ","
            }
            Switch($AF.IAFs[$i].RuleType.ToLower())
            {
                "DIRECT".ToLower()
                {
                    $outLine += ",Direct"
                }
                "EXPRESSION".ToLower()
                {
                    $outLine += ",Expression"
                }
            }
            $outLine += "," + $AF.IAFs[$i].Precedence
            $outLine += "," + $AF.IAFs[$i].disabled
        }
        #required for MV attribute and remaining columns to line up if no Inbound flow exists
        else
        {
            $outLine += ",,,,,,,"
        }
        #endregion
        #Writes out MV attribute and objectType for all flows
        if ($AF.MVAttribute -ne $null)
        {
            $outLine += "," + $AF.MVAttribute + "," + $AF.MVObjectType
        }
        #handles Outbound Code\Constants where no MV attribute is used
        else
        {
            $outLine += ",," + $AF.MVObjectType
        }
        #region handles Outbound Flows
        if ($AF.EAFs.Count -gt $i)
        {
            $outLine += "," + $AF.EAFs[$i].disabled
            $outLine += "," + $AF.EAFs[$i].Precedence
            $outLine += "," + $AF.EAFs[$i].'SR-Name'
            $outLine += "," + $AF.EAFs[$i].DestinationMA
            $outLine += "," + $AF.EAFs[$i].CDObjectType
            $outLine += "," + $($AF.EAFs[$i].CDAttribute)
            if ($null -ne $AF.EAFs[$i].expression)
            {
                $outLine += ",`"" + $AF.EAFs[$i].expression.Replace('"','""') + "`""
            }
            else
            {
                $outLine += ","
            }
            Switch($AF.EAFs[$i].RuleType.ToLower())
            {
                "EXPRESSION".ToLower()
                {
                    if ($AF.EAFs[$i].MVAttribute.Count -gt 1)
                    {
                        $Attributes = "`""
                        foreach ($str in $AF.EAFs[$i].MVAttribute)
                        {
                            $Attributes += $str + "`n"
                        }
                        $Attributes = $Attributes.TrimEnd("`n")
                        $Attributes += "`""
                    }
                    else
                    {
                        $Attributes += $AF.EAFs[$i].MVAttribute
                    }
                    $outLine += ",Expression," + $Attributes
                }
                "DIRECT".ToLower()
                {
                    $outLine += ",Direct"
                }
            }
        }
        #endregion
        $i++
        #out-file requires UTF8 encoding for Excel to respect multi-line cells
        $outLine | Out-File $OutputPath -Append -Encoding utf8
    }
}
Get-Date
