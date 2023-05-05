#region Functions


#Evaluates inputs and determines if defaults are used
Function CheckInputs([ref]$FirstPath,[ref]$SecondPath,[ref]$OutputPath,[ref]$ObjectFilter)
{
    Try
    {
        If ($Inputs.Count -eq 0)
        {
            #No args - output message with example to console and proceed with defaults
            "`r`nThere are no arguments.  Exiting" | Write-Host
            "`r`n`r`nInputs`r`n------" | Write-Host
            "First Folder Path (Required)               Path to exported MIM Sync Engine Configuration `r`n" | Write-Host
            "Second Folder Path (Required)               Path to exported MIM Sync Engine Configuration `r`n" | Write-Host
            "Output File Path (Required)         Path for output CSV file`r`n" | Write-Host
            "Object filter (Optional)               Filter used to select which objects are drawn" | Write-Host
            "     Default - *`r`n" | Write-Host
            "`r`n`r`nExamples`r`n--------" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`" `"c:\out.csv`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`" `"c:\out.csv`" `"Person`"`r`n" | Write-Host
            Exit
        }
        ElseIf ($Inputs.Count -eq 1)
        {
            "`r`nThere's only 1 argument.  Exiting" | Write-Host
            "`r`n`r`nInputs`r`n------" | Write-Host
            "First Folder Path (Required)               Path to exported FIM Sync Engine Configuration `r`n" | Write-Host
            "Second Folder Path (Required)               Path to exported MIM Sync Engine Configuration `r`n" | Write-Host
            "Output File Path (Required)         Path for output CSV file`r`n" | Write-Host
            "Object filter (Optional)               Filter used to select which objects are drawn" | Write-Host
            "     Default - *`r`n" | Write-Host
            "`r`n`r`nExamples`r`n--------" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`" `"c:\out.csv`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`" `"c:\out.csv`" `"Person`"`r`n" | Write-Host
            Exit
        }
        ElseIf ($Inputs.Count -eq 2)
        {
            "`r`nThere's only 2 arguments.  Exiting" | Write-Host
            "`r`n`r`nInputs`r`n------" | Write-Host
            "First Folder Path (Required)               Path to exported FIM Sync Engine Configuration `r`n" | Write-Host
            "Second Folder Path (Required)               Path to exported MIM Sync Engine Configuration `r`n" | Write-Host
            "Output File Path (Required)         Path for output CSV file`r`n" | Write-Host
            "Object filter (Optional)               Filter used to select which objects are drawn" | Write-Host
            "     Default - *`r`n" | Write-Host
            "`r`n`r`nExamples`r`n--------" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`" `"c:\out.csv`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\MIMConfigServer1`" `"C:\MIMConfigServer2`" `"c:\out.csv`" `"Person`"`r`n" | Write-Host
            Exit
        }
        ElseIf ($Inputs.Count -eq 3)
        {
            #3 inputs provided
            $FirstPath.Value = $Inputs[0]
            $SecondPath.Value = $Inputs[1]
            $OutputPath.Value = $Inputs[2]
            "`r`nFirst Folder Path - $($FirstPath.Value)" | Write-Host
            "`r`nSecond Folder Path - $($SecondPath.Value)" | Write-Host
            "`r`nOutput File Path - $($OutputPath.Value)" | Write-Host
            "`r`nObject Filter - $($ObjectFilter.Value)" | Write-Host
        }
        ElseIf ($Inputs.Count -eq 4)
        {
            #3 inputs provided
            $FirstPath.Value = $Inputs[0]
            $SecondPath.Value = $Inputs[1]
            $OutputPath.Value = $Inputs[2]
            $ObjectFilter.Value = $Inputs[3].Split(',')
            "`r`nFirst Folder Path - $($FirstPath.Value)" | Write-Host
            "`r`nSecond Folder Path - $($SecondPath.Value)" | Write-Host
            "`r`nOutput File Path - $($OutputPath.Value)" | Write-Host
            "`r`nObject Filter - $($ObjectFilter.Value)" | Write-Host
        }
        Else
        {
            #More than 4 inputs detected
            "ERROR - more than four arguments were passed" | Write-Host
            Exit
        }
        if (Test-Path -Path $FirstPath.Value -PathType Leaf)
        {
            "ERROR - First Folder Path provided is a file.  Folder path is required" | Write-Host
            Exit
        }
        if (Test-Path -Path $SecondPath.Value -PathType Leaf)
        {
            "ERROR - Second Folder Path provided is a file.  Folder path is required" | Write-Host
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
   	Gets the Import Attribute Flow Rules from Sync Server Configuration

   	.DESCRIPTION
   	Reads the server configuration from the XML files, and outputs the Import Attribute Flow rules as PSObjects

   	.OUTPUTS
   	PSObjects containing the synchronization server import attribute flow rules
   
   	.EXAMPLE
   	Get-ImportAttributeFlow -ServerConfigurationFolder "E:\ServerConfiguration" | out-gridview
#>
Function Get-ImportAttributeFlow
{
   Param
   (        
        [parameter(Mandatory=$false)]
		[ValidateScript({Test-Path $_})]		
		[String]$ServerConfigurationFolder,
        [String[]] $ObjFilter
   ) 
   End
   {   	
   		### This is where the rules will be aggregated before we output them
		$rules = @()

		###
		### Loop through the management agent XMLs to get the Name:GUID mapping
		###
		$maList = @{}
		$maFiles = (get-item (join-path $ServerConfigurationFolder "*.xml"))
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

		###
		### Get:
		###    mv-object-type
		###      mv-attribute
		###        src-ma
		###        cd-object-type
		###          src-attribute
		###
		[xml]$mv = get-content (join-path $ServerConfigurationFolder "MV.xml")
 
		foreach($importFlowSet in $mv.selectNodes("//import-flow-set"))
		{
		    $mvObjectType = $importFlowSet.'mv-object-type'
            if ($ObjFilter -eq "*" -or $ObjFilter -contains $mvObjectType)
            {
		        foreach($importFlows in $importFlowSet.'import-flows')
		        {
		            $mvAttribute = $importFlows.'mv-attribute'        
				    $precedenceType = $importFlows.type
				    $precedenceRank = 0
		           
		            foreach($importFlow in $importFlows.'import-flow')
		            {
		                $cdObjectType = $importFlow.'cd-object-type'
		                $srcMA = $maList[$importFlow.'src-ma']
		                $maID = $importFlow.'src-ma'
		                $maName = $maList[$maID]			
		                        
		                if ($importFlow.'direct-mapping' -ne $null)
		                {
						    if ($precedenceType -eq 'ranked')
						    {
						     $precedenceRank += 1
						    }
						    else
						    {
						     $precedenceRank = $null
						    }
					
                            ###
                            ### Handle src-attribute that are intinsic (<src-attribute intrinsic="true">dn</src-attribute>)
                            ###
                            if ($importFlow.'direct-mapping'.'src-attribute'.intrinsic)
                            {
                                $srcAttribute = "<{0}>" -F $importFlow.'direct-mapping'.'src-attribute'.'#text'
                            }
                            else
                            {
		                        $srcAttribute = $importFlow.'direct-mapping'.'src-attribute'
                            }
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'DIRECT'
		                    $rule | Add-Member -MemberType noteproperty -name 'SourceMA' -value $srcMA
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $srcAttribute
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $mvAttribute
		                    $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $null
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceType' -value $precedenceType
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceRank' -value $precedenceRank
		                
		                    $rules += $rule                               
		                }
		                elseif ($importFlow.'scripted-mapping' -ne $null)
		                {                
		                    $scriptContext = $importFlow.'scripted-mapping'.'script-context'  

                            ###
                            ### Handle src-attribute that are intrinsic (<src-attribute intrinsic="true">dn</src-attribute>)
                            ###              
		                    $srcAttributes = @()
                            $importFlow.'scripted-mapping'.'src-attribute' | ForEach-Object {
                                if ($_.intrinsic)
                                {
                                    $srcAttributes += "<{0}>" -F $_.'#text'
                                }
                                else
                                {
		                            $srcAttributes += $_
                                }
                            }
                            if ($srcAttributes.Count -eq 1)
                            {
                                $srcAttributes = $srcAttributes -as [String]
                            }
						
						    if ($precedenceType -eq 'ranked')
						    {
						      $precedenceRank += 1
						    }
						    else
						    {
						      $precedenceRank = $null
						    }
		                
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'SCRIPTED'
		                    $rule | Add-Member -MemberType noteproperty -name 'SourceMA' -value $srcMA
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $srcAttributes
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $mvAttribute
						    $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $scriptContext.Replace("`"","'")
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceType' -value $precedenceType
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceRank' -value $precedenceRank
		                                
		                    $rules += $rule                        
		                }   
					    elseif ($importFlow.'sync-rule-mapping' -ne $null)
		                {                
		                    $scriptContext = $null 
						    $ruleType = ("ISR-{0}" -f $importFlow.'sync-rule-mapping'.'mapping-type')
		                    $srcAttributes = $importFlow.'sync-rule-mapping'.'src-attribute'    
						
						    if ($precedenceType -eq 'ranked')
						    {
						      $precedenceRank += 1
						    }
						    else
						    {
						      $precedenceRank = $null
						    }
						
		                    $rule = New-Object PSObject

						    if ($importFlow.'sync-rule-mapping'.'mapping-type' -ieq 'expression')
						    {
							    $scriptContext = $importFlow.'sync-rule-mapping'.'sync-rule-value'.'import-flow'.InnerXml
						        $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $scriptContext
						    }
							elseif ($importFlow.'sync-rule-mapping'.'mapping-type' -ieq 'constant')
							{
                                $constantValue = $importFlow.'sync-rule-mapping'.'sync-rule-value'
						        $rule | Add-Member -MemberType noteproperty -name 'ConstantValue' -value $constantValue
							}
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value $ruleType
		                    $rule | Add-Member -MemberType noteproperty -name 'SourceMA' -value $srcMA
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $srcAttributes
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $mvAttribute
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceType' -value $precedenceType
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceRank' -value $precedenceRank
		                                
		                    $rules += $rule                        
		                }
					    elseif ($importFlow.'constant-mapping' -ne $null)
					    {
						    if ($precedenceType -eq 'ranked')
						    {
							     $precedenceRank += 1
						    }
						    else
						    {
							     $precedenceRank = $null
						    }

					
						    $constantValue = $importFlow.'constant-mapping'.'constant-value'
						
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value "CONSTANT"
		                    $rule | Add-Member -MemberType noteproperty -name 'SourceMA' -value $srcMA
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
							$rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $null																						
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $mvAttribute
							$rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $null
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceType' -value $precedenceType
						    $rule | Add-Member -MemberType noteproperty -name 'PrecedenceRank' -value $precedenceRank
						    $rule | Add-Member -MemberType noteproperty -name 'ConstantValue' -value $constantValue
		                                
		                    $rules += $rule
					    }
		            }#foreach($importFlow in $importFlows.'import-flow')
		        }#foreach($importFlows in $importFlowSet.'import-flows')
		    }#foreach($importFlowSet in $mv.selectNodes("//import-flow-set"))
		}
		Write-Output $rules
   }#End
}
<#
   .SYNOPSIS 
   Gets the Export Attribute Flow Rules from Sync Server Configuration

   .DESCRIPTION
   Reads the server configuration from the XML files, and outputs the Export Attribute Flow rules as PSObjects

   .OUTPUTS
   PSObjects containing the synchronization server export attribute flow rules
   
   .EXAMPLE
   Get-ExportAttributeFlow -ServerConfigurationFolder "E:\sd\IAM\ITAuthorize\Source\Configuration\FimSync\ServerConfiguration"

#>
Function Get-ExportAttributeFlow
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
		
		### Export attribute flow rules are contained in the ma-data nodes of the MA*.XML files
		$maFiles = @(get-item (Join-Path $ServerConfigurationFolder "MA-*.xml"))
		
		
		foreach ($maFile in $maFiles)
		{
			### Get the MA Name and MA ID
		   	$maName = (select-xml $maFile -XPath "//ma-data/name").Node.InnerText
		   
		    foreach($exportFlowSet in (Select-Xml -path $maFile -XPath "//export-flow-set" | select -ExpandProperty Node))
		    {
		        $mvObjectType = $exportFlowSet.'mv-object-type'
		        $cdObjectType = $exportFlowSet.'cd-object-type'
		        
                if ($ObjFilter -eq "*" -or $ObjFilter -contains $mvObjectType)
                {
		            foreach($exportFlow in $exportFlowSet.'export-flow')
		            {
		                $cdAttribute = $exportFlow.'cd-attribute'
		                [bool]$allowNulls = $false
					    if ([bool]::TryParse($exportFlow.'suppress-deletions', [ref]$allowNulls))
					    {
						    $allowNulls = -not $allowNulls
					    }
						[string]$initialFlowOnly = $null
						[string]$isExistenceTest = $null
						[string]$syncRuleID = $null
		                if ($exportFlow.'direct-mapping' -ne $null)
		                {
                            ###
                            ### Handle src-attribute that are intrinsic (<src-attribute intrinsic="true">object-id</src-attribute>)
                            ###
                            if ($exportFlow.'direct-mapping'.'src-attribute'.intrinsic)
                            {
                                $srcAttribute = "<{0}>" -F $exportFlow.'direct-mapping'.'src-attribute'.'#text'
                            }
                            else
                            {
		                        $srcAttribute = $exportFlow.'direct-mapping'.'src-attribute'
                            }
		                
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType NoteProperty -Name 'RuleType' -Value 'DIRECT'
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MAName' -Value $maName                
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MVObjectType' -Value $mvObjectType
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MVAttribute' -Value $srcAttribute
		                    $rule | Add-Member -MemberType NoteProperty -Name 'CDObjectType' -Value $cdObjectType
		                    $rule | Add-Member -MemberType NoteProperty -Name 'CDAttribute' -Value $cdAttribute
							$rule | Add-Member -MemberType NoteProperty -Name 'ScriptContext' -Value $null
						    $rule | Add-Member -MemberType NoteProperty -Name 'AllowNulls' -Value $allowNulls
							$rule | Add-Member -MemberType NoteProperty -Name 'InitialFlowOnly' -Value $initialFlowOnly
							$rule | Add-Member -MemberType NoteProperty -Name 'IsExistenceTest' -Value $isExistenceTest
		                
		                    $rules += $rule
		                }
                        if ($exportFlow.'constant-mapping' -ne $null)
		                {
                            ###
                            ### Handle src-attribute that are intrinsic (<src-attribute intrinsic="true">object-id</src-attribute>)
                            ###
                            $constantValue = $exportFlow.'constant-mapping'.'constant-Value'
		                
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType NoteProperty -Name 'RuleType' -Value 'CONSTANT'
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MAName' -Value $maName                
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MVObjectType' -Value $mvObjectType
		                    $rule | Add-Member -MemberType NoteProperty -Name 'CDObjectType' -Value $cdObjectType
		                    $rule | Add-Member -MemberType NoteProperty -Name 'CDAttribute' -Value $cdAttribute
						    $rule | Add-Member -MemberType NoteProperty -Name 'ConstantValue' -Value $constantValue
						    $rule | Add-Member -MemberType NoteProperty -Name 'AllowNulls' -Value $allowNulls
							$rule | Add-Member -MemberType NoteProperty -Name 'InitialFlowOnly' -Value $initialFlowOnly
							$rule | Add-Member -MemberType NoteProperty -Name 'IsExistenceTest' -Value $isExistenceTest
                            $rule | Add-Member -MemberType NoteProperty -Name 'SyncRuleID' -Value $syncRuleID
		                
		                    $rules += $rule
		                }
		                elseif ($exportFlow.'scripted-mapping' -ne $null)
		                {                
		                    $scriptContext = $exportFlow.'scripted-mapping'.'script-context'		                
						    $srcAttributes = @()
						
                            ###
                            ### Handle src-attribute that are intrinsic (<src-attribute intrinsic="true">object-id</src-attribute>)
                            ###
                            $exportFlow.'scripted-mapping'.'src-attribute' | ForEach-Object {
                                if ($_.intrinsic)
                                {
                                    $srcAttributes += "<{0}>" -F $_.'#text'
                                }
                                elseif ($_) # Do not add empty values.
                                {
		                            $srcAttributes += $_
                                }
                            }
                            # (Commented) Leave as collection
                            if ($srcAttributes.Count-eq 1)
                            {
                                $srcAttributes = $srcAttributes -as[String]
                            }
		                    
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType NoteProperty -Name 'RuleType' -Value 'SCRIPTED'
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MAName' -Value $maName
						    $rule | Add-Member -MemberType NoteProperty -Name 'MVObjectType' -Value $mvObjectType
		                    $rule | Add-Member -MemberType NoteProperty -Name 'MVAttribute' -Value $srcAttributes
		                    $rule | Add-Member -MemberType NoteProperty -Name 'CDObjectType' -Value $cdObjectType
		                    $rule | Add-Member -MemberType NoteProperty -Name 'CDAttribute' -Value $cdAttribute	
		                    $rule | Add-Member -MemberType NoteProperty -Name 'ScriptContext' -Value $scriptContext.Replace("`"","'")
						    $rule | Add-Member -MemberType NoteProperty -Name 'AllowNulls' -Value $allowNulls
							$rule | Add-Member -MemberType NoteProperty -Name 'InitialFlowOnly' -Value $initialFlowOnly
							$rule | Add-Member -MemberType NoteProperty -Name 'IsExistenceTest' -Value $isExistenceTest
                            $rule | Add-Member -MemberType NoteProperty -Name 'SyncRuleID' -Value $syncRuleID
		                                
		                    $rules += $rule                        
		                }
					    elseif ($exportFlow.'sync-rule-mapping' -ne $null)
					    {
                            $syncRuleID = $exportFlow.'sync-rule-mapping'.'sync-rule-id'
						    $srcAttribute = $exportFlow.'sync-rule-mapping'.'src-attribute'
                            $initialFlowOnly = $exportFlow.'sync-rule-mapping'.'initial-flow-only'
                            $isExistenceTest = $exportFlow.'sync-rule-mapping'.'is-existence-test'
						    if($exportFlow.'sync-rule-mapping'.'mapping-type' -eq 'direct')
						    {
							    $rule = New-Object PSObject
							    $rule | Add-Member -MemberType NoteProperty -Name 'RuleType' -Value 'OSR-Direct'
							    $rule | Add-Member -MemberType NoteProperty -Name 'MAName' -Value $maName
							    $rule | Add-Member -MemberType NoteProperty -Name 'MVObjectType' -Value $mvObjectType
							    $rule | Add-Member -MemberType NoteProperty -Name 'MVAttribute' -Value $srcAttribute
							    $rule | Add-Member -MemberType NoteProperty -Name 'CDObjectType' -Value $cdObjectType
							    $rule | Add-Member -MemberType NoteProperty -Name 'CDAttribute' -Value $cdAttribute
								$rule | Add-Member -MemberType NoteProperty -Name 'ScriptContext' -Value $null
							    $rule | Add-Member -MemberType NoteProperty -Name 'AllowNulls' -Value $allowNulls
                                $rule | Add-Member -MemberType NoteProperty -Name 'InitialFlowOnly' -Value $initialFlowOnly
                                $rule | Add-Member -MemberType NoteProperty -Name 'IsExistenceTest' -Value $isExistenceTest
                                $rule | Add-Member -MemberType NoteProperty -Name 'SyncRuleID' -Value $syncRuleID
											
							    $rules += $rule             
						    }
						    elseif ($exportFlow.'sync-rule-mapping'.'mapping-type' -eq 'expression')
						    {
							    $scriptContext = $exportFlow.'sync-rule-mapping'.'sync-rule-value'.'export-flow'.InnerXml
							    $cdAttribute = $exportFlow.'sync-rule-mapping'.'sync-rule-value'.'export-flow'.dest
							    $rule = New-Object PSObject
							    $rule | Add-Member -MemberType NoteProperty -Name 'RuleType' -Value 'OSR-Expression'
							    $rule | Add-Member -MemberType NoteProperty -Name 'MAName' -Value $maName
							    $rule | Add-Member -MemberType NoteProperty -Name 'MVObjectType' -Value $mvObjectType
							    $rule | Add-Member -MemberType NoteProperty -Name 'MVAttribute' -Value $srcAttribute
							    $rule | Add-Member -MemberType NoteProperty -Name 'CDObjectType' -Value $cdObjectType
							    $rule | Add-Member -MemberType NoteProperty -Name 'CDAttribute' -Value $cdAttribute														
							    $rule | Add-Member -MemberType NoteProperty -Name 'ScriptContext' -Value $scriptContext
							    $rule | Add-Member -MemberType NoteProperty -Name 'AllowNulls' -Value $allowNulls
                                $rule | Add-Member -MemberType NoteProperty -Name 'InitialFlowOnly' -Value $initialFlowOnly
                                $rule | Add-Member -MemberType NoteProperty -Name 'IsExistenceTest' -Value $isExistenceTest
                                $rule | Add-Member -MemberType NoteProperty -Name 'SyncRuleID' -Value $syncRuleID
											
							    $rules += $rule             
						    }
                            elseif($exportFlow.'sync-rule-mapping'.'mapping-type' -eq 'constant')
						    {
                                $constantValue = $exportFlow.'sync-rule-mapping'.'sync-rule-value'
							    $rule = New-Object PSObject
							    $rule | Add-Member -MemberType NoteProperty -Name 'RuleType' -Value 'OSR-Constant'
							    $rule | Add-Member -MemberType NoteProperty -Name 'MAName' -Value $maName
							    $rule | Add-Member -MemberType NoteProperty -Name 'MVObjectType' -Value $mvObjectType
							    $rule | Add-Member -MemberType NoteProperty -Name 'MVAttribute' -Value $srcAttribute
							    $rule | Add-Member -MemberType NoteProperty -Name 'CDObjectType' -Value $cdObjectType
							    $rule | Add-Member -MemberType NoteProperty -Name 'CDAttribute' -Value $cdAttribute
							    $rule | Add-Member -MemberType NoteProperty -Name 'AllowNulls' -Value $allowNulls
                                $rule | Add-Member -MemberType NoteProperty -Name 'InitialFlowOnly' -Value $initialFlowOnly
                                $rule | Add-Member -MemberType NoteProperty -Name 'IsExistenceTest' -Value $isExistenceTest
							    $rule | Add-Member -MemberType NoteProperty -Name 'ConstantValue' -Value $constantValue
                                $rule | Add-Member -MemberType NoteProperty -Name 'SyncRuleID' -Value $syncRuleID
											
							    $rules += $rule             
						    }
						    else
						    {
                                $exportFlow.'sync-rule-mapping'.'mapping-type' | Write-Host
							    throw "Unsupported Export Flow type"
						    }
			           
					    }
		            }
		        }
            }
		}
		
		Write-Output $rules
   }#End
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
		### Get the Import Attribute Flow Rules
		$IAFs = Get-ImportAttributeFlow -ServerConfigurationFolder $ServerConfigurationFolder $Filter
        
        ### Add extra property for matching IAFs from different MAs
        ### Used to prevent duplicating IAFs that were previously matched
        Foreach ($IAF in $IAFs)
        {
            $IAF | Add-Member -MemberType "NoteProperty" -Name "Matched" -Value $false
        }
		
		### Get the Export Attribute Flow Rules
		$EAFs = Get-ExportAttributeFlow -ServerConfigurationFolder $ServerConfigurationFolder $Filter

        ### Add extra property for matching IAFs from different MAs
        ### Used to prevent duplicating IAFs that were previously matched
        Foreach ($EAF in $EAFs)
        {
            $EAF | Add-Member -MemberType "NoteProperty" -Name "Matched" -Value $false
        }


        ########################################################################
        ### This is where the rules will be aggregated before we output them ###
        ########################################################################

        ### Array holding PSObjects with arrays of IAFs and EAFs
		$e2eFlowRules = @()
        $i = 0

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

        ### For readable output to screen
#        Foreach ($Object in $e2eFlowRules)
#        {
#            "MVObjectType   : " + $Object.MVObjectType
#            "MVAttribute    : " + $Object.MVAttribute
#            "IAFs"
#            "----"
#            $Object.IAFs
#            "EAFs"
#            "----"
#            $Object.EAFs
#            "------------------------------------------------------------------------------"
#        }
	}
}

Function Get-ServerNameFromMV
{
   Param
   (        
        [parameter(Mandatory=$false)]
		[ValidateScript({Test-Path $_})]		
		[String]$ServerConfigurationFolder
   )
   [xml]$mv = get-content (join-path $ServerConfigurationFolder "MV.xml")
   $savedMVConfig = $mv.selectNodes("//saved-mv-configuration")
   return $savedMVConfig.Server

}

Function WriteIAFs-ToFile
{
   Param
   (		
		[String]$OutPath,
        $inboundAFs
   )
   #header
    #out-file requires UTF8 encoding for Excel to respect multi-line cells
    "Comparison Result,Server Name,Inbound MA Name,Inbound ObjectType,Inbound Attributes,Inbound Flow Type,Inbound Precedence,MV Attribute,MV ObjectType" | Out-File $OutPath -Append -Encoding utf8
    #Get MVAttributeName
    $i = 0
    #repeat while Inbound or Outbound flows exist
    While ($inboundAFs.Count -gt $i)
    {
        $outLine = ""
        #region handle Inbound flows
        if ($inboundAFs.Count -gt $i)
        {
            $outLine += $inboundAFs[$i].Result
            $outLine += "," + $inboundAFs[$i].ServerName
            $outLine += "," + $inboundAFs[$i].SourceMA
            $outLine += "," +$inboundAFs[$i].CDObjectType
            $Attributes = ""
            if ($inboundAFs[$i].CDAttribute.Count -gt 1)
            {
                $Attributes += "`""
                foreach ($str in $inboundAFs[$i].CDAttribute)
                {
                    $Attributes += $str + "`n"
                }
                $Attributes = $Attributes.TrimEnd("`n")
                $Attributes += "`""
            }
            else
            {
                $Attributes += $inboundAFs[$i].CDAttribute
            }
            $outLine += "," + $Attributes
            Switch($inboundAFs[$i].RuleType.ToLower())
            {
                "SCRIPTED".ToLower()
                {
                    $outLine += ",`"Code - " + $inboundAFs[$i].ScriptContext + "`""
                }
                "DIRECT".ToLower()
                {
                    $outLine += ",Direct"
                }
                "CONSTANT".ToLower()
                {
                    $outLine += ",CONSTANT - " + $inboundAFs[$i].ConstantValue #$AF.inboundAFs[$i].ScriptContext
                }
                #don't recall if value will be SCRIPTED or EXPRESSION
                "ISR-SCRIPTED".ToLower()
                {
                    $outLine += ",SyncRule-Expression"
                }
                "ISR-EXPRESSION".ToLower()
                {
                    $outLine += ",SyncRule-Expression"
                }
                "ISR-DIRECT".ToLower()
                {
                    $outLine += ",SyncRule-Direct"
                }
                "ISR-CONSTANT".ToLower()
                {
                    $outLine += ",SyncRule-CONSTANT - " + $inboundAFs[$i].ConstantValue
                }
            }
            if ($inboundAFs[$i].PrecedenceType -eq "equal")
            {
                $outLine += ",Equal"
            }
            elseIf ($inboundAFs[$i].PrecedenceType -eq "manual")
            {
                $outLine += ",manual"
            }
            elseIf ($inboundAFs[$i].PrecedenceType -eq "ranked")
            {
                $outLine += "," + $inboundAFs[$i].PrecedenceRank
            }
        }
        #endregion
        #Writes out MV attribute and objectType for all flows
        if ($inboundAFs[$i].MVAttribute -ne $null)
        {
            $outLine += "," + $inboundAFs[$i].MVAttribute + "," + $inboundAFs[$i].MVObjectType
        }
        #handles Outbound Code\Constants where no MV attribute is used
        else
        {
            $outLine += ",," + $inboundAFs[$i].MVObjectType
        }
        #endregion
        $i++
        #out-file requires UTF8 encoding for Excel to respect multi-line cells
        $outLine | Out-File $OutPath -Append -Encoding utf8
    }
}

Function WriteEAFs-ToFile
{
   Param
   (		
		[String]$OutPath,
        $outboundAFs
   )
   #header
    #out-file requires UTF8 encoding for Excel to respect multi-line cells
    "Comparison Result,Server Name,MV Attribute,MV ObjectType,Outbound MA Name,Outbound ObjectType,Outbound Attribute,Outbound Allow Nulls,Outbound Flow Type,All MV Attributes used,Initial Flow Only" | Out-File $OutPath -Append -Encoding utf8
    #Get MVAttributeName
    $i = 0
    #repeat while Inbound or Outbound flows exist
    While ($outboundAFs.Count -gt $i)
    {
        $outLine = ""
        #region handle Outbound flows
        if ($outboundAFs.Count -gt $i)
        {
            $Attributes = ""
            $outLine += $outboundAFs[$i].Result
            $outLine += "," + $outboundAFs[$i].ServerName
            #Writes out MV attribute and objectType for all flows
            if ($outboundAFs[$i].MVAttribute -ne $null)
            {
                if ($outboundAFs[$i].MVAttribute.Count -gt 1)
                {
                    $Attributes += "`""
                    foreach ($str in $outboundAFs[$i].MVAttribute)
                    {
                        $Attributes += $str + "`n"
                    }
                    $Attributes = $Attributes.TrimEnd("`n")
                    $Attributes += "`""
                }
                else
                {
                    $Attributes += $outboundAFs[$i].MVAttribute
                }
                $outLine += "," + $Attributes + "," + $outboundAFs[$i].MVObjectType
            }
            #handles Outbound Code\Constants where no MV attribute is used
            else
            {
                $outLine += ",," + $outboundAFs[$i].MVObjectType
            }
            #endregion
            $outLine += "," + $outboundAFs[$i].MAName
            $outLine += "," + $outboundAFs[$i].CDObjectType
            $outLine += "," + $($outboundAFs[$i].CDAttribute)
            $outLine += "," + $outboundAFs[$i].AllowNulls
            Switch($outboundAFs[$i].RuleType.ToLower())
            {
                "SCRIPTED".ToLower()
                {
                    if ($outboundAFs[$i].MVAttribute.Count -gt 1)
                    {
                        $Attributes = "`""
                        foreach ($str in $outboundAFs[$i].MVAttribute)
                        {
                            $Attributes += $str + "`n"
                        }
                        $Attributes = $Attributes.TrimEnd("`n")
                        $Attributes += "`""
                    }
                    else
                    {
                        $Attributes += $outboundAFs[$i].MVAttribute
                    }
                    $outLine += ",`"Code - " + $outboundAFs[$i].ScriptContext + "`"," + $Attributes
                }
                "DIRECT".ToLower()
                {
                    $outLine += ",Direct"
                }
                "CONSTANT".ToLower()
                {
                    $outLine += ",CONSTANT - " + $outboundAFs[$i].ConstantValue
                }
                "OSR-EXPRESSION".ToLower()
                {
                    $outLine += ",SyncRule-Expression,," + $outboundAFs[$i].InitialFlowOnly
                }
                "OSR-DIRECT".ToLower()
                {
                    $outLine += ",SyncRule-Direct,," + $outboundAFs[$i].InitialFlowOnly
                }
                "OSR-CONSTANT".ToLower()
                {
                    $outLine += ",SyncRule-CONSTANT - " + $outboundAFs[$i].ConstantValue + ",,"  + $outboundAFs[$i].InitialFlowOnly
                }
            }
        }
        #endregion
        $i++
        #out-file requires UTF8 encoding for Excel to respect multi-line cells
        $outLine | Out-File $OutPath -Append -Encoding utf8
    }
}
#endregion


#Set default parameters
$FirstPath = ""
$SecondPath = ""
$OutputPath = ""
$ObjectFilter = @()
$ObjectFilter += "*"

#puts script arguments into array so they can be used
$Inputs = @()
Foreach ($Arg in $args)
{
    $Inputs = $Inputs + $Arg
}

CheckInputs ([ref]$FirstPath) ([ref]$SecondPath) ([ref]$OutputPath) ([ref]$ObjectFilter)

$firstServerName = Get-ServerNameFromMV $FirstPath
$FirstIAFhash = Get-ImportAttributeFlow -ServerConfigurationFolder $FirstPath $ObjectFilter | Group-Object -Property MVObjectType,MVAttribute -AsHashTable -AsString | sort
$FirstEAFhash = Get-ExportAttributeFlow -ServerConfigurationFolder $FirstPath $ObjectFilter | Group-Object -Property CDObjectType,CDAttribute -AsHashTable -AsString



$secondServerName = Get-ServerNameFromMV $secondPath
$SecondIAFhash = Get-ImportAttributeFlow -ServerConfigurationFolder $SecondPath $ObjectFilter | Group-Object -Property MVObjectType,MVAttribute -AsHashTable -AsString
$SecondEAFhash = Get-ExportAttributeFlow -ServerConfigurationFolder $SecondPath $ObjectFilter | Group-Object -Property CDObjectType,CDAttribute -AsHashTable -AsString

$split = $OutputPath.Split(".")
$OutputPathInbound = $OutputPath.Replace($split[$split.count - 2], ($split[$split.count - 2] + "_inbound"))
$OutputPathOutbound = $OutputPath.Replace($split[$split.count - 2], ($split[$split.count - 2] + "_outbound"))


$matches = 0
#loop through first hashtable import keys checking if value exists in second hashtable
"`r`nChecking for maching Inbound flows" | Write-Host
#"`r`n`r`nmvObjbectType,mvAttribute" | Write-Host
#"-------------------------" | Write-Host
#work with sorted DictionaryEntry objects
$dicEntries = $FirstIAFhash.GetEnumerator() | sort -Property name
$IAFs = @()
for ($i = 0; $i -lt $dicEntries.Count; $i++)
{
    #check for matches
    if ($SecondIAFhash[$dicEntries[$i].key] -ne $null)
    {
        #Add Server name to attribute flows
        #Add result attribute to attribute flows
        $AFs = $FirstIAFhash[$dicEntries[$i].key]
        for ($x = 0; $x -lt $AFs.count; $x++)
        {
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $firstServerName        
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Overlap"
        }

        
        $AFs = $SecondIAFhash[$dicEntries[$i].key]
        for ($x = 0; $x -lt $AFs.count; $x++)
        {
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $secondServerName        
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Overlap"
        }

        $matches++
        $IAFs += $FirstIAFhash[$dicEntries[$i].key]
        $IAFs += $SecondIAFhash[$dicEntries[$i].key]
        #remove from dictionary to identify remaining non-matches
        $SecondIAFhash.Remove($dicEntries[$i].Key)
        $FirstIAFhash.Remove($dicEntries[$i].Key)
    }
}

foreach ($key in $FirstIAFhash.Keys)
{
    $AFs = $FirstIAFhash[$key]
    for ($x = 0; $x -lt $AFs.count; $x++)
    {
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $firstServerName        
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Unique"
    }
    $IAFs += $FirstIAFhash[$key]
}
foreach ($key in $SecondIAFhash.Keys)
{
    $AFs = $SecondIAFhash[$key]
    for ($x = 0; $x -lt $AFs.count; $x++)
    {
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $secondServerName        
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Unique"
    }
    $IAFs += $SecondIAFhash[$key]
}
WriteIAFs-ToFile -OutPath $OutputPathInbound -inboundAFs $IAFs

"`r`n`r`nTotal inbound matches - $($matches)" | Write-Host


$matches = 0
#loop through first hashtable import keys checking if value exists in second hashtable
"`r`nChecking for maching Outbound flows" | Write-Host
#"`r`n`r`ncsObjbectType,csAttribute" | Write-Host
#"-------------------------" | Write-Host
#work with sorted DictionaryEntry objects
$dicEntries = $FirstEAFhash.GetEnumerator() | sort -Property name
$EAFs = @()
for ($i = 0; $i -lt $dicEntries.Count; $i++)
{
    #check for matches
    if ($SecondEAFhash[$dicEntries[$i].key] -ne $null)
    {
        #Add Server name to attribute flows
        #Add result attribute to attribute flows
        $AFs = $FirstEAFhash[$dicEntries[$i].key]
        for ($x = 0; $x -lt $AFs.count; $x++)
        {
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $firstServerName        
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Overlap"
        }

        
        $AFs = $SecondEAFhash[$dicEntries[$i].key]
        for ($x = 0; $x -lt $AFs.count; $x++)
        {
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $secondServerName        
            $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Overlap"
        }

        $matches++
        $EAFs += $FirstEAFhash[$dicEntries[$i].key]
        $EAFs += $SecondEAFhash[$dicEntries[$i].key]
        #remove from dictionary to identify remaining non-matches
        $SecondEAFhash.Remove($dicEntries[$i].Key)
        $FirstEAFhash.Remove($dicEntries[$i].Key)
    }
}

foreach ($key in $FirstEAFhash.Keys)
{
    $AFs = $FirstEAFhash[$key]
    for ($x = 0; $x -lt $AFs.count; $x++)
    {
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $firstServerName        
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Unique"
    }
    $EAFs += $FirstEAFhash[$key]
}
foreach ($key in $SecondEAFhash.Keys)
{
    $AFs = $SecondEAFhash[$key]
    for ($x = 0; $x -lt $AFs.count; $x++)
    {
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $secondServerName        
        $AFs[$x] | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Unique"
    }
    $EAFs += $SecondEAFhash[$key]
}
WriteEAFs-ToFile -OutPath $OutputPathOutbound -outboundAFs $EAFs
"`r`n`r`nTotal outbound matches - $($matches)" | Write-Host