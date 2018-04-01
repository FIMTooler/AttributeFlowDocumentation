#region Functions


#Evaluates inputs and determines if defaults are used
Function CheckInputs([ref]$FilePath,[ref]$ObjectFilter)
{
    Try
    {
        If ($Inputs.Count -eq 0)
        {
            #No args - output message with example to console and proceed with defaults
            "`r`nThere are no arguments.  Exiting" | Write-Host
            "`r`n`r`nInputs`r`n------" | Write-Host
            "File Path (Required)               Path to exported FIM Sync Engine Configuration `r`n" | Write-Host
            "Object filter (Optional)               Filter used to select which objects are drawn" | Write-Host
            "     Default - *`r`n" | Write-Host
            "`r`n`r`nExamples`r`n--------" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`"`r`n" | Write-Host
            "ScriptName.ps1 `"C:\FIM Config`" `"Person`"`r`n" | Write-Host
            Exit
        }
        ElseIf ($Inputs.Count -eq 1)
        {
            #Only 1 input - used for File Path
            #defaults used for FollowFilterReferences and MPR Filter
            $FilePath.Value = $Inputs[0]
            "`r`nThere's only 1 argument.  It will be used as the File Path" | Write-Host
            "`r`n`r`nFile Path - $($FilePath.Value)" | Write-Host
            "Object Filter - *" | Write-Host
        }
        ElseIf ($Inputs.Count -eq 2)
        {
            #2 inputs provided
            $FilePath.Value = $Inputs[0]
            $ObjectFilter.Value = $Inputs[1].Split(',')
            "`r`nFile Path - $($FilePath.Value)" | Write-Host
            "Object Filter - $($ObjectFilter.Value)" | Write-Host
        }
        Else
        {
            #More than 2 inputs detected
            "ERROR - more than two arguments were passed" | Write-Host
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

#Starts Visio Application
Function StartVisio
{
    Try
    {
        $application = New-Object -ComObject Visio.Application
        $application
    }
    Catch
    {
        "Error in StartVisio function - Please ensure Visio 2010 is installed on PC where script is located" | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        Exit
    }
}

#Opens Visio template file
Function AddNewPageFromTemplate
{
    Try
    {
        #document class is Microsoft.Office.Interop.Visio.DocumentClass
        $Documents = $Application.Documents
		#Visio 2016
        $Document = $Documents.Add("BASIC_U.VSSX")
        $Document = $Documents.Add("ARROWS_U.VSSX")
		#Visio 2013 or 2010 ????
        #$Document = $Documents.Add("BASIC_U.VSS")
        #Visio 2016 Template BASICD_U.VSTX
        $Document = $Documents.Add("C:\Program Files\Microsoft Office\Office14\Visio Content\1033\BASICD_U.VSTX")
        $Document.printlandscape = $true
    }
    Catch
    {
        "Error in AddNewPageFromTemplate function" | Write-Host
        "BASICD_U.VST template may not be available" | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Gets top edge of shape object passed to it
Function GetTopEdge($Shape)
{
    #Find top of shape
    #Formula breakdown
        #(Shape height * .5) gets distance from center of shape to top/bottom edge
        #PinY is center of the shape
        #center of the shape + distance to top edge is location of top edge
    $Shape.Cells("PinY").ResultIU + ($Shape.Cells("height").ResultIU * .5)
}

#Gets bottom edge of shape object passed to it
Function GetBottomEdge($Shape)
{
    #Find bottom of shape
    #Formula breakdown
        #(Shape height * .5) gets distance from center of shape to top/bottom edge
        #PinY is center of the shape
        #center of the shape - distance to bottom edge is location of bottom edge
    $Shape.Cells("PinY").ResultIU - ($Shape.Cells("height").ResultIU * .5)
}

#Gets left edge of shape object passed to it
Function GetLeftEdge($Shape)
{
    #Find Left edge of shape
    #Formula breakdown
        #(shape width * .5) is distance from center of shape to left/right edge
        #PinX is vertical center of shape
        #Center of shape - distance to left edge is location of left edge
    $Shape.Cells("PinX").ResultIU - ($Shape.Cells("width").ResultIU * .5)
}

#Gets right edge of shape object passed to it
Function GetRightEdge($Shape)
{
    #Find Right edge of shape
    #Formula breakdown
        #(shape width * .5) is distance from center of shape to left/right edge
        #PinX is vertical center of shape
        #Center of shape + distance to right edge is location of right edge
    $Shape.Cells("PinX").ResultIU + ($Shape.Cells("width").ResultIU * .5)
}

#Sizes shape to fix text within it
Function ShapeFitText($Shape)
{
    Try
    {
        #make shape as wide as text box
        $ShapeWidth = $Shape.cells("width")
        $ShapeWidth.Formula = "TEXTWIDTH(theText)"

        #make shape as high as text box
        $ShapeHeight = $Shape.cells("height")
        $ShapeHeight.Formula = "TEXTHEIGHT(theText,width)"
    }
    Catch
    {
        "Error in ShapeFitText function" | Write-Host
        "Shape is " + $Shape | Write-Host
        "Page is " + $Page | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Centers shape horizontally on page
Function HCenterShapeOnPage($LocalPage,$Shape)
{
    Try
    {
        $Width = $LocalPage.pagesheet.cells("pagewidth").resultIU
        $Shape.cells("PinX").resultIU = $Width/2
    }
    Catch
    {
        "Error in HCenterShapeOnPage function" | Write-Host
        "Page is " + $LocalPage.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Centers shape vertically on page
Function VCenterShapeOnPage($LocalPage,$Shape)
{
    Try
    {
        $Height = $LocalPage.pagesheet.cells("pageheight").resultIU
        $Shape.cells("PinY").resultIU = $Height/2
    }
    Catch
    {
        "Error in VCenterShapeOnPage function" | Write-Host
        "Page is " + $LocalPage.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Moves shape's top edge to top page margin
Function AlignShapeTopWithTopMargin($LocalPage,$Shape)
{
    Try
    {
        $ShapeTop = GetTopEdge $Shape

        # -.25 accounts for top/bottom margins
        $PageHeight = ($LocalPage.pagesheet.cells("pageheight").resultIU - .25)

        #PinY is in center of shape
        #formula breakdown
            #(Top of shape - top of page) is distance above top of the page that the shape is
            #PinY - (Top of shape - top of page) is location PinY needs to be for top of shape to equal top of page
        $Shape.cells("pinY").resultIU = $Shape.cells("PinY").resultIU - ($ShapeTop - $Pageheight)
    }
    Catch
    {
        "Error in AlignShapetopWithTopMargin function" | Write-Host
        "Page is " + $LocalPage.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Moves shape's left edge to left margin
Function AlignShapeLeftWithLeftMargin($LocalPage,$Shape)
{
    Try
    {
        $ShapeLeft = GetLeftEdge $Shape

        #PinX is in center of shape
        #formula breakdown
            #(left edge of shape * -1) is distance left of the page that the shape is
            #PinX - (left edge of shape * -1) is location PinX needs to be for left edge of shape to equal left edge of page
            # + .25 moves PinX left enough for left edge of shape to align with left margin
        $Shape.cells("pinX").resultIU = $Shape.cells("PinX").resultIU + ($ShapeLeft * -1) + 0.25
        }
    Catch
    {
        "Error in AlignShapeLeftWithLeftMargin function" | Write-Host
        "Page is " + $LocalPage.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Resets Grig and Ruler so that bottom left corner is always 0,0
Function ResetGridAndRuler($LocalPage)
{
    Try
    {
        $LocalPage.PageSheet.cells("XRulerOrigin").ResultIU = 0
        $LocalPage.PageSheet.cells("XGridOrigin").ResultIU = 0
        $LocalPage.PageSheet.cells("YRulerOrigin").ResultIU = 0
        $LocalPage.PageSheet.cells("YGridOrigin").ResultIU = 0
    }
    Catch
    {
        "Error in ResetGrigAndRuler function" | Write-Host
        "Page is " + $LocalPage.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Creates Shape
Function AddShape($LocalPage, $Text)
{
    Try
    {
        $Shape = $Page.Drop($RectangleStencil,5.5,4.25)

        $Shape.Text = $Text

        $Shape.Cells("Para.HorzAlign").ResultIU = 0

        ShapeFitText $Shape

        $Shape
    }
    Catch
    {
        "Error in AddShape function" | Write-Host
        "Object name is " + ($Text.SubString(0,$Text.Indexof([char]10))) | Write-Host
        "Page is " + $LocalPage.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
        $Shape
    }
}

#Gets left edge of shape with left edge furthest left
Function GetLeftMostEdge
{
    Try
    {
        $LeftEdge = $Page.PageSheet.Cells("pagewidth").ResultIU
        $r = 1
        While (($r-1) -lt $Page.Shapes.Count)
        {
            #connectors exist in page shape list, but root shape isn't shape class
            If ($Page.Shapes[$r].Name.ToString().Contains("Rectangle"))
            {
                If ($LeftEdge -gt (GetLeftEdge $Page.Shapes[$r]))
                {
                    $LeftEdge = (GetLeftEdge $Page.Shapes[$r])
                }
            }
            $r = $r + 1
        }
        $LeftEdge
    }
    Catch
    {
        "Error in GetLeftMostEdge function" | Write-Host
        "Page is " + $Page.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
        $LeftEdge
    }
}

#Gets right edge of shape with right edge furthest right
Function GetRightMostEdge
{
    Try
    {
        $RightEdge = 0
        $r = 0
        While ($r -lt $Page.Shapes.Count)
        {
            If ($RightEdge -lt (GetRightEdge $Page.Shapes[($r+1)]))
            {
                $RightEdge = (GetRightEdge $Page.Shapes[($r+1)])
            }
            $r = $r + 1
        }
        $RightEdge
    }
    Catch
    {
        "Error in GetRightMostEdge function" | Write-Host
        "Page is " + $Page.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
        $RightEdge
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
						    $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $scriptContext
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
						
						    if ($importFlow.'sync-rule-mapping'.'mapping-type' -ieq 'expression')
						    {
							    $scriptContext = $importFlow.'sync-rule-mapping'.'sync-rule-value'.'import-flow'.InnerXml
						    }
							elseif ($importFlow.'sync-rule-mapping'.'mapping-type' -ieq 'constant')
							{
							  $scriptContext = $importFlow.'sync-rule-mapping'.'sync-rule-value'
							}
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value $ruleType
		                    $rule | Add-Member -MemberType noteproperty -name 'SourceMA' -value $srcMA
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $srcAttributes
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $mvAttribute
						    $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $scriptContext
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
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'DIRECT'
		                    $rule | Add-Member -MemberType noteproperty -name 'MAName' -value $maName                
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $srcAttribute
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $cdAttribute
						    $rule | Add-Member -MemberType noteproperty -name 'AllowNulls' -value $allowNulls
		                
		                    $rules += $rule
		                }
                        if ($exportFlow.'constant-mapping' -ne $null)
		                {
                            ###
                            ### Handle src-attribute that are intrinsic (<src-attribute intrinsic="true">object-id</src-attribute>)
                            ###
                            $constant = $exportFlow.'constant-mapping'.'constant-value'
		                
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'CONSTANT'
		                    $rule | Add-Member -MemberType noteproperty -name 'MAName' -value $maName                
		                    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $cdAttribute
						    $rule | Add-Member -MemberType noteproperty -name 'ConstantValue' -value $constant
						    $rule | Add-Member -MemberType noteproperty -name 'AllowNulls' -value $allowNulls
		                
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
                                elseIf ($_) # Do not add empty values.
                                {
		                            $srcAttributes += $_
                                }
                            }
                            if ($srcAttributes.Count -eq 1)
                            {
                                $srcAttributes = $srcAttributes -as [String]
                            }
                            # (Commented) Leave as collection.
                            #if ($srcAttributes.Count-eq 1)
                            #{
                            #    $srcAttributes = $srcAttributes -as[String]
                            #}
		                    
		                    $rule = New-Object PSObject
		                    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'SCRIPTED'
		                    $rule | Add-Member -MemberType noteproperty -name 'MAName' -value $maName
						    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $srcAttributes
		                    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
		                    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $cdAttribute	
		                    $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $scriptContext
						    $rule | Add-Member -MemberType noteproperty -name 'AllowNulls' -value $allowNulls
		                                
		                    $rules += $rule                        
		                }
					    elseif ($exportFlow.'sync-rule-mapping' -ne $null)
					    {
                            $syncRuleID = $exportFlow.'sync-rule-mapping'.'sync-rule-id'
                            $isExistenceTest = $exportFlow.'sync-rule-mapping'.'is-existence-test'
                            $initialFlowOnly = $exportFlow.'sync-rule-mapping'.'initial-flow-only'
						    $srcAttribute = $exportFlow.'sync-rule-mapping'.'src-attribute'
						    if($exportFlow.'sync-rule-mapping'.'mapping-type' -eq 'direct')
						    {
							    $rule = New-Object PSObject
							    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'OSR-Direct'
                                $rule | Add-Member -MemberType NoteProperty -Name 'Sync Rule ID' -Value $syncRuleID
                                $rule | Add-Member -MemberType NoteProperty -Name 'is Existence Test' -Value $isExistenceTest
                                $rule | Add-Member -MemberType NoteProperty -Name 'initial flow only' -Value $initialFlowOnly
							    $rule | Add-Member -MemberType noteproperty -name 'MAName' -value $maName
							    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
							    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $srcAttribute
							    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
							    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $cdAttribute
							    $rule | Add-Member -MemberType noteproperty -name 'AllowNulls' -value $allowNulls
											
							    $rules += $rule             
						    }
						    elseif ($exportFlow.'sync-rule-mapping'.'mapping-type' -eq 'expression')
						    {
							    $scriptContext = $exportFlow.'sync-rule-mapping'.'sync-rule-value'.'export-flow'.InnerXml
							    $cdAttribute = $exportFlow.'sync-rule-mapping'.'sync-rule-value'.'export-flow'.dest
							    $rule = New-Object PSObject
							    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'OSR-Expression'
                                $rule | Add-Member -MemberType NoteProperty -Name 'Sync Rule ID' -Value $syncRuleID
                                $rule | Add-Member -MemberType NoteProperty -Name 'is Existence Test' -Value $isExistenceTest
                                $rule | Add-Member -MemberType NoteProperty -Name 'initial flow only' -Value $initialFlowOnly
							    $rule | Add-Member -MemberType noteproperty -name 'MAName' -value $maName
							    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
							    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $srcAttribute
							    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
							    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $cdAttribute														
							    $rule | Add-Member -MemberType noteproperty -name 'ScriptContext' -value $scriptContext
							    $rule | Add-Member -MemberType noteproperty -name 'AllowNulls' -value $allowNulls
											
							    $rules += $rule             
						    }
                            elseif($exportFlow.'sync-rule-mapping'.'mapping-type' -eq 'constant')
						    {
                                $scriptContext = $exportFlow.'sync-rule-mapping'.'sync-rule-value'
							    $rule = New-Object PSObject
							    $rule | Add-Member -MemberType noteproperty -name 'RuleType' -value 'OSR-Constant'
                                $rule | Add-Member -MemberType NoteProperty -Name 'Sync Rule ID' -Value $syncRuleID
                                $rule | Add-Member -MemberType NoteProperty -Name 'is Existence Test' -Value $isExistenceTest
                                $rule | Add-Member -MemberType NoteProperty -Name 'initial flow only' -Value $initialFlowOnly
							    $rule | Add-Member -MemberType noteproperty -name 'MAName' -value $maName
							    $rule | Add-Member -MemberType noteproperty -name 'MVObjectType' -value $mvObjectType
							    $rule | Add-Member -MemberType noteproperty -name 'MVAttribute' -value $srcAttribute
							    $rule | Add-Member -MemberType noteproperty -name 'CDObjectType' -value $cdObjectType
							    $rule | Add-Member -MemberType noteproperty -name 'CDAttribute' -value $cdAttribute														
							    $rule | Add-Member -MemberType noteproperty -name 'ConstantValue' -value $ScriptContext
							    $rule | Add-Member -MemberType noteproperty -name 'AllowNulls' -value $allowNulls
											
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

#Loops through IAFShapes and positions them on page
Function PositionIAFs
{
    Try
    {
        $Y = 0
        While ($Y -lt $IAFShapes.Count)
        {
            #position left of MV object
            If ($Y -eq 0)
            {
                #height of new top edge
                $NewTopEdge = ($Page.pagesheet.cells("pageheight").resultIU - .25)
        
                #height of current top edge
                $CurrentTopEdge = GetTopEdge $IAFShapes[$Y]
                 
                #Set PinY for first shape to top margin of page
                $IAFShapes[$Y].Cells("PinY").ResultIU = $IAFShapes[$Y].Cells("PinY").ResultIU + ($NewTopEdge - $CurrentTopEdge)

                $NewRightEdge = (GetLeftEdge $MVShape) - 2
                $CurrentRightEdge = GetRightEdge $IAFShapes[$Y]
                $IAFShapes[$Y].Cells("PinX").ResultIU = $IAFShapes[$Y].Cells("PinX").ResultIU + ($NewRightEdge - $CurrentRightEdge)
            }
            #Position under previous set object
            else
            {
                #height of new top edge
                $NewTopEdge = ((GetBottomEdge $IAFShapes[($Y-1)]) -.5)
        
                #height of current top edge
                $CurrentTopEdge = GetTopEdge $IAFShapes[$Y]
        
                #PinY needs to move the difference between the new and current top edges
                $IAFShapes[$Y].Cells("PinY").ResultIU = $IAFShapes[$Y].Cells("PinY").ResultIU + ($NewTopEdge - $CurrentTopEdge)

                $NewRightEdge = GetRightEdge $IAFShapes[($Y-1)]
                $CurrentRightEdge = GetRightEdge $IAFShapes[$Y]
                $IAFShapes[$Y].Cells("PinX").ResultIU = $IAFShapes[$Y].Cells("PinX").ResultIU + ($NewRightEdge - $CurrentRightEdge)
            }
            $Y = $Y + 1
            Start-Sleep -m 500
        }
    }
    Catch
    {
        "Error in PositionIAFs function" | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Loops through EAFShapes and positions them on page
Function PositionEAFs
{
    Try
    {
        $Y = 0
        While ($Y -lt $EAFShapes.Count)
        {
            #position right of MV object
            If ($Y -eq 0)
            {
                 #height of new top edge
                $NewTopEdge = ($Page.pagesheet.cells("pageheight").resultIU - .25)
        
                #height of current top edge
                $CurrentTopEdge = GetTopEdge $EAFShapes[$Y]
                 
                #Set PinY for first shape to top margin of page
                $EAFShapes[$Y].Cells("PinY").ResultIU = $EAFShapes[$Y].Cells("PinY").ResultIU + ($NewTopEdge - $CurrentTopEdge)

                $NewLeftEdge = (GetRightEdge $MVShape) + 2
                $CurrentLeftEdge = GetLeftEdge $EAFShapes[$Y]
                $EAFShapes[$Y].Cells("PinX").ResultIU = $EAFShapes[$Y].Cells("PinX").ResultIU + ($NewLeftEdge - $CurrentLeftEdge)
            }
            #Position under previous set object
            else
            {
                #height of new top edge
                $NewTopEdge = ((GetBottomEdge $EAFShapes[($Y-1)]) -.5)
        
                #height of current top edge
                $CurrentTopEdge = GetTopEdge $EAFShapes[$Y]
        
                #PinY needs to move the difference between the new and current top edges
                $EAFShapes[$Y].Cells("PinY").ResultIU = $EAFShapes[$Y].Cells("PinY").ResultIU + ($NewTopEdge - $CurrentTopEdge)

                $NewLeftEdge = GetLeftEdge $EAFShapes[($Y-1)]
                $CurrentLeftEdge = GetLeftEdge $EAFShapes[$Y]
                $EAFShapes[$Y].Cells("PinX").ResultIU = $EAFShapes[$Y].Cells("PinX").ResultIU + ($NewLeftEdge - $CurrentLeftEdge)
            }
            $Y = $Y + 1
            Start-Sleep -m 500
        }
    }
    Catch
    {
        "Error in PositionEAFs function" | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Formats Text header for each shape as size 12 and bold and underlined
Function FormatTextHeader($Characters,$Length)
{
    Try
    {
        $Characters.Begin = 0
        $Characters.End = $Length
        $Characters.CharProps(2) = 5
        $Characters.CharProps(7) = 14
    }
    Catch
    {
        "Error in FormatTextHeader function" | Write-Host
        "Text is " + $Characters.Text | Write-Host
        "Page is " + $Page.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Formats Text body with 12pt line spacing
Function FormatTextBody($Characters, $Begin)
{
    Try
    {
        $Characters.Begin = $Begin
        $Characters.End = $Characters.End
        $Characters.ParaProps(3) = 12
        $Characters.CharProps(7) = 10
    }
    Catch
    {
        "Error in FormatTextBody function" | Write-Host
        "Text is " + $Characters.Text | Write-Host
        "Page is " + $Page.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#Formats MV Text
Function FormatMVText($Characters, $Begin)
{
    Try
    {
        $Characters.Begin = $Begin
        $Characters.End = $Characters.End + 1
        $Characters.CharProps(2) = 1
        $Characters.CharProps(7) = 16
        $Characters.ParaProps(3) = 18
    }
    Catch
    {
        "Error in FormatTextBody function" | Write-Host
        "Text is " + $Characters.Text | Write-Host
        "Page is " + $Page.Name | Write-Host
        "`r`n`r`nStack trace is`r`n`r`n$($Error)`r`n`r`n" | Write-Host
        $Error.Clear()
    }
}

#endregion






#Set default parameters
$FilePath = ""
$ObjectFilter = @()
$ObjectFilter += "*"

#puts script arguments into array so they can be used
$Inputs = @()
Foreach ($Arg in $args)
{
    $Inputs = $Inputs + $Arg
}

CheckInputs ([ref]$FilePath) ([ref]$ObjectFilter)

$AFs = Get-ImportToExportAttributeFlow $FilePath $ObjectFilter #'.\FIM Export Config'

#open and setup Viso Application
$Application = StartVisio
AddNewPageFromTemplate
$BasicStencils = $application.Documents[1]
#Visio 2016
$ArrowStencils = $application.Documents[2]
$RectangleStencil = $BasicStencils.Masters.Item("Rectangle")
$ArrowStencil = $ArrowStencils.Masters.Item("Simple Arrow")
#Visio 2010 or 2013 ???
#$RectangleStencil = $BasicStencils.Masters.Item("Rectangle")
#$ArrowStencil = $BasicStencils.Masters.Item("45 degree single")

$Pages = $application.ActiveDocument.Pages
$Page = $Pages.Item(1)
$Page.AutoSizeDrawing()
ResetGridAndRuler $Page
$Application.ActiveWindow.ShowPageBreaks = 0
$Application.ActiveWindow.ShowConnectPoints = 0


Foreach ($AF in $AFs)
{
    if ($AF.MVAttribute -ne $null)
    {
        $MVText = [char]10 + "Metaverse Object Type:    " + $AF.MVObjectType + [char]10 + [char]10 + "Metaverse Attribute:         " + $AF.MVAttribute + [char]10 + " "
        $ProposedName = $AF.MVObjectType + " : " + $AF.MVAttribute
        [string[]]$UsedNames = @()
        $Pages.GetNames([ref]$UsedNames)
        $i = 1
        While ($UsedNames.ToLower().Contains($ProposedName.ToLower()))
        {
            if ($i -eq 1)
            {
                $ProposedName += "_$i"
            }
            else
            {
                $ProposedName = $ProposedName.Remove($ProposedName.IndexOf("_"))
                $ProposedName += "_$i"
            }
            $i += 1
        }
        $Page.Name = $ProposedName
    }
    else
    {
        $i = 1
        $MVText = [char]10 + "Metaverse Object Type:    " + $AF.MVObjectType + [char]10 + " "
        $ProposedName = $AF.MVObjectType + " : CONSTANT"
        [string[]]$UsedNames = @()
        $Pages.GetNames([ref]$UsedNames)
        $i = 1
        While ($UsedNames.ToLower().Contains($ProposedName.ToLower()))
        {
            if ($i -eq 1)
            {
                $ProposedName += "_$i"
            }
            else
            {
                $ProposedName = $ProposedName.Remove($ProposedName.IndexOf("_"))
                $ProposedName += "_$i"
            }
            $i += 1
        }
        $Page.Name = $ProposedName
    }

    $MVShape = AddShape $Page $MVText

    FormatMVText $MVShape.Characters 0
    ShapeFitText $MVShape
    Start-Sleep -m 500

    if ($AF.IAFs.Count -gt 0)
    {
        $IAFShapes = @()
        Foreach ($IAF in $AF.IAFs)
        {
            $IAFHeader = $IAF.SourceMA + "     " + $IAF.RuleType + [char]10 + [char]10
            $IAFText = $IAFHeader
            $IAFProperties = $IAF | Get-Member | Where { $_.MemberType -eq "NoteProperty" } | select Name
            Foreach ($IAFProperty in $IAFProperties)
            {
                if ($IAFProperty.Name -ne "SourceMA" -and $IAFProperty.Name -ne "RuleType")
                {
                    $IAFText += $IAFProperty.Name + ": " + $IAF.$($IAFProperty.Name) + [char]10
                }
            }
            $IAFShapes += AddShape $Page $IAFText
            #Format Text
            FormatTextHeader $IAFShapes[($IAFShapes.Length-1)].Characters $IAFHeader.Length
            FormatTextBody $IAFShapes[($IAFShapes.Length-1)].Characters ($IAFHeader.Length + 1)
            ShapeFitText $IAFShapes[($IAFShapes.Length-1)]
        }
        if ($IAFShapes.Count -gt 1)
        {
            PositionIAFs
            $Application.ActiveWindow.DeselectAll()
            $Selection = $Application.ActiveWindow.Selection
            Foreach ($IAFShape in $IAFShapes)
            {
                $Selection.Select($IAFShape,2)
            }
            Start-Sleep -m 500
            $GroupShape = $Selection.Group()
            VCenterShapeOnPage $Page $GroupShape
            $GroupShape.Ungroup()

            if (($MVShape.Cells("Height").ResultIU/.5) -lt $IAFShapes.Count)
            {
                $MVShape.Cells("Height").ResultIU = $IAFShapes.Count * .5
            }

            $PreviousEndY = 0
            Foreach ($IAFShape in $IAFShapes)
            {
                $BeginX = GetRightEdge $IAFShape
                $EndX =  GetLeftEdge $MVShape
                $BeginY = $IAFShape.Cells("PinY").ResultIU
                if ($PreviousEndY -eq 0)
                {
                    $EndY = (GetTopEdge $MVShape) - (($MVShape.Cells("Height").ResultIU/$IAFShapes.Count)/2)
                }
                else
                {
                    $EndY = $PreviousEndY - ($MVShape.Cells("Height").ResultIU/$IAFShapes.Count)#$MVShape.Cells("PinY").ResultIU
                }
                $PreviousEndY = $EndY
                $ArrowShape = $Page.Drop($ArrowStencil,(($BeginX + $EndX)/2),(($BeginY + $EndY)/2))
                Start-Sleep -m 500
                $ArrowShape.Cells("BeginX").ResultIU = $BeginX
                $ArrowShape.Cells("EndX").ResultIU = $EndX
                $ArrowShape.Cells("BeginY").ResultIU = $BeginY
                $ArrowShape.Cells("EndY").ResultIU = $EndY
            }
        }
        else
        {
            $NewRightEdge = (GetLeftEdge $MVShape) - 2
            $CurrentRightEdge = GetRightEdge $IAFShapes[0]
            $IAFShapes[0].Cells("PinX").ResultIU = $IAFShapes[0].Cells("PinX").ResultIU + ($NewRightEdge - $CurrentRightEdge)
            
            Start-Sleep -m 500
            $BeginX = GetRightEdge $IAFShapes[0]
            $EndX = GetLeftEdge $MVShape
            $BeginY = $IAFShapes[0].Cells("PinY").ResultIU
            $EndY = $MVShape.Cells("PinY").ResultIU
            $ArrowShape = $Page.Drop($ArrowStencil,(($BeginX + $EndX)/2),(($BeginY + $EndY)/2))
            Start-Sleep -m 500
            $ArrowShape.Cells("BeginX").ResultIU = $BeginX
            $ArrowShape.Cells("EndX").ResultIU = $EndX
            $ArrowShape.Cells("BeginY").ResultIU = $BeginY
            $ArrowShape.Cells("EndY").ResultIU = $EndY
        }
    }

    if ($AF.EAFs.Count -gt 0)
    {
        $EAFShapes = @()
        Foreach ($EAF in $AF.EAFs)
        {
            $EAFHeader = $EAF.MAName + "     " + $EAF.RuleType + [char]10 + [char]10
            $EAFText = $EAFHeader
            $EAFProperties = $EAF | Get-Member | Where { $_.MemberType -eq "NoteProperty" } | select Name
            Foreach ($EAFProperty in $EAFProperties)
            {
                if ($EAFProperty.Name -ne "MAName" -and $EAFProperty.Name -ne "RuleType")
                {
                    $EAFText += $EAFProperty.Name + ": " + $EAF.$($EAFProperty.Name) + [char]10
                }
            }
            $EAFShapes += AddShape $Page $EAFText
            
            #Format Text
            FormatTextHeader $EAFShapes[($EAFShapes.Length-1)].Characters $EAFHeader.Length
            FormatTextBody $EAFShapes[($EAFShapes.Length-1)].Characters ($EAFHeader.Length + 1)
            ShapeFitText $EAFShapes[($EAFShapes.Length-1)]
        }
        if ($EAFShapes.Count -gt 1)
        {
            PositionEAFs
            $Application.ActiveWindow.DeselectAll()
            $Selection = $Application.ActiveWindow.Selection
            Foreach ($EAFShape in $EAFShapes)
            {
                $Selection.Select($EAFShape,2)
            }
            Start-Sleep -m 500
            $GroupShape = $Selection.Group()
            VCenterShapeOnPage $Page $GroupShape
            $GroupShape.Ungroup()

            $PreviousBeginY = 0
            Foreach ($EAFShape in $EAFShapes)
            {
                $BeginX = GetRightEdge $MVShape
                $EndX = GetLeftEdge $EAFShape
                if ($PreviousBeginY -eq 0)
                {
                    $BeginY = (GetTopEdge $MVShape) - (($MVShape.Cells("Height").ResultIU/$EAFShapes.Count)/2)
                }
                else
                {
                    $BeginY = $PreviousBeginY - ($MVShape.Cells("Height").ResultIU/$EAFShapes.Count)#$MVShape.Cells("PinY").ResultIU
                }
                $PreviousBeginY = $BeginY
                $EndY = $EAFShape.Cells("PinY").ResultIU
                $ArrowShape = $Page.Drop($ArrowStencil,(($BeginX + $EndX)/2),(($BeginY + $EndY)/2))
                Start-Sleep -m 500
                $ArrowShape.Cells("BeginX").ResultIU = $BeginX
                $ArrowShape.Cells("EndX").ResultIU = $EndX
                $ArrowShape.Cells("BeginY").ResultIU = $BeginY
                $ArrowShape.Cells("EndY").ResultIU = $EndY
            }
        }
        else
        {
            $NewLeftEdge = (GetRightEdge $MVShape) + 2
            $CurrentLeftEdge = GetLeftEdge $EAFShapes[0]
            $EAFShapes[0].Cells("PinX").ResultIU = $EAFShapes[0].Cells("PinX").ResultIU + ($NewLeftEdge - $CurrentLeftEdge)
            
            Start-Sleep -m 500
            $BeginX = GetRightEdge $MVShape
            $EndX = GetLeftEdge $EAFShapes[0]
            $BeginY = $MVShape.Cells("PinY").ResultIU
            $EndY = $EAFShapes[0].Cells("PinY").ResultIU
            $ArrowShape = $Page.Drop($ArrowStencil,(($BeginX + $EndX)/2),($BeginY + $EndY/2))
            Start-Sleep -m 500
            $ArrowShape.Cells("BeginX").ResultIU = $BeginX
            $ArrowShape.Cells("EndX").ResultIU = $EndX
            $ArrowShape.Cells("BeginY").ResultIU = $BeginY
            $ArrowShape.Cells("EndY").ResultIU = $EndY
        }
    }

    #resizes drawing to fix diagram
    $Page.AutoSizeDrawing()
    ResetGridAndRuler $Page

    #Groups all shapes and aligns single grouped shape
    $Application.ActiveWindow.SelectAll()
    $Selection = $Application.ActiveWindow.Selection
    Start-Sleep -m 500
    $GroupShape = $Selection.Group()
    AlignShapeTopWithTopMargin $page $GroupShape
    AlignShapeLeftWithLeftMargin $page $GroupShape
    $page.AutoSizeDrawing()
    ResetGridAndRuler $page
    Start-Sleep -m 500
    HCenterShapeOnPage $Page $GroupShape
    VCenterShapeOnPage $Page $GroupShape

    #Sets zoom fix to window
    $Application.ActiveWindow.Zoom = -1
    
    #Create new page for next MPR
    $Page = $Pages.Add()
    $Page.AutoSizeDrawing()
    ResetGridAndRuler $Page

    #Sets zoom fix to window
    $Application.ActiveWindow.Zoom = -1
}

$Page.Delete(0)

#clean up all variables
#get-item Variable:* | forEach { Remove-Variable $_.Name -ErrorAction SilentlyContinue; $Error.Clear(); }
