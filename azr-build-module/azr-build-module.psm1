########Null out variables
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0
####################################################

$password = "" | ConvertTo-SecureString -AsPlainText -Force
$username = ""
$credential = New-Object System.Management.Automation.PSCredential($username,$password)

#Build Sheet location
$BSFile = 'C:\Users\Jeremy Whiting\OneDrive - Rackspace Inc\azure\Customer Builds\Hawthorne\Build Sheet-170926-06700.xlsx'

$SheetArray = @('Subscriptions','Environments','ResourceGroups','VirtualNetworks','OMSWOrkspaces','WebApps','TrafficManager','StorageAccounts')

foreach($sheetName in $SheetArray){
    $objExcel = New-Object -ComObject Excel.Application
    $workbook = $objExcel.Workbooks.Open($BSFile) 
    $sheet = $workbook.Worksheets.Item($sheetName)
    $objExcel.Visible=$false

    $rowMax = ($Sheet.UsedRange.Rows).count

    switch($sheetName){
        "Subscriptions" {

                            $rowSubName,$colSubName = 1,1
                            $rowSubID,$colSubID = 1,2
                            $rowTenantID,$colTenantID = 1,3
                            $rowCOREID,$colCOREID = 1,4
                            $rowDeploymentName,$colDeploymentName = 1,5
                            $rowBuildBy,$colBuildBy = 1,6
                            $rowBuildDate,$colBuildDate = 1,7
                            $rowTicket,$colTicket = 1,8

                            $SubHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $SubHash += @{
                                                SubName = $sheet.Cells.Item($rowSubName+$i,$colSubName).text; `
                                                SubID=$sheet.Cells.Item($rowSubID+$i,$colSubID).text; `
                                                TenID=$sheet.Cells.Item($rowTenantID+$i,$colTenantID).text; `
                                                CoreID=$sheet.Cells.Item($rowCOREID+$i,$colCOREID).text; `
                                                DeploymentName=$sheet.Cells.Item($rowDeploymentName+$i,$colDeploymentName).text; `
                                                BuildBy=$sheet.Cells.Item($rowBuildBy+$i,$colBuildBy).text; `
                                                BuildDate=$sheet.Cells.Item($rowBuildDate+$i,$colBuildDate).text; `
                                                Ticket=$sheet.Cells.Item($rowTicket+$i,$colTicket).text 
                                            }

                            }

                        }
        "Environments" {

                            $rowEnvironment,$colEnvironment = 1,1
                            $rowLocation,$colLocation = 1,2 
                            $rowSubName,$colSubName = 1,3 

                            $ENVHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $ENVHash += @{
                                                ENV = $sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text; `
                                                ENVLocation = $sheet.Cells.Item($rowLocation+$i,$colLocation).text; `
                                                ENVSubName = $sheet.Cells.Item($rowSubName+$i,$colSubName).text 
                                            }
                            }

                        }
        "ResourceGroups" {
                            
                            $rowResourceGroupName,$colResourceGroupName = 1,1
                            $rowLocation,$colLocation = 1,2 
                            $rowSubName,$colSubName = 1,3 
                            $rowEnvironment,$colEnvironment = 1,4
                             
                            $RSGHash += @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $RSGHash += @{
                                                RSG = $sheet.Cells.Item($rowResourceGroupName+$i,$colResourceGroupName).text; `
                                                RSGLocation = $sheet.Cells.Item($rowLocation+$i,$colLocation).text; `
                                                RSGSubName = $sheet.Cells.Item($rowSubName+$i,$colSubName).text; ` 
                                                RSGENV = $sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text
                                            }
                            }

                        }
        "VirtualNetworks" {

                            $rowVirtualNetworkName,$colVirtualNetworkName = 1,1
                            $rowVirtualNetworkCIDR,$colVirtualNetworkCIDR = 1,2 
                            $rowVirutalNetworkRSG,$colVirutalNetworkRSG = 1,3
                            $rowVirutalNetworkLocation,$colVirutalNetworkLocation = 1,4 
                            $rowVirutalNetworkEnvironment,$colVirutalNetworkEnvironment = 1,5 
                            $rowVirutalNetworkSubName,$colVirutalNetworkSubName = 1,6 

                            $VNETHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $VNETHash += @{
                                                VNET = $sheet.Cells.Item($rowVirtualNetworkName+$i,$colVirtualNetworkName).text; `
                                                VNETCIDR = $sheet.Cells.Item($rowVirtualNetworkCIDR+$i,$colVirtualNetworkCIDR).text; `
                                                VNETRSG = $sheet.Cells.Item($rowVirutalNetworkRSG+$i,$colVirutalNetworkRSG).text; `
                                                VNETLocation = $sheet.Cells.Item($rowVirutalNetworkLocation+$i,$colVirutalNetworkLocation).text; `
                                                VNETENV = $sheet.Cells.Item($rowVirutalNetworkEnvironment+$i,$colVirutalNetworkEnvironment).text; ` 
                                                VNETSubName = $sheet.Cells.Item($rowVirutalNetworkSubName+$i,$colVirutalNetworkSubName).text 
                                            }
                            }

                        }
        "OMSWorkspaces" {

                            $rowOMSWorkspaceName,$colOMSWorkspaceName = 1,1
                            $rowOMSserviceTier,$colOMSserviceTier = 1,2 
                            $rowOMSlocation,$colOMSlocation = 1,3
                            $rowEnvironment,$colEnvironment = 1,4
                            $rowOMSRSG,$colOMSRSG = 1,5 
                            $rowbuildDate,$colbuildDate = 1,6 
                            $rowbuildBy,$colbuildBy = 1,7 
                            $rowTemplate,$colTemplate = 1,8 
                            $rowSAS,$colSAS = 1,9 

                            $OMSHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $OMSHash += @{
                                                OMSworkspaceName = $sheet.Cells.Item($rowOMSWorkspaceName+$i,$colOMSWorkspaceName).text; `
                                                OMSserviceTier = $sheet.Cells.Item($rowOMSserviceTier+$i,$colOMSserviceTier).text; `
                                                OMSlocation = $sheet.Cells.Item($rowOMSlocation+$i,$colOMSlocation).text; `
                                                environment = $sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text; `
                                                OMSRSG = $sheet.Cells.Item($rowOMSRSG+$i,$colOMSRSG).text; `
                                                buildDate = $sheet.Cells.Item($rowbuildDate+$i,$colbuildDate).text; `
                                                buildBy = $sheet.Cells.Item($rowbuildBy+$i,$colbuildBy).text; `
                                                template = $sheet.Cells.Item($rowTemplate+$i,$colTemplate).text; `
                                                sas = $sheet.Cells.Item($rowSAS+$i,$colSAS).text  
                                            }
                            }

                        }
        "WebApps"       {
                            $rowwebAppNames,$colwebAppNames = 1,1
                            $rowskuName,$colskuName = 1,2 
                            $rowskuCapacity,$colskuCapacity = 1,3
                            $rowAppServicePlanName,$colAppServicePlanName = 1,4
                            $rownetFrameworkVersion,$colnetFrameworkVersion = 1,5 
                            $rowphpVersion,$colphpVersion = 1,6
                            $rowpythonVersion,$colpythonVersion = 1,7
                            $row32Bit,$col32Bit = 1,8 
                            $rowwebSockets,$colwebSockets = 1,9
                            $rowalwaysOn,$colalwaysOn = 1,10
                            $rowwebServerLogging,$colwebServerLogging = 1,11 
                            $rowdetailedErrors,$coldetailedErrors = 1,12 
                            $rowfailedRequestTrace,$colfailedRequestTrace = 1,13
                            $rowclientAffinityEnabled,$colclientAffinityEnabled = 1,14 
                            $rowlogSize,$collogSize = 1,15
                            $rowEnvironment,$colEnvironment = 1,16
                            $rowWebAppRSG,$colWebAppRSG = 1,17
                            $rowbuildDate,$colbuildDate = 1,18
                            $rowbuildBy,$colbuildBy = 1,19
                            $rowTemplate,$colTemplate = 1,20
                            $rowSAS,$colSAS = 1,21

                            $WebAppHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $WebAppHash += @{
                                                    AppServicePlanName = $sheet.Cells.Item($rowwebAppNames+$i,$colwebAppNames).text; `
                                                    skuName = $sheet.Cells.Item($rowskuName+$i,$colskuName).text; `
                                                    skuCapacity = [Convert]::ToInt32($sheet.Cells.Item($rowskuCapacity+$i,$colskuCapacity).text); `
                                                    webAppNames = @($sheet.Cells.Item($rowAppServicePlanName+$i,$colAppServicePlanName).text); `
                                                    netFrameworkVersion = $sheet.Cells.Item($rownetFrameworkVersion+$i,$colnetFrameworkVersion).text; `
                                                    phpVersion = $sheet.Cells.Item($rowphpVersion+$i,$colphpVersion).text; `
                                                    pythonVersion = $sheet.Cells.Item($rowpythonVersion+$i,$colpythonVersion).text; `
                                                    '32Bit' = ($sheet.Cells.Item($row32Bit+$i,$col32Bit).text).tolower(); `
                                                    webSockets = ($sheet.Cells.Item($rowwebSockets+$i,$colwebSockets).text).tolower(); `
                                                    alwaysOn = ($sheet.Cells.Item($rowalwaysOn+$i,$colalwaysOn).text).tolower(); `
                                                    webServerLogging = ($sheet.Cells.Item($rowwebServerLogging+$i,$colwebServerLogging).text).tolower(); `
                                                    detailedErrors = ($sheet.Cells.Item($rowdetailedErrors+$i,$coldetailedErrors).text).tolower(); ` 
                                                    failedRequestTrace = ($sheet.Cells.Item($rowfailedRequestTrace+$i,$colfailedRequestTrace).text).tolower(); `
                                                    clientAffinityEnabled = ($sheet.Cells.Item($rowclientAffinityEnabled+$i,$colclientAffinityEnabled).text).tolower(); `
                                                    logSize = [Convert]::ToInt32($sheet.Cells.Item($rowlogSize+$i,$collogSize).text); `
                                                    environment = $sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text; `
                                                    WebAppRSG = $sheet.Cells.Item($rowWebAppRSG+$i,$colWebAppRSG).text; `
                                                    buildDate = $sheet.Cells.Item($rowbuildDate+$i,$colbuildDate).text; `
                                                    buildBy = $sheet.Cells.Item($rowbuildBy+$i,$colbuildBy).text; `
                                                    template = $sheet.Cells.Item($rowTemplate+$i,$colTemplate).text; `
                                                    sas = $sheet.Cells.Item($rowSAS+$i,$colSAS).text
                                            }
                            }


                        }
       "TrafficManager" {

                            $rowTrafficManagerName,$colTrafficManagerName = 1,1
                            $rowTrafficManagerRelative,$colTrafficManagerRelative = 1,2 
                            $rowTrafficRoutingMethod,$coltrafficRoutingMethod = 1,3
                            $rowTrafficManagerTTL,$colTrafficManagerTTL = 1,4
                            $rowTrafficManagerMonitorProtocol,$colTrafficManagerMonitorProtocol = 1,5 
                            $rowTrafficManagerMonitorPort,$colTrafficManagerMonitorPort = 1,6
                            $rowTrafficManagerMonitorPath,$colTrafficManagerMonitorPath = 1,7
                            $rowTrafficManagerRSG,$colTrafficManagerRSG = 1,8 


                            $TMHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $TMHash += @{
                                                name = $sheet.Cells.Item($rowTrafficManagerName+$i,$colTrafficManagerName).text; `
                                                relativeName = $sheet.Cells.Item($rowTrafficManagerRelative+$i,$colTrafficManagerRelative).text; `
                                                trafficRoutingMethod = $sheet.Cells.Item($rowTrafficRoutingMethod+$i,$coltrafficRoutingMethod).text; `
                                                ttl = [convert]::ToInt32($sheet.Cells.Item($rowTrafficManagerTTL+$i,$colTrafficManagerTTL).text); `
                                                MonitorProtocol = $sheet.Cells.Item($rowTrafficManagerMonitorProtocol+$i,$colTrafficManagerMonitorProtocol).text; ` 
                                                MonitorPort = [convert]::ToInt32($sheet.Cells.Item($rowTrafficManagerMonitorPort+$i,$colTrafficManagerMonitorPort).text); `
                                                MonitorPath = $sheet.Cells.Item($rowTrafficManagerMonitorPath+$i,$colTrafficManagerMonitorPath).text; ` 
                                                ResourceGroup = $sheet.Cells.Item($rowTrafficManagerRSG+$i,$colTrafficManagerRSG).text 
                                            }
                            }

                        }
       "StorageAccounts" {

                            $rowStorageAccountName,$colStorageAccountName = 1,1
                            $rowSAResourceGroupName,$colSAResourceGroupName = 1,2 
                            $rowSASkuName,$colSASkuName = 1,3
                            $rowSALocation,$colSALocation = 1,4
                            $rowSAKind,$colSAKind = 1,5 
                            $rowSAAccessTier,$colSAAccessTier = 1,6
                            $rowSAEncryptionService,$colSAEncryptionService = 1,7
                            $rowSAEnvironment,$colSAEnvironment = 1,8

                            $SAHash = @(@{})

                            for ($i=1; $i -le $rowMax-1; $i++){ 
                                $SAHash += @{
                                                name = ($sheet.Cells.Item($rowStorageAccountName+$i,$colStorageAccountName).text).tolower(); `
                                                ResourceGroupName = $sheet.Cells.Item($rowSAResourceGroupName+$i,$colSAResourceGroupName).text; `
                                                SkuName = $sheet.Cells.Item($rowSASkuName+$i,$colSASkuName).text; `
                                                Location = $sheet.Cells.Item($rowSALocation+$i,$colSALocation).text; `
                                                Kind = $sheet.Cells.Item($rowSAKind+$i,$colSAKind).text; ` 
                                                AccessTier = $sheet.Cells.Item($rowSAAccessTier+$i,$colSAAccessTier).text; `
                                                EnableEncryptionService = $sheet.Cells.Item($rowSAEncryptionService+$i,$colSAEncryptionService).text; ` 
                                                Environment = $sheet.Cells.Item($rowSAEnvironment+$i,$colSAEnvironment).text

                                            }
                            }

                        }

    }

$objExcel.quit() 

}


Login-AzureRmAccount -Credential $Credential -SubscriptionId $SubHash.SubID -TenantId $SubHash.TenID

# Create Resource Groups
foreach($RSG in $RSGHash){
        if($RSG.RSG -ne $null){
            
            if(Get-AzureRmResourceGroup -Name $RSG.RSG -ErrorAction SilentlyContinue){
                Write-Host "Resource Group: $($RSG.RSG) in $($RSG.RSGLocation) already exists" -ForegroundColor Yellow
            }
            else{
                
                try{
                    Write-Host "Creating Resource Group: $($RSG.RSG) in $($RSG.RSGLocation)" -ForegroundColor Green
                    $status = New-AzureRmResourceGroup -Name $RSG.RSG -Location $RSG.RSGLocation -Tag @{ BuildBy=$SubHash.BuildBy;BuildDate=$SubHash.BuildDate;Ticket=$SubHash.Ticket;Environment=$RSG.RSGENV }
                    if($status.ProvisioningState -eq 'Succeeded'){
                        Write-Host "Success: Resource Group: $($RSG.RSG) in $($RSG.RSGLocation)" -ForegroundColor Green
                    }
                    else{
                        Write-Host "Warning: Resource Group: $($RSG.RSG) in $($RSG.RSGLocation) not in a Succeeded state, please validate" -ForegroundColor Yellow
                        break
                    }
                }
                catch{
                    Write-Host "Error: Resource Group: $($RSG.RSG) in $($RSG.RSGLocation)" -ForegroundColor Red
                    break
                }

            }
        }   
}


# Create OMS Workspace
foreach($OMS in $OMSHash){
    if($OMS.OMSworkspaceName -ne $null){
        $ResourceGroupName = $OMS.OMSRSG
        $template = $OMS.Template
        $SAS = $OMS.SAS
        $OMS.Remove('OMSRSG')
        $OMS.Remove('Template')
        $OMS.Remove('SAS')
        try{
            Write-Host "Creating OMS Workspace: $($OMS.OMSworkspaceName) in $ResourceGroupName" -ForegroundColor Green
            $status = New-AzureRmResourceGroupDeployment -Name $SubHash.DeploymentName -ResourceGroupName $ResourceGroupName `
                                           -Mode Incremental `
                                           -TemplateParameterObject $OMS `
                                           -TemplateFile ("$template" + "$SAS") `
                                           -Force
            if($status.ProvisioningState -eq 'Succeeded'){
                Write-Host "Success: Creating OMS Workspace: $($OMS.OMSworkspaceName) in $ResourceGroupName" -ForegroundColor Green
            }
            else{
                  Write-Host "Warning: Creating OMS Workspace: $($OMS.OMSworkspaceName) in $ResourceGroupName is not in a Succeeded state, please validate" -ForegroundColor Yellow
                  break
            }
        }
        catch{
            Write-Host "Error: Creating OMS Workspace: $($OMS.OMSworkspaceName) in $ResourceGroupName" -ForegroundColor Red
            break

        }

    }

}

# Create Web Apps
foreach($WebApp in $WebAppHash){
    if($WebApp.webAppNames -ne $null){
        $ResourceGroupName = $WebApp.WebAppRSG
        $template = $WebApp.Template
        $SAS = $WebApp.SAS

        $WebApp.Remove('WebAppRSG')
        $WebApp.Remove('Template')
        $WebApp.Remove('SAS')
        try{
            Write-Host "Creating Web APP: $($WebApp.webAppNames) in $ResourceGroupName" -ForegroundColor Green
            $status = New-AzureRmResourceGroupDeployment -Name $SubHash.DeploymentName -ResourceGroupName $ResourceGroupName `
                                           -Mode Incremental `
                                           -TemplateParameterObject $WebApp `
                                           -TemplateFile ("$template" + "$SAS") `
                                           -Force
            if($status.ProvisioningState -eq 'Succeeded'){
                Write-Host "Success: Creating Web APP: $($WebApp.webAppNames) in $ResourceGroupName" -ForegroundColor Green
            }
            else{
                  Write-Host "Warning: Creating Web APP: $($WebApp.webAppNames) in $ResourceGroupName is not in a Succeeded state, please validate" -ForegroundColor Yellow
                  break
            }
        }
        catch{
            Write-Host "Error: Creating Web APP: $($WebApp.webAppNames) in $ResourceGroupName" -ForegroundColor Red
            break

        }

    }

}

# Create Traffic Manager Profile
foreach($TM in $TMHash){

    if($TM.name -ne $null){
        
        if(Get-AzureRmTrafficManagerProfile -Name $TM.name -ResourceGroupName $TM.ResourceGroup -ErrorAction SilentlyContinue){
            Write-Host "Traffic Manager Profile: $($TM.name) in $($TM.ResourceGroup) already exists" -ForegroundColor Yellow
        }
        else{
            try{
                Write-Host "Creating Traffic Manager: $($TM.name) in $($TM.ResourceGroup)" -ForegroundColor Green
                $status = New-AzureRmTrafficManagerProfile -Name $TM.name -ResourceGroupName $TM.ResourceGroup -TrafficRoutingMethod $TM.trafficRoutingMethod -RelativeDnsName $TM.relativeName `
                                                            -Ttl $TM.ttl -MonitorProtocol $TM.MonitorProtocol -MonitorPort $TM.MonitorPort -MonitorPath $TM.MonitorPath -Tag @{ BuildBy=$SubHash.BuildBy;BuildDate=$SubHash.BuildDate;Ticket=$SubHash.Ticket }
            
                if($status.ProfileStatus -eq 'Enabled'){
                    Write-Host "Success: Creating Traffic Manager: $($TM.name) in $($TM.ResourceGroup)" -ForegroundColor Green
                }
                else{
                    Write-Host "Warning: Creating Traffic Manager: $($TM.name) in $($TM.ResourceGroup) is not in a Succeeded state, please validate" -ForegroundColor Yellow
                    break
                }
            }
            catch{
                    Write-Host "Error: Creating Traffic Manager: $($TM.name) in $($TM.ResourceGroup)" -ForegroundColor Red
                    break

            }

        }
    }
}

# Create Storage Account
foreach($SA in $SAHash){

    if($SA.name -ne $null){
        
        if(Get-AzureRmStorageAccount -Name $SA.name -ResourceGroupName $SA.ResourceGroupName -ErrorAction SilentlyContinue){
            Write-Host "Storage Account: $($SA.name) in $($SA.ResourceGroupName) already exists" -ForegroundColor Yellow
        }
        else{
            if(Get-AzureRmStorageAccountNameAvailability -Name $SA.name){
                try{
                    Write-Host "Creating Storage Account: $($SA.name) in $($SA.ResourceGroupName)" -ForegroundColor Green

                    $status = New-AzureRmStorageAccount -ResourceGroupName $SA.ResourceGroupName -Name $SA.name -SkuName $SA.SkuName -Location $SA.Location `
                                                        -Kind $SA.kind -AccessTier $SA.AccessTier -EnableEncryptionService $SA.EnableEncryptionService `
                                                        -Tag @{ BuildBy=$SubHash.BuildBy;BuildDate=$SubHash.BuildDate;Ticket=$SubHash.Ticket;Environment=$SA.Environment }
            
                    if($status.ProvisioningState -eq 'Succeeded'){
                        Write-Host "Success: Creating Storage Account: $($SA.name) in $($SA.ResourceGroupName)" -ForegroundColor Green
                    }
                    else{
                        Write-Host "Warning: Creating Storage Account: $($SA.name) in $($SA.ResourceGroupName) is not in a Succeeded state, please validate" -ForegroundColor Yellow
                        break
                    }
                }
                catch{
                        Write-Host "Error: Creating Storage Account: $($SA.name) in $($SA.ResourceGroupName)" -ForegroundColor Red
                        break

                }
            }
            else{
                Write-Host "Error: Creating Storage Account: $($TM.name) is already in use" -ForegroundColor Red
                break
            }
        }
    }
}