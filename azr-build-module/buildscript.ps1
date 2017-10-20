
#function Build-RAXAzureEnvironment

#{

        param(

#            [Parameter(Mandatory=$true)]

            [string]$BuildSheet = '',
            [string]$DomainPassword = '',
			[bool]$License = $false

         )


        $OMSWorkspace = $null
        $OMSRSG = $null
        $VNETDNS = @(@{})
        #Build Sheet location
        $BSFile = $BuildSheet
        $SheetArray = @('Subscriptions','Environments','ResourceGroups','VirtualNetworks','OMSWorkspaces','RecoveryServicesVault','ActiveDirectory','VirtualGateway','VPNConnections','NSGs','VirtualMachines','WebApps','TrafficManager','StorageAccounts')
   
   try{
       $ErrorActionPreference = 'Stop'
       $Error.Clear()

        $objExcel = New-Object -ComObject Excel.Application
        $workbook = $objExcel.Workbooks.Open($BSFile)
         
        foreach($sheetName in $SheetArray){
            
            
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
                                        if($sheet.Cells.Item($rowSubName+$i,$colSubName).text -ne ""){  
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

                                }
                "Environments" {

                                    $rowEnvironment,$colEnvironment = 1,1
                                    $rowLocation,$colLocation = 1,2 
                                    $rowSubName,$colSubName = 1,3 

                                    $ENVHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){
                                        if($sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text -ne ""){   
                                            $ENVHash += @{
                                                        ENV = $sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text; `
                                                        ENVLocation = $sheet.Cells.Item($rowLocation+$i,$colLocation).text; `
                                                        ENVSubName = $sheet.Cells.Item($rowSubName+$i,$colSubName).text 
                                                    }
                                        }
                                    }

                                }
                "ResourceGroups" {
                            
                                    $rowResourceGroupName,$colResourceGroupName = 1,1
                                    $rowLocation,$colLocation = 1,2 
                                    $rowSubName,$colSubName = 1,3 
                                    $rowEnvironment,$colEnvironment = 1,4
                             
                                    $RSGHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){
                                        if($sheet.Cells.Item($rowResourceGroupName+$i,$colResourceGroupName).text -ne ""){  
                                        $RSGHash += @{
                                                        RSG = $sheet.Cells.Item($rowResourceGroupName+$i,$colResourceGroupName).text; `
                                                        RSGLocation = $sheet.Cells.Item($rowLocation+$i,$colLocation).text; `
                                                        RSGSubName = $sheet.Cells.Item($rowSubName+$i,$colSubName).text; ` 
                                                        RSGENV = $sheet.Cells.Item($rowEnvironment+$i,$colEnvironment).text
                                                    }
                                        }
                                    }

                                }
                "VirtualNetworks" {

                                    $rowvnetSize,$colvnetSize = 1,1
                                    $rowvnetRSG,$colvnetRSG = 1,2 
                                    $rowvnetName,$colvnetName = 1,3
                                    $rowName,$colName = 1,4
                                    $rowvnetPrimaryDNS,$colvnetPrimaryDNS = 1,5
                                    $rowvnetSecondaryDNS,$colvnetSecondaryDNS = 1,6
                                    $rowvnetCIDR,$colvnetCIDR = 1,7 
                                    $rowvnetEnvironmentA,$colvnetEnvironmentA = 1,8 
                                    $rowvnetsubnetDMZCIDRA,$colvnetsubnetDMZCIDRA = 1,9                                    
                                    $rowvnetsubnetAPPCIDRA,$colvnetsubnetAPPCIDRA = 1,10
                                    $rowvnetsubnetINSCIDRA,$colvnetsubnetINSCIDRA = 1,11 
                                    $rowvnetsubnetADCIDRA,$colvnetsubnetADCIDRA = 1,12
                                    $rowvnetsubnetBASCIDR,$colvnetsubnetBASCIDR = 1,13 
                                    $rowvnetsubnetAGWCIDRA,$colvnetsubnetAGWCIDRA = 1,14 
                                    $rowvnetsubnetGWCIDR,$colvnetsubnetGWCIDR = 1,15                                     
                                    $rowvnetEnvironmentB,$colvnetEnvironmentB = 1,16
                                    $rowvnetsubnetDMZCIDRB,$colvnetsubnetDMZCIDRB = 1,17 
                                    $rowvnetsubnetAPPCIDRB,$colvnetsubnetAPPCIDRB = 1,18
                                    $rowvnetsubnetINSCIDRB,$colvnetsubnetINSCIDRB = 1,19 
                                    $rowvnetsubnetADCIDRB,$colvnetsubnetADCIDRBt = 1,20 
                                    $rowvnetsubnetAGWCIDRB,$colvnetsubnetAGWCIDRB = 1,21                                     
                                    $rowvnetenvironmentC,$colvnetenvironmentC = 1,22
                                    $rowvnetsubnetDMZCIDRC,$colvnetsubnetDMZCIDRC = 1,23
                                    $rowvnetsubnetAPPCIDRC,$colvnetsubnetAPPCIDRC = 1,24
                                    $rowvnetsubnetINSCIDRC,$colvnetsubnetINSCIDR = 1,25 
                                    $rowvnetsubnetADCIDRC,$colvnetsubnetADCIDRC = 1,26 
                                    $rowvnetsubnetAGWCIDRC,$colvnetsubnetAGWCIDRC = 1,27 
                                    $rowTemplate,$colTemplate = 1,28 
                                    $rowSAS,$colSAS = 1,29

                                    $VNETHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){
                                        if($sheet.Cells.Item($rowvnetName+$i,$colvnetName).text -ne ""){  
                                            $VNETHash += @{
                                                        VNETSize = $sheet.Cells.Item($rowvnetSize+$i,$colvnetSize).text; `
                                                        VNETRSG = $sheet.Cells.Item($rowvnetRSG+$i,$colvnetRSG).text; `
                                                        vnetName = $sheet.Cells.Item($rowvnetName+$i,$colvnetName).text; `
                                                        Name = $sheet.Cells.Item($rowName+$i,$colName).text; `
                                                        PrimaryDNS = $sheet.Cells.Item($rowvnetPrimaryDNS+$i,$colvnetPrimaryDNS).text; `
                                                        SecondaryDNS = $sheet.Cells.Item($rowvnetSecondaryDNS+$i,$colvnetSecondaryDNS).text; `
                                                        vnetCIDR = $sheet.Cells.Item($rowvnetCIDR+$i,$colvnetCIDR).text; `
                                                        environmentA = $sheet.Cells.Item($rowvnetEnvironmentA+$i,$colvnetEnvironmentA).text; ` 
                                                        subnetDMZCIDRA = $sheet.Cells.Item($rowvnetsubnetDMZCIDRA+$i,$colvnetsubnetDMZCIDRA).text; `                                                      
                                                        subnetAPPCIDRA = $sheet.Cells.Item($rowvnetsubnetAPPCIDRA+$i,$colvnetsubnetAPPCIDRA).text; `
                                                        subnetINSCIDRA = $sheet.Cells.Item($rowvnetsubnetINSCIDRA+$i,$colvnetsubnetINSCIDRA).text; `
                                                        subnetADCIDRA = $sheet.Cells.Item($rowvnetsubnetADCIDRA+$i,$colvnetsubnetADCIDRA).text; `
                                                        subnetBASCIDR = $sheet.Cells.Item($rowvnetsubnetBASCIDR+$i,$colvnetsubnetBASCIDR).text; `
                                                        subnetAGWCIDRA = $sheet.Cells.Item($rowvnetsubnetAGWCIDRA+$i,$colvnetsubnetAGWCIDRA).text; ` 
                                                        subnetGWCIDR = $sheet.Cells.Item($rowvnetsubnetGWCIDR+$i,$colvnetsubnetGWCIDR).text; ` 
                                                        environmentB = $sheet.Cells.Item($rowvnetEnvironmentB+$i,$colvnetEnvironmentB).text; `
                                                        subnetDMZCIDRB = $sheet.Cells.Item($rowvnetsubnetDMZCIDRB+$i,$colvnetsubnetDMZCIDRB).text; `
                                                        subnetAPPCIDRB = $sheet.Cells.Item($rowvnetsubnetAPPCIDRB+$i,$colvnetsubnetAPPCIDRB).text; `
                                                        subnetINSCIDRB = $sheet.Cells.Item($rowvnetsubnetINSCIDRB+$i,$colvnetsubnetINSCIDRB).text; `
                                                        subnetADCIDRB = $sheet.Cells.Item($rowvnetsubnetADCIDRB+$i,$colvnetsubnetADCIDRBt).text; ` 
                                                        subnetAGWCIDRB = $sheet.Cells.Item($rowvnetsubnetAGWCIDRB+$i,$colvnetsubnetAGWCIDRB).text; `                                                       
                                                        environmentC = $sheet.Cells.Item($rowvnetenvironmentC+$i,$colvnetenvironmentC).text; `
                                                        subnetDMZCIDRC = $sheet.Cells.Item($rowvnetsubnetDMZCIDRC+$i,$colvnetsubnetDMZCIDRC).text; `
                                                        subnetAPPCIDRC = $sheet.Cells.Item($rowvnetsubnetAPPCIDRC+$i,$colvnetsubnetAPPCIDRC).text; `
                                                        subnetINSCIDRC = $sheet.Cells.Item($rowvnetsubnetINSCIDRC+$i,$colvnetsubnetINSCIDR).text; `
                                                        subnetADCIDRC = $sheet.Cells.Item($rowvnetsubnetADCIDRC+$i,$colvnetsubnetADCIDRC).text; ` 
                                                        subnetAGWCIDRC = $sheet.Cells.Item($rowvnetsubnetAGWCIDRC+$i,$colvnetsubnetAGWCIDRC).text; `
                                                        Template = $sheet.Cells.Item($rowTemplate+$i,$colTemplate).text; ` 
                                                        SAS = $sheet.Cells.Item($rowSAS+$i,$colSAS).text
                                                    }
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
                                        if($sheet.Cells.Item($rowOMSWorkspaceName+$i,$colOMSWorkspaceName).text -ne ""){ 
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

                                }
                "RecoveryServicesVault" {

                                    $rowRSVName,$colRSVName = 1,1
                                    $rowRSVPolicyName,$colRSVPolicyName = 1,2 
                                    $rowRSVscheduleRunTimes,$colRSVscheduleRunTimes = 1,3
                                    $rowRSVdailyRetentionDurationCount,$colRSVdailyRetentionDurationCount = 1,4
                                    $rowRSVdaysOfTheWeek,$colRSVdaysOfTheWeek = 1,5 
                                    $rowRSVweeklyRetentionDurationCount,$colRSVweeklyRetentionDurationCount = 1,6 
                                    $rowRSVRSG,$colRSVRSG = 1,7 
                                    $rowTemplate,$colTemplate = 1,8 
                                    $rowSAS,$colSAS = 1,9 

                                    $RSVHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){
                                        if($sheet.Cells.Item($rowRSVName+$i,$colRSVName).text -ne ""){ 
                                            $RSVHash += @{
                                                        vaultName = $sheet.Cells.Item($rowRSVName+$i,$colRSVName).text; `
                                                        policyName = $sheet.Cells.Item($rowRSVPolicyName+$i,$colRSVPolicyName).text; `
                                                        scheduleRunTimes = @($sheet.Cells.Item($rowRSVscheduleRunTimes+$i,$colRSVscheduleRunTimes).text); `
                                                        dailyRetentionDurationCount = [Convert]::ToInt32($sheet.Cells.Item($rowRSVdailyRetentionDurationCount+$i,$colRSVdailyRetentionDurationCount).text); `
                                                        daysOfTheWeek = @($sheet.Cells.Item($rowRSVdaysOfTheWeek+$i,$colRSVdaysOfTheWeek).text); `
                                                        weeklyRetentionDurationCount = [Convert]::ToInt32($sheet.Cells.Item($rowRSVweeklyRetentionDurationCount+$i,$colRSVweeklyRetentionDurationCount).text); `
                                                        ResourceGroupName = $sheet.Cells.Item($rowRSVRSG+$i,$colRSVRSG).text; `
                                                        Template = $sheet.Cells.Item($rowTemplate+$i,$colTemplate).text; `
                                                        SAS = $sheet.Cells.Item($rowSAS+$i,$colSAS).text  
                                                    }
                                        }
                                    }

                                }
                "ActiveDirectory"       {
                                    $rowADResourceGroupName,$colADResourceGroupName = 1,1
                                    $rowADType,$colADType = 1,2 
                                    $rowADoperatingSystem,$colADoperatingSystem = 1,3
                                    $rowADvmNamePrefix,$colADvmNamePrefix = 1,4
                                    $rowADvmSize,$colADvmSize = 1,5 
                                    $rowADstorageAccountOption,$colADstorageAccountOption = 1,6
                                    $rowADdataDiskSize,$colADdataDiskSize = 1,7
                                    $rowADblobEncryptionEnabled,$colADblobEncryptionEnabled = 1,8 
                                    $rowADavailabilitySetName,$colADavailabilitySetName = 1,9
                                    $rowADvnetRG,$colADvnetRG = 1,10
                                    $rowADvnetName,$colADvnetName = 1,11 
                                    $rowADsubnetName,$colADsubnetName = 1,12 
                                    $rowADprimaryDCIp,$colADprimaryDCIp = 1,13
                                    $rowADsecondaryDCIp,$colADsecondaryDCIp = 1,14 
                                    $rowADdomainName,$colADdomainName = 1,15
                                    $rowADnetbiosName,$colADnetbiosName = 1,16
                                    $rowADtimeZone,$colADtimeZone = 1,17
                                    $rowADapplyOSPatches,$colADapplyOSPatches = 1,18
                                    $rowADantiMalware,$colADantiMalware = 1,19
                                    $rowADlogAnalytics,$colADlogAnalytics = 1,20
                                    $rowADenvironment,$colADenvironment = 1,21
                                    $rowADRaxAutomationExclude,$colADRaxAutomationExclude = 1,22
                                    $rowADassetLocation,$colADassetLocation = 1,23
                                    $rowADTemplate,$colADTemplate = 1,24
                                    $rowADsasToken,$colADsasToken = 1,25

                                    $ADHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){ 
                                        
                                        if($sheet.Cells.Item($rowADvmNamePrefix+$i,$colADvmNamePrefix).text -ne ""){
                                            $ADHash += @{
                                                                ResourceGroupName = $sheet.Cells.Item($rowADResourceGroupName+$i,$colADResourceGroupName).text; `
                                                                Type = $sheet.Cells.Item($rowADType+$i,$colADType).text; `
                                                                operatingSystem = $sheet.Cells.Item($rowADoperatingSystem+$i,$colADoperatingSystem).text; `
                                                                vmNamePrefix = $sheet.Cells.Item($rowADvmNamePrefix+$i,$colADvmNamePrefix).text; `
                                                                vmSize = $sheet.Cells.Item($rowADvmSize+$i,$colADvmSize).text; `
                                                                storageAccountOption = $sheet.Cells.Item($rowADstorageAccountOption+$i,$colADstorageAccountOption).text; `
                                                                dataDiskSize = [convert]::ToInt32($sheet.Cells.Item($rowADdataDiskSize+$i,$colADdataDiskSize).text); `
                                                                blobEncryptionEnabled = [convert]::ToBoolean($sheet.Cells.Item($rowADblobEncryptionEnabled+$i,$colADblobEncryptionEnabled).text); `
                                                                availabilitySetName = $sheet.Cells.Item($rowADavailabilitySetName+$i,$colADavailabilitySetName).text; `
                                                                vnetRG = $sheet.Cells.Item($rowADvnetRG+$i,$colADvnetRG).text; `
                                                                vnetName = $sheet.Cells.Item($rowADvnetName+$i,$colADvnetName).text; `
                                                                subnetName = $sheet.Cells.Item($rowADsubnetName+$i,$colADsubnetName).text; ` 
                                                                primaryDCIp = $sheet.Cells.Item($rowADprimaryDCIp+$i,$colADprimaryDCIp).text; `
                                                                secondaryDCIp = $sheet.Cells.Item($rowADsecondaryDCIp+$i,$colADsecondaryDCIp).text; `
                                                                domainName = $sheet.Cells.Item($rowADdomainName+$i,$colADdomainName).text; `
                                                                netbiosName = $sheet.Cells.Item($rowADnetbiosName+$i,$colADnetbiosName).text; ` 
                                                                timeZone = $sheet.Cells.Item($rowADtimeZone+$i,$colADtimeZone).text; `
                                                                applyOSPatches = $sheet.Cells.Item($rowADapplyOSPatches+$i,$colADapplyOSPatches).text; `
                                                                antiMalware = $sheet.Cells.Item($rowADantiMalware+$i,$colADantiMalware).text; `
                                                                logAnalytics = $sheet.Cells.Item($rowADlogAnalytics+$i,$colADlogAnalytics).text; `
                                                                environment = $sheet.Cells.Item($rowADenvironment+$i,$colADenvironment).text; `
                                                                RaxAutomationExclude = $sheet.Cells.Item($rowADRaxAutomationExclude+$i,$colADRaxAutomationExclude).text; `
                                                                assetLocation = $sheet.Cells.Item($rowADassetLocation+$i,$colADassetLocation).text; `
                                                                Template = $sheet.Cells.Item($rowADTemplate+$i,$colADTemplate).text; `
                                                                sasToken = $sheet.Cells.Item($rowADsasToken+$i,$colADsasToken).text
                                                        }
                                                    
                                        }
                                    }


                                }

               "VirtualGateway" {

                                    $rowVPNName,$colVPNName= 1,1
                                    $rowResourceGroupName,$colResourceGroupName = 1,2 
                                    $rowLocation,$colLocation = 1,3
                                    $rowVNET,$colVNET = 1,4
                                    $rowVNETRSG,$colVNETRSG = 1,5 
                                    $rowGatewaySubnet,$colGatewaySubnet= 1,6
                                    $rowGatewayType,$colGatewayType = 1,7
                                    $rowVPNType,$colVPNType = 1,8 
                                    $rowGatewaySku,$colGatewaySku = 1,9 

                                    $VPNHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){ 
                                        if($sheet.Cells.Item($rowVPNName+$i,$colVPNName).text -ne ""){
                                            $VPNHash += @{
                                                        VPNName = $sheet.Cells.Item($rowVPNName+$i,$colVPNName).text; `
                                                        ResourceGroupName = $sheet.Cells.Item($rowResourceGroupName+$i,$colResourceGroupName).text; `
                                                        Location = $sheet.Cells.Item($rowLocation+$i,$colLocation).text; `
                                                        VNET = $sheet.Cells.Item($rowVNET+$i,$colVNET).text; `
                                                        VNETRSG = $sheet.Cells.Item($rowVNETRSG+$i,$colVNETRSG).text; ` 
                                                        GatewaySubnet = $sheet.Cells.Item($rowGatewaySubnet+$i,$colGatewaySubnet).text; `
                                                        GatewayType = $sheet.Cells.Item($rowGatewayType+$i,$colGatewayType).text; ` 
                                                        VPNType = $sheet.Cells.Item($rowVPNType+$i,$colVPNType).text; `
                                                        GatewaySku = $sheet.Cells.Item($rowGatewaySku+$i,$colGatewaySku).text 
                                                    }
                                        }
                                    }

                                }
                                
               "VPNConnections" {

                                    $rowName,$colName = 1,1
                                    $rowResourceGroupName,$colResourceGroupName = 1,2 
                                    $rowVirtualGateway,$colVirtualGateway = 1,3
                                    $rowConnectionType,$colConnectionType = 1,4
                                    $rowLocation,$colLocation = 1,5 
                                    $rowLNGName,$colLNGName= 1,6
                                    $rowLNGRSG,$colrowLNGRSG = 1,7
                                    $rowLNGGatewayIP,$colLNGGatewayIP = 1,8
                                    $rowLNGAddressPrefix,$colLNGAddressPrefix = 1,9 
                                    $rowRoutingWeight,$colRoutingWeight = 1,10 
                                    $rowSharedKey,$colSharedKey = 1,11 

                                    $VPNCONHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){ 
                                        if($sheet.Cells.Item($rowName+$i,$colName).text -ne ""){
                                            $VPNCONHash += @{
                                                        Name = $sheet.Cells.Item($rowName+$i,$colName).text; `
                                                        ResourceGroupName = $sheet.Cells.Item($rowResourceGroupName+$i,$colResourceGroupName).text; `
                                                        VirtualGateway = $sheet.Cells.Item($rowVirtualGateway+$i,$colVirtualGateway).text; `
                                                        ConnectionType = $sheet.Cells.Item($rowConnectionType+$i,$colConnectionType).text; `
                                                        Location = $sheet.Cells.Item($rowLocation+$i,$colLocation).text; ` 
                                                        LNGName = $sheet.Cells.Item($rowLNGName+$i,$colLNGName).text; `
                                                        LNGRSG = $sheet.Cells.Item($rowLNGRSG+$i,$colrowLNGRSG).text; `
                                                        LNGGatewayIP = $sheet.Cells.Item($rowLNGGatewayIP+$i,$colLNGGatewayIP).text; ` 
                                                        LNGAddressPrefix = @($sheet.Cells.Item($rowLNGAddressPrefix+$i,$colLNGAddressPrefix).text); `
                                                        RoutingWeight = $sheet.Cells.Item($rowRoutingWeight+$i,$colRoutingWeight).text; `
                                                        SharedKey = $sheet.Cells.Item($rowSharedKey+$i,$colSharedKey).text  
                                                    }
                                        }
                                    }

                                }

               "NSGs" {

                                    $rowNSGName,$colNSGName = 1,1
                                    $rowNSGRule,$colNSGRule = 1,2 
                                    $rowNSGDescription,$colNSGDescription = 1,3
                                    $rowNSGAccess,$colNSGAccess = 1,4
                                    $rowNSGProtocol,$colNSGProtocol = 1,5 
                                    $rowNSGDirection,$colNSGDirection= 1,6
                                    $rowNSGPriority,$colNSGPriority = 1,7
                                    $rowNSGSourceAddressPrefix,$colNSGSourceAddressPrefix = 1,8
                                    $rowNSGSourcePortRange,$colNSGSourcePortRange = 1,9 
                                    $rowNSGDestinationAddressPrefix,$colNSGDestinationAddressPrefix = 1,10 
                                    $rowNSGDestinationPortRange,$colNSGDestinationPortRange = 1,11 
                                    $rowNSGResourceGroupName,$colNSGResourceGroupName = 1,12

                                    $NSGHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){ 
                                        if($sheet.Cells.Item($rowNSGName+$i,$colNSGName).text -ne ""){
                                            $NSGHash += @{
                                                        Name = $sheet.Cells.Item($rowNSGName+$i,$colNSGName).text; `
                                                        Rule = $sheet.Cells.Item($rowNSGRule+$i,$colNSGRule).text; `
                                                        Description = $sheet.Cells.Item($rowNSGDescription+$i,$colNSGDescription).text; `
                                                        Access = $sheet.Cells.Item($rowNSGAccess+$i,$colNSGAccess).text; `
                                                        Protocol = $sheet.Cells.Item($rowNSGProtocol+$i,$colNSGProtocol).text; ` 
                                                        Direction = $sheet.Cells.Item($rowNSGDirection+$i,$colNSGDirection).text; `
                                                        Priority = $sheet.Cells.Item($rowNSGPriority+$i,$colNSGPriority).text; `
                                                        SourceAddressPrefix = $sheet.Cells.Item($rowNSGSourceAddressPrefix+$i,$colNSGSourceAddressPrefix).text; ` 
                                                        SourcePortRange = @($sheet.Cells.Item($rowNSGSourcePortRange+$i,$colNSGSourcePortRange).text); `
                                                        DestinationAddressPrefix = $sheet.Cells.Item($rowNSGDestinationAddressPrefix+$i,$colNSGDestinationAddressPrefix).text; `
                                                        DestinationPortRange = $sheet.Cells.Item($rowNSGDestinationPortRange+$i,$colNSGDestinationPortRange).text; `
                                                        ResourceGroupName = $sheet.Cells.Item($rowNSGResourceGroupName+$i,$colNSGResourceGroupName).text  
                                                    }
                                        }
                                    }

                                }

        "VirtualMachines"       {
                                    $rowVMOS,$colVMOS = 1,1
                                    $rowVMRSG,$colVMRSG = 1,2
                                    $rowVMName,$colVMName = 1,3 
                                    $rowVMSize,$colVMSize = 1,4
                                    $rowVMAVSETOption,$colVMAVSETOption = 1,5
                                    $rowVMAVSETName,$colVMAVSETName = 1,6 
                                    $rowVMSAOption,$colVMSAOption = 1,7
                                    $rowVMexistingSA,$colVMexistingSA = 1,8
                                    $rowVMSAType,$colVMSAType = 1,9 
                                    $rowVMBlobEncryption,$colVMBlobEncryption = 1,10
                                    $rowVMnumDataDisks,$colVMnumDataDisks = 1,11
                                    $rowVMdataDiskSize,$colVMdataDiskSize = 1,12 
                                    $rowVMdiskCaching,$colVMdiskCaching = 1,13 
                                    $rowVMVNETRG,$colVMVNETRG = 1,14
                                    $rowVMVNETName,$colVMVNETName = 1,15 
                                    $rowVMsubnetName,$colVMsubnetName = 1,16
                                    $rowVMLBOption,$colVMLBOption = 1,17
                                    $rowVMLBPort,$colVMLBPort = 1,18
                                    $rowVMLBProtocol,$colVMLBProtocol = 1,19
                                    $rowVMLBProbePath,$colVMLBProbePath = 1,20
                                    $rowVMDNSLabel,$colVMDNSLabel = 1,21
                                    $rowVMjoinADDomain,$colVMjoinADDomain = 1,22
                                    $rowVMDomainToJoin,$colVMDomainToJoin = 1,23
                                    $rowVMDomainOU,$colVMDomainOU = 1,24
                                    $rowVMDomainAdminUsername,$colVMDomainAdminUsername = 1,25
                                    $rowVMdeployWebServer,$colVMdeployWebServer = 1,26
                                    $rowVMtimeZone,$colVMtimeZone = 1,27
                                    $rowVMapplyOSPatches,$colVMapplyOSPatches = 1,28
                                    $rowVMantiMalware,$colVMantiMalware = 1,29
                                    $rowVMlogAnalytics,$colVMlogAnalytics = 1,30
                                    $rowVMenvironment,$colVMenvironment = 1,31
                                    $rowVMRaxAutomationExclude,$colVMRaxAutomationExclude = 1,32
                                    $rowVMassetLocation,$colVMassetLocation = 1,33
                                    $rowVMTemplate,$colVMTemplate = 1,34
                                    $rowVMsasToken,$colVMsasToken = 1,35
									if ($License)
									{
										$rowVMHub,$colVMHub = 1,36
									}
                                   

                                    $VMHash = @(@{})

                                    for ($i=1; $i -le $rowMax-1; $i++){ 
                                        
                                        if($sheet.Cells.Item($rowVMName+$i,$colVMName).text -ne ""){
                                            $VMHash += @{
                                                                operatingSystem = $sheet.Cells.Item($rowVMOS+$i,$colVMOS).text; `
                                                                VMRSG = $sheet.Cells.Item($rowVMRSG+$i,$colVMRSG).text; `
                                                                vmName = $sheet.Cells.Item($rowVMName+$i,$colVMName).text; `
                                                                vmSize = $sheet.Cells.Item($rowVMSize+$i,$colVMSize).text; `
                                                                availabilitySetOption = $sheet.Cells.Item($rowVMAVSETOption+$i,$colVMAVSETOption).text; `
                                                                availabilitySetName = $sheet.Cells.Item($rowVMAVSETName+$i,$colVMAVSETName).text; `
                                                                storageAccountOption = $sheet.Cells.Item($rowVMSAOption+$i,$colVMSAOption).text; `
                                                                existingStorageAccount = $sheet.Cells.Item($rowVMexistingSA+$i,$colVMexistingSA).text; `
                                                                storageAccountType = $sheet.Cells.Item($rowVMSAType+$i,$colVMSAType).text; `
                                                                blobEncryptionEnabled = [convert]::ToBoolean($sheet.Cells.Item($rowVMBlobEncryption+$i,$colVMBlobEncryption).text); `
                                                                numDataDisks = [convert]::ToInt32($sheet.Cells.Item($rowVMnumDataDisks+$i,$colVMnumDataDisks).text); ` 
                                                                dataDiskSize = [convert]::ToInt32($sheet.Cells.Item($rowVMdataDiskSize+$i,$colVMdataDiskSize).text); `
                                                                diskCaching = $sheet.Cells.Item($rowVMdiskCaching+$i,$colVMdiskCaching).text; `
                                                                vnetRG = $sheet.Cells.Item($rowVMVNETRG+$i,$colVMVNETRG).text; `
                                                                vnetName = $sheet.Cells.Item($rowVMVNETName+$i,$colVMVNETName).text; `
                                                                subnetName = $sheet.Cells.Item($rowVMsubnetName+$i,$colVMsubnetName).text; `
                                                                lbOption = $sheet.Cells.Item($rowVMLBOption+$i,$colVMLBOption).text; `
                                                                lbPort = [convert]::ToInt32($sheet.Cells.Item($rowVMLBPort+$i,$colVMLBPort).text); `
                                                                lbProtocol = $sheet.Cells.Item($rowVMLBProtocol+$i,$colVMLBProtocol).text; `
                                                                lbProbePath = $sheet.Cells.Item($rowVMLBProbePath+$i,$colVMLBProbePath).text; `
                                                                dnsLabel = $sheet.Cells.Item($rowVMDNSLabel+$i,$colVMDNSLabel).text; `
                                                                joinADDomain = $sheet.Cells.Item($rowVMjoinADDomain+$i,$colVMjoinADDomain).text; `
                                                                domainToJoin = $sheet.Cells.Item($rowVMDomainToJoin+$i,$colVMDomainToJoin).text; `
                                                                organizationalUnit = $sheet.Cells.Item($rowVMDomainOU+$i,$colVMDomainOU).text; `
                                                                domainAdminUsername = $sheet.Cells.Item($rowVMDomainAdminUsername+$i,$colVMDomainAdminUsername).text; `
                                                                deployWebServer = $sheet.Cells.Item($rowVMdeployWebServer+$i,$colVMdeployWebServer).text; `
                                                                timeZone = $sheet.Cells.Item($rowVMtimeZone+$i,$colVMtimeZone).text; `
                                                                applyOSPatches = $sheet.Cells.Item($rowVMapplyOSPatches+$i,$colVMapplyOSPatches).text; `
                                                                antiMalware = $sheet.Cells.Item($rowVMantiMalware+$i,$colVMantiMalware).text; `
                                                                logAnalytics = $sheet.Cells.Item($rowVMlogAnalytics+$i,$colVMlogAnalytics).text; `
                                                                environment = $sheet.Cells.Item($rowVMenvironment+$i,$colVMenvironment).text; `
                                                                RaxAutomationExclude = $sheet.Cells.Item($rowVMRaxAutomationExclude+$i,$colVMRaxAutomationExclude).text; `
                                                                assetLocation = $sheet.Cells.Item($rowVMassetLocation+$i,$colVMassetLocation).text; `
                                                                Template = $sheet.Cells.Item($rowVMTemplate+$i,$colVMTemplate).text; `
                                                                sasToken = $sheet.Cells.Item($rowVMsasToken+$i,$colVMsasToken).text; 
                                                        }
											if ($License)
											{
												$VMHash += @{HubLicense = $sheet.Cells.Item($rowVMHub+$i,$colVMHub).text;}
											}
                                                    
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
                                        
                                        if($sheet.Cells.Item($rowwebAppNames+$i,$colwebAppNames).text -ne ""){
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
                                        if($sheet.Cells.Item($rowTrafficManagerName+$i,$colTrafficManagerName).text -ne ""){
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
                                        if($sheet.Cells.Item($rowStorageAccountName+$i,$colStorageAccountName).text -ne ""){ 
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

            }
        
      
        }
    
    }
    catch{
        throw $_;
        Write-Host "Something is happening"
        break
    }
    Finally{
        $workbook.Save()
        $workbook.Close() | Out-Null
        $objExcel.quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
        [System.GC]::Collect() | Out-Null
        [System.GC]::WaitForPendingFinalizers() | Out-Null
    
    
    } 

        Login-AzureRmAccount 
            
        Select-AzureRmSubscription -SubscriptionId $SubHash.SubID -TenantId $SubHash.TenID

        function Invoke-Parallel {
    <#
    .SYNOPSIS
        Function to control parallel processing using runspaces

    .DESCRIPTION
        Function to control parallel processing using runspaces

            Note that each runspace will not have access to variables and commands loaded in your session or in other runspaces by default.  
            This behaviour can be changed with parameters.

    .PARAMETER ScriptFile
        File to run against all input objects.  Must include parameter to take in the input object, or use $args.  Optionally, include parameter to take in parameter.  Example: C:\script.ps1

    .PARAMETER ScriptBlock
        Scriptblock to run against all computers.

        You may use $Using:<Variable> language in PowerShell 3 and later.
        
            The parameter block is added for you, allowing behaviour similar to foreach-object:
                Refer to the input object as $_.
                Refer to the parameter parameter as $parameter

    .PARAMETER InputObject
        Run script against these specified objects.

    .PARAMETER Parameter
        This object is passed to every script block.  You can use it to pass information to the script block; for example, the path to a logging folder
        
            Reference this object as $parameter if using the scriptblock parameterset.

    .PARAMETER ImportVariables
        If specified, get user session variables and add them to the initial session state

    .PARAMETER ImportModules
        If specified, get loaded modules and pssnapins, add them to the initial session state

    .PARAMETER Throttle
        Maximum number of threads to run at a single time.

    .PARAMETER SleepTimer
        Milliseconds to sleep after checking for completed runspaces and in a few other spots.  I would not recommend dropping below 200 or increasing above 500

    .PARAMETER RunspaceTimeout
        Maximum time in seconds a single thread can run.  If execution of your code takes longer than this, it is disposed.  Default: 0 (seconds)

        WARNING:  Using this parameter requires that maxQueue be set to throttle (it will be by default) for accurate timing.  Details here:
        http://gallery.technet.microsoft.com/Run-Parallel-Parallel-377fd430

    .PARAMETER NoCloseOnTimeout
		Do not dispose of timed out tasks or attempt to close the runspace if threads have timed out. This will prevent the script from hanging in certain situations where threads become non-responsive, at the expense of leaking memory within the PowerShell host.

    .PARAMETER MaxQueue
        Maximum number of powershell instances to add to runspace pool.  If this is higher than $throttle, $timeout will be inaccurate
        
        If this is equal or less than throttle, there will be a performance impact

        The default value is $throttle times 3, if $runspaceTimeout is not specified
        The default value is $throttle, if $runspaceTimeout is specified

    .PARAMETER LogFile
        Path to a file where we can log results, including run time for each thread, whether it completes, completes with errors, or times out.

	.PARAMETER Quiet
		Disable progress bar.

    .EXAMPLE
        Each example uses Test-ForPacs.ps1 which includes the following code:
            param($computer)

            if(test-connection $computer -count 1 -quiet -BufferSize 16){
                $object = [pscustomobject] @{
                    Computer=$computer;
                    Available=1;
                    Kodak=$(
                        if((test-path "\\$computer\c$\users\public\desktop\Kodak Direct View Pacs.url") -or (test-path "\\$computer\c$\documents and settings\all users

        \desktop\Kodak Direct View Pacs.url") ){"1"}else{"0"}
                    )
                }
            }
            else{
                $object = [pscustomobject] @{
                    Computer=$computer;
                    Available=0;
                    Kodak="NA"
                }
            }

            $object

    .EXAMPLE
        Invoke-Parallel -scriptfile C:\public\Test-ForPacs.ps1 -inputobject $(get-content C:\pcs.txt) -runspaceTimeout 10 -throttle 10

            Pulls list of PCs from C:\pcs.txt,
            Runs Test-ForPacs against each
            If any query takes longer than 10 seconds, it is disposed
            Only run 10 threads at a time

    .EXAMPLE
        Invoke-Parallel -scriptfile C:\public\Test-ForPacs.ps1 -inputobject c-is-ts-91, c-is-ts-95

            Runs against c-is-ts-91, c-is-ts-95 (-computername)
            Runs Test-ForPacs against each

    .EXAMPLE
        $stuff = [pscustomobject] @{
            ContentFile = "windows\system32\drivers\etc\hosts"
            Logfile = "C:\temp\log.txt"
        }
    
        $computers | Invoke-Parallel -parameter $stuff {
            $contentFile = join-path "\\$_\c$" $parameter.contentfile
            Get-Content $contentFile |
                set-content $parameter.logfile
        }

        This example uses the parameter argument.  This parameter is a single object.  To pass multiple items into the script block, we create a custom object (using a PowerShell v3 language) with properties we want to pass in.

        Inside the script block, $parameter is used to reference this parameter object.  This example sets a content file, gets content from that file, and sets it to a predefined log file.

    .EXAMPLE
        $test = 5
        1..2 | Invoke-Parallel -ImportVariables {$_ * $test}

        Add variables from the current session to the session state.  Without -ImportVariables $Test would not be accessible

    .EXAMPLE
        $test = 5
        1..2 | Invoke-Parallel {$_ * $Using:test}

        Reference a variable from the current session with the $Using:<Variable> syntax.  Requires PowerShell 3 or later. Note that -ImportVariables parameter is no longer necessary.

    .FUNCTIONALITY
        PowerShell Language

    .NOTES
        Credit to Boe Prox for the base runspace code and $Using implementation
            http://learn-powershell.net/2012/05/10/speedy-network-information-query-using-powershell/
            http://gallery.technet.microsoft.com/scriptcenter/Speedy-Network-Information-5b1406fb#content
            https://github.com/proxb/PoshRSJob/

        Credit to T Bryce Yehl for the Quiet and NoCloseOnTimeout implementations

        Credit to Sergei Vorobev for the many ideas and contributions that have improved functionality, reliability, and ease of use

    .LINK
        https://github.com/RamblingCookieMonster/Invoke-Parallel
    #>
    [cmdletbinding(DefaultParameterSetName='ScriptBlock')]
    Param (   
        [Parameter(Mandatory=$false,position=0,ParameterSetName='ScriptBlock')]
            [System.Management.Automation.ScriptBlock]$ScriptBlock,

        [Parameter(Mandatory=$false,ParameterSetName='ScriptFile')]
        [ValidateScript({test-path $_ -pathtype leaf})]
            $ScriptFile,

        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Alias('CN','__Server','IPAddress','Server','ComputerName')]    
            [PSObject]$InputObject,

            [PSObject]$Parameter,

            [switch]$ImportVariables,

            [switch]$ImportModules,

            [int]$Throttle = 20,

            [int]$SleepTimer = 200,

            [int]$RunspaceTimeout = 0,

			[switch]$NoCloseOnTimeout = $false,

            [int]$MaxQueue,

        [validatescript({Test-Path (Split-Path $_ -parent)})]
            [string]$LogFile = "C:\temp\log.log",

			[switch] $Quiet = $false
    )
    
    Begin {
                
        #No max queue specified?  Estimate one.
        #We use the script scope to resolve an odd PowerShell 2 issue where MaxQueue isn't seen later in the function
        if( -not $PSBoundParameters.ContainsKey('MaxQueue') )
        {
            if($RunspaceTimeout -ne 0){ $script:MaxQueue = $Throttle }
            else{ $script:MaxQueue = $Throttle * 3 }
        }
        else
        {
            $script:MaxQueue = $MaxQueue
        }

        Write-Verbose "Throttle: '$throttle' SleepTimer '$sleepTimer' runSpaceTimeout '$runspaceTimeout' maxQueue '$maxQueue' logFile '$logFile'"

        #If they want to import variables or modules, create a clean runspace, get loaded items, use those to exclude items
        if ($ImportVariables -or $ImportModules)
        {
            $StandardUserEnv = [powershell]::Create().addscript({

                #Get modules and snapins in this clean runspace
                $Modules = Get-Module | Select -ExpandProperty Name
                $Snapins = Get-PSSnapin | Select -ExpandProperty Name

                #Get variables in this clean runspace
                #Called last to get vars like $? into session
                $Variables = Get-Variable | Select -ExpandProperty Name
                
                #Return a hashtable where we can access each.
                @{
                    Variables = $Variables
                    Modules = $Modules
                    Snapins = $Snapins
                }
            }).invoke()[0]
            
            if ($ImportVariables) {
                #Exclude common parameters, bound parameters, and automatic variables
                Function _temp {[cmdletbinding()] param() }
                $VariablesToExclude = @( (Get-Command _temp | Select -ExpandProperty parameters).Keys + $PSBoundParameters.Keys + $StandardUserEnv.Variables )
                Write-Verbose "Excluding variables $( ($VariablesToExclude | sort ) -join ", ")"

                # we don't use 'Get-Variable -Exclude', because it uses regexps. 
                # One of the veriables that we pass is '$?'. 
                # There could be other variables with such problems.
                # Scope 2 required if we move to a real module
                $UserVariables = @( Get-Variable | Where { -not ($VariablesToExclude -contains $_.Name) } ) 
                Write-Verbose "Found variables to import: $( ($UserVariables | Select -expandproperty Name | Sort ) -join ", " | Out-String).`n"

            }

            if ($ImportModules) 
            {
                $UserModules = @( Get-Module | Where {$StandardUserEnv.Modules -notcontains $_.Name -and (Test-Path $_.Path -ErrorAction SilentlyContinue)} | Select -ExpandProperty Path )
                $UserSnapins = @( Get-PSSnapin | Select -ExpandProperty Name | Where {$StandardUserEnv.Snapins -notcontains $_ } ) 
            }
        }

        #region functions
            
            Function Get-RunspaceData {
                [cmdletbinding()]
                param( [switch]$Wait )

                #loop through runspaces
                #if $wait is specified, keep looping until all complete
                Do {

                    #set more to false for tracking completion
                    $more = $false

                    #Progress bar if we have inputobject count (bound parameter)
                    if (-not $Quiet) {
						Write-Progress  -Activity "Running Query" -Status "Starting threads"`
							-CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
							-PercentComplete $( Try { $script:completedCount / $totalCount * 100 } Catch {0} )
					}

                    #run through each runspace.           
                    Foreach($runspace in $runspaces) {
                    
                        #get the duration - inaccurate
                        $currentdate = Get-Date
                        $runtime = $currentdate - $runspace.startTime
                        $runMin = [math]::Round( $runtime.totalminutes ,2 )

                        #set up log object
                        $log = "" | select Date, Action, Runtime, Status, Details
                        $log.Action = "Removing:'$($runspace.object)'"
                        $log.Date = $currentdate
                        $log.Runtime = "$runMin minutes"

                        #If runspace completed, end invoke, dispose, recycle, counter++
                        If ($runspace.Runspace.isCompleted) {
                            
                            $script:completedCount++
                        
                            #check if there were errors
                            if($runspace.powershell.Streams.Error.Count -gt 0) {
                                
                                #set the logging info and move the file to completed
                                $log.status = "CompletedWithErrors"
                                Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                                foreach($ErrorRecord in $runspace.powershell.Streams.Error) {
                                    Write-Error -ErrorRecord $ErrorRecord
                                }
                            }
                            else {
                                
                                #add logging details and cleanup
                                $log.status = "Completed"
                                Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                            }

                            #everything is logged, clean up the runspace
                            $runspace.powershell.EndInvoke($runspace.Runspace)
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null

                        }

                        #If runtime exceeds max, dispose the runspace
                        ElseIf ( $runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout) {
                            
                            $script:completedCount++
                            $timedOutTasks = $true
                            
							#add logging details and cleanup
                            $log.status = "TimedOut"
                            Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                            Write-Error "Runspace timed out at $($runtime.totalseconds) seconds for the object:`n$($runspace.object | out-string)"

                            #Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the powershell instance
                            if (!$noCloseOnTimeout) { $runspace.powershell.dispose() }
                            $runspace.Runspace = $null
                            $runspace.powershell = $null
                            $completedCount++

                        }
                   
                        #If runspace isn't null set more to true  
                        ElseIf ($runspace.Runspace -ne $null ) {
                            $log = $null
                            $more = $true
                        }

                        #log the results if a log file was indicated
                        if($logFile -and $log){
                            ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1] | out-file $LogFile -append
                        }
                    }

                    #Clean out unused runspace jobs
                    $temphash = $runspaces.clone()
                    $temphash | Where { $_.runspace -eq $Null } | ForEach {
                        $Runspaces.remove($_)
                    }

                    #sleep for a bit if we will loop again
                    if($PSBoundParameters['Wait']){ Start-Sleep -milliseconds $SleepTimer }

                #Loop again only if -wait parameter and there are more runspaces to process
                } while ($more -and $PSBoundParameters['Wait'])
                
            #End of runspace function
            }

        #endregion functions
        
        #region Init

            if($PSCmdlet.ParameterSetName -eq 'ScriptFile')
            {
                $ScriptBlock = [scriptblock]::Create( $(Get-Content $ScriptFile | out-string) )
            }
            elseif($PSCmdlet.ParameterSetName -eq 'ScriptBlock')
            {
                #Start building parameter names for the param block
                [string[]]$ParamsToAdd = '$_'
                if( $PSBoundParameters.ContainsKey('Parameter') )
                {
                    $ParamsToAdd += '$Parameter'
                }

                $UsingVariableData = $Null
                

                # This code enables $Using support through the AST.
                # This is entirely from  Boe Prox, and his https://github.com/proxb/PoshRSJob module; all credit to Boe!
                
                if($PSVersionTable.PSVersion.Major -gt 2)
                {
                    #Extract using references
                    $UsingVariables = $ScriptBlock.ast.FindAll({$args[0] -is [System.Management.Automation.Language.UsingExpressionAst]},$True)    

                    If ($UsingVariables)
                    {
                        $List = New-Object 'System.Collections.Generic.List`1[System.Management.Automation.Language.VariableExpressionAst]'
                        ForEach ($Ast in $UsingVariables)
                        {
                            [void]$list.Add($Ast.SubExpression)
                        }

                        $UsingVar = $UsingVariables | Group SubExpression | ForEach {$_.Group | Select -First 1}
        
                        #Extract the name, value, and create replacements for each
                        $UsingVariableData = ForEach ($Var in $UsingVar) {
                            Try
                            {
                                $Value = Get-Variable -Name $Var.SubExpression.VariablePath.UserPath -ErrorAction Stop
                                [pscustomobject]@{
                                    Name = $Var.SubExpression.Extent.Text
                                    Value = $Value.Value
                                    NewName = ('$__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
                                    NewVarName = ('__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
                                }
                            }
                            Catch
                            {
                                Write-Error "$($Var.SubExpression.Extent.Text) is not a valid Using: variable!"
                            }
                        }
                        $ParamsToAdd += $UsingVariableData | Select -ExpandProperty NewName -Unique

                        $NewParams = $UsingVariableData.NewName -join ', '
                        $Tuple = [Tuple]::Create($list, $NewParams)
                        $bindingFlags = [Reflection.BindingFlags]"Default,NonPublic,Instance"
                        $GetWithInputHandlingForInvokeCommandImpl = ($ScriptBlock.ast.gettype().GetMethod('GetWithInputHandlingForInvokeCommandImpl',$bindingFlags))
        
                        $StringScriptBlock = $GetWithInputHandlingForInvokeCommandImpl.Invoke($ScriptBlock.ast,@($Tuple))

                        $ScriptBlock = [scriptblock]::Create($StringScriptBlock)

                        Write-Verbose $StringScriptBlock
                    }
                }
                
                $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param($($ParamsToAdd -Join ", "))`r`n" + $Scriptblock.ToString())
            }
            else
            {
                Throw "Must provide ScriptBlock or ScriptFile"; Break
            }

            Write-Debug "`$ScriptBlock: $($ScriptBlock | Out-String)"
            Write-Verbose "Creating runspace pool and session states"

            #If specified, add variables and modules/snapins to session state
            $sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            if ($ImportVariables)
            {
                if($UserVariables.count -gt 0)
                {
                    foreach($Variable in $UserVariables)
                    {
                        $sessionstate.Variables.Add( (New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Variable.Name, $Variable.Value, $null) )
                    }
                }
            }
            if ($ImportModules)
            {
                if($UserModules.count -gt 0)
                {
                    foreach($ModulePath in $UserModules)
                    {
                        $sessionstate.ImportPSModule($ModulePath)
                    }
                }
                if($UserSnapins.count -gt 0)
                {
                    foreach($PSSnapin in $UserSnapins)
                    {
                        [void]$sessionstate.ImportPSSnapIn($PSSnapin, [ref]$null)
                    }
                }
            }

            #Create runspace pool
            $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
            $runspacepool.Open() 

            Write-Verbose "Creating empty collection to hold runspace jobs"
            $Script:runspaces = New-Object System.Collections.ArrayList        
        
            #If inputObject is bound get a total count and set bound to true
            $bound = $PSBoundParameters.keys -contains "InputObject"
            if(-not $bound)
            {
                [System.Collections.ArrayList]$allObjects = @()
            }

            #Set up log file if specified
            if( $LogFile ){
                New-Item -ItemType file -path $logFile -force | Out-Null
                ("" | Select Date, Action, Runtime, Status, Details | ConvertTo-Csv -NoTypeInformation -Delimiter ";")[0] | Out-File $LogFile
            }

            #write initial log entry
            $log = "" | Select Date, Action, Runtime, Status, Details
                $log.Date = Get-Date
                $log.Action = "Batch processing started"
                $log.Runtime = $null
                $log.Status = "Started"
                $log.Details = $null
                if($logFile) {
                    ($log | convertto-csv -Delimiter ";" -NoTypeInformation)[1] | Out-File $LogFile -Append
                }

			$timedOutTasks = $false

        #endregion INIT
    }

    Process {

        #add piped objects to all objects or set all objects to bound input object parameter
        if($bound)
        {
            $allObjects = $InputObject
        }
        Else
        {
            [void]$allObjects.add( $InputObject )
        }
    }

    End {
        
        #Use Try/Finally to catch Ctrl+C and clean up.
        Try
        {
            #counts for progress
            $totalCount = $allObjects.count
            $script:completedCount = 0
            $startedCount = 0

            foreach($object in $allObjects){
        
                #region add scripts to runspace pool
                    
                    #Create the powershell instance, set verbose if needed, supply the scriptblock and parameters
                    $powershell = [powershell]::Create()
                    
                    if ($VerbosePreference -eq 'Continue')
                    {
                        [void]$PowerShell.AddScript({$VerbosePreference = 'Continue'})
                    }

                    [void]$PowerShell.AddScript($ScriptBlock).AddArgument($object)

                    if ($parameter)
                    {
                        [void]$PowerShell.AddArgument($parameter)
                    }

                    # $Using support from Boe Prox
                    if ($UsingVariableData)
                    {
                        Foreach($UsingVariable in $UsingVariableData) {
                            Write-Verbose "Adding $($UsingVariable.Name) with value: $($UsingVariable.Value)"
                            [void]$PowerShell.AddArgument($UsingVariable.Value)
                        }
                    }

                    #Add the runspace into the powershell instance
                    $powershell.RunspacePool = $runspacepool
    
                    #Create a temporary collection for each runspace
                    $temp = "" | Select-Object PowerShell, StartTime, object, Runspace
                    $temp.PowerShell = $powershell
                    $temp.StartTime = Get-Date
                    $temp.object = $object
    
                    #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                    $temp.Runspace = $powershell.BeginInvoke()
                    $startedCount++

                    #Add the temp tracking info to $runspaces collection
                    Write-Verbose ( "Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring() )
                    $runspaces.Add($temp) | Out-Null
            
                    #loop through existing runspaces one time
                    Get-RunspaceData

                    #If we have more running than max queue (used to control timeout accuracy)
                    #Script scope resolves odd PowerShell 2 issue
                    $firstRun = $true
                    while ($runspaces.count -ge $Script:MaxQueue) {

                        #give verbose output
                        if($firstRun){
                            Write-Verbose "$($runspaces.count) items running - exceeded $Script:MaxQueue limit."
                        }
                        $firstRun = $false
                    
                        #run get-runspace data and sleep for a short while
                        Get-RunspaceData
                        Start-Sleep -Milliseconds $sleepTimer
                    
                    }

                #endregion add scripts to runspace pool
            }
                     
            Write-Verbose ( "Finish processing the remaining runspace jobs: {0}" -f ( @($runspaces | Where {$_.Runspace -ne $Null}).Count) )
            Get-RunspaceData -wait

            if (-not $quiet) {
			    Write-Progress -Activity "Running Query" -Status "Starting threads" -Completed
		    }
        }
        Finally
        {
            #Close the runspace pool, unless we specified no close on timeout and something timed out
            if ( ($timedOutTasks -eq $false) -or ( ($timedOutTasks -eq $true) -and ($noCloseOnTimeout -eq $false) ) ) {
	            Write-Verbose "Closing the runspace pool"
			    $runspacepool.close()
            }

            #collect garbage
            [gc]::Collect()
        }       
    }
}

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

        # Create Virtual Network
        foreach($VNET in $VNETHash){
            if($VNET.VNETName -ne $null){
                $ResourceGroupName = $VNET.VNETRSG
                $VNETDNS += @{Name=$VNET.Name;vnetRSG=$VNET.VNETRSG;PrimaryDNS=$VNET.PrimaryDNS;SecondaryDNS=$VNET.SecondaryDNS}
                $template = $VNET.Template
                $SAS = $VNET.SAS
                $Size = $VNET.VNETSize
                
                $VNET.Remove('Name')
                $VNET.Remove('PrimaryDNS')
                $VNET.Remove('SecondaryDNS')
                $VNET.Remove('VNETRSG')
                $VNET.Remove('Template')
                $VNET.Remove('SAS')
                $VNET.Remove('VNETSize')
                $VNET.ADD('buildDate',$SubHash.BuildDate)
                $VNET.ADD('buildBy',$SubHash.BuildBy)
                
                switch($Size){

                    "Small" {
                                $VNET.Remove('environmentB')
                                $VNET.Remove('subnetDMZCIDRB')
                                $VNET.Remove('subnetAPPCIDRB')
                                $VNET.Remove('subnetINSCIDRB')
                                $VNET.Remove('subnetADCIDRB')
                                $VNET.Remove('subnetAGWCIDRB')
                                $VNET.Remove('environmentC')
                                $VNET.Remove('subnetDMZCIDRC')
                                $VNET.Remove('subnetAPPCIDRC')
                                $VNET.Remove('subnetINSCIDRC')
                                $VNET.Remove('subnetADCIDRC')
                                $VNET.Remove('subnetAGWCIDRC')                            
                            
                            }
                    "Medium" {
                                $VNET.Remove('environmentC')
                                $VNET.Remove('subnetDMZCIDRC')
                                $VNET.Remove('subnetAPPCIDRC')
                                $VNET.Remove('subnetINSCIDRC')
                                $VNET.Remove('subnetADCIDRC')
                                $VNET.Remove('subnetAGWCIDRC')     
                                                         
                            }

                }

                try{
                    Write-Host "Creating Virtual Network: $($VNET.VNETName) in $ResourceGroupName" -ForegroundColor Green
                    $status = New-AzureRmResourceGroupDeployment -Name ($SubHash.DeploymentName + "-VNET") -ResourceGroupName $ResourceGroupName `
                                                   -Mode Incremental `
                                                   -TemplateParameterObject $VNET `
                                                   -TemplateFile ("$template" + "$SAS") `
                                                   -Force
                    if($status.ProvisioningState -eq 'Succeeded'){
                        Write-Host "Success: Creating Virtual Network: $($VNET.VNETName) in $ResourceGroupName" -ForegroundColor Green
                    }
                    else{
                          Write-Host "Warning: Creating Virtual Network: $($VNET.VNETName) in $ResourceGroupName is not in a Succeeded state, please validate" -ForegroundColor Yellow
                          break
                    }
                }
                catch{
                    Write-Host "Error: Creating Virtual Network: $($VNET.VNETName) in $ResourceGroupName" -ForegroundColor Red
                    break

                }

            }

        }

        # Create OMS Workspace
        foreach($OMS in $OMSHash){
            if($OMS.OMSworkspaceName -ne $null){
                $OMSWorkspace = $OMS.OMSworkspaceName
                $OMSRSG = $OMS.OMSRSG
                $ResourceGroupName = $OMS.OMSRSG
                $template = $OMS.Template
                $SAS = $OMS.SAS
                $OMS.Remove('OMSRSG')
                $OMS.Remove('Template')
                $OMS.Remove('SAS')
                try{
                    Write-Host "Creating OMS Workspace: $($OMS.OMSworkspaceName) in $ResourceGroupName" -ForegroundColor Green
                    $status = New-AzureRmResourceGroupDeployment -Name ($SubHash.DeploymentName + "-OMS") -ResourceGroupName $ResourceGroupName `
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

        # Create Recovery Services Vault
        foreach($RSV in $RSVHash){
            if($RSV.vaultName -ne $null){
                $ResourceGroupName = $RSV.ResourceGroupName
                $template = $RSV.Template
                $SAS = $RSV.SAS
                $RSV.Remove('ResourceGroupName')
                $RSV.Remove('Template')
                $RSV.Remove('SAS')
                try{
                    Write-Host "Creating Recovery Services Vault Policy: $($RSV.policyName) in $($RSV.vaultName) in $ResourceGroupName" -ForegroundColor Green
                    $status = New-AzureRmResourceGroupDeployment -Name ($SubHash.DeploymentName + "-RSV") -ResourceGroupName $ResourceGroupName `
                                                   -Mode Incremental `
                                                   -TemplateParameterObject $RSV `
                                                   -TemplateFile ("$template" + "$SAS") `
                                                   -Force
                    if($status.ProvisioningState -eq 'Succeeded'){
                        Write-Host "Success: Creating Recovery Services Vault Policy: $($RSV.policyName) in $($RSV.vaultName) in $ResourceGroupName" -ForegroundColor Green
                    }
                    else{
                          Write-Host "Warning: Creating Recovery Services Vault Policy: $($RSV.policyName) in $($RSV.vaultName) in $ResourceGroupName is not in a Succeeded state, please validate" -ForegroundColor Yellow
                          break
                    }
                }
                catch{
                    Write-Host "Error: Creating Recovery Services Vault Policy: $($RSV.policyName) in $($RSV.vaultName) in $ResourceGroupName" -ForegroundColor Red
                    break

                }

            }

        }

        
        # Create Domain Controllers
        foreach($AD in $ADHash){
            if($AD.vmNamePrefix -ne $null){
                $ResourceGroupName = $AD.ResourceGroupName
                $template = $AD.Template
                $SAS = $AD.sasToken
                $Type = $AD.Type
                $domainAdminPassword = $DomainPassword | ConvertTo-SecureString -AsPlainText -Force
                $AD.Add('domainAdminPassword',$domainAdminPassword)
                $AD.Remove('ResourceGroupName')
                $AD.Remove('Template')
                $AD.Remove('Type')
                $AD.Add('buildDate',$SubHash.BuildDate)
                $AD.Add('buildBy',$SubHash.BuildBy)

                $workspaceID = (Get-AzureRmOperationalInsightsWorkspace -ResourceGroupName $OMSRSG -Name $OMSWorkspace).CustomerId.ToString()
                $workspaceKey = (Get-AzureRmOperationalInsightsWorkspaceSharedKeys -ResourceGroupName $OMSRSG -Name $OMSWorkspace).PrimarySharedKey.ToString()

                $AD.Add('workspaceId',$workspaceID)
                $AD.Add('workspaceKey',$workspaceKey)

                if($Type -eq 'Single'){
                    $AD.Remove('availabilitySetName')
                    $AD.Add('vmName',$AD.vmNamePrefix)
                    $AD.Remove('vmNamePrefix')
                    $AD.Remove('secondaryDCIp')
                }

                try{
                    Write-Host "Creating Active Directory. Type: $Type Domain: $($AD.domainName) in $ResourceGroupName -- This will take over 30 minutes" -ForegroundColor Green
                    $status = New-AzureRmResourceGroupDeployment -Name ($SubHash.DeploymentName + "-AD") -ResourceGroupName $ResourceGroupName `
                                                   -Mode Incremental `
                                                   -TemplateParameterObject $AD `
                                                   -TemplateFile ("$template" + "$SAS") `
                                                   -Force
                    if($status.ProvisioningState -eq 'Succeeded'){
                        Write-Host "Success: Creating Active Directory. Type: $Type Domain: $($AD.domainName) in $ResourceGroupName" -ForegroundColor Green
                        Write-Host "Changing DNS on VNETS" -ForegroundColor Green
                        foreach($VNET in $VNETDNS){
                            if($VNET.Name -ne $null){
                                $vn = Get-AzureRmVirtualNetwork -Name $VNET.Name -ResourceGroupName $VNET.vnetRSG
                                $vn.DhcpOptions.DnsServers = $VNET.primaryDNS
                                if($Type -eq 'Multi'){
                                    $vn.DhcpOptions.DnsServers += $VNET.SecondaryDNS
                                }
                                try{
                                    $Status = Set-AzureRmVirtualNetwork -VirtualNetwork $vn
                                    if($Status = 'Succeeded'){
                                        Write-Host "Success: DNS modified on $($VNET.Name)" -ForegroundColor Green

                                    }
                                    else{
                                        Write-Host "Warning: DNS modified on $($VNET.Name)" -ForegroundColor Yellow 
                                        break
                                    }
                                }
                                catch{
                                        Write-Host "Error: DNS modified on $($VNET.Name)" -ForegroundColor Red
                                        break
                                }
                            }
                        }
       
                        $VM = Get-AzureRmVM -ResourceGroupName $ResourceGroupName
                        foreach($V in $VM){
                            Write-Host "Restarting VM: $($V.name)" -ForegroundColor Green
                            $status = Restart-AzureRmVM -ResourceGroupName $ResourceGroupName -Name $V.name
                        }
                        
                    }
                    else{
                          Write-Host "Warning: Creating Active Directory. Type: $Type Domain: $($AD.domainName) in $ResourceGroupName is not in a Succeeded state, please validate" -ForegroundColor Yellow
                          break
                    }
                }
                catch{
                    Write-Host "Error: Creating Active Directory. Type: $Type Domain: $($AD.domainName) in $ResourceGroupName" -ForegroundColor Red
                    break

                }

            }

        }
        
        # Create Virtual Network Gateway
        foreach($VPN in $VPNHash){

            if($VPN.VPNname -ne $null){
        
                if(Get-AzureRmVirtualNetworkGateway -Name $VPN.VPNname -ResourceGroupName $VPN.ResourceGroupName -ErrorAction SilentlyContinue){
                    Write-Host "Virtual Network Gateway: $($VPN.VPNname) in $($VPN.ResourceGroupName) already exists" -ForegroundColor Yellow
                }
                else{

                        try{
                            Write-Host "Creating Virtual Network Gateway: $($VPN.VPNname) in $($VPN.ResourceGroupName) - This will take over 20 mins" -ForegroundColor Green

                            $gwpip= New-AzureRmPublicIpAddress -Name ($VPN.VPNName + "-PIP") -ResourceGroupName $VPN.ResourceGroupName -Location $VPN.Location -AllocationMethod Dynamic
                            $vnet = Get-AzureRmVirtualNetwork -Name $VPN.VNET -ResourceGroupName $VPN.VNETRSG
                            $subnet = Get-AzureRmVirtualNetworkSubnetConfig -Name 'GatewaySubnet' -VirtualNetwork $vnet
                            $gwipconfig = New-AzureRmVirtualNetworkGatewayIpConfig -Name gwipconfig1 -SubnetId $subnet.Id -PublicIpAddressId $gwpip.Id
                            $Status = New-AzureRmVirtualNetworkGateway -Name $VPN.VPNName -ResourceGroupName $VPN.ResourceGroupName `
                                                             -Location $VPN.Location -IpConfigurations $gwipconfig -GatewayType $VPN.GatewayType `
                                                             -VpnType $VPN.VPNType -GatewaySku $VPN.GatewaySku -Tag @{ BuildBy=$SubHash.BuildBy;BuildDate=$SubHash.BuildDate;Ticket=$SubHash.Ticket }

                            if($status.ProvisioningState -eq 'Succeeded'){
                                Write-Host "Success: Creating Virtual Network Gateway: $($VPN.VPNname) in $($VPN.ResourceGroupName)" -ForegroundColor Green
                            }
                            else{
                                Write-Host "Warning: Creating Virtual Network Gateway: $($VPN.VPNname) in $($VPN.ResourceGroupName) is not in a Succeeded state, please validate" -ForegroundColor Yellow
                                break
                            }
                        }
                        catch{
                                Write-Host "Error: Creating Virtual Network Gateway: $($VPN.VPNname) in $($VPN.ResourceGroupName)" -ForegroundColor Red
                                break

                        }        

                }
            }
        }

        # Create Virtual Network Connections
        foreach($VPNCON in $VPNCONHash){

            if($VPNCON.Name -ne $null){
        
                if(Get-AzureRmVirtualNetworkGatewayConnection -Name $VPNCON.Name -ResourceGroupName $VPNCON.ResourceGroupName -ErrorAction SilentlyContinue){
                    Write-Host "Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName) already exists" -ForegroundColor Yellow
                }
                else{
                        if($VPNCON.ConnectionType -eq 'IPSec'){

                            try{
                                Write-Host "Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName)" -ForegroundColor Green
                                    $Status = New-AzureRmLocalNetworkGateway -Name $VPNCON.LNGName -ResourceGroupName $VPNCON.ResourceGroupName `
                                                                   -Location $VPNCON.Location -GatewayIpAddress $VPNCON.LNGGatewayIP -AddressPrefix $VPNCON.LNGAddressPrefix

                                    $gateway1 = Get-AzureRmVirtualNetworkGateway -Name $VPNCON.VirtualGateway -ResourceGroupName $VPNCON.ResourceGroupName
                                    $local = Get-AzureRmLocalNetworkGateway -Name $VPNCON.LNGName -ResourceGroupName $VPNCON.ResourceGroupName

                                    $Status = New-AzureRmVirtualNetworkGatewayConnection -Name $VPNCON.Name -ResourceGroupName $VPNCON.ResourceGroupName `
                                                                               -Location $VPNCON.Location -VirtualNetworkGateway1 $gateway1 -LocalNetworkGateway2 $local `
                                                                               -ConnectionType $VPNCON.ConnectionType -RoutingWeight $VPNCON.RoutingWeight -SharedKey $VPNCON.SharedKey

                                    if($status.ProvisioningState -eq 'Succeeded'){
                                        Write-Host "Success: Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName)" -ForegroundColor Green
                                    }
                                    else{
                                        Write-Host "Warning: Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName) is not in a Succeeded state, please validate" -ForegroundColor Yellow
                                        break
                                    }
                            }
                            catch{
                                    Write-Host "Error: Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName)" -ForegroundColor Red
                                    break

                            }        
                        }
                        elseif($VPNCON.ConnectionType -eq 'Vnet2Vnet'){
                            try{
                                    Write-Host "Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName)" -ForegroundColor Green

                                        $gateway1 = Get-AzureRmVirtualNetworkGateway -Name $VPNCON.VirtualGateway -ResourceGroupName $VPNCON.ResourceGroupName
                                        $local = Get-AzureRmVirtualNetworkGateway -Name $VPNCON.LNGName -ResourceGroupName $VPNCON.LNGRSG

                                        $Status = New-AzureRmVirtualNetworkGatewayConnection -Name $VPNCON.Name -ResourceGroupName $VPNCON.ResourceGroupName `
                                                                                   -Location $VPNCON.Location -VirtualNetworkGateway1 $gateway1 -VirtualNetworkGateway2 $local `
                                                                                   -ConnectionType $VPNCON.ConnectionType -SharedKey $VPNCON.SharedKey

                                        if($status.ProvisioningState -eq 'Succeeded'){
                                            Write-Host "Success: Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName)" -ForegroundColor Green
                                        }
                                        else{
                                            Write-Host "Warning: Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName) is not in a Succeeded state, please validate" -ForegroundColor Yellow
                                            break
                                        }
                                }
                                catch{
                                        Write-Host "Error: Creating Virtual Network Connection: $($VPNCON.Name) in $($VPNCON.ResourceGroupName)" -ForegroundColor Red
                                        break

                                }     

                        }
                        else{ 
                            Write-Host "Connection Type not recognized for $($VPNCON.Name)" -ForegroundColor Red
                            break
                        }
                }
            }
        }

        # Create Network Security Groups
        $NSGConfigList = $null
        $NSGConfigList = Get-AzureRmNetworkSecurityGroup
        
        foreach($NSG in $NSGConfigList){
            
            $NSGRuleHash = $NSGHash | Where { $_.Name -eq $NSG.Name }
            foreach($NSGRule in $NSGRuleHash){
                if($NSGRule.Name -ne $null){
                    try{
                        Write-Host "Adding NSG rule: $($NSGRule.Rule) to $($NSG.Name)" -ForegroundColor Green
                        $status = Add-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG -Name $NSGRule.Rule -Description $NSGRule.Description `
                                                                      -Access $NSGRule.Access -Protocol $NSGRule.Protocol -Direction $NSGRule.Direction -Priority $NSGRule.Priority `
                                                                      -SourceAddressPrefix $NSGRule.SourceAddressPrefix -SourcePortRange "$($NSGRule.SourcePortRange)" `
                                                                      -DestinationAddressPrefix $NSGRule.DestinationAddressPrefix -DestinationPortRange "$($NSGRule.DestinationPortRange)" | out-Null
                    }
                    catch{
                        Write-Host "Conflicting NSG rule: $($NSGRule.Rule) already exists in $($NSG.Name)" -ForegroundColor Yellow
                        
                    }                                              
                }
            }
            try{
                Write-Host "Setting NSG: $($NSG.Name)" -ForegroundColor Green
                $Status = Set-AzureRmNetworkSecurityGroup -NetworkSecurityGroup $NSG
            }
            catch{
                Write-Host "Error: Setting NSG: $($NSG.Name)" -ForegroundColor Red
                break
            }
        }

        # Create Virtual Servers

        $workspaceID = (Get-AzureRmOperationalInsightsWorkspace -ResourceGroupName $OMSRSG -Name $OMSWorkspace).CustomerId.ToString()
        $workspaceKey = (Get-AzureRmOperationalInsightsWorkspaceSharedKeys -ResourceGroupName $OMSRSG -Name $OMSWorkspace).PrimarySharedKey.ToString()

        $stuff = [pscustomobject] @{

            workspaceID = $workspaceID
            workspaceKey = $workspaceKey
            buildDate = $SubHash.BuildDate
            buildBy = $SubHash.BuildBy
	        DomainPassword = $DomainPassword | ConvertTo-SecureString -AsPlainText -Force
            DeploymentName = $SubHash.DeploymentName
        }

        $VMHash | Invoke-Parallel -Quiet -NoCloseOnTimeout -Parameter $stuff {
        
            if($_.vmName -ne $null){
                $ResourceGroupName = $_.VMRSG
                $template = $_.Template
                $SAS = $_.sasToken
                
                $_.Add('buildDate',$Parameter.buildDate)
                $_.Add('buildBy',$Parameter.buildBy)
                $_.Add('adminPassword',$Parameter.DomainPassword)
                $_.Add('domainAdminPassword',$Parameter.DomainPassword)
                $_.Remove('VMRSG')
                $_.Remove('Template')
                $_.Add('workspaceId',$Parameter.workspaceID)
                $_.Add('workspaceKey',$Parameter.workspaceKey)
                
                if($_.operatingSystem -match "SQL"){
                    $_.Remove('lbOption')
                    $_.Remove('lbPort')
                    $_.Remove('lbProtocol')
                    $_.Remove('lbProbePath')
                    $_.Remove('dnsLabel')
                    $_.Remove('deployWebServer')
                }
                try{
                    Write-Host "Creating Virtual Machine: $($_.vmName) in $ResourceGroupName" -ForegroundColor Green
                    $status = New-AzureRmResourceGroupDeployment -Name ($Parameter.DeploymentName + "-$($_.vmName)") -ResourceGroupName $ResourceGroupName `
                                                   -Mode Incremental `
                                                   -TemplateParameterObject $_ `
                                                   -TemplateFile ("$template" + "$SAS") `
                                                   -Force
                    if($status.ProvisioningState -eq 'Succeeded'){
                        Write-Host "Success: Creating Virtual Machine: $($_.vmName) in $ResourceGroupName" -ForegroundColor Green
                    }
                    else{
                          Write-Host "Warning: Creating Virtual Machine: $($_.vmName) in $ResourceGroupName is not in a Succeeded state, please validate" -ForegroundColor Yellow
                          break
                    }
                }
                catch{
                    Write-Host "Error: Creating Virtual Machine: $($_.vmName) in $ResourceGroupName" -ForegroundColor Red
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
                    $status = New-AzureRmResourceGroupDeployment -Name ($SubHash.DeploymentName + "-WebApps") -ResourceGroupName $ResourceGroupName `
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

#}

