# Clear Screen
cls
# CyberArk Integration
$cyberark= Invoke-RestMethod -Method Get -Uri http://10.150.36.150/AIMWebService/api/accounts -Body @{AppID = "QEC4E"; Safe = "FIDEV-QEC4E"; Object = "Website-GenericWebApp-httpsapp.genrocket.com-CAF_GenRocket_UserName"}
$cyberarkuser=$cyberark.Content


$JSONOUT= Invoke-RestMethod -Method Get -Uri http://10.150.36.150/AIMWebService/api/accounts -Body @{AppID = "QEC4E"; Safe = "FIDEV-QEC4E"; Object = "GENROCKET_USER_PWD"}
$cyberarkpwd=$JSONOUT.Content
# Disable the Proxy (Global)
# Filename: ProxyDisable.ps1
# Requires an Elevated PowerShell

# Disable the Proxy (Global)
Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\' -Name 'ProxyEnable' -Value 0
# Unset any Proxy Variables (Local)
$pvars=@('http_proxy','https_proxy', 'no_proxy')
foreach ($pvar in $pvars) {
  Remove-Item "ENV:\${pvar}" -ErrorAction SilentlyContinue
}

# Might have to search for and Disable the McAfee Proxy
#Stop-Service -Name 'mcpservice' -Force
#netsh winhttp show proxy

$ContentPath = "F:\Test\"
$_url1 = "https://app.genrocket.com/rest/login"
$jbody1 = Get-Content -Path "$($ContentPath)GenReqBody.txt"

$jbody1 = $jbody1.Replace("#user#",$cyberarkuser).Replace("#password#",$cyberarkpwd)


# Login to GenRocket and get Auth Taken
$Resplogin = Invoke-WebRequest -UseBasicParsing $_url1 -Body $jbody1  -Method Post -ContentType "application/json" 



$respjson = $Resplogin | ConvertFrom-Json



#Determine AccessToken
$AccessToken =""
foreach ($name in $respjson) 
{
    if ( $name.username -match "$cyberarkuser" ) {    
        $AccessToken= $name.accessToken    
             
    }   
}



#GetAllProjects
$_url2 = "https://app.genrocket.com/rest/project/list"
# External 
$orgId = "117872ef-c6cb-44e0-9a23-2f69b5384335"
$JBody = '{
           "organizationId": "#orgId"
           }'

$JBody=$JBody.replace("#orgId",$orgId)

#Write-Output $JBody

#$ProjectList = Invoke-WebRequest -UseBasicParsing $_url2 -Body $JBody  -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}
$ProjectList = Get-Content -path F:\test\Incremental.json

$Proplist = $ProjectList | ConvertFrom-Json 
#Write-output $Proplist
# External Organization ID
$orgId = "117872ef-c6cb-44e0-9a23-2f69b5384335"
# Base Directory
$baseDir = "F:\GenRocketHome\genrocket\bin\Files"

#Remove-Item -LiteralPath $baseDir -Force -Recurse

 #Write-Output $Proplist.scenarios.Count
foreach ($name in $Proplist) 
{

    $lcount = $name.projects.Count
    #Write-Output $lcount
    for($i=0;$i -lt $lcount; $i++)
    {
       $subdirname= $name.projects.GetValue($i).name
       Write-Output "$baseDir\$subdirname"
       
       $vcount = $name.projects.GetValue($i).projectVersions.Count
       for($j=0;$j -lt $vcount; $j++)
       {
            
           # Write-Output $name.projects.GetValue($i).name  " " $name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber
            # ########################################################
            $JBody = '{
           "organizationId": "#orgId",
           "projectName": "#ProjectName",
           "versionNumber": "#versionNumber"
           }'

            # Replace #org id with organizaqtion id            
            $JBody=$JBody.replace("#orgId",$orgId)
            # Replace #ProjectName with project name
            $JBody=$JBody.replace("#ProjectName",$name.projects.GetValue($i).name  )
            # Replace #versionNumber with project version number
            $JBody=$JBody.replace("#versionNumber",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
            
            $url2 = "https://app.genrocket.com/rest/scenario/list"

            $ScenList = Invoke-WebRequest -UseBasicParsing $url2 -Body $JBody  -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}

            $url2x = "https://app.genrocket.com/rest/chain/list"
            $ScenChainList = Invoke-WebRequest -UseBasicParsing $url2x -Body $JBody  -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}

            $url2xx = "https://app.genrocket.com/rest/chainSet/list"
            $ScenChainSetList = Invoke-WebRequest -UseBasicParsing $url2xx -Body $JBody  -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}

            $url3gq = "https://app.genrocket.com/rest/gQuery/list"
            $GQueryList = Invoke-WebRequest -UseBasicParsing $url3gq -Body $JBody  -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}    
            
            $urltdc ="https://app.genrocket.com/rest/testDataCase/list"    
            $GCaseList = Invoke-WebRequest -UseBasicParsing $urltdc -Body $JBody  -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}    

            $ScenList = $ScenList | ConvertFrom-Json
            $ScenChainList = $ScenChainList | ConvertFrom-Json
            $ScenChainSetList = $ScenChainSetList | ConvertFrom-Json
            $GQueryList = $GQueryList | ConvertFrom-Json
            $GCaseList = $GCaseList | ConvertFrom-Json

            #GQuery List code change on 4/17/2024
            foreach ($name_X in $GQueryList) 
            {
                $lcount2 = $name_X.gQueries.Count
                #Write-Output $lcount
                for($k=0;$k -lt $lcount2; $k++)
                {
                    $gQueryName = $name_X.gQueries.GetValue($k).name
                    # Write-Output "GQUERY: $gQueryName"
                    # Download gQuery
                    $_url3 = "https://app.genrocket.com/rest/gQuery/download"
                    $JBody2 = '{"organizationId": "#orgId","projectName": "#ProjectName","versionNumber": "#versionNumber","gQueryName": "#gQueryName"}'
                    $JBody2 = $JBody2.Replace("#gQueryName",$gQueryName)
                    # Replace #org id with organizaqtion id            
                    $JBody2=$JBody2.replace("#orgId",$orgId)
                    # Replace #ProjectName with project name
                    $JBody2=$JBody2.replace("#ProjectName",$name.projects.GetValue($i).name  )
                    # Replace #versionNumber with project version number
                    $JBody2=$JBody2.replace("#versionNumber",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)

                    $destfile = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\#gQueryName.gtdq"
                    $destdir = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\"

                    $destdir = $destdir.Replace("#Project",$name.projects.GetValue($i).name )
                    $destdir = $destdir.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)

                    if($k -eq 0)
                    {
                        Remove-Item -LiteralPath $destdir -Force -Recurse
                    }
                    #Write-Output $destfile
                    if (-Not(Test-Path -Path $destdir))
                    {
                        [system.io.directory]::CreateDirectory("$($destdir)")
                    }

                    $destfile = $destfile.Replace("#Project",$name.projects.GetValue($i).name )
                    $destfile = $destfile.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
                    $destfile = $destfile.Replace("#gQueryName",$gQueryName)

                    
                    try
                    {
                        Invoke-WebRequest -UseBasicParsing $_url3 -Body $JBody2 -Method Post -ContentType 'application/json' -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken} -OutFile $destfile

                    }
                    catch [System.Net.WebException] 
                    {
                         Write-Output $_.Exception.Response
                         Write-Output $($ResponseQueue)

                    }
                    #Write-Output $destfile
                }    
            }
            #GQuery List

            #GCase List code change on 4/17/2024
            foreach ($name_X in $GCaseList) 
            {
                $lcount2 = $name_X.testDataCases.Count
                #Write-Output $lcount
                for($k=0;$k -lt $lcount2; $k++)
                {
                    $tdcname = $name_X.testDataCases.GetValue($k).testDataSuiteName
                    # Write-Output "GQUERY: $gQueryName"
                    # Download gQuery
                    $_url3 = "https://app.genrocket.com/rest/testDataCase/download"
                    $JBody2 = '{"organizationId": "#orgId","projectName": "#ProjectName","versionNumber": "#versionNumber","testDataSuiteName": "#tdcname"}'
                    $JBody2 = $JBody2.Replace("#tdcname",$tdcname)
                    # Replace #org id with organizaqtion id            
                    $JBody2=$JBody2.replace("#orgId",$orgId)
                    # Replace #ProjectName with project name
                    $JBody2=$JBody2.replace("#ProjectName",$name.projects.GetValue($i).name  )
                    # Replace #versionNumber with project version number
                    $JBody2=$JBody2.replace("#versionNumber",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)

                    $destfile = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\#tdcname.gtdc"
                    $destdir = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\"

                    $destdir = $destdir.Replace("#Project",$name.projects.GetValue($i).name )
                    $destdir = $destdir.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)


                    Write-Output "GCASE: $destfile"
                    if (-Not(Test-Path -Path $destdir))
                    {
                        [system.io.directory]::CreateDirectory("$($destdir)")
                    }

                    $destfile = $destfile.Replace("#Project",$name.projects.GetValue($i).name )
                    $destfile = $destfile.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
                    $destfile = $destfile.Replace("#tdcname",$tdcname)

                    
                    try
                    {
                        Invoke-WebRequest -UseBasicParsing $_url3 -Body $JBody2 -Method Post -ContentType 'application/json' -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken} -OutFile $destfile

                    }
                    catch [System.Net.WebException] 
                    {
                         Write-Output $_.Exception.Response
                         Write-Output $($ResponseQueue)

                    }
                    Write-Output $destfile
                }    
            }
            #GCase List

            $ScenarioID =""

            foreach ($name_X in $ScenList) 
            {
                $lcount2 = $name_X.scenarios.Count
                #Write-Output $lcount
                for($k=0;$k -lt $lcount2; $k++)
                {
                    $ScenarioID = $name_X.scenarios.GetValue($k).externalId
                    #Write-Output $ScenarioID

                    # Download Scenario
                    $_url3 = "https://app.genrocket.com/rest/scenario/download"
                    $JBody2 = '{"scenarioId": "#ScenarioID"}'
                    $JBody2 = $JBody2.Replace("#ScenarioID",$ScenarioID)

                    $destfile = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\#sname.grs"
                    $destdir = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\"
                    
                    $destdir = $destdir.Replace("#Project",$name.projects.GetValue($i).name )
                    $destdir = $destdir.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
                    
                    #Write-Output $destfile
                    if (-Not(Test-Path -Path $destdir))
                    {
                        [system.io.directory]::CreateDirectory("$($destdir)")
                    }

                    $destfile = $destfile.Replace("#Project",$name.projects.GetValue($i).name )
                    $destfile = $destfile.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
                    $destfile = $destfile.Replace("#sname",$name_X.scenarios.GetValue($k).name)

                    # Queue Check
                    $JBodyQueue = '{"scenarioId": "#ScenarioID"}'
                    $JBodyQueue = $JBodyQueue.Replace("#ScenarioID",$ScenarioID)
                    $ResponseQueue =""

                    while(($ResponseQueue.Trim() -NotLike "*fileReady*") -and ($ResponseQueue.Trim() -NotLike "*isReady*"))
                    {
                        
                        $ResponseQueue = (Invoke-WebRequest -UseBasicParsing "https://app.genrocket.com/rest/scenario/verify" -Method Post -Body $JBodyQueue -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}).Content

                        Start-Sleep -s 1
                    }


                    # End Queue Check

                    #Write-Output $destfile
                    try
                    {
                        Invoke-WebRequest -UseBasicParsing $_url3 -Body $JBody2 -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken} -OutFile $destfile
                    }
                    catch [System.Net.WebException] 
                    {
                         Write-Output $_.Exception.Response
                         Write-Output $($ResponseQueue)

                    }
                    #Write-Output $destfile
                }    
            }
            foreach ($name_X in $ScenChainList) 
            {
                $lcount2 = $name_X.chains.Count
                #Write-Output $lcount
                for($k=0;$k -lt $lcount2; $k++)
                {
                    $ScenarioID = $name_X.chains.GetValue($k).externalId
                    #Write-Output $ScenarioID

                    # Download Scenario
                    $_url3x = "https://app.genrocket.com/rest/chain/download"
                    $JBody2 = '{"chainId": "#chainID"}'
                    $JBody2 = $JBody2.Replace("#chainID",$ScenarioID)

                    $destfile = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\#sname.grs"
                    $destdir = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\"

                    $destdir = $destdir.Replace("#Project",$name.projects.GetValue($i).name )
                    $destdir = $destdir.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)

                    #Write-Output $destfile
                    if (-Not(Test-Path -Path $destdir))
                    {
                        [system.io.directory]::CreateDirectory("$($destdir)")
                    }

                    $destfile = $destfile.Replace("#Project",$name.projects.GetValue($i).name )
                    $destfile = $destfile.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
                    $destfile = $destfile.Replace("#sname",$name_X.chains.GetValue($k).name)

                    # Queue Check
                    $JBodyQueue = '{"chainId": "#ScenarioID"}'
                    $JBodyQueue = $JBodyQueue.Replace("#ScenarioID",$ScenarioID)
                    $ResponseQueue =""

                    while(($ResponseQueue.Trim() -NotLike "*fileReady*") -and ($ResponseQueue.Trim() -NotLike "*isReady*"))
                    {
                        #Write-Output $ResponseQueue
                        $ResponseQueue = (Invoke-WebRequest -UseBasicParsing "https://app.genrocket.com/rest/chain/verify" -Method Post -Body $JBodyQueue -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}).Content
                        Start-Sleep -s 1
                    }


                    # End Queue Check

                    #Write-Output $destfile
                    try
                    {
                        Invoke-WebRequest -UseBasicParsing $_url3x -Body $JBody2 -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken} -OutFile $destfile
                    
                    }
                    catch [System.Net.WebException] 
                    {
                         Write-Output $_.Exception.Response
                         Write-Output $ResponseQueue

                    }
                }    
            }

            foreach ($name_X in $ScenChainSetList) 
            {
                $lcount2 = $name_X.chainSets.Count
                #Write-Output $lcount
                for($k=0;$k -lt $lcount2; $k++)
                {
                    $ScenarioID = $name_X.chainSets.GetValue($k).externalId
                    #Write-Output "Chain Set"
                    

                    # Download Scenario
                    $_url3x = "https://app.genrocket.com/rest/chainSet/download"
                    $JBody2 = '{"chainSetId": "#chainSetID"}'
                    $JBody2 = $JBody2.Replace("#chainSetID",$ScenarioID)
                    
                    $destfile = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\#sname.grs"
                    $destdir = "F:\GenRocketHome\genrocket\bin\Files\#Project\#version\"

                    $destdir = $destdir.Replace("#Project",$name.projects.GetValue($i).name )
                    $destdir = $destdir.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)

                    # Queue Check
                    $JBodyQueue = '{"chainSetId": "#ScenarioID"}'
                    $JBodyQueue = $JBodyQueue.Replace("#ScenarioID",$ScenarioID)
                    $ResponseQueue =""

                    while(($ResponseQueue.Trim() -NotLike "*fileReady*") -and ($ResponseQueue.Trim() -NotLike "*isReady*"))
                    {
                        #Write-Output $ResponseQueue
                        $ResponseQueue = (Invoke-WebRequest -UseBasicParsing "https://app.genrocket.com/rest/chainSet/verify" -Method Post -Body $JBodyQueue -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken}).Content
                        #Write-Output $ResponseQueue
                        Start-Sleep -s 1
                    }


                    # End Queue Check

                    #Write-Output $destfile
                    if (-Not(Test-Path -Path $destdir))
                    {
                        [system.io.directory]::CreateDirectory("$($destdir)")
                    }

                    $destfile = $destfile.Replace("#Project",$name.projects.GetValue($i).name )
                    $destfile = $destfile.Replace("#version",$name.projects.GetValue($i).projectVersions.GetValue($j).versionNumber)
                    $destfile = $destfile.Replace("#sname",$name_X.chainSets.GetValue($k).name)
                    #Write-Output $JBody2
                    #Write-Output $destfile
                    try
                    {
                        Invoke-WebRequest -UseBasicParsing $_url3x -Body $JBody2 -Method Post -ContentType "application/json" -Headers @{'Accept' = 'application/json';'x-auth-token' = $AccessToken} -OutFile $destfile
                    }
                    catch [System.Net.WebException] 
                    {
                         Write-Output $_.Exception.Response
                         Write-Output $($ResponseQueue)

                    }
                }    
            }
            # ########################################################
       }
    }    
}
