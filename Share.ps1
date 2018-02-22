#this function uss a commandline tool to query the Distributed File System links and return the results in a PSobject.
Function Get-DFSLinkDetails{
    Param([Parameter(Mandatory=$True,Position=1)][string]$DFSlink)
    $dfslink= $dfslink.toupper() -replace "^G:\\","\\DFSSERVER\Dept\"
    $target=@()
    $target=@(dfsutil link ""$DFSlink"")
    if($target.count -eq 1)
        {
        if($target.trim() -match "^Could not execute the command successfully")
            {
            $Returnobject=New-Object PSObject -Property @{
                    LinkState="ZZZ_ELEMENT_NOT_FOUND"
                    Linktarget="ZZZ_ELEMENT_NOT_FOUND"
                    LinkName=$($target)
                    }
            }
        }
    else{
    foreach ($line in $target)
        {
        if($line.toupper().trim() -match "^DESCRIPTION:")
            {
                $Returnobject=New-Object PSObject -Property @{
                LinkState="ZZZ_Link had failed"
                Linktarget="ZZZ_Link had failed"
                LinkName=$("\\DFSSERVER\Dept\"+$LinkName)
                }
            }
        elseif($line.trim() -match "^Link Name")
            {
            $LinkName=$line.substring($line.indexof("Link Name=")+$("Link Name=").length+1,$line.indexof(" State=")-($line.indexof("Link Name=")+$("Link Name=").length+1)-1)
            }
        elseif ($line.trim() -match "^Target=" -and $line.trim() -match "State=")
            {
            $Linktarget=$line.substring($line.indexof("Target=")+$("Target=").length+1,$line.indexof(" State=")-($line.indexof("Target=")+$("Target=").length+1)-1)
            $linkstate=$line.substring($line.indexof("State=")+$("State=").length+1,$line.indexof("  [Site")-($line.indexof("State=")+$("State=").length+1)-1)
            
                $Returnobject=New-Object PSObject -Property @{
                LinkState=$linkstate
                Linktarget=$Linktarget
                LinkName=$("\\DFSSERVER\Dept\"+$LinkName)
                }
            }
        }
        }
        return $Returnobject
    }
#This function collects the subfolders using the built-in PS functions.
Function GetSubFolders{
    Param([Parameter(Mandatory=$True,Position=1)][string]$Targetfolder)
    
    Get-ChildItem $Targetfolder -force | ?{ $_.PSIsContainer } | Select-Object FullName
    $file= Get-ChildItem $Targetfolder -force| ?{ $_.PSIsContainer } | foreach {$_.FullName}
    $File
    }

#THis is a logger function. It throws the message on the screen, immediately writes it out into a file and produces a completes file at the end of the session 
# In order to solve threadsafety, I made it hold the messages until the filei s 
Function LOGME{
    Param(
    [Parameter(Mandatory=$True,Position=1)][string]$Message,
    [Parameter(Mandatory=$False,Position=2)][Bool]$printscreen = $true
    )
    $date=get-date
    If(-not $script:logdata){
        $script:logdata=@()
        }
    $IsErrorMsg=$false
    if ($message.length -ge 5)
        {
            if ($Message.substring(0,5) -eq "ERROR") {$IsErrorMsg=$true}
        }
    if ($printscreen -eq $true)
    {
        if ($IsErrorMsg)
            {
              Write-Host $Date.ToString("yyyy-MM-dd hh:mm:ss")" - $Message" -Backgroundcolor "black" -foregroundcolor  "Red"
            }
            
        Else{
            Write-Host $Date.ToString("yyyy-MM-dd hh:mm:ss")" - $Message" }
    }
    $script:logdata+="$($Date.ToString('yyyy-MM-dd hh:mm:ss')) - $Message"
    $script:fax+="$($Date.ToString('yyyy-MM-dd hh:mm:ss')) - $Message"
    #FOFP is the FolderOutputFilePath. 
    If (@($Null, "") -notcontains $script:FOFP)
        {
        Try
        {$script:fax| out-file -filepath "$($Script:FOFP)\RollingLogMe.txt" -append  -force -ea stop
        $script:fax=@()
        }
        catch 
            {
            }
        }
    else
        {
        }
    }

#This function uses the return an rmtshare.exe to query the attributes of the file shares. Rather quite similar to the DFS query one. Lots of parsing and flow control.
Function Get-ShareDetails{
    Param([Parameter(Mandatory=$True,Position=1)][array][string]$TargetShares)
            $script:Valid=@()
            $script:goodshares=@()
            $script:BadShares=@()
            $script:Badlines=@()
            $script:Noshares=@()
            foreach ($_ in $TargetShares)
                    {
                    [array]$Results =.\rmtshare.exe  $_
                    if (@($null,"") -contains $results)
                        {$status+="The command failed rmtshare.exe was blocked by system policies."
                        $ShareTempObject=New-Object PSObject -Property @{
                                Status=$Status
                                SharePath=$_
                                Path=$Pathline
                                Remark=$RemarkLine
                                MaxUser=$MaximumLine
                                CurrentUsers=$Usersline
                                DFSLink=$Null
                                AdminSharePath=$Null				
                                }
                        
                        Continue}
                    elseIf ($Results -match "^The command failed")
                        {
                        $ShareTempObject=New-Object PSObject -Property @{
                                Status=$results[0]
                                SharePath=$_
                                Path=$Pathline
                                Remark=$RemarkLine
                                MaxUser=$MaximumLine
                                CurrentUsers=$Usersline
                                DFSLink=$Null
                                AdminSharePath=$Null				
                                }
                        continue
                        }
                    Else
                    {
                    :Outerloop	
                    Foreach ($line in $results)
                        {if ($line -match "^Share name\w*")
                            {
                            $Startline=$line
                            $Endline=$null
                            $Pathline=$null
                            $RemarkLine=$Null
                            $MaximumLine=$Null
                            $UsersLine=$Null
                            }
                        elseif($line -match "^Path\w*")
                            {$Pathline=($Line -replace "Path","").trim()
                            }
                        elseif($line -match "^Remark\w*")
                            {$RemarkLine=($Line -replace "Remark","").trim()
                            }	
                        elseif($line -match "^Maximum users\w*")
                            {$MaximumLine=($Line -replace "Maximum users", "").trim()
                            }	
                        elseif($line -match "^Users\w*")
                            {$UsersLine=($Line -replace "Users", "").trim()
                            }	
                        elseif($line -match"^The command completed successfully.\w*")
                            {
                            $Endine=$Line
                            If ([array]::IndexOf($results, $Line) -eq [array]::IndexOf($results, $PermissionLine)+1)
                                {$script:Noshares+=$_
                                $PermissionIdentities=$null
                                }
                                                
                            $ShareTempObject=New-Object PSObject -Property @{
                                SharePath=$_
                                status=$line
                                Path=$Pathline
                                Remark=$RemarkLine
                                MaxUser=$MaximumLine
                                CurrentUsers=$Usersline
                                DFSLink=$Null
                                AdminSharePath=$Null
                                ADIdentities=$PermissionIdentities
                                }
                                $Returnvar=$ShareTempObject
    
                            }
                        elseif($line -match "^Permissions:\w*")
                            {$Endine=$Line
                            $Startline=$Null
                            $PermissionLine=$Line
                            }
                        elseif ($line -match "No permissions specified.")
                            {
                            $PermissionIdentities="No Permissions Specified."
                            $ShareTempObject=New-Object PSObject -Property @{
                                SharePath=$_
                                Path=$Pathline
                                Remark=$RemarkLine
                                MaxUser=$MaximumLine
                                CurrentUsers=$Usersline
                                ADIdentities=$PermissionIdentities
                                RightType=$Null
                                DFSLink=$Null
                                AdminSharePath=$Null
                                }
                            Returnvar=$ShareTempObject						
                            $script:Noshares+=$_
                            }
                        else{
                            $script:Badlines+=$Line
                            }
                        }
                    }
                    
                        
                    }
    $returnvar = $ShareTempObject
    Return $returnVar
    }

 
    


#this function uss a commandline tool to query the Distributed File System links and return the results in a PSobject.
Function Get-DFSLinkDetails{
    Param([Parameter(Mandatory=$True,Position=1)][string]$DFSlink)
    $dfslink= $dfslink.toupper() -replace "^G:\\","\\DFSSERVER\Dept\"
    $target=@()
    $target=@(dfsutil link ""$DFSlink"")
    if($target.count -eq 1)
        {
        if($target.trim() -match "^Could not execute the command successfully")
            {
            $Returnobject=New-Object PSObject -Property @{
                    LinkState="ZZZ_ELEMENT_NOT_FOUND"
                    Linktarget="ZZZ_ELEMENT_NOT_FOUND"
                    LinkName=$($target)
                    }
            }
        }
    else{
    foreach ($line in $target)
        {
        if($line.toupper().trim() -match "^DESCRIPTION:")
            {
                $Returnobject=New-Object PSObject -Property @{
                LinkState="ZZZ_Link had failed"
                Linktarget="ZZZ_Link had failed"
                LinkName=$("\\DFSSERVER\Dept\"+$LinkName)
                }
            }
        elseif($line.trim() -match "^Link Name")
            {
            $LinkName=$line.substring($line.indexof("Link Name=")+$("Link Name=").length+1,$line.indexof(" State=")-($line.indexof("Link Name=")+$("Link Name=").length+1)-1)
            }
        elseif ($line.trim() -match "^Target=" -and $line.trim() -match "State=")
            {
            $Linktarget=$line.substring($line.indexof("Target=")+$("Target=").length+1,$line.indexof(" State=")-($line.indexof("Target=")+$("Target=").length+1)-1)
            $linkstate=$line.substring($line.indexof("State=")+$("State=").length+1,$line.indexof("  [Site")-($line.indexof("State=")+$("State=").length+1)-1)
            
                $Returnobject=New-Object PSObject -Property @{
                LinkState=$linkstate
                Linktarget=$Linktarget
                LinkName=$("\\DFSSERVER\Dept\"+$LinkName)
                }
            }
        }
        }
        return $Returnobject
    }
    
    
    
    
    # This function used a webservice to query data of service IDs. 
    function get-fidDetails {
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true, position=1)][string]$fid,
    [Parameter(Mandatory=$False, position=2)][string]$Domain
    )
    #initiall
    
        $hasMain=@()
        [string]$uri="https://mywebservice.com/doLookup.jsp?svc=ServiceName=xml&fidName=$fid&systemName=$Domain&systemId=&fidOwner="
        #web Service Query#
        $PS2Request = new-object System.Net.WebClient
        $PS2Request.UseDefaultCredentials = $true
        $web_response = $PS2Request.downloadstring($uri) 
    
        #convert output to XML
        [xml]$wyniki = $web_response
        $XOB = New-Object psobject
        If (@($Wyniki.resultroot.resultlist.result|?{$_.name -eq $FID}).count -eq 1)
        {
        $Wyniki.resultroot.resultlist.result|?{$_.name -eq $FID}|%{$_.exresult.ex}|select attr, value|%{Add-Member -InputObject $XOB -Type NoteProperty -Name $_.Attr -Value $_.value}
        }
        ELSEIf (@($Wyniki.resultroot.resultlist.result|?{$_.name -eq $FID}).count -eq 0)
        {}
        elseif (@($Wyniki.resultroot.resultlist.result|?{$_.name -eq $FID}).count -gt 1)
            {
            $OK=@($Wyniki.resultroot.resultlist.result|?{$_.name -eq $FID}|?{$_.exresult.ex.attr -eq "systemname" -and $_.exresult.ex.value -eq "EUR"})
            
            $OK|%{$_.exresult.ex}|select attr, value|%{Add-Member -InputObject $XOB -Type NoteProperty -Name $_.Attr -Value $_.value}
            }
            else {
                  }
    
    #$dane = $wyniki.ResultRoot.ResultList.Result.ExResult
    return $XOB
    }
        
#This function is activated if the script received the $useini file. It is a simplified function that walks trough the lines, looks for variable assignents and if found, Create\update the variable in the script's scope.  
if($useini){
    if ($(test-path $ParameterPSfile -ea 'silentlycontinue') -eq $true)
        {	
            try{
                $baseparameters=get-ini $ParameterPSfile "[BASEPARAMETERS]"
            foreach ($line in $baseparameters)
                {
                $NewVarValue=$null
                $newvarname=$null
                if(@($null, "") -notcontains $Line.trim()) 
                    {
                    if ($line.trim().substring(0,1) -eq "$" -and $line.trim() -match "=")
                        {
                        $newvarname=$($line.trim().substring(1,$line.trim().indexof("=")-1)).trim()
                        #variables
                        $NewVarValue=$($line.trim().substring($line.trim().indexof("=")+1)).trim()
                        If ($newvarvalue -match "\`"\S+\`"")
                        {#write-host "$line is a string."
                        $Newvarvalue=$newvarvalue.substring(1,$newvarvalue.length-2)
                        }
                        elseif($newvarvalue -match "\`$\S+$")
                            {
                            If (@('$true','$False','$Null','True','False') -contains $newvarvalue)
                                {#write-host "$newvarname is recognised true/false/boolean"
                                $Newvarvalue =[System.Convert]::ToBoolean($($newvarvalue -replace "\`$",""))}            
                            }
                        else
                        {LOGME "$line is unknown type."}            
    
                        }
                    else{
                         LOGME ("$line is bad")
                        }
                    }
                            if(Test-Path variable:script:$($line.trim().substring(1,$line.trim().indexof("=")-1)))
                            {#write-host "Variable $newvarname found "
                            $(Get-Variable -Name "$newvarname").value=$newvarvalue}
                        elseif(-not (Test-Path variable:script:$($line.trim().substring(1,$line.trim().indexof("=")-1))))
                            {#write-host "Variable $newvarname doesn't found. Creating "
                            $testvar=New-Variable -Name $line.trim().substring(1,$line.trim().indexof("=")-1) -Value $newvarvalue
                            }
                            $testvar
                }
            $temparray=@()
            $Temparray=Get-Ini $ParameterPSfile "[CUSTOMFOLDERLIST]"
            if ($temparray.count -gt 0)
                {
                LOGME "Customfolderlist is not empty"
                $CUSTOMFOLDERLIST=$temparray
                }
                    $temparray=@()
            $Temparray=Get-Ini $ParameterPSfile "[SWITCHSHARESERVERLIST]"
            if ($temparray.count -gt 0)
                {
                LOGME " temparray is large"
                $servernames=$temparray
                }	
            }
            catch{
            logme "failed to load the parameterfile $ParameterPSfile. Exiting"
            sendmeanemail
            exit		
            }		
        }
    else
        {
        LOGME "ERROR: $ParameterPSfile was not found.Exiting"
        sendmeanemail
        exit
        }
    }    
 #this attempt tried to get folder sizes using the old DIR command. 
 Function Get-Size {
    Param([Parameter(Mandatory=$True,Position=1)][string]$Targetpath)
    $returnarray+=cmd.exe /c dir $Targetpath  /-c /s /d
    $boofound = $false
    [double]$Tempvar = $null
    foreach ($Line in $returnarray)
    {
    if  ($boofound -eq $true) {
    $start="File(s)"
    $finish = "bytes"
    $cleared=$Line -replace " "
    $TempVar= ($cleared.substring($cleared.indexof($start)+$start.length,($cleared.indexof($finish)-($cleared.indexof($start)+$start.length))))/1024/1024/1024
    break
    }
        if($line -match "Total Files Listed:"){
            $boofound = $true
        }
    }
    #[bool]($tempvar -as [double] -is [double])
    #$tempvar.gettype()
    return ,$tempvar
    }
        


# Several of the tasks required to get verious types of user data. This query below a tag in the users's name that referred to the organisational unit according to a certain type of hierarchy and level. 
Function Get-Universal-LDAPResults{
    [CmdletBinding()]
    param(
         [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [string[]] $SOEID
    )
    
    $strFilter = "(&(objectCategory=User)(Name=$SOEID))"
    
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = $strFilter
    $objSearcher.SearchScope = "Subtree"
    
    $colProplist = "name", "extensionattribute4", "extensionattribute5","displayname","samaccountname"
    
    foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)| Out-Null} 
    
    $colResults = $objSearcher.FindAll()
    
    foreach ($objResult in $colResults)
        {$objItem = $objResult.Properties
        $tagstartloc=$($objItem.displayname).tostring().lastindexof("[")
        if ($tagstartloc -ne -1)
        {
        $taghypenloc=$($objItem.displayname).tostring().lastindexof("-")
        $tagendloc=$($objItem.displayname).tostring().lastindexof("]")
        $tag=$($objItem.displayname).tostring().substring($tagstartloc+1,$tagendloc-$tagstartloc-1)
        switch ($tag.tostring().tolower()){
        "ch-lcl" {
            $tag="lcl"
            }
        default {
        
            if ($taghypenloc -ne -1 -and $taghypenloc -gt $tagstartloc)
                {	
                $tagendloc=$taghypenloc
                $tag=$($objItem.displayname).tostring().substring($tagstartloc+1,$tagendloc-$tagstartloc-1)
                }
            }
        }
        }
        }
     
    Return 	$tag
    }