
## One Time Use Program to Download all of the existing rooms from EMS and match them up against Continuing Ed's Xenegrade spreadsheet.
## Add in Missing Rooms and their EMS ID Number/Room Name, but also put the EMS ID Number/Room Name on rooms that already exist in the Xenegrade sheet.


Clear-Host

## EMS API Connection Variables
$url = "http://ems.yourdomain.com/EMSAPI/Service.asmx"
$username = "youremsapiuser"
$password = "youremsapipassword"

## User Changeable Items
$xenegradeincomingfilename = "C:\Users\shardwic\Desktop\EMS\LSUS-Rooms-Xenegrade.csv"
$xenegradeoutgoingfilename = "C:\Users\shardwic\Desktop\EMS\LSUS-Rooms-Xenegrade-corrected2.csv"
$global:searchstartdate = "1/1/2018" #Earliest potential date - used in searching to see if has already been booked
$global:searchenddate = "12/31/2018" #Last potential date - used in searching to see if has already been booked

## Constants
$global:allBuildings  = -1  # Building Code -1 is ALL buildings

$EmsConnector = New-WebServiceProxy -Uri $url

[xml]$ApiVersion = $EmsConnector.GetAPIVersion()
Write-Host "EMS Version" $ApiVersion.API.APIVersion.Version $ApiVersion.API.APIVersion.License
Write-Host ""

## Store in Arrays these items that are needed for browsing Bookings

[xml]$Statuses = $EmsConnector.GetStatuses($username, $password)
$StatusArray = $Statuses.Statuses.Data.id

[xml]$EventTypes=$EmsConnector.GetEventTypes($username, $password)
$EventTypeArray = $EventTypes.EventTypes.Data.id

[xml]$GroupTypes = $EmsConnector.GetGroupTypes($username, $password)
$GroupArray = $GroupTypes.GroupTypes.Data.id

# Store in Array this item to be able to use as a room lookup

[xml]$AllRooms = $EmsConnector.GetAllRooms($username, $password, $allbuildings)
$global:RoomTranslation = $AllRooms.Rooms.Data | Select-Object -Property Id, Room, Description, Building


## Class File Import
$XenegradeHeader = "EMSRoomID","EMSRoomCode","rooID","rooRoom","facID","facName","buiID","buiName","Location"
$XenegradeFile = Import-CSV -Path $xenegradeincomingfilename -Header $xenegradeHeader


$XenegradeCorrected = @()

$RoomTranslation | ForEach-Object {
    
    $CurrentID = $_.Id
    $CurrentRoom = $_.Room
    $CurrentDesc = $_.Description
    $CurrentBldg = $_.Building
    $CurrentBldgNumber = $null

    if ($CurrentDesc.StartsWith("SC")) { $CurrentDesc = $CurrentDesc.Replace("SC", "Science Building,")
    $CurrentBldgNumber = "9"
     }
    if ($CurrentDesc.StartsWith("TC")) { $CurrentDesc = $CurrentDesc.Replace("TC", "Technology Center,") 
    $CurrentBldgNumber = "2"
    }
    if ($CurrentDesc.StartsWith("BH")) { $CurrentDesc = $CurrentDesc.Replace("BH", "Bronson Hall,") 
    $CurrentBldgNumber = "3"
    }
    if ($CurrentDesc.StartsWith("NL")) { $CurrentDesc = $CurrentDesc.Replace("NL", "Noel Library,") }
    if ($CurrentDesc.StartsWith("BE")) { $CurrentDesc = $CurrentDesc.Replace("BE", "Business Education,") 
    $CurrentBldgNumber = "4"
    }

    $XenObject = $XenegradeFile | Where-Object {$_.EMSRoomCode -eq $CurrentRoom}

    if ($XenObject -ne $null) { 
            #Write-Host $CurrentRoom "is in the xenegrade file. Add the code $CurrentID" 

            $RoomProperties = [pscustomobject]@{
                EMSRoomID = $CurrentID  
                EMSRoomCode  = $CurrentRoom         
                rooID     = $XenObject.rooID     
                rooRoom   = $XenObject.RooRoom    
                facID     = $XenObject.FacId     
                facName   = $XenObject.FacName        
                buiID     = $XenObject.BuiID
                buiName   = $XenObject.BuiName
                Location  = $XenObject.Location
              }

            $XenegradeCorrected += $RoomProperties
            
            }
    else {
             #Write-Host $CurrentRoom "IS MISSING FROM XENEGRADE FILE. Add Record with Code $CurrentID" 

            $RoomProperties = [pscustomobject]@{
                EMSRoomID = $CurrentID  
                EMSRoomCode  = $CurrentRoom         
                rooID     = $null
                rooRoom   = $null  
                facID     = "27"  
                facName   = "LSU Shreveport"       
                buiID     = $CurrentBldgNumber
                buiName   = $CurrentBldg
                Location  = ("LSU Shreveport, " + $CurrentDesc)
              }

            $XenegradeCorrected += $RoomProperties

        }


  
}


$xenegradecorrected | Format-Table -Property EMSRoomID, EMSRoomCode, rooId, rooRoom, Location, buiID
$XenegradeCorrected | Export-Csv -Path $xenegradeoutgoingfilename