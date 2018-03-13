# Reproduction of this Gist for XML and using New-WebServiceProxy
# https://gist.github.com/jdmills-edu/6b8492ce34081eb31ed824c69af11dd9
######################################################################################################################################################

Function Get-BookingAtTime {
# A PowerShell script that uses the EMS Software, LLC (formerly Dean Evans & Associates) EMS API to return the list of bookings in a given room at a given time.
param(
    [string]$room,
    [string]$time
)

# Constants
$allBuildings = -1  # Building Code -1 is ALL buildings
$ViewComboRoomComponents = $false

# Set up the Date Variables
$dateAndTime = $time | Get-Date

## EMS API Connection Variables
$url = "http://ems.yourdomain.com/EMSAPI/Service.asmx"
$username = "youremsapiuser"
$password = "youremsapipassword"

# Set up the EMS Connection
$EmsConnector = New-WebServiceProxy -Uri $url

# Get a Listing of all rooms in all buildings
[xml]$Rooms = $EmsConnector.GetAllRooms($username, $password, $allBuildings)
$RoomArray = $Rooms.Rooms.Data

# Find the specific room ID we are looking for
$thisRoom = $RoomArray | Where-Object {$_.Room -like "*$room*"}
$roomID = $thisRoom.ID

# Get Bookings for this specific room from a specific day
[xml]$BookingArray = $EmsConnector.GetAllRoomBookings($username, $password, $dateAndTime, $dateAndTime, $roomID, $ViewComboRoomComponents)

# Narrow down the list to include only items that match our time parameter
if ($BookingArray.Bookings.Data) {
    $booking = $BookingArray.Bookings.Data | Where-Object { $dateAndTime -ge $([datetime]$_.TimeEventStart) -and $dateAndTime -lt $([datetime]$_.TimeEventEnd) }
    }
else { $booking = $null }

return $booking

}

Get-BookingAtTime -room "BH101" -time "8:30AM"
