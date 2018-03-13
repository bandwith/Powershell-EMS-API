# Reproduction of this Gist for XML and using New-WebServiceProxy
# https://gist.github.com/jdmills-edu/b1467237644ed8383858aa0b9765b1e2
######################################################################################################################################################

Function Get-RoomsAndBookings {
# A PowerShell script that uses the EMS Software, LLC (formerly Dean Evans & Associates) EMS API to retrieve a complete list of all buildings, rooms, and bookings in those rooms for the next calendar year.

# Constants
$allBuildings = -1  # Building Code -1 is ALL buildings
$ViewComboRoomComponents = $false

# Set up the Date Variables
$today = Get-Date
$nextYear = $today.AddYears(1)

# EMS API Connection Variables
$url = "http://ems.yourdomain.com/EMSAPI/Service.asmx"
$username = "youremsapiuser"
$password = "youremsapipassword"

# Set up the EMS Connection
$EmsConnector = New-WebServiceProxy -Uri $url

# Get a Listing of all rooms in all buildings
[xml]$Rooms = $EmsConnector.GetAllRooms($username, $password, $allBuildings)
$RoomArray = $Rooms.Rooms.Data

# Get a Listing of all bookings in all rooms
$AllBookingsArray = @()
$RoomArray | ForEach-Object {
    $roomID = $_.ID
    [xml]$bookingsReturn = $EmsConnector.GetAllRoomBookings($username, $password, $today, $nextyear, $roomID, $ViewComboRoomComponents)
    $bookings = $bookingsReturn.Bookings.Data
    $AllBookingsArray += $bookings
}

return $AllBookingsArray

}

Get-RoomsAndBookings
