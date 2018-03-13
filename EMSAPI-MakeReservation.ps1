# Example of Checking Availability and Creating a Reservation with a single Booking in EMS
######################################################################################################################################################

Clear-Host

# EMS API Connection Variables
$url = "http://ems.yourdomain.com/EMSAPI/Service.asmx"
$username = "youremsapiuser"
$password = "youremsapipassword"

## Constants
$global:groupCE        = 1 # Group 1 is Continuing Education
$global:confirmed      = 1   # Reservation Status 1 is Confirmed
$global:allBuildings   = -1  # Building Code -1 is ALL buildings
$global:noncreditevent = 36 # Event Type 36 is Class-NonCredit
$global:WebTemplate    = "1"  # The standard Everyday User Template

# Reservation and Booking Variables
[string]$eventName = "Scott's Test Event"
[string]$room = "AD215"
[string]$bookingdate = "3/12/2018"
[string]$starttime = "10:00 PM"
[string]$endtime = "11:00 PM"
[string]$groupId = $groupCE 
[string]$eventTypeId = $noncreditevent
[string]$statusId = $confirmed

# Set up the EMS Connection
$EmsConnector = New-WebServiceProxy -Uri $url

# Store in Array the EMS Room IDs converted to our Naming
[xml]$AllRooms = $EmsConnector.GetAllRooms($username, $password, $allbuildings)
$global:RoomTranslation = $AllRooms.Rooms.Data | Select-Object -Property Id, Room

# Check Room Availability
[datetime]$bookingdateConverted = $bookingdate
[string]$roomId = ($RoomTranslation | Where-Object { $_.Room -eq $room }).ID
[xml]$IsAvailableReturn = $EmsConnector.GetRoomAvailability($username, $password, $roomid, $bookingdateConv, $starttime, $endtime)
$IsAvailable = $IsAvailableReturn.RoomAvailability.Data.Available

# Get a Web Users Email and EMS ID Number Based on their Name
[xml]$CurrentEMSUser = $EmsConnector.GetWebUsers($username, $password, "Scott Hardwick", "", "", "")
$EmsUserId = $CurrentEMSUser.WebUsers.Data.ID

# Create the Reservation and First Booking (Additional Bookings would need to be done via separate calls to AddBooking)
[xml]$Reservation = $EmsConnector.AddReservation2($username, $password, $groupId, $roomId, $bookingdateConverted, $starttime, $endtime, $eventName, $statusId, $eventTypeId, $EmsUserId, $WebTemplate)
$ReservationId = $Reservation.Reservation.Data.ReservationID
$BookingId = $Reservation.Reservation.Data.BookingID
if (!($Reservation.Errors.Error)) {Write-Host "Your Reservation ID is $ReservationID and your booking ID is $BookingId" }
else {Write-Host "Error: $($Reservation.Errors.Error.Message) "}



