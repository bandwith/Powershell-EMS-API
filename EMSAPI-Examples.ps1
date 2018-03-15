<#
Random group of notes on 
EMS API created by EMS Software, LLC (formerly Dean Evans & Associates)

Note: RiseVision has a Dean Evans EMS Widget https://github.com/Rise-Vision/dean-evans
For Visix https://oit-axis-server.ohio.edu/Help/Dean_Evans.htm  

Can be used as XML or JSON. I stuck with XML as that is what our Xenegrade connection needed.


#>

# EMS API Connection Variables
$url = "http://ems.yourdomain.com/EMSAPI/Service.asmx"
$username = "youremsapiuser"
$password = "youremsapipassword"

$EmsConnector = New-WebServiceProxy -Uri $url

[xml]$ApiVersion = $EmsConnector.GetAPIVersion()
$ApiVersion.API.APIVersion

## Store in Arrays these items that are needed for browsing Bookings

[xml]$Statuses = $EmsConnector.GetStatuses($username, $password)
$StatusArray = $Statuses.Statuses.Data.id

[xml]$EventTypes=$EmsConnector.GetEventTypes($username, $password)
$EventTypeArray = $EventTypes.EventTypes.Data.id

[xml]$GroupTypes = $EmsConnector.GetGroupTypes($username, $password)
$GroupArray = $GroupTypes.GroupTypes.Data.id

[xml]$Buildings = $EmsConnector.GetBuildings($username, $password)
$BuildingArray = $Buildings.Buildings.Data.Id

# Store in Array the EMS Room IDs converted to our Naming
[xml]$AllRooms = $EmsConnector.GetAllRooms($username, $password, $allbuildings)
$global:RoomTranslation = $AllRooms.Rooms.Data | Select-Object -Property Id, Room

## Functions

[xml]$RoomDetail = $EmsConnector.GetRoomDetails($username, $password, 21)  #Integer of the Room
$RoomDetail

# Get a Specific Room's bookings
$roomid = 21
$startdate = "1/1/2018"
$enddate = "2/26/2018"
[xml]$RoomBookings = $EmsConnector.GetRoomBookings($username, $password, $startdate, $enddate, $roomid, $StatusArray, $EventTypeArray, $GroupArray, $false)

#$RoomBookings.Bookings.Data | Format-Table -Property RoomCode, EventName, TimeEventStart, TimeEventEnd, GroupName

# Determine if a Room is available
$roomId = 21
$bookingdate = "1/26/2018"
$starttime = "1:00 PM"
$endtime = "2:00 PM"
[xml]$IsAvailable = $EmsConnector.GetRoomAvailability($username, $password, $roomid, $bookingdate, $starttime, $endtime)
$IsAvailable

# Add a Reservation without a User Name or event type attached.
$groupID = 2 #aka On Campus  (May want to do one for Credit Class)
$statusID = 1 #aka Confirmed
$eventname = "Scott Test Rez"
$bookingdate = "1/26/2018"
$starttime = "5:00 PM"
$endtime = "6:00 PM"
$roomId = 21
[xml]$AddReserverationResult = $EmsConnector.AddReservation($username, $password, $groupid, $roomid, $bookingdate, $starttime, $endtime, $eventname, $statusID)
$AddReserverationResult

# Get a Web Users Email and EMS ID Number
# This may at first blush look like you can't run it, however you can leave several of the parameters blank.
[xml]$CurrentEMSUser = $EmsConnector.GetWebUsers($username, $password, "Scott Hardwick", "", "", "")
$EmsUserId = $CurrentEMSUser.WebUsers.Data.ID

# Add a Reservation WITH a User Name and event type attached.
# Can do it with both of the last two parameters blank.  If put in a EmsUserId, then the WebTemplate must be 1 or 2 in our system. Can use a method to get the WebTemplate possible values.
$groupID = 2 #aka On Campus  (May want to do one for Credit Class)
$statusID = 1 #aka Confirmed
$eventname = "Scott Test Rez"
$bookingdate = "1/26/2018"
$starttime = "5:00 PM"
$endtime = "6:00 PM"
$roomId = 21
$WebTemplate = "1"
#[xml]$Reservation = $EmsConnector.AddReservation2($username, $password, $groupId, $roomId, $bookingdate, $starttime, $endtime, $eventName, $statusId, $eventTypeId, "", "")
[xml]$Reservation = $EmsConnector.AddReservation2($username, $password, $groupId, $roomId, $bookingdate, $starttime, $endtime, $eventName, $statusId, $eventTypeId, $EmsUserId, $WebTemplate)


<#
<Reservation>
  <Data>
    <ReservationID>4844</ReservationID>
    <BookingID>80050</BookingID>
  </Data>
</Reservation>
#>

# Add in Reference Number for Lookup
[xml]$UDFResult = $EmsConnector.AddUDF($username, $password, 0,$reservationid, $udfid, $udfvalue)


# Get Bookings
$reservationid = 4844
$startdate = "1/1/2018"
$enddate = "12/31/2018"
$buildings = -1
[xml]$SemesterBookings = $EmsConnector.GetBookings2($username, $password, $reservationid, $startdate, $enddate, $buildings, $StatusArray, $EventTypeArray, $GroupArray, $false)
$semesterbookings

# Add a Booking
$reservationid = 4844
$roomId = 21
$bookingdate = "1/27/2018"
$starttime = "5:00 PM"
$endtime = "6:00 PM"
$eventname = "Scott Test Rez"
$statusID = 1 #aka Confirmed
[xml]$AddBookingResult = $EmsConnector.AddBooking($username, $password, $reservationid, $roomid, $bookingdate, $starttime, $endtime, $eventname, $statusid)
$AddBookingResult



[xml]$UDFdefs = $EmsConnector.GetUDFDefinitions($username, $password)


# Other Attempts

Your Reservation ID is 82137 and your booking ID is 4966
 $EmsConnector.GetBooking($username, $password, "82137")

 $EmsConnector.UpdateBooking2($username,$password,"82137","","","","","","New Name")
 $EmsConnector.UpdateBooking2($username,$password,"82137",$null,$null,$null,$null,$null,"New Name")
 $EmsConnector.UpdateReservation($username, $password, "4966", "82137", "51", "1", "","")

 #Missing
 # Given a Reference Number, return all the bookings associated with it.
 # Given a Reference Number, ability to change Status to Cancelled (or other) -- Instead you have to cancel each booking.
 # Given a Reference Number, ability to cancel all of its related bookings. (may not change status, as you may want to put in new bookings)


  82145
  #Cancelled = 3

# Find all bookings associated with a Reservation Number
$EmsConnector.GetBookings2($username, $password, "4970", "1/1/2017", "1/1/2099", $BuildingArray, $StatusArray, $EventTypeArray, $GroupArray, "true")
# Cancel a booking
$EmsConnector.UpdateBooking($username, $password, "82145", "1/1/2017", "00:00", "00:00","3", "")

$EmsConnector.GetBookings2($username, $password, "4970", "1/1/2017", "1/1/2099", "0,1", $StatusArray, $EventTypeArray, $GroupArray, "true")


<#

Searching for existing stuff

On the desktop client, you can use filters to do all kinds of searches, by group, by UDF Reference Number, by UDF Semester, or any combination of.
On the desktop client, you can filter out by EventType from many reports

***Via API, Can I search for all bookings from a group?   (DOESNT LOOK LIKE) NEED THIS.
***Via API, Can I search for all bookings from a group AND a UDF filter? (No) Might be useful to filter down to Integrow group and spring2018 semester.
***Via API, When I pull getallbookings2 (or really any of these), would like to return the UDFs (including the one I searched for).  This would allow me to search for all Integrow,
    all a specific semester, and all a specific ref num (if all are set up as UDFs)


# Determine if a Room is available
$roomId = ($RoomTranslation | Where-Object { $_.Room -eq "BE203" }).ID
$bookingdate = "1/26/2018"
$starttime = "1:00 PM"
$endtime = "2:00 PM"
[xml]$IsAvailable = $EmsConnector.GetRoomAvailability($username, $password, $roomid, $bookingdate, $starttime, $endtime)
$IsAvailable.RoomAvailability.Data.Available

# Look up all bookings for a specific Reservation Id
$reservationid = 4844
[xml]$SemesterBookings = $EmsConnector.GetBookings2($username, $password, $reservationid, $searchstartdate, $searchenddate, $allBuildings, $StatusArray, $EventTypeArray, $GroupArray, $false)
$semesterbookings.Bookings.Data | Format-Table -Property ReservationID, BookingID, UDF, EventName, RoomCode, @{expression={[datetime]$_.TimeBookingStart};label="BookingStart"}, @{expression={[datetime]$_.TimeBookingEnd};label="BookingEnd"}
$semesterbookings.Bookings.Data.Count


# Look up all bookings for a specific Reference Number
$referencenumber = "123456789"
[xml]$RefLookup = $EmsConnector.GetAllBookings2($username, $password, $searchstartdate, $searchenddate, $allBuildings, $false, $udfRefNum, $referencenumber)
$RefLookup.Bookings.Data | Format-Table -Property ReservationID, BookingID, UDF, EventName, RoomCode, @{expression={[datetime]$_.TimeBookingStart};label="BookingStart"}, @{expression={[datetime]$_.TimeBookingEnd};label="BookingEnd"}
$RefLookup.Bookings.Data.Count


# Room Sign Type Info
$roomId = ($RoomTranslation | Where-Object { $_.Room -eq "AD215" }).ID
$startdate = "1/28/2018"
$enddate = "1/28/2018"
$starttime = "1:00 PM"
$endtime = "2:00 PM"
[xml]$RoomData = $EmsConnector.GetRoomBookings($username, $password, $startdate, $enddate, $roomid, $StatusArray, $EventTypeArray, $GroupArray, $false)
$RoomData.Bookings.Data | Format-Table -Property ReservationID, BookingID, UDF, EventName, RoomCode, @{expression={[datetime]$_.TimeBookingStart};label="BookingStart"}, @{expression={[datetime]$_.TimeBookingEnd};label="BookingEnd"}
$RoomData.Bookings.Data.Count


List of possibly useful Methods:

GetAPIVersion
GetStatuses
GetBuildings
GetAllRooms -- List by Building -1 for all buildings
GetRooms -- Returns list of rooms for multiple buildings

GetRoomBookings
GetRoomDetails

GetAllBookings -- List by Building (startdate and end date filter)  -1 for all buildings.
GetAllBookings2 -- List by UDF Value (startdate, enddate, and building filters)
GetAllRoomBookings -- List by Room (startdate and end date filter)

GetBookings -- svc.GetBookings(“UserID”, “Password”, Date.Today, Date.Today, strBuildings, strStatuses, strEventTypes, strGroupTypes, False))

GetRoomAvailability -- based on a date/time/room, returns a yes/no the room is available.  GetRoomsAvailable(“UserID”, “Password”, 3199, Date.Today, CDate(“1/1/1900 8:00 AM”), CDate(“1/1/1900 9:00 AM”)))

AddReservation
svc.AddReservation(“UserID”, “Password”, 11, 3199, Date.Today, CDate(“1/1/1900 8:00 AM”), CDate(“1/1/1900 9:00 AM”),”Meeting”,1))

AddBooking (Requires EMS API Advanced License) -- AddBooking(“UserID”, “Password”, 1234, 3199, Date.Today, CDate(“1/1/1900 8:00 AM”), CDate(“1/1/1900 9:00 AM”),”Meeting”,1))
AddBooking2 (Requires EMS API Advanced License)  -- this one adds an event type and user to the booking.

UpdateBooking (Requires EMS API Advanced License)
UpdateBooking2 (Requires EMS API Advanced License)
GetChangedBookings
GetBookingHistory

GetBookings2 -- requires reservation id to return.
GetBookings3 -- requires a date range
GetBooking -- based on specific booking id  GetBooking(“UserID”, “Password”, 222008))

GetRoomsAvailable
GetRoomsAvailable2
GetRoomsAvailable3
GetRoomsAvailable4

#>