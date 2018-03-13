## EMS API CODE TO IMPORT CLASSES
## Incomplete code sample
##
## Created when I initially couldn't get the Academic Import utility within EMS to work.
## This script uses several of the API Methods.
## The trick in one case is that while these methods have multiple parameters, you can often just send a null in for the value and they happily go 
## along (even though the method documentation says they are required.)

#### TODOS
# TODO: Courses with Multiple Day lines. Each coming in separately. But ref number would appear twice in UDF. Do we really need to use the UDFs at all. Could
#       just put the ref number in the Event Name like the Academic Import utility does.
# TODO: Cross listing. Could use Academic Import logic and put in these classes instead of with Confirmed but with Academic-Confirmed. If the response comes back
#       that the room/date/time is unavailable to get the booking and if it is an Academic-Confirmed booking then redo the booking with an Academic-CrossList status instead.

function Test-DateAcceptable
{
# Ensure that date is not in the holiday array and IS one of the pattern dates
param(
    [Parameter(ValueFromPipeline=$true)][datetime]$incomingdate,
    [array]$holidayArray,
    [array]$dayArray
    )

    if ($dayarray -contains $incomingdate.DayOfWeek) {
        if ($holidayarray -notcontains $incomingdate.ToShortDateString() ) {
            return $true
            }
    }

}

function Find-DaysToBook
{
param(
    [datetime]$startdate,
    [datetime]$enddate,
    [array]$holidayArray,
    [array]$dayArray
    )

[array]$datestobook = @()    
while($startDate -le $endDate){
    [datetime]$nextDate = $startDate
    if ($nextDate | Test-DateAcceptable -dayArray $dayarray -holidayArray $holidayarray) {$datestobook += $nextdate}
    $startDate = $startDate.AddDays(1)
}

    return $datestobook
}	

function Find-Reservation {
    param(
        [Parameter(ValueFromPipeline=$true)][string]$integrowRefNumber
        )

    [xml]$RefLookup = $EmsConnector.GetAllBookings2($username, $password, $searchstartdate, $searchenddate, $allbuildings, $false, $udfRefNum, $integrowRefNumber)
    #$RefLookup.Bookings.Data | Format-Table -Property ReservationID, BookingID, UDF, EventName, RoomCode, @{expression={[datetime]$_.TimeBookingStart};label="BookingStart"}, @{expression={[datetime]$_.TimeBookingEnd};label="BookingEnd"}, Teacher
    if ($RefLookup.Bookings.Data.Count -gt 0) { return $true }
    else { return $false }
}

function Create-Reservation {
param(
    [string]$eventname, #255 max chars
    [string]$bookingdate,
    [string]$starttime,
    [string]$endtime,
    [string]$room,
    [string]$referencenumber,
    [string]$semester
    )

$roomId = ($RoomTranslation | Where-Object { $_.Room -eq $room }).ID

# Add a Reservation
## NOTE: While the spec says WebUserID and WebTemplateID are required for AddReservation2, you can just leave them blank
[xml]$AddReserverationResult = $EmsConnector.AddReservation2($username, $password, $groupIntegrow, $roomId, $bookingdate, $startTime, $endTime, $eventName, $confirmed, $creditevent, "", "")
$reservationid = $AddReserverationResult.Reservation.Data.ReservationID

# Add an SIS Reference Number to a Reservation
[xml]$AddRef = $EmsConnector.AddUDF($username, $password, 0, $reservationid, $udfRefNum, $referencenumber)
#$AddRef.UDF.Data.UDFID

# Add an SIS Semester to a Reservation
[xml]$AddSemester = $EmsConnector.AddUDF($username, $password, 0, $reservationid, $udfSemester, $semester)
#$AddSemester.UDF.Data.UDFID

return $reservationid
}

function Add-Booking {
param(
    [string]$reservationid,
    [string]$eventname, #255 max chars
    [string]$bookingdate,
    [string]$starttime,
    [string]$endtime,
    [string]$room
    )

    # TODO:
    # On http://xxx.xxx.xxx/EMSAPI/Service.asmx it says there is only AddBooking and it has 10 arguments.
    # On https://success.emssoftware.com/Content/OptionalFeatures/EMS_API/A_V44.1/API_EMSAPIfunctions.html it says there is an AddBooking and an AddBooking2. The first has 9 arguments and the second has 10 arguments.
    # When attempting to run AddBooking with 9 arguments, it fails saying it can't find an overload for AddBooking and the argument count of 9.
    # When attempting to run AddBooking with 10 arguments, it runs and uses the 10th argument to insert an eventtype on the booking.
    # When attempting to run AddBooking2, it fails saying there is no method named AddBooking2.

$roomId = ($RoomTranslation | Where-Object { $_.Room -eq $room }).ID

# Add a Booking to an existing Reservation
[xml]$AddBookingResult = $EmsConnector.AddBooking($username, $password, $reservationid, $roomId, $bookingdate, $startTime, $endTime, $eventName, $confirmed, $creditevent)
$bookingid = $AddBookingResult.Booking.Data.BookingId

return $bookingid
}



######################################################################################################################################################
Clear-Host

# EMS API Connection Variables
$url = "http://ems.yourdomain.com/EMSAPI/Service.asmx"
$username = "youremsapiuser"
$password = "youremsapipassword"

## User Changeable Items
$holidayfilename = "C:\Users\shardwic\Desktop\EMS\LSUS-DaysClassesNotInSession.csv"
$classfilename = "C:\Users\shardwic\Desktop\EMS\LSUS-SpringClasses.csv"
$global:searchstartdate = "1/1/2018" #Earliest potential date - used in searching to see if has already been booked
$global:searchenddate = "12/31/2018" #Last potential date - used in searching to see if has already been booked

## Constants
$global:groupIntegrow  = 627 # Group 627 is Integrow Student System
$global:groupCE        = 1 # Group 1 is Continuing Education
$global:udfRefNum      = 14  # UDF 14 is SIS Reference Number (Reserveration UDF)
$global:udfSemester    = 15  # UDF 15 is SIS Semester (Reservation UDF)
$global:confirmed      = 1   # Reservation Status 1 is Confirmed
$global:allBuildings   = -1  # Building Code -1 is ALL buildings
$global:creditevent    = 35 # Event Type 35 is Class-Credit
$global:noncreditevent = 36 # Event Type 36 is Class-NonCredit

$EmsConnector = New-WebServiceProxy -Uri $url

[xml]$ApiVersion = $EmsConnector.GetAPIVersion()
Write-Host "EMS Version" $ApiVersion.API.APIVersion.Version $ApiVersion.API.APIVersion.License
Write-Host ""

## Store in Arrays these items that are needed for browsing Bookings

[xml]$Statuses = $EmsConnector.GetStatuses($username, $password)
$StatusArray = $Statuses.statuses.data.id

[xml]$EventTypes=$EmsConnector.GetEventTypes($username, $password)
$EventTypeArray = $EventTypes.EventTypes.Data.id

[xml]$GroupTypes = $EmsConnector.GetGroupTypes($username, $password)
$GroupArray = $GroupTypes.GroupTypes.Data.id

# Store in Array this item to be able to use as a room lookup

[xml]$AllRooms = $EmsConnector.GetAllRooms($username, $password, $allbuildings)
$global:RoomTranslation = $AllRooms.Rooms.Data | Select-Object -Property Id, Room

## Semester holidays Import
$HolidayHeader = "DateNotInSession"
$HolidayFile = Import-CSV -Path $holidayfilename  -Header $HolidayHeader
[array]$holidayarray  = $HolidayFile.DateNotInSession

## Class File Import
$ClassHeader = "Reference","Name", "Title","Instructor","Room","StartDate","EndDate","StartTime","EndTime","DaysMeet","Semester"
$ClassFile = Import-CSV -Path $classfilename -Header $ClassHeader

$ClassFile | ForEach-Object {

    ## Class variables
    [string]$referencenumber = $_.Reference
    [string]$eventname = $_.Name + "-" + $_.Title
    [string]$classname = $_.Name 
    [string]$classtitle = $_.Title
    [string]$teacher = $_.Instructor
    [string]$room = $_.Room
    [datetime]$startDate = $_.StartDate
    [datetime]$endDate = $_.EndDate
    [string]$starttime = $_.StartTime
    [string]$endtime = $_.EndTime
    [string]$days = $_.DaysMeet
    [string]$semester = $_.Semester

    ## Convert the days string to an Array, and then to the full day names, so we can compare to DayOfWeek
    $daysarray = $days.ToCharArray()
    $daystoprocess = @()

    foreach ($day in $daysarray) {
      switch ($day) {
           "M" {$daystoprocess += "Monday"}
           "T" {$daystoprocess += "Tuesday"}
           "W" {$daystoprocess += "Wednesday"}
           "R" {$daystoprocess += "Thursday"}
           "F" {$daystoprocess += "Friday"}
           "S" {$daystoprocess += "Saturday"}
           default {"ERROR"}
        }
    }

    ## Find Information to Create Reservation and Bookings
    $DaysToBook = Find-DaysToBook -startdate $startDate -enddate $endDate -dayArray $daystoprocess -holidayArray $holidayarray

    $ClassReservation = @()

    foreach ($daytobook in $DaysToBook) {

        $ClassProperties = [pscustomobject]@{
            Reference = $referencenumber  # Stored in UDF 14 is Integrow Reference Number (Reserveration UDF)
            Semester  = $Semester         # Stored in UDF 15 is Integrow Semester (Reservation UDF)
            EventName = $eventname        # KHS 281 - Personal Health
            ClassName = $classname        # Not Used
            Title     = $classtitle       # Not Used
            Teacher   = $teacher          # Not Used
            Room      = $room
            DayName   = $daytobook.DayOfWeek
            Bookday   = $daytobook.ToShortDateString()
            StartTime = $starttime
            EndTime   = $endtime
            EventType = $creditevent      # Event Type 35 is Class-Credit
          }

        $ClassReservation += $ClassProperties
    }
    $ClassReservation.Bookday

    # Make the Reservation and first booking
    Write-Host "Reservation and First Booking"
    $returnedRezid= $ClassReservation | Select-Object -First 1 | Create-Reservation -eventname $_.eventname -bookingdate $_.bookday -starttime $_.starttime -endtime $_.endtime -roomId $_.room -referencenumber $_.reference -semester $_.semester

    # Make the Additional Bookings
    Write-Host "Additional Bookings for Reservation # " $returnedRezid
    $ClassReservation | Select-Object -Skip 1 |  Add-Booking -reservationid $returnedRezid -eventname $_.eventname -bookingdate $_.bookday -starttime $_.starttime -endtime $_.endTime -room $_.room



}



