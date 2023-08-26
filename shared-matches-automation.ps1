<#
.
. Open Shared matches, select new matches only and public linked trees in a browser window with no other tabs.
. When that's ready run the script (ie: play button in VSCode or Powershell ISE)
. then switch to browser window WITHIN 5 seconds. It will beep once when completed a batch.
. Go to the second browser tab (first profile page) and use [ctrl]+[w] keys to close tabs with no shared matches
. repeatedly until you find one with matches you can identify as being a specific grandparent line. 
. When the last one is completed, run the script again for another batch.
.
. Note: if your broadband or device is slow you may need to adjust the timings with the seconds in the Start-Sleep commands below.
.
#>

$wshell = New-object -com wscript.shell

Start-Sleep 5

### Part 1 ###

# Refresh confirms tab stop focus in browser
$wshell.SendKeys("^R")
# Give this time to refresh
Start-Sleep 20

# Navigate to first person which takes 16 tab stops:
1..16 | ForEach-Object {
   $wshell.SendKeys("{TAB}")
}

# Open that person's profile in new tab
$wshell.SendKeys("^{ENTER}")


# open more people from that page
function OpenProfiles {
    # toggle 8 for public linked trees and 7 for private trees. Cannot do both!
    1..8 | ForEach-Object {
       $wshell.SendKeys("{TAB}")
    }
    $wshell.SendKeys("^{ENTER}")
}

1..15 | ForEach-Object {
    Start-Sleep 1 
    OpenProfiles

}



### Part two: ###

Start-Sleep 10
$wshell.SendKeys("^{TAB}")

function shared {

Start-Sleep 1

1..9 | ForEach-Object {
    $wshell.SendKeys("{TAB}")
}
#Start-Sleep 1
$wshell.SendKeys(" ")
$wshell.SendKeys(" ")
Start-sleep 1
$wshell.SendKeys("^{TAB}")
Start-Sleep 1
}


# Number of browser tabs of profile pages to open - 15 works well for me as a batch.
1..15 | ForEach-Object {
   shared
}

# Issue a system beep to notify batch completion
[console]::beep(600,1000)

