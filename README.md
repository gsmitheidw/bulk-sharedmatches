# bulk-sharedmatches

Powershell script to partially-automate the opening of shared DNA matches on [Ancestry.com](https://www.ancestry.com)

This is to aid anyone using the [Leeds Method](https://www.danaleeds.com/the-leeds-method/) to categorise DNA matches.
I have written this to minimise the number of keyboard presses and mouse clicks to open a batch of relatives profile pages
in seperate browser tabs, then browse to the shared matches section of each person's page. 

Ancestry.com does not provide any public api access. It may be possible to achieve this with [Selenium](https://www.selenium.dev/), sending keypresses is somewhat primitive - however it is not likely to overload Ancestry's servers particularly with timings and small batches as currently set. This script runs without any software installation or 3rd party products. 


## Requirements:

- Powershell on Windows (default), it won't work on Linux with Powershell due to dependency on the sendkeys method.
- Any browser with tabs should do, have tested this with Chrome and Firefox. Edge Chromium should be fine too. 

## Setup and use:

1. Open you web browser on the shared matches page. 
2. Select **new matches only** and ***public linked trees.*** - private trees is handled seperately
3. Open the script from within Visual Studio Code or Powershell ISE (or command line powershell if you prefer)
4. Run the script (play button :arrow_forward: ) and **immediately** switch to your own browser window
5. Hands off!! Let this run until it has opened 15 profile pages and cycled through each one opening the shared matches. . It will beep once when completed a batch. Do not use the mouse or keyboard whilst a batch is in progress.
6. You will have 16 browser tabs, the first is your list of shared matches, go to the second one <kbd>Ctrl</kbd> + <kbd>2</kbd>. If there are no shared matches close the tab with <kbd>ctrl</kbd> + <kbd>w</kbd>... continue until you see matches. Tag as appropiately and continue closing tabs until you reach the end of the batch. 
7. To start the next batch, go back to the script, run again and switch straight back to the browser.

## Private trees:

In this section of code:
```powershell
function OpenProfiles {
    # toggle 8 for public linked trees and 7 for private trees. Cannot do both!
    1..8 | ForEach-Object {
       $wshell.SendKeys("{TAB}")
    }
    $wshell.SendKeys("^{ENTER}")
}
```

Change 1..8 to 1..7 to do private trees, switch it back to 8 for public trees. 
***Remember to change it in the browser too or the script will open/click randomly!***
These must be changed in tandem.

## Final Notes:

- If Ancestry changes their web design the tab stops will change and the code will break.
- This does not work with other sites like MyHeritage, FT DNA, 23andMe etc. 
- If the script fails or you forget to keep the public/private tree profiles in tandem with the 7/8 toggle above, just quickly switch from browser window to powershell and stop the script: :red_square:

