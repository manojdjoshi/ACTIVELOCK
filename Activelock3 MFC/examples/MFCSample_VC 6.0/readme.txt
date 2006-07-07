MFCSample

The licensing mechanism of MFCSample is to say the least a little strange but was developed before
the latest Hard Disk Serial Lock Code of Activelock3 was developed.

The final installation code consists of 2 items of data
1. What is conventionally called the installation code. ie if you select disk you will typically
   receive a code like NgHjuknl6u
2. The lock type as an index 0-5. This is appended to the code in 1 to typically produce NgHjuknl6u5

The above 2 items of information are shown in the about box as the Installation Code

(In earlier version there was a user name, this perform no function, so has been removed, in fact
commented out incase it is wanted)

Production of a license in Alugen

In the License Key Generator tab select TestApp 1.0 and time locked
Set Using Lock Type=LockNone Only
Enter into the user name (not installation code)
the installation code from the about box.
(earlier this also was followed by an ! exclaimation mark followed by the user name
ie NgHjuknl6u5!David Weatherall but has been removed)

Generate a license
and save to c:\testapp.all

restart the application, which should now be licensed

A cheap trial license has been implemented. This is not the new Trial license implemented in
ActiveLock3. We are waiting for the activeLock version to mature, MFCSample has the Activelock3
Trial version commented out, I had reasonable success with it ( ie not 100%).

For cheapo trial license, use a short time lock and for username enter
trial0
that is trial followed by a zero (0 is a number)
and continue as described.


Note: 
There is a testapp.all in the MFCSample directory, created as a cheapo trial license. Place
testapp.all in c:\ and run MFCSample.
