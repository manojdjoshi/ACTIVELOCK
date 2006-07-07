MFCSample

The latest changes to MFCSample are aimed at seperating the ActiveLock code from the application
code. The aims are

1. Seperation of code as described above
2. In the case of having multiple application this enables one set of activelock code to be 
   maintained.
3. ActiveLock V3.4

4. Utilises Jereon's #import "ActiveLock3.4.dll" technique instead of the poor one I was using.
This gives full access to the dll, including the enumerations from the dll which were missing
from my versions.

5. Djordie provided a class which he developed under my old technique of accessing the dll.
I have updated this to use the #import "ActiveLock3.4.dll" technique.

6. There was a weakness of design in the old MFCSample in regard to obtaining the Installation
code.
 
      