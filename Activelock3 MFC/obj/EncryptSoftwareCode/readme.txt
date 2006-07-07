Encrypt Software Code
---------------------

Developed under Microsoft Visual Studio 2003
---------------------------------------------
C/C++ is pure C++
Visual Basic is in VB.NET

StringHider
-----------

This program takes the license key

"AAAAB3NzaC1yc2EAAAABJQAAAIB8/KWB2oai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q=="

and hides it by splitting it into sections, cyclically shifting the sections and then randomly
shuffling the sections before storing the sections as integers.

At the same time it produces the complete code, C/C++ or VB, in a file temp.c to decode the array
of integers back to the license key.
The code contains data structures as described above and code to reform the license key.

This has been implemented in MFCSample an example in C++.


Details
-------

The license key is is made a multiple of an ineger size by adding 'x' to key. 
Maybe this is nor necessary, as always a multiple, but just in case!

The license key is split into small strings of integer size and placed in an array.

Each string is converted to an integer and rotated cyclically 1 place and then added to an array.

The last array is randomly shuffled. How it is randomly shuffled is recorded so that we can
unshuffle later

The rest of the program outputs code to file temp.c

Data
Constants such as charsPerInt

An array describing how the license key was shuffled
The shuffled integer array representing the license

The code for converting the above structures back to the license key


Comments
--------

After temp.c has been produced it is imported into the program StringHiderCTest for testing.

As this is open source and can be read by anyone, this is intended more as an example.
I have rotated cyclically 1 place, this can be varied.
The data is inside the unshuffling procedure, the procedure and data should be split and
hidden, etc etc

StringHiderBasic is a similiar test program to test the Visual Basic output of StringHider.

Problems
--------

As I finish writing this it dawns on me that if I produce the code on a 32 bit machine and this is 
run on a 64 bit machine I may be in trouble ( I do not have a 64 bit machine, so I do nt know for 
certain). On reflection it should be no problem !!!!!