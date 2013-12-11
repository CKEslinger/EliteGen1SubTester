PassTime Elite II Tester README.txt
 Build Date: 7/9/2013
 File Version: 2.4.0.20
 Product version 2.04.0020
 Splash Screen Version 2.4, Build 20

This file contains information on how to execute the PassTime Elite II Tester, 
dialog based application (PassTimeTester.exe) for validating the correct 
programming and functioning of a PassTime Elite II device.

Software:
This validation application is a form based application written in Microsoft 
Visual Basic 6.0. Along with various activeX controls and dynamic link 
libraries (installed as part of the PassTime Tester Setup Package) the 
application uses an xml file for storing and retrieving the serial numbers of 
the current PassTime device under test. The xml fils is also used to store and 
retrieve application settings. A seperate text file, PassTimeInv.txt, installed
in the C:\PassTimeData folder, contains the serial numbers of units that have 
passed the tests.

Hardware:
This application is used in conjunction with a PassTime Elite II Functional 
Fixture (with a PassTime Elite II device properly inserted); a PassTime 
interconnect board; a cell site simulator (attached to PC via RS232 cable on 
COM port 5 and attached to the fixture); a GPS antenna (attached to the fixture); 
a standard PC microphone (attached to PC's integrated or external audio and 
attached to the fixture); a standard 9-pin serial cable (attached to PC using COM 
port 4 and attached to the fixture); a digital data acquisition board (PCI DAQ 
installed in the PC and attached to the interconnect board); a variable power 
supply (attached to PC via RS232 cable on COM port 1); a label printer (attached 
to PC via 25-pin cable using LPT1 port); and a barcode scanner (attached to PC 
using any available USB port). See the system wiring diagram for details on 
connecting the components.

Operating the PassTime Tester:
-Before Running the First Test:
Assuming that the devices listed in the hardware section are properly connected 
and configured, make sure that the label printer (Zebra 110xiIII Thermal 
Transfer Printer), the power supply (Agilent Power supply E3646A) and the call 
box (Wiltek 4202s) are powered on.

Additionally, the the gps simulator (Aeroflex GPS-101) should be powered on. 
After the GPS Simulator self-check completes a message displays that warns of 
an out of date almanac. Push the "ESC" button on the front panel to clear the 
message. Use the up and down "Select" arrow keys to select the data fields and 
the SLEW/STEP knob to set the values of data fields as follows:

 -dBm: "-85 dBm"
 SV: "SV14"
 T: "T2"
 DPLR: "DPLR0"
 RF: "RFON"

Finally, start the application by double-clicking the desktop link, 
"PassTime Tester"

-Running a Test
The following is a step-by-step procedure for testing a "known good" PassTime 
Elite II device. These procedures may vary from the actual test procedures used 
by the manufacturer and are listed here for demonstration purposes only:

1) Open the tester. The tester is a double enclosure and requires two handles
   to be released and raised up in order to insert or remove a device
2) Select the type of Elite II unit from the "Model Type Under Test" 
   drop-down list box
3) Select the type of SIM card that is inserted in 
   the Elite II unit from the "Installed SIM Card" drop-down box
4) Select the "Operations to perform". For a full test select all operations
   except "Trial Run"
5) Insert the PassTime Elite II Device Under Test (DUT) into the jig and make
   sure that it is seated correctly in the test fixture
6) Close the tester. Make sure the inner tester encolsure's handle is in the
   locked down position before closing the outer enclosure lid. The test run
   will automatically start
7) After successful completion of the test and subsequent label printing, scan 
   the printed labels with the barcode scanner when prompted
8) Repeat steps 1 through 7 as necessary

-Tools->Setup
Changes to the tester application settings can be made by clicking Tools->Setup
and entering the Elite II Tester Application Password. The most common setup 
items that change are the expected firmware version, the Passtime serial numbers 
that are written to each board and the modem IMEI numbers.

-->Application Firmware Selection
The tester application can download firmware to the DUT based on the selected 
Model Type Under Test. To specify what firmware file the tester application will
download to the DUT, first select the model type in the tester application then
click Tools->Setup, enter the password and click on the Application Firmware 
browse button to open a dialog box that allows you to select the firmware file.
Double click the file that should be used for the selected model and click OK.

The tester application will only download firmware to the DUT when the "Download 
Firmware" check box is selected. The application will, howvever, verify that the 
selected firmware version matches the firmware version on the DUT regardless of 
whether or not the firmware is actually downloaded.

If a newer or older version of firmware needs to be downloaded, simply copy the 
new firmware to the C:\PassTimeData folder and use the instructions above to 
specify which models will use this firmware.

Firmware files must contain the firmware version in the filename and should be 
renamed so that they are of the form, "EliteX.Y.hex" where "X" is the major
version number and "Y" is the minor versrion number. For example, if the firmware
version number is v3.9, the firmware file name should be "Elite3.9.hex".

If firmware won't be downloaded to the DUT, dummy firmware files can be created 
so that the tester will know what version of firmware to expect on the already
programmed chip. Simply use notepad to save a file with no data in it to the file
name as specified above and located in the C:\PassTimeData folder.

-->Serial Number Selection
click Tools->Serial Numbers and enter the password to display the "Update Serial
Numbers" dialog box. Make the necessarry changes (typically just update the "Next 
available" serial number), click OK, click OK again and click "Yes" to verify 
that the changes should be saved.

-->IMEI Number Selection
click Tools->IMEI Numbers and enter the password to display the "Update Serial
Numbers" dialog box. Clicking on a row that corresponds to a model type that 
needs to have its IMEI numbers adjusted brings up a second dialog box. Make
the necessarry changes (typically just update the "Next available" IMEI number),
click OK, click OK again and click "Yes" to verify that the changes should be 
saved.
