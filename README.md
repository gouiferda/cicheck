# ciCHECK Project :computer: :credit_card: (Attendance tracking system) 

### ciCHECK: Attendance tracking system, Scanning RFID Cards or tags and saving text and excel records on connected computer (Arduino and Processing Mini Project)

> Links to download latest ciCHECK app and requirements:

* Download ciCHECK app for MacOs (Requires Java): http://bit.ly/getciCHECK2

* Download ciCHECK Box drivers : http://bit.ly/getciCHECKBoxDrivers

* Download Java : http://bit.ly/downloadJava1073


### ↪️ User guide instructions:

> Install Instructions for ciCHECK app:
1) Download "ciCHECK Box Drivers.zip" then extract/decompress the folder to Desktop, Double click on "run.command" (Requires password), then Double click on "CH34x_Install_V1.4.pkg" Install then reboot.
2) Download and run "ciCHECK Installer.pkg" and follow Instructions to install (insert user password if needed)
3) ciCHECK app is ready to run from Launchpad and Applications folder

* Troubleshooting:
if you connect ciCHECK box and run the ciCHECK app , and the ciCHECK box is not detected you should probably follow this guide https://www.igeeksblog.com/how-to-disable-system-integrity-protection-on-mac/
and then repeat step (1) in Install Instructions.


> Instructions to use ciCHECK Box and app:
* To add new users: 
	- Run the app "ciCHECK.app"
	- Plug the ciCHECK box ,Check if the box is connected in the app
	- Then click on "New card" button, follow instructions.
	
* More guides coming soon...

	
Side note:
* It is safe to unplug/plug the ciCHECK box into another usb port

### 🖼 Screenshots:

> ScreenShot of the ciCHECK application (MacOS)
![screen shot 1](https://i.imgur.com/J81IKHv.png)
![screen shot 2](https://i.imgur.com/KoUadsH.png)

> The ciCHECK box Plugged to the computer
![screen shot Attendance Pad](https://i.imgur.com/Uil2mJB.jpg)

> ScreenShot of generated excel sheet file records
![screen shot Attendance Pad Records](https://i.imgur.com/nhZl0s5.png)


### ⚙️ Electronics Developement Instructions (for developers):

* Componenets and tutorial:
https://www.mschoeffler.de/2017/02/07/how-to-use-the-rfid-rc522-module-rfid-reader-with-the-arduino-uno/

Side note: I developed the project little bit further by adding LEDS to the project to notify when the card is scanned and when the arduino is on.


### 🖥 App Developement Instructions (for developers):

Used Developement Envirement softwares:

* Arduino IDE : https://www.arduino.cc/en/Main/Software
* Processing : https://processing.org/download/
	
Both code sources for Arduino and for Processing are shared on the master repository,
to add more features and enhacements work on Processing source code.

* the Processing code uses:
	> Java Excel API library (Included on the code source of Processing)
	/ Java Mail Library (Builtin with Processing).
* the Arduino code uses MFRC522 library (Available in Arduino library manager).
* the Arduino pitch sounds library (included) for future sound feature developement.


# Credit:
2018 (c) Soufiane Gouiferda 

# Thanks to:
Connect Institute http://connectinstitute.ma/
