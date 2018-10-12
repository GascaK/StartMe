#StartMe
**StartMe** is an AutoIt script that utilizes UDFs (User Defined Functions) and the standard library to manage and control a plethora of 
programs needed to work in the hotel industry. Several mundane and repetitive tasks can be easily completed with guaranteed success and accuracy using the
built-in GUI (Graphic User Interface).

##Getting Started
StartMe began because of the frustration with several *Passwords* and *Usernames* needed to complete routine tasks and metrics. This started as a password 
management tool, but ended up being a reliable and all-inclusive app instead. StarMe currently works only for the *HotSOS* and Lodging Management Systems (*LMS*) utilized at
the hotel I am currently employed in. 

###Prerequisites
A functional version of AutoIt is required. Unfortunately, the security systems within most businesses are very strong and a working copy of an .EXE is not possible.
Head over to [AutoItScript.com](https://www.AutoItscript.com/site/AutoIt/downloads/) and download the latest version (currently **3.3.14.5**) at time of writing.

###Installation
Download the repository (Or Clone), and install to a secure location on your Work PC. If this is the first time running this script, make sure to link the .au3 
extension in the Operating System. This was turned off on my Workbench, so a workaround of a run.bat was created to link the different scripts for clarity.
<br>
*runSnitch.bat*
`@echo off
Start("snitchMe.au3")
`
<br>
The main .au3 file [StartMeGUI.au3](StartMeGUI.au3) will create the .smg files and the .ini if they weren't created on bootup. 
That's it!

##How to Use
Launch StartMeGUI.au3. You may get a warning to set up your passwords and usernames before using the application, and this will help prevent locking anyone out of LMS or HotSOS.
Navigate to "Res/StartMe.ini" and edit the HotSOS=*Change Here* and the LMS=*ChangeHere* to your currently assigned usernames.

###Password Management
StartME utilizes the <crypt.au3> UDF and AES (128 Bit) encryption to obfuscate the program. However, the private key is made available in the .au3 file itself.
As long as no one gains physical access to your PC, you should be fine. 
<br>
When first loading the program, navigate to File -> LMS Password and File -> HotSOS Password. A popup will instruct you to change the passwords to
the corresponding programs. Everytime you need to change your password in either HotSOS or LMS, the password must be changed via this method as well. At this time,
clicking the large HotSOS or LMS buttons in the main window will automatically log you in to the corresponding application.

###SnitchMe
A detailed report of the hotel room calls must be completed at the end of every shift. This report, called the Snitch report, is designed to verify completion of
all calls in any given room by the engineer. HotSOS is unforgivingly slow at this, but a .csv (Comma Seperated Value) 
file can be exported via HotSOS and manipulated by AutoIt. **hotelOrders.csv** must be the file name to get the report to cooperate fully. A full list of every 
qualifying room and engineer along with the issue will be in an array listed on a new window. Drag and Drop the Row # to the Input Field and hit "Clipboard" to 
copy the information and paste it in the SnitchMe report.

##Version

###2.0
- Git support started here
- SnitchMe report generator with Drag and Drop feature
- AES encryption for LMS and HotSOS passwords
- Functioning Autolog feature prevention

###1.0
- Password managment
- Prevent Autolog feature (not functioning)
- Clear and fullproof GUI