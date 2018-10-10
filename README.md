#StartMe
**StartMe** is an Autoit script that utilizes UDFs (User Defined Functions) and the robust standard library to manage and securely control the plethora of 
programs needed to work in the Hotel Industry. Several mundane and repetitive tasks can be easily completed with guaranteed success and accuracy using the
built in GUI (Graphic User Interface).

##Getting Started
StartMe started because of a frustration for the several *Passwords* and *Usernames* needed to complete routine tasks to complete metrics. This started as a password 
managment tool but ended up being a reliable all inclusive app instead. Currently works only for the *HotSOS* and Lodging Manament Systems (*LMS*) utilized at
the Hotel I am currenlty employed in. Honestly this probably won't help you very much and is really only used to minimize the frustration at my job.

###Prerequisites
A functional version of Autoit is required, unfortunately the Security Systems in place at most businesses is very strong and a working copy of a .EXE is 
most likely not going to be possible. Head over to [AutoitScript.com](https://www.autoitscript.com/site/autoit/downloads/) and download the latest version currently
**3.3.14.5** at time of writing.

###Installation
Download the repository (Or Clone), and install to a secure location on your Work PC. If this is the first time running this script make sure to link the .au3 
extension in the Operating System. This was turned off on my Workbench so a workaround of a run.bat was created to link the differnet scripts for clarity.
<br>
*runSnitch.bat*
`@echo off
Start("snitchMe.au3")
`
<br>
The main .au3 file [StartMeGUI.au3](StartMeGUI.au3) will create the .smg files and the .ini in case they were not created at bootup. That's it!

##How to Use
Launch StartMeGUI.au3, you might get a warning to set up your passwords and usernames before using, this is to prevent locking anyone out of LMS or HotSOS which
are a pain to get IT to reset. "I'm talking to you Brian.." Navigate over to "Res/StartMe.ini" and edit the HotSOS=*Change Here* and the LMS=*ChangeHere* to your
currently assigned usernames.

###Password managment
StartME utilizes the <crypt.au3> UDF and AES (128 Bit) encryption to obfuscate the program. However the private key is made available in the .au3 file itself.
As long as no one gains physical access to your PC you should be fine. 
<br>
When first loading up the program, navigate over to File -> LMS Password and File -> HotSOS Password. A popup should instruct you to change the passwords to
the corresponding programs. Everytime you need to change your password in either HotSOS or LMS the password must be changed via this method as well. At this time
clicking either large HotSOS or LMS buttons in the main window will autologin the corresponding application.

###Snitch Me
A detailed report of the completed hotel room calls must be completed at the end of my shift everyday. This report, lovingly called the Snitch report is to verify 
all calls in any given room is completed by the engineer responding to the issue. HotSOS is unforgivnely slow at this. Fortunately a .csv (Comma Seperated Value) 
file can be exported via HotSOS and manipulated by Autoit. **hotelOrders.csv** must be the file name to get the report to cooperate fully. A full list of every 
qualifying room and engineer along with the issue will be listed in an array on a new window. Drag and Drop the Row # to the Input Field and hit Clipboard to 
copy the information to be pasted in the Snitch Me report.

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