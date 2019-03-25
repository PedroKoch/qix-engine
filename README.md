This software is provided 'as is'.  
Created by Pedro Koch  
January 4th, 2019  
  
Last Update on March 20th, 2019  
Starting to support Qlik Sense Desktop.  
  
  
  Tested on Qlik Sense Server - February 2019 - 13.9.1
  Tested on Qlik Sense Desktop - February 2019 - 13.9.1
  
  
  
### **Prerequisites**
>>You must have nodejs installed in your pc.  
>>>>https://nodejs.org/en/download/  

Preparation:  
>>1) Run the init.cmd by double-clicking the file. This will create two necessary folders and install some nodejs modules.  

>Alternative  
>>1) In a Command Prompt window, navigate to the folder in which you plan to mantain all the related files  
>>>>1.1) Type 'init' (without the quotes) in the prompt, press ENTER  

### **Configuration**

**Desktop**
>>1) If you are working with Qlik Sense Desktop, navigate to src/ folder.  
>>>>1.1) In the config.json file, you'll have to configure some specific settings, usually the standard:  

>>>>>>`"engineHost": "localhost",`  
>>>>>>`"enginePort": 4848,`  
>>>>>>`"appId": "engineData",`  
>>>>>>`"server": false,`  
>>>>>>`"userDirectory": "",`  
>>>>>>`"userId": "",`  
>>>>>>`"certificatesPath": ""`  

>**NOTE** the double quotes !!!  
>Save the changes made to config.json.   

 ---

**Server**
>>1) Go to the QMC > Certificates  
>>>>1.1) Click on 'Add a machine' and input your machine's name  

>>>>>>1.1.1) If you don't know your machine's name, in Windows 10, go to Settings > System > About  

>>>>1.2) In the 'Export file format for certificates' option set 'Platform independent PEM-format'  

>>>>>>1.2.1) If you like and know what you're doing, you can set the other options  

>>>>1.3) Click 'Export certificates'  

>>2) Navigate to the folder where the certificates were exported, then copy 
(or move) the files inside the folder with the same name of your machine's name 
(the one you've entered in the QMC > Certificates page) to the Certificates/ folder 
of this project. Do not modify these files.  

>>3) In the src/config.json file, you'll have to configure some specific settings of your QS Server:

>>>>>>`"engineHost": "",`  
>>>>>>`"enginePort": 4747,`  
>>>>>>`"appId": "engineData",`  
>>>>>>`"server": true,`  
>>>>>>`"userDirectory": "",`  
>>>>>>`"userId": "",`  
>>>>>>`"certificatesPath": "../Certificates"`  

>>>Follow this example to properly configure these settings:
>>>>My company's name is Company X and has a QS Server on its intranet. When I 
logon to the QMC (or Hub) it will be requested an username and password. The 
first one is in the form Domain\username (e.g. CompanyX\john.smith). My 
machine's name is CX123. And the URL to access the QMC should be someting 
like https://cx123.companyx.intranet/qmc/. Then the result is someting like:

>>>>>>`"engineHost": "cx123.companyx.intranet",`  
>>>>>>`"enginePort": 4747,`  
>>>>>>`"appId": "engineData",`  
>>>>>>`"server": true,`  
>>>>>>`"userDirectory": "CompanyX",`  
>>>>>>`"userId": "john.smith",`  
>>>>>>`"certificatesPath": "../Certificates"`  

>**NOTE** the double quotes !!!  
>The enginePort is the same configured during the QS Server Installation 
(should be the same 4747 if you did not change it). The certificatesPath
should be the folder with the exported certificates. The appId should be kept "engineData".
>Save the changes made to config.json.

---  

### **Final Considerations**
>>1) Now we're ready to start using the program. Four (4) basic operations are available:
>>>- Backup: will create 3 csv files in the Backup folder, with Master Measures, Master Dimensions and Variables.
>>>- Load: load Master Measures, Master Dimensions and Variables contained in an excel file. Currently, hierarchic dimensions are not properly handled. See Step 6).
>>>- Erase: it will delete all the app's master objects (Dimensions, Measures and Variables).
>>>- Create Adhoc Table: creates a master Visualization with all master Dimensions and master Measures.

>>>Both QS Server Services or QS Desktop must be running in order to achieve the expected results.
>>>Alternatively, in the Command Prompt, in the src/ folder, type node main.js. 
>>>This will enter an interactive style console, where you can do the above operations in a less automated way.

>>2) The excel file used must have the .xlsx extension. It should follow the template file. 
It is of utmost importance to follow columns order and sheet names. 
The sheet order is irrelevant, however, the sheets must be the (3) first ones in the workbook.
Using the given template excel file, you should enter Template OR ./Template OR ./Template.xlsx

>>3) Warning: when creating something on your app, it can be necessary to reload the app a few 
times to see the changes you made. Be aware that delete operations are persistent.
