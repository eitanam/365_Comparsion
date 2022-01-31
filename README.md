# 365_Comparison

This tool will assist you comparing between the HR active users lists to the Office 365 active users. 
You will get a list of all the users which are eanabled on the 365 but are not appearubg on the HR active users list.

*********************************************************************************
This script is provided AS-IS without any warranty to any damage that may occured.
If you are using it it's AT YOUR OWN RISK!
*********************************************************************************

Version 1.0
Inital release


How to use:
In order to use it you will need the list of users from your HR department. Make sure that there is column named Business Email.
Export the 365 users. From the 365 portal click on Users --> Active users --> Export users

Run the Comparsion.ps1 

![image](https://user-images.githubusercontent.com/71331120/151758917-aac525a0-f170-4297-8ab4-c462ec845a6d.png)

Choose the file you got from the HR department
Choose the file you generated from the 365
Click on the Search button

once done you will get a report of all the acrive users in the 365 who are not appeared in the HR lists.
