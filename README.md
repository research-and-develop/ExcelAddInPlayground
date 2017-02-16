# ExcelAddInPlayground
This repository contains C# project that is showing sample VSTO excel add-in implementation 

# Requirements
To run this project you need:
- Visual Studio 2015 (I've used Community Edition);
- Office Tools For Visual Studio 2015 (which I downloaded from here https://www.visualstudio.com/vs/office-tools/);
- You need to have installed Office 2013 or 2016;
    
Project Type is **Excel 2013 and 2016 VSTO Add-in**

# Info

The project is a simple playground.
Mainly it shows:
  - background processing sample using TaskScheduler class;
  - adding custom menu items in cell's context menu(right click);
  - programatically creating of VBA module and functions inside of it;
  - making a proxy between the excel application and add-in's C# code that allows you to call C# functions from VBA;
  
Cheers!
