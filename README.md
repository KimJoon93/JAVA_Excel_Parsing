# JAVA_Excel_Parsing
POI Library by Apache

## Before get starting...
I had a Excel file that has to be upload to my database. Each cells were mixed by  imformations, blanks, and formulas. And I had to parse ExcelFile in forms like "1","apple","red","","Round"...etc. So first, I decided to use Java and I found POI Library.

## What Library?
There were some Libraries to use when we handle Excel files with Java. When I surf in Internet most people used POI Library. Because there are version of Excel we could use. Most of people use higher version Excel, so I have 2007 version and I used POI Library. 

### POI Library
- Download here : https://poi.apache.org/

![POI](https://user-images.githubusercontent.com/32008149/60108008-dffe6c80-97a2-11e9-963f-7d87a7cf7d5a.PNG)

- We have to Add Red Rectangular files in BuildPath(Libraries)\
if you are going to use xlsx file you should add ooxml-lib directory files.
![POI2](https://user-images.githubusercontent.com/32008149/60109091-afb7cd80-97a4-11e9-99f9-56b4ec8a9a40.PNG)

- In Maven, you could put dependency to POM.xml (Beacareful for the Version)\
You can find Here: https://mvnrepository.com/artifact/org.apache.poi/poi/3.17
![POI3](https://user-images.githubusercontent.com/32008149/60109311-1937dc00-97a5-11e9-8ef5-db98598edaad.PNG)

Now What I have to do is importing csv file to my DataBase.\
But there were serious Problem in imporing to DataBase.


## CSV
I had serious problem with importing csv file to Database. There were many ways to import Data to DataBase, but I need to import csv File because I need a function that has to be imported by button in Web. So I parse xlsx file and made new xlsw file in it's own form.  
File looks just like this.
[Image]
I was using Dbeaver tools, and I found importing function. I need to use Query but, I am not used to Database so, I use the function. Then I found that I had to change xlsx files to csv file. So I just changed extension. But then, problem occured. They shows how it works but, it seems something wrong.
So I tried to change Column delimiter, Quote char, and Quote char. But that doesn't work at all. I tried to think what makes problem. Then I thought what is CSV?

### What is CSV?
CSV (Comma separarted version) is file that has been separated by comma. It is not same as xls,xlsx. First, I thought it's same as xls or xlsx. But it wasn't. It is not Excel file. There are meta data in files. In program I made xlsx file. So it's xlsx file even if I changed extension. And what is CSV? It is text file! I was wrong! So I changed some code in program. And I made csv file in program. Then it works! 

### Import CSV to Database(MariaDB)
I used DBeaver tool to handle MariaDB. And There were some problem to think about. First I should change "o", "x", " " text to "Y", or "N".
```
if (columnindex ==11 || )
```
DB Import was successful, but we have to think about some issue.
- First, 

