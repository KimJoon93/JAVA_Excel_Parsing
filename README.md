# JAVA_Excel_Parsing
POI Library by Apache

## Before get starting...
I had a Excel file that has to be upload to my database. 
Each cells were mixed by  imformations, blanks,
 and formulas. And I had to parse ExcelFile in forms
  like "1","apple","red","","Round"...etc. So first, 
  I decided to use Java and I found POI Library.

## What Library?
There were some Libraries to use when 
we handle Excel files with Java. When I 
surf in Internet most people used POI Library. 
Because there are versions of Excel we could use. 
Most of people use higher version Excel, and I have 2007 version so I used POI Library. 

### POI Library
- Before we use library
    Check this out if you have problem with using xssfWorkbook.\
    [Can Apache be compiled / used Java 10 or newer?](https://poi.apache.org/help/faq.html#faq-java10)\
    Because I had problem using xssfworkbook in new Mac book. I installed java 12 version in mac. 
    So my code worked in windows notebook, but it had problem in compiling in Mac.
    If you think there are no problems in your code and adding jar files, and have compile problem like this
    ![스크린샷 2019-07-02 오후 10 00 38](https://user-images.githubusercontent.com/32008149/60514655-e5186a00-9d14-11e9-9f5a-eab1df34fae1.png)
    
    Try use maven project. You could easily copy and paste dependency in there.

- XSSF / HSSF / SS
  
  Name | Feature 
  ----- | ------      
  HSSF | Excel 97 ~ 2003
  XSSF | Excel 2007 ~
  SS   | XSSF Straming version (Low memory and fits to mass data)  
   
- Max Data

    Excel 2003 | Excel 2007 
    ----- | -----
    265 Column | 16,384 Column
    65,536 Line | 1,048,576 Line

- Download here : https://poi.apache.org \
Window OS : Download Zip\
Linux or Unix OS : Download tar

![POI](https://user-images.githubusercontent.com/32008149/60108008-dffe6c80-97a2-11e9-963f-7d87a7cf7d5a.PNG)

- We have to Add Red Rectangular files in BuildPath(Libraries)\
if you are going to use xlsx file you should add ooxml-lib directory files.
![POI2](https://user-images.githubusercontent.com/32008149/60109091-afb7cd80-97a4-11e9-99f9-56b4ec8a9a40.PNG)

- In Maven, you could put dependency to POM.xml (Beacareful for the Version)\
You can find Here: https://mvnrepository.com/artifact/org.apache.poi/poi/3.17
![POI3](https://user-images.githubusercontent.com/32008149/60109311-1937dc00-97a5-11e9-8ef5-db98598edaad.PNG)

- Make new Maven Project\
![스크린샷 2019-07-02 오후 10 09 22](https://user-images.githubusercontent.com/32008149/60515203-180f2d80-9d16-11e9-881e-960d3d0c3fe4.png)

- Put dependency in to pom.xml and install maven.(Becareful for the version, try to use one that is mostly used.) 
![스크린샷 2019-07-02 오후 10 46 19](https://user-images.githubusercontent.com/32008149/60517857-5b1fcf80-9d1b-11e9-93f2-946ca674d862.png)

- Check it works!\
![스크린샷 2019-07-02 오후 10 50 31](https://user-images.githubusercontent.com/32008149/60518118-d6818100-9d1b-11e9-88bc-716f5217d214.png)

- Put code in jsp file.\

Now What I have to do is importing csv file to my DataBase.\
But there were serious Problem in imporing to DataBase.

## CSV
I had serious problem with importing csv file to Database.
There were many ways to import Data to DataBase, but I need
 to import csv File because I need a function that has to be 
 imported by button in Web. So I parse xlsx file and made new 
 xlsw file in it's own form.  
File looks just like this.  
[Image]  
I was using Dbeaver tools, and I found importing function. 
I need to use Query but, I am not used to Database so, 
I use the function. Then I found that I had to change xlsx files 
to csv file. So I just changed extension. But then, problem occured.
They shows how it works but, it seems something wrong.
So I tried to change Column delimiter, Quote char, and Quote char. 
But that doesn't work at all. I tried to think what makes problem. 
Then I thought what is CSV?

### What is CSV?
CSV (Comma separarted version) is file that has been separated 
by comma. It is not same as xls,xlsx. First, I thought it's same 
as xls or xlsx. But it wasn't. It is not Excel file. There are meta 
data in files. In program I made xlsx file. So it's xlsx file, even 
if I changed extension. And what is CSV? It is text file! I was wrong! 
So I changed some code in program. And I made csv file in program. 
Then it works! 

### Import CSV to Database(MariaDB)
I used DBeaver tool to handle MariaDB. And There were some 
problem to think about. 
- First, I should change "o", "x", " " text to "Y", or "N".  
So I found index that has to be converted.(Code can be weired about "A,B,C .." values can be "Y", But Data that I 
received doesn't need to think about that issues.)
  ~~~
  if(columnindex==11 || columnindex==15){                   		 
         if(value=="x" || value.isEmpty()){
              buff.append("\"N\","); 
         }else{
              buff.append("\"Y\",");                 			
         } 
  ~~~
- Second, there are some formulas, in data. When I parse data from file, it brings formula such as "x1 + y1".
What I want was result but, it brings me formula. So I changed value to get Numericvalue. 
    ```
    switch (cell.getCellType()){  
    
         case XSSFCell.CELL_TYPE_FORMULA:                        
    
    	 value=cell.getFormulaCellValue()+"";
    	// we should use cell.getNumericCellValue!
    
    	break;
    		                        
    }
   ``` 
- Third, if Formula cell's value has problem that makes "#value", 
         problem occurs. So I decided to choose NumericValue rather than to 
         show both String Value and NumericValue by "if else" syntax. Because what is formula? It has to make
         NumericValue. I can make them to show "-1" if there are problem in Formula, but 
         
- Lastly, date has to be converted. First I didn't check the date but when I tried to import to Database, I saw weired number.Then
I realized I should convert to data format. 
    ```
    case XSSFCell.CELL_TYPE_NUMERIC:
                        	if(columnindex==14 || columnindex==27){
                        		SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
                        		value = format.format(cell.getDateCellValue());
                        		break;
                        	}
                            value=cell.getNumericCellValue()+"";
                            break;

    ```         
DB Import was successful, but we have to think about some issue.

