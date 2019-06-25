import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFROW;
import org.apache.poi.xssf.usermodel.XSSFSHEET;


public class ExcelParse {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		try {

		    FileInputStream file = new FileInputStream("D:\\Users\\...\\Parsing\\Filename.xlsx");
		    XSSFWorkbook workbook = new XSSFWorkbook(file);

		    int rowindex=0;

		    int columnindex=0;

		    int sheetindex=4;

		    StringBuffer buff = new StringBuffer();

		    XSSFSheet sheet=workbook.getSheetAt(sheetindex);

		    

		    /*생성할 새로운 파일의 Info*/

		    XSSFRow newrow;

		    XSSFCell newcell;    

		    XSSFWorkbook newworkbook = new XSSFWorkbook();

		    XSSFSheet newsheet = newworkbook.createSheet("data");

		       

		    int rows=sheet.getPhysicalNumberOfRows();

		    for(rowindex=1;rowindex<rows;rowindex++){

		 

		        XSSFRow row=sheet.getRow(rowindex);

		        if(row !=null){

					

		        	/*전부 끝 값까지 나와야하는데 getPhysicalNumberOfColumns() 사용하면 빈칸때문에 값이 다르게 나와서 끝 값까지 제대로 안나올 경우 발생하므로 하드코딩*/

		            int cells= 44;

		            for(columnindex=0; columnindex<=cells; columnindex++){

		 

		                XSSFCell cell=row.getCell(columnindex);

		                String value="";

		       

		                if(cell==null){              	

		                	 if(columnindex==cells){

		                      	buff.append("\"");

		                      }else{                   

		                        buff.append("\",");                    

		                      }

		                	 

		                }else{

		              

		                	switch (cell.getCellType()){              	

		                	/*수식을 가져올 경우 값만 가져오기 위해 NumericCellValue를 가져오도록 하자*/

		                    case XSSFCell.CELL_TYPE_FORMULA:                        

		                    	value=cell.getNumericCellValue()+"";

		                        break;

		                    case XSSFCell.CELL_TYPE_NUMERIC:

		                        value=cell.getNumericCellValue()+"";

		                        break;

		                    case XSSFCell.CELL_TYPE_STRING:

		                        value=cell.getStringCellValue()+"";

		                        break;

		                    case XSSFCell.CELL_TYPE_ERROR:

		                        value=cell.getErrorCellValue()+"";

		                        break;

		                    }

		                	

		                	if(columnindex==cells){

		                     	buff.append("\""+value+"\"");

		                     }else{                   

		                         buff.append("\""+value+ "\",");                    

		                     }

		                }               

		            }

		        }

		        

		        System.out.println(buff);       

		        /*새로운 시트에 Row 만들어서 버퍼 데이터 올리기*/

		        newrow = newsheet.createRow(rowindex-1);

		        newrow.createCell(0).setCellValue(""+buff);

		        

		        /*버퍼에 있는 데이터 지워주기*/

		        buff.delete(0, buff.length());

		    }

		    	/*파일 생성하기*/

		    	FileOutputStream outFile;    	   

				outFile = new FileOutputStream("D:\\Users\\...\\Parsing\\parseDatas.xlsx");

				newworkbook.write(outFile);

				outFile.close();

				System.out.println("File Creation Complete");

			

		}catch(Exception e) {

		    e.printStackTrace();

		}
	}

}
