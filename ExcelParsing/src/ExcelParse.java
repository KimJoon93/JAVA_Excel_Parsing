import java.io.File;
import java.io.FileWriter;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.BufferedWriter;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ExcelParse {
		
	public static void main(String[] args) throws IOException{
		// TODO Auto-generated method stub
		
		try {
			
		    
		    XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("/Users/joon/Desktop/JAVA_Excel_Parsing/sample.xlsx")); 
		    int rowindex=0;

		    int columnindex=0;

		    int sheetindex=0;
		    
		    StringBuffer buff = new StringBuffer();

		    XSSFSheet sheet = workbook.getSheetAt(sheetindex);
		    
		    File csvfile = new File("/Users/joon/Desktop/JAVA_Excel_Parsing/parseSample.csv");
		    BufferedWriter fw = new BufferedWriter(new FileWriter(csvfile,true));

		    int rows=sheet.getPhysicalNumberOfRows();

		    for(rowindex=0;rowindex<rows;rowindex++){

		        XSSFRow row=sheet.getRow(rowindex);

		        if(row !=null){

		        	/*전부 끝 값까지 나와야하는데 getPhysicalNumberOfColumns() 사용하면 빈칸때문에 값이 다르게 나와서 끝 값까지 제대로 안나올 경우 발생하므로 하드코딩*/

		            int cells= 5;

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

		                    case FORMULA:                        

		                    	value=cell.getNumericCellValue()+"";

		                        break;

		                    case NUMERIC:

		                        value=cell.getNumericCellValue()+"";

		                        break;

		                    case STRING:

		                        value=cell.getStringCellValue()+"";

		                        break;

		                    case ERROR:

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

		        fw.write(buff.toString());
		        fw.newLine();
		        /*버퍼에 있는 데이터 지워주기*/
		        buff.delete(0, buff.length());

		    }

		    	fw.flush();
		    	workbook.close();
				System.out.println("File Creation Complete");

			

		}catch(Exception e) {

		    e.printStackTrace();

		}
	}

}
