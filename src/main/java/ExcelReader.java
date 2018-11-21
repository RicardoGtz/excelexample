import java.io.*;
import java.rmi.server.ExportException;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

public class ExcelReader{
    private XSSFWorkbook workbook;
    private XSSFWorkbook newWorkbook;

    public ExcelReader(String filePath)throws IOException,ExcelReaderException{
        File file=new File(filePath);
        FileInputStream fIP = new FileInputStream(file);
        workbook = new XSSFWorkbook(fIP);
    }

    public void clearExcel(FileOutputStream fos, int sheetIndex)throws IOException,ExcelReaderException{
        //Creates the new workbook
        newWorkbook= new XSSFWorkbook();
        //Creates a new sheet
        XSSFSheet newSheet=newWorkbook.createSheet("Sheet 1");
        //Gets a sheet form the workbook
        XSSFSheet sheet=workbook.getSheetAt(sheetIndex);
        //Gets the last index row from the sheet
        int lastRow=sheet.getLastRowNum();
        //Stores de header row
        XSSFRow headerRow=null;
        //Stores the number of columns
        int numOfColumns=0;
        //Verifies if the sheet isn't empty
        if(sheet.getPhysicalNumberOfRows()>0){
            //Gives the correct format to the header if necessary
            headerRow=formatTitleHeader(sheet.getRow(0),newSheet);
            numOfColumns=headerRow.getLastCellNum();
        }else{
            throw new ExcelReaderException("Empty file sheet exception");
        }
        XSSFRow auxRow=null;
        XSSFRow newRow=null;
        //Iterates through the rows
        for(int i=1;i<lastRow+1;i++){
            auxRow=sheet.getRow(i);
            //Verifies if has the same number of cells
            if (auxRow.getLastCellNum() == numOfColumns) {
                newRow=newSheet.createRow(i);
                //Iterates through the cells
                XSSFCell auxCell=null;
                for(int j=0;j<auxRow.getLastCellNum();j++){
                    auxCell=auxRow.getCell(j);
                    //Verifies if cell isn't blank
                    if(auxCell!=null){
                       /* XSSFCellStyle style = workbook.createCellStyle();
                        style.setDataFormat();*/
                        //XSSFCellStyle style=auxCell.getCellStyle();

                        //Gets the type of the cell
                       CellType type=auxCell.getCellType();
                       //Verifies if cell is Numeric type(numbers, fractions, dates)
                       if(type == CellType.NUMERIC) {
                           //System.out.println(auxCell.toString());

                           //Verifies if is date formatted
                           if(HSSFDateUtil.isCellDateFormatted(auxCell)) {
                               //System.out.println("Es Fecha");
                               newRow.createCell(j).setCellValue(auxCell.toString());
                           }else{
                               newRow.createCell(j).setCellValue(Double.parseDouble(auxCell.toString()));
                           }
                       //Verifies if is a Formula
                       }else if(type == CellType.FORMULA){
                           newRow.createCell(j).setCellValue(auxCell.getNumericCellValue());
                       }else{
                           newRow.createCell(j).setCellValue(auxCell.toString());
                       }
                    }else{
                        throw new ExcelReaderException("Blank cell at row:"+(i+1)+" column:"+(j+1));
                    }
                }
            } else {
              throw new ExcelReaderException("Blank row or wrong number of cells at row: " + (i + 1));
            }
        }
        //Writes the new file
        newWorkbook.write(fos);
        fos.close();
    }

    private XSSFRow formatTitleHeader(XSSFRow row,XSSFSheet nsheet) throws ExcelReaderException{
        //Creates a new row
        XSSFRow newRow=nsheet.createRow(0);
        //Iterates through the cells
        for(int i=0;i<row.getLastCellNum();i++){
            XSSFCell cell=row.getCell(i);
            //Verifies if cell isn't empty
            if(cell!=null) {
                /* Verifies format i.e. "123.2" or "123", if true, throws an
                 * exception to notify this error  */
                if (cell.toString().matches("\\d*\\.*\\d*"))
                    throw new ExcelReaderException("Title cells must be alphanumeric. Format problem at row:0 column:"+(i+1));
                else
                    newRow.createCell(i).setCellValue(cell.toString());
            }else{
                throw new ExcelReaderException("Blank cell at row:0 column:"+(i+1));
            }
        }
        return newRow;
    }
}
