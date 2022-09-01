import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import static java.sql.Types.NUMERIC;

public class XlsxReader {

    public static void main(String[] args) {

        try {

            // Get data from the file
            FileInputStream file = new FileInputStream(new File("vzorek_dat.xlsx"));

            // Created WorkBook instance to refer xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Created sheet object to get the object
            XSSFSheet sheet = workbook.getSheetAt(0);

            // i will use for each loop to iterate over row
            for (Row row : sheet) {

                System.out.println(row.getCell(1));

                }
                for(Row row : sheet){

                    Iterator<Cell> cellItr = row.iterator();
                    while(cellItr.hasNext()){
                        System.out.println(cellItr.next().toString());
                    }
                }


        } catch (FileNotFoundException e) {

            e.printStackTrace();

        } catch (IOException ex) {

            ex.printStackTrace();

        }


    }

}
