import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class XlsxReader {

    public static void main(String[] args) {

        try {
            // The new file instance created
            File myFile = new File("D:\\Prime Numbers Project\\vzorek_dat.xlsx");

            // Get data from the file
            FileInputStream file = new FileInputStream(new File(String.valueOf(myFile)));

            // Created WorkBook instance to refer xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Created sheet object to get the object
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Evaluating cell type
            FormulaEvaluator formEva = workbook.getCreationHelper().createFormulaEvaluator();

            // i will use for each loop to iterate over row
            for (Row row : sheet) {

                // i will use for each loop to iterate over cell
                for (Cell cell : row) {

                    switch (cell.getCellType()) {

                        case STRING:

                            System.out.println(cell.getRichStringCellValue().getString());

                            break;

                        case NUMERIC:

                            System.out.print(cell.getNumericCellValue());

                            break;

                        case BOOLEAN:

                            System.out.println(cell.getBooleanCellValue());

                            break;

                        case FORMULA:

                            System.out.println(cell.getCellFormula());

                            break;

                        case BLANK:

                            System.out.println();

                            break;

                    }


                }
            }

        } catch (FileNotFoundException e) {

            e.printStackTrace();

        } catch (IOException ex) {

            ex.printStackTrace();

        } /*catch (InvalidFormatException exc) {

            exc.printStackTrace();
        }*/
        

    }

}
