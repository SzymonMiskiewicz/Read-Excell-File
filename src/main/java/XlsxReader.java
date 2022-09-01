import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class XlsxReader {

    public static void main(String[] args) {

        // The new file instance created
        File myFile = new File("D:\\Prime Numbers Project\\vzorek_dat.xlsx");

        // Get data from the file
        FileInputStream file = new FileInputStream(new File(myFile));

        // Created WorkBook instance to refer xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        // Created sheet object to get the object
        XSSFSheet sheet = workbook.getSheetAt(0);

        // Evaluating cell type
        FormulaEvaluator formEva = workbook.getCreationHelper().createFormulaEvaluator();

        // i will use for each loop to iterate over row
        for (Row row : sheet) {

            // i will use for each loop to iterate over cell
            for(Cell cell : row) {

                switch (cell.getCellType()) {

                    case CellType.STRING:

                        System.out.println(cell.getRichStringCellValue().getString());

                        break;

                    case CellType.NUMERIC:

                        System.out.print(cell.getNumericCellValue() );

                        break;

                    case BOOLEAN:

                        System.out.println(cell.getBooleanCellValue() );

                        break;

                    case CellType.FORMULA:

                        System.out.println(cell.getCellFormula());

                        break;

                    case CellType.BLANK:

                        System.out.println();

                        break;

                }


            }
        }
        

    }

}
