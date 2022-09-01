import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import static java.sql.Types.NUMERIC;

public class XlsxReader {

    public static void main(String[] args) {

        try {

            // Get data from the file
            FileInputStream file = new FileInputStream(new File("D:\\Read Excell File\\Read-Excell-File\\vzorek_dat.xlsx"));

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

                    CellType type = cell.getCellType();

                    if (type == CellType.NUMERIC) {

                        System.out.printf("[%d, %d] = NUMERIC; Value = %f%n",
                                cell.getRowIndex(), cell.getColumnIndex(),
                                cell.getNumericCellValue());

                    }

                }
            }

        } catch (FileNotFoundException e) {

            e.printStackTrace();

        } catch (IOException ex) {

            ex.printStackTrace();

        }


    }

}
