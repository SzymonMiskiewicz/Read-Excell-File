import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class XlsxReader {

    public static void main(String[] args) {


        try {

            // Get data from the file
            FileInputStream file = new FileInputStream(new File("vzorek_dat.xlsx"));

            // Created WorkBook instance to refer xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Created sheet object to get the object
            XSSFSheet sheet = workbook.getSheetAt(0);

            // We can use DataFormatter to fetch the string value of an Excel cell.
            DataFormatter formatter = new DataFormatter();

            // i will use for each loop to iterate over row
            try {

            for (Row row : sheet) {

                for (Cell cell : row) {

                    String strValue = formatter.formatCellValue(cell);
                    int number;
                    number = Integer.parseInt(strValue);

                    // Check if number is less than
                    // equal to 1
                    if (number<= 1)

                        System.out.println(number + "it is not prime number");

                        // Check if number is 2
                    else if (number== 2)
                        System.out.println( number + "it is prime number");

                        // Check if n is a multiple of 2
                    else if (number% 2 == 0)
                        System.out.println(number + "it is not prime number");

                    // If not, then just check the odds
                    for (int i = 3; i <= Math.sqrt(number); i += 2)
                    {
                        if (number % i == 0)
                            System.out.println(number + "it is not prime number");
                    }
                    System.out.println(row.getCell(1));

                }


            }

            } catch (NumberFormatException exc) {

                exc.printStackTrace();
            }

            for (Row row : sheet) {

                Iterator<Cell> cellItr = row.iterator();

                while (cellItr.hasNext()) {

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
