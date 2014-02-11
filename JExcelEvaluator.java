import java.io.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.ss.usermodel.*;

public class JExcelEvaluator {

    public static void main(String[] args) {
        if (args.length != 3) {
            System.err.println("Usage: JExcelEvaluator <model> <inputs> <outputs>");
            System.exit(-1);
        }

        try {
            InputStream model = new FileInputStream(args[0]);

            Workbook wb = new HSSFWorkbook(model);
            Sheet sheet = wb.getSheetAt(0);
            String line;
            String pos;

            InputStream input = new FileInputStream(args[1]);
            BufferedReader  bri = new BufferedReader(new InputStreamReader(input));
            while ((line = bri.readLine()) != null) {
                String[] parts = line.split(" ");
                pos = parts[0];
                Double val = null;

                try {
                    val = Double.parseDouble(parts[1]);
                } catch (NumberFormatException e) {
                    System.err.println("Cannot parse " + parts[1]);
                    System.exit(-1);
                }

                CellReference cr = new CellReference(pos);
                Row row = sheet.getRow(cr.getRow());
                Cell cell = row.getCell(cr.getCol());
                cell.setCellValue(val);
            }
            bri.close();
            input.close();

            // Excel caches previously calculated results and you need to trigger recalculation to updated them
            // Ref: http://poi.apache.org/spreadsheet/eval.html
            wb.getCreationHelper().createFormulaEvaluator().evaluateAll();

            InputStream output = new FileInputStream(args[2]);
            BufferedReader bro = new BufferedReader(new InputStreamReader(output));
            while ((pos = bro.readLine()) != null) {
                CellReference cr = new CellReference(pos);
                Row row = sheet.getRow(cr.getRow());
                Cell cell = row.getCell(cr.getCol());
                try {
                    System.out.println(pos+" "+cell.getNumericCellValue());
                } catch (NumberFormatException e) {
                    System.err.println("Value of " + pos + "isnot a valid number");
                    System.exit(-1);
                }
            }
            bro.close();
            output.close();

            FileOutputStream fileOut = new FileOutputStream(args[0]);
            wb.write(fileOut);
            fileOut.close();

        } catch (IOException e) {
            System.err.println("Error: " + e);
            System.exit(-1);
        }
    }

}
