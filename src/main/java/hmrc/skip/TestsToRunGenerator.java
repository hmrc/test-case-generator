package hmrc.skip;

import hmrc.BaseSheetGenerator;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class TestsToRunGenerator extends BaseSheetGenerator {

    private DataFormatter dataFormatter;
    private Workbook workbook;

    public TestsToRunGenerator(DataFormatter dataFormatter, Workbook workbook) {
        this.dataFormatter = dataFormatter;
        this.workbook = workbook;
    }

    @Override
    public Workbook generate() {

        Sheet testCaseIndex = workbook.getSheet("Test case index");

        if(testCaseIndex != null) {

            List<String> testCases = new ArrayList<>();

            for (Row row : testCaseIndex) {
               Cell testCaseCell = row.getCell(0);
               Cell testCaseIsUsedCell = row.getCell(6);

               if(testCaseCell != null && testCaseIsUsedCell != null){
                   String testCaseNumber = dataFormatter.formatCellValue(testCaseCell);
                   String isUsed = dataFormatter.formatCellValue(testCaseIsUsedCell);

                   if(isUsed.equalsIgnoreCase("YES")){
                       testCases.add(testCaseNumber);
                   }
               }
            }

            print(String.join(",", testCases));
        }

        return workbook;
    }
}
