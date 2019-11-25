package hmrc;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Collectors;

public abstract class BaseSheetGenerator implements SheetGenerator {

    protected boolean generateNewSheetValues(Sheet sheet, Map<String, SheetAddress> dataMap, boolean first){

        for (SheetAddress sheetAddress : sortByCode(dataMap).values()){

            int newRowNumber = 0;

            if(!first)
                newRowNumber = sheet.getLastRowNum() + 1;

            Row newRow = sheet.createRow(newRowNumber);
            Cell stageCell = newRow.createCell(0);
            stageCell.setCellValue(sheetAddress.getCode());
            Cell valueCell = newRow.createCell(1);
            valueCell.setCellType(CellType.FORMULA);
            valueCell.setCellFormula(createFormula(sheetAddress));

            first = false;
        }

        return first;
    }

    protected String createFormula(SheetAddress sheetAddress){
        if(sheetAddress.getAddress() != null)
            return "\'" + sheetAddress.getSheetName() + "\'!" + sheetAddress.getAddress();
        return "NA()";
    }

    protected static Map<String, SheetAddress> sortByCode(final Map<String, SheetAddress> sheetAddressMap) {
        return sheetAddressMap.entrySet()
                .stream()
                //.sorted((e1, e2) -> e2.getValue().getCode().compareTo(e1.getValue().getCode()))
                .sorted((e1, e2) -> e1.getKey().compareTo(e2.getKey()))
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e1, LinkedHashMap::new));
    }



    protected int getMaxColumnIndex(Sheet sheet){
        int max = 0;

        for(Row row: sheet){
            if(row.getLastCellNum() >  max)
                max = row.getLastCellNum();
        }

        return max;
    }

    protected boolean isEmptyString(String string) {
        return string == null || string.isEmpty();
    }
    protected static void print(String value){
        System.out.println(value);
    }
}
