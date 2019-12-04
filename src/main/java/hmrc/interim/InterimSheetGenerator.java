package hmrc.interim;

import hmrc.BaseSheetGenerator;
import hmrc.SheetAddress;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class InterimSheetGenerator extends BaseSheetGenerator {
    private DataFormatter dataFormatter;
    private Workbook workbook;

    public InterimSheetGenerator(DataFormatter dataFormatter, Workbook workbook) {
        this.dataFormatter = dataFormatter;
        this.workbook = workbook;
    }

    @Override
    public Workbook generate(){
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Sheet interimSheet = createSheet(workbook,"TaxCalc_Interim_JUNIT");

        boolean first = true;

        for(Sheet sheet: workbook){

            if(!isStageSheet(sheet.getSheetName()))
                continue;

            Map<String, SheetAddress> stageMap = new HashMap<>();

            for (Row row : sheet) {
                for (Cell cell : row) {

                    if(cell.getAddress().getColumn() != 0)
                        continue;

                    String val = dataFormatter.formatCellValue(cell);

                    if(!isEmptyString(val) && isStageRef(val)){
                        if(!stageMap.containsKey(cell.toString())){
                            SheetAddress sheetAddress = new SheetAddress();
                            sheetAddress.setCode(val);
                            sheetAddress.setSheetName(sheet.getSheetName());

                            stageMap.put(val, sheetAddress);
                        }
                    }
                }
            }

            if(sheet.getSheetName().equalsIgnoreCase("NS income (c1)"))
                stageMap.putAll(addSecondEmployments());

            for (Row row : sheet) {
                for (Cell cell : row) {

                    if(cell.getAddress().getColumn() == 0)
                        continue;

                    CellValue cellValue = evaluator.evaluate(cell);

                    if(cellValue == null)
                        continue;

                    String val = cellValue.getStringValue();

                    if(stageMap.containsKey(val))
                    {
                        SheetAddress sheetAddress = stageMap.get(val);
                        Cell dataCell = row.getCell(cell.getColumnIndex() + 1);

                        if(dataCell != null){
                            sheetAddress.setAddress(dataCell.getAddress().formatAsString());
                        }

                        if(val.equalsIgnoreCase("c1.5a") ||
                                val.equalsIgnoreCase("c1.5b") ||
                                val.equalsIgnoreCase("c1.5c")){
                            SheetAddress sheetAddress1 = stageMap.get(val+"-2");
                            Cell dataCell1 = row.getCell(cell.getColumnIndex() + 2);

                            if(dataCell1 != null){
                                sheetAddress1.setAddress(dataCell1.getAddress().formatAsString());
                            }
                        }
                    }
                }
            }

            first = generateNewSheetValues(interimSheet, stageMap, first);
        }

        return workbook;
    }

    private Map<String,SheetAddress> addSecondEmployments(){
        Map<String,SheetAddress> sheetAddresses = new HashMap<>();

        SheetAddress sheetAddress = new SheetAddress();
        sheetAddress.setCode("c1.5a-2");
        sheetAddress.setSheetName("NS income (c1)");

        SheetAddress sheetAddress1 = new SheetAddress();
        sheetAddress1.setCode("c1.5b-2");
        sheetAddress1.setSheetName("NS income (c1)");

        SheetAddress sheetAddress2 = new SheetAddress();
        sheetAddress2.setCode("c1.5c-2");
        sheetAddress2.setSheetName("NS income (c1)");

        sheetAddresses.put("c1.5a-2", sheetAddress);
        sheetAddresses.put("c1.5b-2", sheetAddress1);
        sheetAddresses.put("c1.5c-2", sheetAddress2);

        return sheetAddresses;
    }

    private boolean isStageSheet(String sheetName){
        String pattern = "(.*)(\\(([c|C](\\d|[a-zA-Z])*)*\\))";
        return isMatch(sheetName, pattern);
    }

    private boolean isStageRef(String cellText){
        String pattern = "^(c(\\d|[a-zA-Z])*\\.[^\\s]*)$";
        return isMatch(cellText, pattern);
    }

    private boolean isMatch(String value, String pattern){
        Pattern r = Pattern.compile(pattern);
        Matcher m = r.matcher(value);
        return m.find();
    }
}
