package hmrc;

import org.apache.poi.ss.usermodel.Workbook;

public interface SheetGenerator {
    Workbook generate();
}
