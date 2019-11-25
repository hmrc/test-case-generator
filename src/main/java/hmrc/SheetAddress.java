package hmrc;

public class SheetAddress implements Comparable<SheetAddress> {
    private String sheetName;
    private String code;
    private String address;
    private String instance;
    private int order;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getInstance() {
        return instance;
    }

    @Override
    public int compareTo(SheetAddress d) {
        return this.order - d.getOrder();
    }

    public void setInstance(String instance) {
        this.instance = instance;
    }

    public int getOrder() {
        return order;
    }

    public void setOrder(int order) {
        this.order = order;
    }
}
