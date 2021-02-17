package result;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public abstract class ResultXLSXFile implements Closeable {

    private final String root;
    private final String country;
    private final ResultFileType fileType;
    private final String fileName;
    private Integer currentRow = 0;

    public ResultXLSXFile(String root, String country, ResultFileType fileType) {

        this.root = root;
        this.country = country;
        this.fileType = fileType;
        this.fileName = root + "/" + this.country + " " + this.fileType.asString() + " shareholder analysis to do.xlsx";
    }

    public void addRow(String companyName, String securityName, String shareholderName, Double percentOwned, String irLink) {

        var workbook = getCurrentWorkBook();

        var sheet = workbook.getSheetAt(0);
        var row = sheet.createRow(this.currentRow++);
        this.fillRow(row, companyName, securityName, shareholderName, percentOwned, irLink);
        this.updateWorkbook();
    }

    protected abstract void fillRow(Row row, String companyName, String securityName, String shareholderName, Double percentOwned, String irLink);

    private Workbook _currentWorkbook;

    protected abstract String[] getTableHeaders();

    private Workbook getCurrentWorkBook() {

        if (_currentWorkbook == null && currentRow == 0) {

            _currentWorkbook = new XSSFWorkbook();

            var sheet = _currentWorkbook.createSheet();
            var row = sheet.createRow(currentRow++);
            var index = 0;

            for (var header : this.getTableHeaders()) {
                var cell = row.createCell(index++);
                cell.setCellValue(header);
            }
        }

        this.updateWorkbook();

        return _currentWorkbook;
    }

    private void updateWorkbook() {

        try (var stream = new FileOutputStream(this.fileName)) {
            this._currentWorkbook.write(stream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    protected CellStyle getBuiltInCellStyle(Integer number) {
        var style = this._currentWorkbook.createCellStyle();
        style.setDataFormat(this._currentWorkbook.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(number)));

        return style;
    }

    @Override
    public void close() throws IOException {
        this._currentWorkbook.close();
    }
}

