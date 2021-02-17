package result;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class OrganisationResultFile extends ResultXLSXFile {
    public OrganisationResultFile(String root, String country) {
        super(root, country, ResultFileType.ORGANISATION);
    }

    private Integer currentCellIndex = 0;
    private Cell getCurrentCell(Row row) {
        return row.createCell(currentCellIndex++);
    }
    @Override
    protected void fillRow(Row row, String companyName, String securityName, String shareholderName, Double percentOwned, String irLink) {

        currentCellIndex = 0;
        this.getCurrentCell(row).setCellValue(companyName);
        this.getCurrentCell(row).setCellValue(securityName);
        this.getCurrentCell(row).setCellValue(shareholderName);

        if (percentOwned != null) {
            var percentCell = this.getCurrentCell(row);
            percentCell.setCellValue(percentOwned);
            percentCell.setCellStyle(this.getBuiltInCellStyle(10));
        }
        else
            currentCellIndex++;

        this.getCurrentCell(row).setCellValue("");
        this.getCurrentCell(row).setCellValue("");
        this.getCurrentCell(row).setCellValue("");
    }

    @Override
    protected String[] getTableHeaders() {
        return new String[] {"Company name", "Security name", "Shareholder name", "% owned", "", "Analysis result", "Supporting link"};
    }
}

