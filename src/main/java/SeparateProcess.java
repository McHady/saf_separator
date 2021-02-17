import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import result.IndividualResultFile;
import result.OrganisationResultFile;
import result.ResultXLSXFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Arrays;

public class SeparateProcess implements Runnable {

    private final String[] companyNameParts;
    private final String file;
    private final String filesRoot;
    private final String country;

    public SeparateProcess(String[] companyNameParts, String filesRoot, String file) {
        this.companyNameParts = companyNameParts;
        this.filesRoot = filesRoot;
        this.file = file;
        this.country = this.getCountryByFileName(file);
        System.out.println("Starting for " + country);
    }

    public String getCountryByFileName(String filename){
        return new File(filename).getName().replace("list.xlsx", "").trim();
    }

    @Override
    public void run() {

        try (var xlsxListStream = new FileInputStream(this.file)) {

            try (var listWorkBook = new XSSFWorkbook(xlsxListStream)) {

                var sheet = listWorkBook.getSheetAt(0);
                var rowIterator = sheet.rowIterator();

                var firstRow = true;
                while (rowIterator.hasNext()) {

                    var row = rowIterator.next();
                    if (firstRow) {
                        firstRow = false;
                        continue;
                    }

                    var companyName = row.getCell(CellReference.convertColStringToIndex("A")).getStringCellValue();
                    var securityName = row.getCell(CellReference.convertColStringToIndex("B")).getStringCellValue();
                    var irLink = row.getCell(CellReference.convertColStringToIndex("F")).getStringCellValue();

                    if (companyName != null && securityName != null) {

                        try (var shareholderFileStream = new FileInputStream(this.filesRoot + "/" + this.country + "/" + securityName + ".xlsx")) {

                            try (var shareholderWorkbook = new XSSFWorkbook(shareholderFileStream)) {

                                System.out.println("Getting data for " + securityName);
                                var shSheet = shareholderWorkbook.getSheetAt(0);
                                var shRowIterator = shSheet.rowIterator();

                                var first = true;
                                while (shRowIterator.hasNext()) {
                                    var shRow = shRowIterator.next();

                                    if (first) {
                                        first = false;
                                        continue;
                                    }

                                    var shareholderName = shRow.getCell(CellReference.convertColStringToIndex("A")).getStringCellValue();

                                    Double percent = null;
                                    var percentCell = shRow.getCell(CellReference.convertColStringToIndex("C"));

                                    if (percentCell != null && percentCell.getCellType() != CellType.BLANK)
                                        percent = percentCell.getNumericCellValue();

                                    var resultFile = getResultFileByShareholderName(shareholderName);

                                    if (shareholderName != null && percent != null) {
                                        resultFile.addRow(companyName, securityName, shareholderName, percent, irLink);
                                    }
                                }
                            }
                        } catch (FileNotFoundException e) {
                            System.out.println("Can't find file for " + securityName);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            this.individual.close();
            this.organisation.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private ResultXLSXFile individual;
    private ResultXLSXFile organisation;

    private ResultXLSXFile getResultFileByShareholderName(String securityName) {

        var isCompanyName = Arrays.stream(this.companyNameParts).anyMatch(securityName::contains);
        return isCompanyName ? getOrganisation() : getIndividual();
    }

    private ResultXLSXFile getIndividual(){

        return this.individual != null ? this.individual : (this.individual = new IndividualResultFile(filesRoot + "/result", country));
    }

    private ResultXLSXFile getOrganisation(){

        return this.organisation != null ? this.organisation : (this.organisation = new OrganisationResultFile(filesRoot + "/result", country));
    }
}
