package io.nanovc.certificates.office.excel;

import com.jacob.com.Variant;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Tests {@link Excel}.
 */
class ExcelTests
{
    @Test
    public void startQuitTest() throws Exception
    {
        try (var excel = new Excel())
        {
            excel.start();
            excel.quit();
        }
    }

    @Test
    public void processTable() throws Exception
    {
        // Define the paths:
        Path excelPath = Paths.get("../certificate-generator-folders/Certificate Generator Data.xlsx");

        try (var excel = new Excel())
        {
            excel.start();

            Workbook workbook = excel.openWorkbook(excelPath);
            int workbookNameCount = workbook.names.getPropertyAsInt("Count");
            //var range = workbook.names.getProperty("Item");


            var worksheet = workbook.worksheets.invokeGetComponent("Item", new Variant(1));

            var worksheetNames = worksheet.getPropertyAsComponent("Names");
            int worksheetNameCount = worksheetNames.getPropertyAsInt("Count");

            //var range2 = worksheet.invokeGetComponent("Range", new Variant("Certificates_for_Printing"));
            var range2 = worksheet.invokeGetComponent("Range", new Variant("Certificates_for_Printing[#All]"));
            var range2RowCount = range2.getPropertyAsComponent("Rows").getPropertyAsInt("Count");
            var range2ColumnCount = range2.getPropertyAsComponent("Columns").getPropertyAsInt("Count");
            var cells11 = range2.invokeGetComponent("Cells", new Variant(1), new Variant(1));
            var value = cells11.getProperty("Value");

            //var tableRange = worksheetNames.invokeGetComponent("Item", new Variant("Certificates_for_Printing"));
            var tableRange = worksheet.invokeGetComponent("Range", new Variant("A1"));

            //var tableRange = workbook.names.invokeGetComponent("Item", new Variant("Certificates_for_Printing[#All]"));
            //var tableRange = workbook.names.invokeGetComponent("Item", new Variant("Certificates_for_Printing"));
            //var tableRange = workbook.names.invokeGetComponent("Item", new Variant("A2"));
            //var tableRange = workbook.names.invoke("Item", "Certificates_for_Printing");
            //var tableRange = workbook.names.invoke("Item", new Variant("Certificates_for_Printing"), Variant.VT_MISSING, Variant.VT_MISSING);



            workbook.close();

            excel.quit();
        }
    }
}
