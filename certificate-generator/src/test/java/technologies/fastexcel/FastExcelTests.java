package technologies.fastexcel;


import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.stream.Stream;

/**
 * Tests the Fast Excel library.
 * https://github.com/dhatim/fastexcel
 */
public class FastExcelTests
{
    @Test
    public void printCurrentWorkingDirectory()
    {
        System.out.println(Paths.get(".").normalize().toAbsolutePath());
    }

    @Test
    public void simpleReading() throws IOException
    {
        try (InputStream is = Files.newInputStream(Paths.get("../example-template/data.xlsx")); ReadableWorkbook wb = new ReadableWorkbook(is))
        {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream())
            {
                rows.forEach(r -> {
                    Cell cell = r.getCell(0);
                    System.out.println("cell.asString() = " + cell.asString());
                    //                    BigDecimal num = r.getCellAsNumber(0).orElse(null);
                    //                    String str = r.getCellAsString(1).orElse(null);
                    //                    LocalDateTime date = r.getCellAsDate(2).orElse(null);
                });
            }
        }
    }

    @Test
    public void readingCertificateData() throws IOException
    {
        try (InputStream is = Files.newInputStream(Paths.get("../certificate-generator-folders/Certificate Generator Data.xlsx")); ReadableWorkbook wb = new ReadableWorkbook(is))
        {

            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream())
            {
                rows.forEach(r -> {
                    for (Cell cell : r)
                    {
                        System.out.print("|");
                        System.out.print(cell.getText());
                    }
                    System.out.print("|");
                    System.out.println();
                });
            }
        }
    }
}
