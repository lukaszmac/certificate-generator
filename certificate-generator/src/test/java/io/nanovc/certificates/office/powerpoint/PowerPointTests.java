package io.nanovc.certificates.office.powerpoint;

import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Tests {@link PowerPoint}.
 */
class PowerPointTests
{
    @Test
    public void startQuitTest() throws Exception
    {
        try (var powerPoint = new PowerPoint())
        {
            powerPoint.start();
            powerPoint.quit();
        }
    }

    @Test
    public void savePowerPointToPDF() throws Exception
    {
        // Define the paths:
        Path presentationPath = Paths.get("../example-template/example-template.pptx");
        Path pdfFullPath = presentationPath.getParent().resolve("java-pptx-to-pdf.pdf");

        try (var powerPoint = new PowerPoint())
        {
            powerPoint.start();

            Presentation presentation = powerPoint.openPresentation(presentationPath);

            presentation.saveAs(pdfFullPath, SaveAsFileType.ppSaveAsPDF, true);

            presentation.close();

            powerPoint.quit();
        }
    }
}
