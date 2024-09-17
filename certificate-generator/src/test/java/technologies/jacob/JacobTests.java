package technologies.jacob;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Tests the JACOB library which is a java to COM bridge and allows ActiveX automation in Windows.
 * https://github.com/freemansoft/jacob-project
 * https://github.com/freemansoft/jacob-project/blob/main/samples/com/jacob/samples/office/WordDocumentProperties.java
 * https://central.sonatype.com/artifact/net.sf.jacob-project/jacob
 */
public class JacobTests
{
    @Test
    public void startAndStopPowerPoint()
    {
        // https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint
        ActiveXComponent activeXComponent = new ActiveXComponent("PowerPoint.Application");
        activeXComponent.invoke("Quit");
    }

    @Test
    public void savePowerPointToPDF()
    {
        /*
        PP = new ActiveXObject("PowerPoint.Application");
        PRSNT = PP.presentations.Open(source,0,0,0)
        //PRSNT.SaveCopyAs(target,32);
        //https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/ppsaveasfiletype-enumeration-powerpoint
        PRSNT.SaveAs(target,32);
        PRSNT.Close();
        PP.Quit();
         */

        // Get the path that we want:
        Path presentationPath = Paths.get("../example-template/example-template.pptx");
        String presentationFullPath = presentationPath.toAbsolutePath().normalize().toString();
        String pdfFullPath = presentationPath.getParent().resolve("example-template.pdf").toAbsolutePath().normalize().toString();

        // https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint
        ActiveXComponent powerPointApplication = new ActiveXComponent("PowerPoint.Application");

        // Get the presentations collection:
        var presentations = powerPointApplication.getPropertyAsComponent("Presentations");

        // Load a presentation:
        var presentation = presentations.invokeGetComponent("Open", new Variant(presentationFullPath), new Variant(0), new Variant(0), new Variant(0));

        // Save the presentation
        presentation.invoke("SaveAs", pdfFullPath, 32);

        // Close the presentation:
        presentation.invoke("Close");

        // Close PowerPoint:
        powerPointApplication.invoke("Quit");
    }
}
