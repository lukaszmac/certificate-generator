package io.nanovc.certificates.office.powerpoint;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;

import java.nio.file.Path;

/**
 * A PowerPoint Presentation.
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation">Presentation Object</a>
 */
public class Presentation
{
    /**
     * The presentation ActiveX Component that we are wrapping.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation">Presentation Object</a>
     */
    protected ActiveXComponent presentationObject;

    /**
     * Creates a new presentation wrapper for the given presentation object.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation">Presentation Object</a>
     * @param presentationObject The presentation ActiveX Component that we are wrapping.
     */
    public Presentation(ActiveXComponent presentationObject)
    {
        this.presentationObject = presentationObject;
    }

    /**
     * Saves the presentation as a new file.
     * @param path The path to save the presentation to.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas">SaveAs Method</a>
     */
    public void saveAs(Path path)
    {
        // Get the absolute path:
        String fullPath = path.toAbsolutePath().normalize().toString();

        // Save the presentation:
        this.presentationObject.invoke("SaveAs", fullPath);
    }

    /**
     * Saves the presentation as a new file, possibly changing the file format.
     * @param path The path to save the presentation to.
     * @param saveAsFileType The type of file to save this presentation as.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas">SaveAs Method</a>
     */
    public void saveAs(Path path, SaveAsFileType saveAsFileType)
    {
        // Get the absolute path:
        String fullPath = path.toAbsolutePath().normalize().toString();

        // Save the presentation:
        this.presentationObject.invoke("SaveAs", fullPath, saveAsFileType.getConstant());
    }

    /**
     * Saves the presentation as a new file, possibly changing the file format.
     * @param path The path to save the presentation to.
     * @param saveAsFileType The type of file to save this presentation as.
     * @param embedFonts     Defines whether to embed fonts or not.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas">SaveAs Method</a>
     */
    public void saveAs(Path path, SaveAsFileType saveAsFileType, boolean embedFonts)
    {
        // Get the absolute path:
        String fullPath = path.toAbsolutePath().normalize().toString();

        // Save the presentation:
        this.presentationObject.invoke("SaveAs", new Variant(fullPath), new Variant(saveAsFileType.getConstant()), new Variant(embedFonts ? TriState.msoTrue.getConstant() : TriState.msoFalse.getConstant()));
    }

    /**
     * Closes the presentation.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.close">Close Method</a>
     */
    public void close()
    {
        this.presentationObject.invoke("Close");
    }
}
