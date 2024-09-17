package io.nanovc.certificates.office.powerpoint;

/**
 * The Save As File Type Enumeration.
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype">SaveAs Enumeration</a>
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas">SaveAs Method</a>
 */
public enum SaveAsFileType
{
    ppSaveAsAddIn	(8),
    ppSaveAsAnimatedGIF	(40),
    ppSaveAsBMP	(19),
    ppSaveAsDefault	(11),
    ppSaveAsEMF	(23),
    ppSaveAsExternalConverter	(64000),
    ppSaveAsGIF	(16),
    ppSaveAsJPG	(17),
    ppSaveAsMetaFile	(15),
    ppSaveAsMP4	(39),
    ppSaveAsOpenDocumentPresentation	(35),
    ppSaveAsOpenXMLAddin	(30),
    ppSaveAsOpenXMLPicturePresentation	(36),
    ppSaveAsOpenXMLPresentation	(24),
    ppSaveAsOpenXMLPresentationMacroEnabled	(25),
    ppSaveAsOpenXMLShow	(28),
    ppSaveAsOpenXMLShowMacroEnabled	(29),
    ppSaveAsOpenXMLTemplate	(26),
    ppSaveAsOpenXMLTemplateMacroEnabled	(27),
    ppSaveAsOpenXMLTheme	(31),
    ppSaveAsPDF	(32),
    ppSaveAsPNG	(18),
    ppSaveAsPresentation	(1),
    ppSaveAsRTF	(6),
    ppSaveAsShow	(7),
    ppSaveAsStrictOpenXMLPresentation	(38),
    ppSaveAsTemplate	(5),
    ppSaveAsTIF	(21),
    ppSaveAsWMV	(37),
    ppSaveAsXMLPresentation	(34),
    ppSaveAsXPS	(33),

    ;

    /**
     * The constant to use for PowerPoint interop.
     */
    private int constant;

    SaveAsFileType(int constant)
    {
        this.constant = constant;
    }

    /**
     * Gets the constant to use for PowerPoint interop.
     * @return The constant to use for PowerPoint interop.
     */
    public int getConstant()
    {
        return this.constant;
    }
}
