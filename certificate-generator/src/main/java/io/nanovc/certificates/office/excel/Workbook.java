package io.nanovc.certificates.office.excel;

import com.jacob.activeX.ActiveXComponent;

/**
 * An Excel Workbook.
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.workbook">Workbook Object</a>
 */
public class Workbook
{
    /**
     * The workbook ActiveX Component that we are wrapping.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.workbook">Workbook Object</a>
     */
    protected ActiveXComponent workbookObject;

    /**
     * The worksheets for the workbook.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.worksheets">Worksheets Object</a>
     */
    public ActiveXComponent worksheets;

    /**
     * The worksheets for the workbook.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.names">Names Object</a>
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.names">Names Collection</a>
     */
    public ActiveXComponent names;

    /**
     * Creates a new workbook wrapper for the given workbook object.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.workbook">Workbook Object</a>
     * @param workbookObject The workbook ActiveX Component that we are wrapping.
     */
    public Workbook(ActiveXComponent workbookObject)
    {
        this.workbookObject = workbookObject;

        // Get items of interest:
        this.worksheets = this.workbookObject.getPropertyAsComponent("Worksheets");
        this.names = this.workbookObject.getPropertyAsComponent("Names");
    }

    /**
     * Closes the workbook.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.close">Close Method</a>
     */
    public void close()
    {
        this.workbookObject.invoke("Close");
    }
}
