package io.nanovc.certificates.office.powerpoint;

/**
 * The Tri State Enumeration.
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/office.msotristate">TriState Enumeration</a>
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas">SaveAs Method</a>
 */
public enum TriState
{
    msoTrue(	-1, "True"),
    msoFalse(	0, "False"),
    msoCTrue(	1,"Not supported"),
    msoTriStateMixed(	-2, "Not supported"),
    msoTriStateToggle(	-3,"Not supported"),

    ;

    /**
     * The constant to use for PowerPoint interop.
     */
    private int constant;

    /**
     * The comment associated with the value.
     */
    private String comment;

    /**
     * @param constant The constant to use for PowerPoint interop.
     * @param comment The comment associated with the value.
     */
    TriState(int constant, String comment)
    {
        this.constant = constant;
        this.comment = comment;
    }

    /**
     * Gets the constant to use for PowerPoint interop.
     * @return The constant to use for PowerPoint interop.
     */
    public int getConstant()
    {
        return this.constant;
    }

    /**
     * Gets the comment associated with the value.
     * @return The comment associated with the value.
     */
    public String getComment()
    {
        return comment;
    }
}
