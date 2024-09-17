package io.nanovc.certificates.generation;

import java.util.*;

/**
 * A row in a table.
 */
public class Row implements Iterable<Object>
{
    /**
     * The table that this row belongs to.
     */
    private final Table table;

    /**
     * The values of the row as an array.
     * Use this when we have rectangular data.
     */
    private Object[] valueArray;

    /**
     * The values of the row as an array list.
     * Use this when we don't have rectangular data.
     */
    private ArrayList<Object> valueList;

    /**
     * The index of the row in the table.
     */
    private int rowIndex;

    /**
     * Creates a new row.
     * @param table The table that the row belongs to.
     * @param rowIndex The index of the row in the table.
     */
    public Row(Table table, int rowIndex) {this.table = table;}

    /**
     * Gets the value list.
     * It initializes it first if necessary.
     * @return The value list.
     */
    protected ArrayList<Object> getValueList()
    {
        if (this.valueList == null)
        {
            // Initialize the value list:
            this.valueList = new ArrayList<>();
        }
        return valueList;
    }

    /**
     * Checks whether we have a value list (without creating one unnecessarily).
     * @return True if we have a value list. False if we don't.
     */
    protected boolean hasValueList()
    {
        return this.valueList != null;
    }

    /**
     * Checks whether we have a value array (without creating one unnecessarily).
     * @return True if we have a value array. False if we don't.
     */
    protected boolean hasValueArray()
    {
        return this.valueArray != null;
    }

    /**
     * Appends a cell value at the end of the row.
     * @param value The value to add to the row.
     */
    public void appendCell(String value)
    {
        this.getValueList().add(value);
    }

    @Override public String toString()
    {
        StringJoiner joiner = new StringJoiner("|");
        if (this.hasValueList())
        {
            // We have a value list.

            // Add all the values:
            this.getValueList().forEach(value -> joiner.add(Objects.toString( value)));
        }
        else
        {
            // We don't have a value list.

            // Check if we have a value array:
            if (this.hasValueArray())
            {
                // We have a value array.

                // Add all the values:
                for (int i = 0; i < this.valueArray.length; i++)
                {
                    joiner.add(Objects.toString(this.valueArray[i]));
                }
            }
            else
            {
                // We don't have any values.
                joiner.add("EMPTY ROW");
            }
        }
        return joiner.toString();
    }

    /**
     * Gets the width (number of cells) in this row.
     * @return The width (number of cells) in this row.
     */
    public int getWidth()
    {
        if (hasValueArray()) return this.valueArray.length;
        else if (hasValueList()) return this.getValueList().size();
        else return 0;
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Object> iterator()
    {
        if (hasValueList()) return this.getValueList().iterator();
        else if (hasValueArray()) return Arrays.stream(this.valueArray).iterator();
        else return Collections.emptyIterator();
    }

    /**
     * Copies the values from another row.
     * @param row The row to copy values from.
     */
    public void copyValuesFromAnotherRow(Row row)
    {
        // Check whether we are rectangular:
        if (this.table.columns.hasColumns())
        {
            // We are a rectangular row because we have columns.

            // Get the row width we are expecting to be:
            int expectedRowWidth = this.table.columns.getColumnCount();

            // Check whether the other row matches exactly:
            if (row.getWidth() == expectedRowWidth)
            {
                // The other row matches the expected width exactly.

                // Set the value array for the row:
                this.valueArray = row.getSnapshotOfValues();

                // Clear the value list if any:
                this.valueList = null;
            }
            else
            {
                // The other row doesn't match the expected width exactly.

                // Add the values to our list:
                ArrayList<Object> list = this.getValueList();
                list.clear();
                if (row.hasValueArray())
                {
                    // The other row has an array.
                    for (Object cellValue : row.valueArray)
                    {
                        list.add(cellValue);
                    }
                }
                else if (row.hasValueList())
                {
                    // The other row has a value list.
                    list.addAll(row.getValueList());
                }

                // Clear the value array if any:
                this.valueArray = null;
            }
        }
        else
        {
            // We are not a rectangular row because we don't have columns.
        }

    }

    /**
     * Gets a snapshot of the values for the row.
     * @return A snapshot of the values for the row.
     */
    public Object[] getSnapshotOfValues()
    {
        if (this.hasValueArray())
        {
            return Arrays.copyOf(this.valueArray, this.valueArray.length);
        }
        else if (this.hasValueList())
        {
            return this.getValueList().toArray();
        }
        else return new Object[0];
    }

    /**
     * Gets the value of the column as a string.
     * @param columnName The name of the column that we want to get.
     * @return The value of the given column as a string.
     */
    public String getCellByColumnNameAsString(String columnName)
    {
        // Get the column:
        Column column = this.table.columns.getColumn(columnName);
        if (column != null)
        {
            // We have the column.
            // Get the value by index:
            return this.getCellByColumnIndexAsString(column.index);
        }
        else return "";
    }

    /**
     * Gets the value of the column as a string.
     * @param columnIndex The index of the column that we want to get.
     * @return The value of the given column as a string.
     */
    public String getCellByColumnIndexAsString(int columnIndex)
    {
        // Make sure the column index is in range:
        if (columnIndex >= this.getWidth() || columnIndex < 0) return "";

        // Check if we have values:
        if (hasValueArray())
        {
            // We have a value array.

            // Get the cell value:
            Object cellValue = this.valueArray[columnIndex];

            return Objects.toString(cellValue);
        }
        else if (hasValueList())
        {
            // We have a value list.

            // Get the cell value:
            Object cellValue = this.valueList.get(columnIndex);

            return Objects.toString(cellValue);
        }
        else
        {
            return "";
        }
    }
}
