package io.nanovc.certificates.generation;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;

public class ColumnCollection implements Iterable<Column>
{
    /**
     * The columns indexed by the column name.
     */
    private final LinkedHashMap<String, Column> columnsByName = new LinkedHashMap<>();

    /**
     * The columns indexed by their index.
     */
    private final ArrayList<Column> columnsByIndex = new ArrayList<>();

    /**
     * The table that these columns belong to.
     */
    private final Table table;

    public ColumnCollection(Table table) {this.table = table;}

    /**
     * Adds a column with the given name.
     * @param name The name of the column to add.
     * @return The column that was added.
     */
    public Column addColumn(String name)
    {
        // Create the column:
        Column column = new Column();
        column.index = this.columnsByIndex.size();
        column.name = name;
        this.columnsByIndex.add(column);
        this.columnsByName.put(name, column);
        return column;
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Column> iterator()
    {
        return this.columnsByIndex.iterator();
    }

    /**
     * Returns true if we have columns, making us a rectangular table. False if we don't have columns, meaning that we are a non-rectangular table.
     * @return True if we have columns, making us a rectangular table. False if we don't have columns, meaning that we are a non-rectangular table.
     */
    public boolean hasColumns()
    {
        return !this.columnsByIndex.isEmpty();
    }

    /**
     * Gets the number of columns that we have.
     * @return The number of columns that we have.
     */
    public int getColumnCount()
    {
        return this.columnsByIndex.size();
    }

    /**
     * Gets the column by name.
     * @param columnName The name of the column to get.
     * @return The column with the given name. Null if it's not found.
     */
    public Column getColumn(String columnName)
    {
        return this.columnsByName.get(columnName);
    }
}
