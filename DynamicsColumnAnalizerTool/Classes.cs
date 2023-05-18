using ClosedXML.Excel;

namespace Classes;

public class Column
{
    //data from the schema
    public string DisplayName { get; set; } 
    public string LogicalName { get; set; }
    public string SchemaName { get; set; }
    public string AttributeType { get; set; } //the datatype (can be Whole Number, Lookup, Text, ect)
    public string Type { get; set; } //simple / rollup (possibly more)
    public bool HasCustomAttribute { get; set; } 
    public string AdditionalData { get; set; } 
    public string Description { get; set; } 

    public IXLColumn ColumnData { get; set; } //IXLColumn from IXLWorksheet
    public int NumFilledRows { get; } //number of filled rows in the column
    public int NumRows { get; set; }
    public int SchemaRow { get; set; } //row on the schema table
    public string PercentFilled { get; }

    public Column(IXLColumn col)
    {
        ColumnData = col;
        
        NumFilledRows = findNumFulledRows();
        PercentFilled = ((float)NumFilledRows / (float)NumRows).ToString("#.##") + "%"

        //adjust the width of the column to match its contents (no more cutoff column names)
        col.AdjustToContents();
    }

    public string cell(int row)
    {
        return ColumnData.Cell(row);
    }

    //gets total rows with data (ignores whitespaces)
    private int findNumFulledRows()
    {
        int count = 0;
        for (int i = 1; i < NumRows; i++)
        {
            if (!string.IsNullOrWhiteSpace(this.cell(i)))
            {
                count++;
            }
        }
        return count;
    }
   

}

public class Table
{
    //data from the schema
    public string PluralDisplayName { get; set; }
    public string Entity { get; set; }
    public string Description { get; set; }
    public string SchemaName { get; set; }
    public string LogicalName { get; set; }
    public int ObjectTypeCode { get; set; }
    public bool IsCustomEntity { get; set; }
    public string OwnershipType { get; set; }

    
    public IXLWorksheet Schema { get; } //schema worksheet, rotated "90 degees" (column names iterate by row instead of by column)
    public List<Column> Columns { get; set; } //list of columns, Column[index].ColumnData.Cell(row) OR Column[index].Cell(row)
    public int NumRows { get; } //number of data entries, used to analyze the usefulnes of any column or table
    public int NumCols { get; } //number of columns in table
    public DateTime LastUpdated { get; } //last time the table was updated, not all tables have this

    public Table(IXLWorkbook dataTable, IXLWorkbook schema)
    {
        IXLWorksheet data = dataTable.Worksheet(1);
        Schema = schema.Worksheet(1);
        Entity = Schema.Cell(1, 2).Value.ToString();
        PluralDisplayName = Schema.Cell(2, 2).Value.ToString();
        Description = Schema.Cell(3, 2).Value.ToString();
        SchemaName = Schema.Cell(4, 2).Value.ToString();
        LogicalName = Schema.Cell(5, 2).Value.ToString();
        ObjectTypeCode = Int32.Parse(Schema.Cell(6, 2).Value);
        IsCustomEntity = Boolean.Parse(Schema.Cell(7, 2).Value);
        OwnershipType = Schema.Cell(8, 2).Value.ToString();

        //add IXLColumns to Columns list
        foreach (IXLColumn col in data.Columns())
        {
            Columns.Add(new Column(col))
        }
       
        NumCols = Columns.Count;
        NumRows = findNumRow();


        //read through schema and populate Column with metadata
        foreach (Column col in Columns)
        {
            col.DisplayName = col.cell(1); //get displayname from the top of the column
            col.SchemaRow = findMatchingSchemaRow(col.DisplayName); //gets the row in which the column metadata is stored in the schema file
            //gets all other column data 
            col.LogicalName = schema.Cell(col.schemaRow, 1).Value.ToString();  
            col.SchemaName = schema.Cell(col.schemaRow, 2).Value.ToString();
            col.AttributeType = schema.Cell(col.schemaRow, 4).Value.ToString();
            col.Description = schema.Cell(col.schemaRow, 5).Value.ToString();
            col.HasCustomAttribute = Boolean.Parse(schema.Cell(col.schemaRow, 6).Value);
            col.Type = schema.Cell(col.schemaRow, 7).Value.ToString();
            col.AdditionalData = schema.Cell(col.schemaRow, 8).Value.ToString();
        }
    }
    
    //assumes the Display Names are on column 3
    private int findMatchingSchemaRow(string columnDisplayName)
    {
        //skips the first ten rows
        //compares display name in schema to display name stored in the Column class
        for (int i = 10; i < MaxRow; i++)
        {
            if(Schema.Cell(i, 3).Value.ToString().Equals(columnDisplayName))
            {
                return i;
            }
        }
    }


    /* ignore this functionality for now, get it later when working on features 
    public void addColumns(IXLColumn col)
    {
        //this method is going to have to find all the data related to columns (display name logical name ect

        //relate columns by name
        //while()
        Column column = new Column(col);

        Columns.Add(column);
        
    }

    //update column by index
    public void updateColumn(int index)
    {

    }

    //update column by column object
    public void updateColumn(Column column)
    {

    }
    */
    public void analyzeTable()
    {
        //copy paste old code from v2 to analyze the table
    }

    //get the data from a specific cell, because Columns is a list, do col-1 otherwise if a user tries to access cell(row, 1) it will return the second column not the first, row should never be below 1 either 
    public string cell(int row, int col)
    {
        return Columns[col-1].cell(row).Value.ToString()
    }

    //internal method for finding the max row in a data table
    private int findNumRow()
    {
        int num = 0;
        // iterate through column (objects) 
        foreach (Column col in Columns)
        {
            num = 0;
            //while the cell in that column in that row is not empty iterate i and num 
            //this code assumes that there is at least one column that is fully populated (100% of rows filled)
            for (int i = 1; !string.IsNullOrWhiteSpace(col.cell(i).Value.ToString()); i++)
            {
                num++;
            }
        }
        return num;
    }

}