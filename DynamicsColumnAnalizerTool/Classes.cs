using ClosedXML.Excel;

namespace Classes;

public class Column
{
    //data from the data excel, not the schema
    public string DisplayName { get; set; } 
    public string LogicalName { get; set; }
    public string SchemaName { get; set; }
    public string AttributeType { get; set; } //the datatype (can be Whole Number, Lookup, Text, BigInt, Whole number, Uniqueidentifier, Virtual, Choice, Miltiline Text, DateTime, EntityName, Owner, Currency, Two options, Double, ect)
    public bool HasCustomAttribute { get; set; } 
    public string AdditionalData { get; set; } 
    public string Description { get; set; } 

    public IXLColumn ColumnData { get; set; } //IXLColumn from IXLWorksheet
    public int numFilledRows { get; } //number of filled rows in the column

    public Column(IXLColumn col)
    {
        ColumnData = col;
    }
   

}

public class Table
{
    //data directly from the excel schema file 
    public string PluralDisplayName { get; set; }
    public string Entity { get; set; }
    public string Description { get; set; }
    public string SchemaName { get; set; }
    public string LogicalName { get; set; }
    public int ObjectTypeCode { get; set; }
    public bool IsCustomEntity { get; set; }
    public string OwnershipType { get; set; }

    //data not directly from excel
    public List<Column> Columns { get; set; }
    public int MaxRow { get; }

    public Table(IXLWorkbook dataTable, IXLWorkbook schema)
    {
        IXLWorksheet  = dataTable.Worksheet(1);
        IXLWorksheet 
    }
    
    
    public void addColumn(IXLColumn col)
    {

    }

    public void analyzeTable()
    {
        //copy paste old code from v2 to analyze the table
    }

    //get the data from a specific cell (ignore use worksheet.Cell(row, col)
    public var cell(int row, int col)
    {
        return 
    }


}