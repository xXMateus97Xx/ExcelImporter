# ExcelImporter

ExcelImporter it's a simple API programmed in .Net Standard 2.0 using EPPlus 4.5.1 to import excel tables to C# objects.
ExcelImporter is able to generate a sample excel file.

# How to Configure the Importer

Consider this class for the examples

```
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
    public DateTime Birthday { get; set; }
    public decimal Weight { get; set; }
    public bool HasJob { get; set; }
    public string Nickname { get; set; }
    public bool HasPet { get; set; }
    public string PetName { get; set; }
    public int? PetAge { get; set; }
}
```

On your project initialization call the mapper

```
ImporterMapper<Person>.Create()
    .AddField(x => x.Age, "Age", 3, "20")
    .AddField(x => x.Birthday, "Birthday Date", 4, "07/25/1997")
    .AddField(x => x.HasJob, "Has Job", 5, "True")
    .AddField(x => x.HasPet, "Has Pet", 7, "False")
    .AddField(x => x.Name, "Name", 1, "John Smith")
    .AddField(x => x.Nickname, "Nickname", 2, "Jojo")
    .AddField(x => x.PetAge, "Pet Age", 9, "5")
    .AddField(x => x.PetName, "Pet Name", 8, "Rex")
    .AddField(x => x.Weight, "Weight", 6, "85")
    .AddExtraField("Marriage", typeof(bool), "Is Marriage", 11, "True")
    .AddExtraField("ChildrenQuantity", typeof(int), "Children(s)", 10, "2")
    .SetWorksheetName("People")
    .Map();
```

Explaning the code:
The AddField method maps a field to be read by the importer, it's receive 4 parameters
* An lambda expression that represents the C# class property
* The name used on excel column
* A display order number used to order the sample excel file
* A sample data used to generate the sample excel file

The AddExtraField method maps a field that it's not included in C# class, it's receive 5 parameters
* An name used as key in extra fields dictionary
* The type of the field
* The name used on excel column
* A display order number used to order the sample excel file
* A sample data used to generate the sample excel file

The SetWorksheetName method sets the name used on the sample excel file worksheet

# How to Get the Sample Excel File

```
IImporter<Person> importer = new Importer<Person>();

byte[] file = importer.GetSampleExcelFile();
```

# How to Import a File

```
byte[] file = File.ReadAllBytes("path to excel file");
IImporter<Person> importer = new Importer<Person>();
IList<ImportedItem<Person>> items = importer.Import(file);
```
or
```
IList<ImportedItem<Person>> items;
IImporter<Person> importer = new Importer<Person>();

using (FileStream file = File.Open("path to excel file", FileMode.Open))
    items = importer.Import(file);

foreach (var personItem in items)
{
    Person person = personItem.Value;
    IDictionary<string, object> extraInfos = personItem.ExtraFields;
    bool hasErros = personItem.HasError;
    IList<string> errors = personItem.Errors;
    
    //Use the results here
}
```

# How to run the tests

1. Install the Nunit Test Adapter 3 on your Visual Studio
2. Move the excel files inside the folder "ExcelSamples" to "C:\ExcelImporterTests"
3. On Visual Studio open the TestExplorer and run the tests.
