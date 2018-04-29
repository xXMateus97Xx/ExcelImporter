using System;

namespace ExcelImporter.Tests
{
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
}
