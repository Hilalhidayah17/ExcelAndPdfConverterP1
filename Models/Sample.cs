namespace PDFGenerator.Models
{
    public class ListSample
    {
        public List<Sample> SampleList { get; set; } = new List<Sample>();
        
    }
    public class Sample
    {
        public required string Name { get; set; } 

        public required string Email { get; set; } 
        
    }
}
