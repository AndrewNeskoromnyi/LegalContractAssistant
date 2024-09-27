namespace LegalContractWebAPI.Models
{
    public class ProcessDTO
    {
        public string article { get; set; }
        public string[] annotations { get; set; }
        public string assistantId { get; set; }
        public string fileContent { get; set; }
    }
}
