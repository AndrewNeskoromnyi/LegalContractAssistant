namespace LegalContractWebAPI.Models
{
    public class ArticleDTO
    {
        public string article { get; set; }
        public string[] annotations { get; set; }
    }

    public class ProcessDTO
    {
        public ArticleDTO[] articles { get; set; }
        public string assistantId { get; set; }
        public string fileContent { get; set; }
        public string scope { get; set; }
    }


}
