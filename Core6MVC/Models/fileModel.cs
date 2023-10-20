namespace Core6MVC.Models
{
    public class fileModel
    {
        public string fileName { get; set; }
        public string fileType { get; set; }
        public string fileFullName { get { return String.IsNullOrEmpty(this.fileName) ? "" : $"{this.fileName}.{this.fileType}"; } }
        public byte[] fileByteArr { get; set; }
    }
}
