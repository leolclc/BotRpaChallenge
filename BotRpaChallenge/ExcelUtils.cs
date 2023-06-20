using ExcelDataReader;
using System.Data;


namespace BotRpaChallenge
{
    public class ExcelUtils
    {
        public static readonly int IndexCabecarioExcel = 0; 
        public static List<Pessoa> LerExcel(string diretorioArquivo)
        {
            IExcelDataReader reader;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream = File.Open(diretorioArquivo, FileMode.Open, FileAccess.Read);
            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            System.Data.DataSet result = reader.AsDataSet();
            try
            {
                List<Pessoa> listaPessoas = new();
                foreach (DataRow linha in result.Tables[0].Rows)
                {
                    Pessoa pessoa = new()
                    {
                        FirstName = (string)linha["Column0"],
                        LastName = (string)linha["Column1"],
                        CompanyName = (string)linha["Column2"],
                        Role = (string)linha["Column3"],
                        Address = (string)linha["Column4"],
                        Email = (string)linha["Column5"],
                        PhoneNumber = linha["Column6"].ToString()
                    };
                    listaPessoas.Add(pessoa);
                }
                listaPessoas.RemoveAt(IndexCabecarioExcel);
                return listaPessoas;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}