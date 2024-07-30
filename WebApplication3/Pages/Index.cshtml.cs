using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Threading.Tasks;
using System.Data;
using Python.Runtime;
using System;
using WebApplication3.Models;
using Microsoft.Extensions.Options;

namespace WebApplication3.Pages
{
    public class IndexModel : PageModel
    {
        private readonly PythonSettings _pythonSettings;

        public IndexModel(IOptions<PythonSettings> pythonSettings)
        {
            _pythonSettings = pythonSettings.Value;
        }

        [BindProperty]
        public IFormFile UploadedFile { get; set; }

        public string Message { get; set; }

        public DataTable Transactions { get; set; }

        public void OnGet()
        {

        }

        public async Task<IActionResult> OnPostAsync(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                Message = "No file selected.";
                return Page();
            }

            // Read the uploaded file into a memory stream
            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream);
                memoryStream.Position = 0; // Reset stream position

                string pythonDll = _pythonSettings.PythonDll;
                Runtime.PythonDLL = pythonDll;

                PythonEngine.Initialize();

                // Read the file using Python.NET
                Transactions = ReadExcelFile(memoryStream);

                AppContext.SetSwitch("System.Runtime.Serialization.EnableUnsafeBinaryFormatterSerialization", true);

                // Shutdown Python.NET
                try
                {
                    PythonEngine.Shutdown();
                }
                catch (NotSupportedException _) { }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                AppContext.SetSwitch("System.Runtime.Serialization.EnableUnsafeBinaryFormatterSerialization", false);
            }

            Message = "File uploaded successfully!";
            return Page();
        }

        static void SetEnvironmentVariable(string variable, string value)
        {
            Environment.SetEnvironmentVariable(variable, value, EnvironmentVariableTarget.Process);
        }

        private DataTable ReadExcelFile(MemoryStream memoryStream)
        {
            DataTable transactions = new DataTable();
            transactions.Columns.Add("Date", typeof(string));
            transactions.Columns.Add("Amount", typeof(double));
            transactions.Columns.Add("Category", typeof(string));
            transactions.Columns.Add("Description", typeof(string));

            using (Py.GIL())
            {

                // Ensure the site-packages directory is in sys.path
                dynamic sys = Py.Import("sys");
                sys.path.append(_pythonSettings.PackagesDir);

                dynamic pandas = Py.Import("pandas");
                dynamic io = Py.Import("io");

                // Convert the MemoryStream to a byte array and create a BytesIO object
                memoryStream.Position = 0;
                byte[] byteArray = memoryStream.ToArray();
                dynamic buffer = io.BytesIO(byteArray);

                // Read the Excel file from the BytesIO object
                dynamic df = pandas.read_excel(buffer);

                foreach (var row in df.itertuples())
                {
                    DateTime date = Convert.ToDateTime(row.Date.ToString());
                    if (date.Month == 7) // Filter for July
                    {
                        transactions.Rows.Add(date.ToString("yyyy-MM-dd"), row.Amount, row.Category, row.Description);

                    }
                }
            }

            return transactions;
        }
    }
}
