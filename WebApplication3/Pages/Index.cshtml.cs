using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Threading.Tasks;
using System.Data;
using Python.Runtime;
using System;

namespace WebApplication3.Pages
{
    public class IndexModel : PageModel
    {
        [BindProperty]
        public IFormFile UploadedFile { get; set; }

        public string Message { get; set; }

        public DataTable Transactions { get; set; }

        public void OnGet()
        {
            // Log the value of PYTHONNET_PYDLL
            string pythonDll = Environment.GetEnvironmentVariable("PYTHONNET_PYDLL");
            Console.WriteLine($"PYTHONNET_PYDLL is set to: {pythonDll}");
        }

        public async Task<IActionResult> OnPostAsync(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                Message = "No file selected.";
                return Page();
            }

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "uploads", file.FileName);

            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // Set Python DLL path
            string pythonDll = @"C:\Users\Tom Lynch\AppData\Local\Programs\Python\Python312\Python312.dll";
            Runtime.PythonDLL = pythonDll;

            // Log the value of PYTHONNET_PYDLL again to confirm
            string currentPythonDll = Environment.GetEnvironmentVariable("PYTHONNET_PYDLL");
            Console.WriteLine($"PYTHONNET_PYDLL is now set to: {currentPythonDll}");

            // Set the Python home to the virtual environment
            string pythonHome = @"C:\temp\ironpython-env";
            SetEnvironmentVariable("PYTHONHOME", pythonHome);

            // Set the Python path to the site-packages directory of the virtual environment
            string pythonPath = System.IO.Path.Combine(pythonHome, "Lib", "site-packages");
            SetEnvironmentVariable("PYTHONPATH", pythonPath);

            // Additional path settings
            string scriptsPath = System.IO.Path.Combine(pythonHome, "Scripts");
            SetEnvironmentVariable("PATH", scriptsPath + ";" + Environment.GetEnvironmentVariable("PATH"));

            PythonEngine.Initialize();

            // Read the file using Python.NET
            Transactions = ReadExcelFile(filePath);

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

            Message = "File uploaded successfully!";
            return Page();
        }

        static void SetEnvironmentVariable(string variable, string value)
        {
            Environment.SetEnvironmentVariable(variable, value, EnvironmentVariableTarget.Process);
        }

        private DataTable ReadExcelFile(string filePath)
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
                sys.path.append(@"C:\temp\ironpython-env\Lib\site-packages");

                dynamic pandas = Py.Import("pandas");
                dynamic df = pandas.read_excel(filePath);

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
