using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Net;

class Program
{
    static async Task Main(string[] args)
    {
        string jsonFilePath = "json_file.json";
        string excelFilePath = "output_excel.xlsx";

        // Deserialize JSON data into a list of Root objects
        List<Root> data = JsonConvert.DeserializeObject<List<Root>>(File.ReadAllText(jsonFilePath));
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("EmployeeData");

            // Add headers to the Excel sheet
            int col = 1;
            foreach (var propertyInfo in typeof(Root).GetProperties())
            {
                worksheet.Cells[1, col].Value = propertyInfo.Name;
                col++;
            }

            // Download and embed profile pictures in parallel
            var downloadTasks = new List<Task>();
            for (int row = 0; row < data.Count; row++)
            {
                string imageUrl = data[row].Profile_Pic;
                Console.WriteLine(imageUrl);
                if (!string.IsNullOrEmpty(imageUrl))
                {
                    downloadTasks.Add(DownloadAndEmbedImageAsync(imageUrl, worksheet, data, row));
                }
            }

            await Task.WhenAll(downloadTasks);

            // Save the Excel package to a file
            FileInfo excelFile = new FileInfo(excelFilePath);
            package.SaveAs(excelFile);
        }

        Console.WriteLine("Excel file created successfully.");
    }

    static async Task DownloadAndEmbedImageAsync(string imageUrl, ExcelWorksheet worksheet, List<Root> data, int row)
    {
        using (WebClient webClient = new WebClient())
        {
            byte[] imageBytes = await webClient.DownloadDataTaskAsync(imageUrl);

            // Get the EmployeeName and Employee_ID from the data list
            string employeeName = data[row].Employee_Name;
            int employeeId = data[row].Employee_ID;

            // Generate the subfolder path based on EmployeeName and Employee_ID
            string subfolder = Path.Combine("C:\\Users\\tahir\\source\\repos\\ConsoleApp4\\ConsoleApp4\\bin\\Debug\\net6.0\\Images");

            // Create the subfolder if it doesn't exist
            if (!Directory.Exists(subfolder))
            {
                Directory.CreateDirectory(subfolder);
            }

            string fileExtension = Path.GetExtension(imageUrl);

            // Save the image in the subfolder with a unique name
            string tempImagePath = Path.Combine(subfolder, $"ProfilePic{row}{fileExtension}");
            File.WriteAllBytes(tempImagePath, imageBytes);

            string imageName = $"{employeeName}_{employeeId}{fileExtension}";
            FileInfo imageFile = new FileInfo(tempImagePath);
            ExcelPicture picture = worksheet.Drawings.AddPicture(imageName, imageFile);

            picture.SetPosition(row + 2, 0, 2, 0);

            Console.WriteLine($"Image for row {row} created successfully.");
        }
    }
}

public class Root
{
    public int Employee_ID { get; set; }
    public string Employee_Name { get; set; }
    public string Department_Name { get; set; }
    public string Month_Year { get; set; }
    public string Image_file { get; set; }
    public string Profile_Pic { get; set; }
    public string Designation_Name { get; set; }
    public string Title { get; set; }
    public string Description { get; set; }
    public DateTime Joining_Date { get; set; }
    public string Joining_Date_string { get; set; }
    public bool status { get; set; }
    public object Modified_By { get; set; }
    public DateTime Modified_Date { get; set; }
    public int office_id { get; set; }
    public string Gender { get; set; }
}