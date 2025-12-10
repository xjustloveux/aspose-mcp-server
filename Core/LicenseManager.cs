using Aspose.Words;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Pdf;

namespace AsposeMcpServer.Core;

public static class LicenseManager
{
    public static void SetLicense(ServerConfig config)
    {
        // Suppress any potential stdout output from Aspose during license loading
        var originalOut = Console.Out;
        try
        {
            // Temporarily redirect stdout to null during license loading
            Console.SetOut(TextWriter.Null);
            
            var baseDirectory = AppContext.BaseDirectory;
            var currentDirectory = Directory.GetCurrentDirectory();
            
            // Build list of possible license file names to search
            var licenseFileNames = new List<string>();
            
            // If license path is explicitly specified, use it first
            if (!string.IsNullOrWhiteSpace(config.LicensePath))
            {
                licenseFileNames.Add(config.LicensePath);
                // Also try with base directory and current directory prefixes
                if (!Path.IsPathRooted(config.LicensePath))
                {
                    licenseFileNames.Add(Path.Combine(baseDirectory, config.LicensePath));
                    licenseFileNames.Add(Path.Combine(currentDirectory, config.LicensePath));
                }
            }
            
            // Add common license file names based on enabled components
            if (config.EnableWord)
            {
                licenseFileNames.Add("Aspose.Words.lic");
                licenseFileNames.Add(Path.Combine(baseDirectory, "Aspose.Words.lic"));
                licenseFileNames.Add(Path.Combine(currentDirectory, "Aspose.Words.lic"));
            }
            
            if (config.EnableExcel)
            {
                licenseFileNames.Add("Aspose.Cells.lic");
                licenseFileNames.Add(Path.Combine(baseDirectory, "Aspose.Cells.lic"));
                licenseFileNames.Add(Path.Combine(currentDirectory, "Aspose.Cells.lic"));
            }
            
            if (config.EnablePowerPoint)
            {
                licenseFileNames.Add("Aspose.Slides.lic");
                licenseFileNames.Add(Path.Combine(baseDirectory, "Aspose.Slides.lic"));
                licenseFileNames.Add(Path.Combine(currentDirectory, "Aspose.Slides.lic"));
            }
            
            if (config.EnablePdf)
            {
                licenseFileNames.Add("Aspose.Pdf.lic");
                licenseFileNames.Add(Path.Combine(baseDirectory, "Aspose.Pdf.lic"));
                licenseFileNames.Add(Path.Combine(currentDirectory, "Aspose.Pdf.lic"));
            }
            
            // Add Total license as fallback
            licenseFileNames.Add("Aspose.Total.lic");
            licenseFileNames.Add(Path.Combine(baseDirectory, "Aspose.Total.lic"));
            licenseFileNames.Add(Path.Combine(currentDirectory, "Aspose.Total.lic"));
            
            // Also search for any .lic files in the directories
            var searchDirectories = new[] { baseDirectory, currentDirectory };
            foreach (var dir in searchDirectories)
            {
                try
                {
                    var licFiles = Directory.GetFiles(dir, "*.lic", SearchOption.TopDirectoryOnly);
                    foreach (var licFile in licFiles)
                    {
                        var fileName = Path.GetFileName(licFile);
                        if (!licenseFileNames.Contains(licFile) && !licenseFileNames.Contains(fileName))
                        {
                            licenseFileNames.Add(licFile);
                        }
                    }
                }
                catch
                {
                    // Ignore directory access errors
                }
            }
            
            // Find the first existing license file
            string? licensePath = null;
            foreach (var path in licenseFileNames)
            {
                if (File.Exists(path))
                {
                    licensePath = path;
                    break;
                }
            }
            
            var loadedLicenses = new List<string>();
            
            if (licensePath != null)
            {
                // Try to set license for each enabled component
                if (config.EnableWord)
                {
                    try
                    {
                        var wordsLicense = new Aspose.Words.License();
                        wordsLicense.SetLicense(licensePath);
                        loadedLicenses.Add("Words");
                    }
                    catch
                    {
                        // License file might not contain Words license, continue
                    }
                }
                
                if (config.EnableExcel)
                {
                    try
                    {
                        var cellsLicense = new Aspose.Cells.License();
                        cellsLicense.SetLicense(licensePath);
                        loadedLicenses.Add("Cells");
                    }
                    catch
                    {
                        // License file might not contain Cells license, continue
                    }
                }
                
                if (config.EnablePowerPoint)
                {
                    try
                    {
                        var slidesLicense = new Aspose.Slides.License();
                        slidesLicense.SetLicense(licensePath);
                        loadedLicenses.Add("Slides");
                    }
                    catch
                    {
                        // License file might not contain Slides license, continue
                    }
                }
                
                if (config.EnablePdf)
                {
                    try
                    {
                        var pdfLicense = new Aspose.Pdf.License();
                        pdfLicense.SetLicense(licensePath);
                        loadedLicenses.Add("Pdf");
                    }
                    catch
                    {
                        // License file might not contain Pdf license, continue
                    }
                }
                
                // Restore stdout before logging
                Console.SetOut(originalOut);
                
                if (loadedLicenses.Count > 0)
                {
                    Console.Error.WriteLine($"[INFO] Aspose licenses loaded successfully from: {licensePath}");
                    Console.Error.WriteLine($"[INFO] Loaded licenses: {string.Join(", ", loadedLicenses)}");
                }
                else
                {
                    Console.Error.WriteLine($"[WARN] License file found but no valid licenses loaded: {licensePath}");
                    Console.Error.WriteLine("[WARN] Running in evaluation mode.");
                }
            }
            else
            {
                Console.SetOut(originalOut);
                Console.Error.WriteLine("[WARN] No Aspose license file found. Searched locations:");
                var searchedPaths = licenseFileNames.Distinct().Take(10); // Limit output
                foreach (var path in searchedPaths)
                {
                    Console.Error.WriteLine($"[WARN]   - {Path.GetFullPath(path)}");
                }
                if (licenseFileNames.Count > 10)
                {
                    Console.Error.WriteLine($"[WARN]   ... and {licenseFileNames.Count - 10} more locations");
                }
                Console.Error.WriteLine("[WARN] Running in evaluation mode.");
                Console.Error.WriteLine("[INFO] You can specify license file via:");
                Console.Error.WriteLine("[INFO]   - Environment variable: ASPOSE_LICENSE_PATH");
                Console.Error.WriteLine("[INFO]   - Command line: --license:path/to/license.lic");
            }
        }
        catch (Exception ex)
        {
            Console.SetOut(originalOut);
            Console.Error.WriteLine($"[ERROR] Error loading Aspose license: {ex.Message}");
            Console.Error.WriteLine("[WARN] Running in evaluation mode.");
        }
    }
}

