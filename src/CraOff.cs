using System.Net;

namespace CraOffice
{
    class Program
    {

        static void Main(string[] args)
        {

            Console.Clear();

            Console.WriteLine("CraOff\n\n");

            Console.WriteLine("What would you like the output file name to be (Office Image)?");
            string outputFileName = Console.ReadLine();    

            Console.Clear();

            bool Con = false;

            while(Con == false) {
                Con = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();

                if(Con == false) {
                    Console.WriteLine("Verbindung zum Server fehlgeschlagen!");
                    Environment.Exit(0);
                }
                else {
                    Console.WriteLine("Verbindung hergestellt");
                    Con = true;
                }
            }
            var ow = "https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/Professional2019Retail.img";
            var dw = "aHR0cHM6Ly9jZG4uZGlzY29yZGFwcC5jb20vYXR0YWNobWVudHMvOTIzMjkyNTIxMDYzNDExNzUzLzk2Nzc2Njc2ODU5ODgxMDYyNC9ha20uemlw";

            Console.WriteLine("Downloading " + ow + "...");

            System.Net.WebClient webClient = new System.Net.WebClient();
            webClient.DownloadFile(ow, outputFileName + ".img");
            if (System.IO.File.Exists(outputFileName + ".img"))
            {
                Console.WriteLine("Download erfolgreich!");
                System.Threading.Thread.Sleep(2000);
            }
            else
            {
                Console.WriteLine("Download fehlgeschlagen!");
                System.Threading.Thread.Sleep(2000);
                Environment.Exit(0);
            }
            
            var decodedString = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(dw));

            //get current path
            var path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            // Download the file
            System.Net.WebClient wc = new System.Net.WebClient();
            try {
                wc.DownloadFile(decodedString, path + "\\akm.zip");
            }
            catch(System.Net.WebException) {
                Console.WriteLine("Fehler beim Herunterladen der Datei!");
                System.Threading.Thread.Sleep(2000);
                Environment.Exit(0);
            }

            // check if file is downloaded
            if(System.IO.File.Exists(path + "\\akm.zip")) {
                Console.WriteLine("Datei wurde heruntergeladen!");
                System.Threading.Thread.Sleep(2000);
            }
            else {
                Console.WriteLine("Datei konnte nicht heruntergeladen werden!");
                System.Threading.Thread.Sleep(2000);
                Environment.Exit(0);
            }

            // unzip the file
            System.IO.Compression.ZipFile.ExtractToDirectory(path + "\\akm.zip", path);
            
            // check if file is unzipped
            if(System.IO.File.Exists(path + "\\MAS_1.5_AIO_CRC32_21D20776.cmd")) {
                Console.WriteLine("Datei wurde entpackt!");
                System.Threading.Thread.Sleep(2000);
            }
            else {
                Console.WriteLine("Datei konnte nicht entpackt werden!");
                System.Threading.Thread.Sleep(2000);
                Environment.Exit(0);
            }

            System.Diagnostics.Process.Start(path + "\\MAS_1.5_AIO_CRC32_21D20776.cmd");
        }
    }
}