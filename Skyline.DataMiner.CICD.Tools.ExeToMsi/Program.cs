namespace Skyline.DataMiner.CICD.Tools.ExeToMsi
{
	using System;
	using System.CommandLine;
	using System.Diagnostics;
	using System.IO;
	using System.IO.Compression;
	using System.Linq;
	using System.Reflection;
	using System.Security.Cryptography;
	using System.Text;
	using System.Threading.Tasks;

	using Microsoft.Extensions.Logging;

	using Serilog;

	/// <summary>
	/// Creates a .msi from a provided .exe installer..
	/// </summary>
	public static class Program
	{
		/*
         * Design guidelines for command line tools: https://learn.microsoft.com/en-us/dotnet/standard/commandline/syntax#design-guidance
         */

		/// <summary>
		/// Code that will be called when running the tool.
		/// </summary>
		/// <param name="args">Extra arguments.</param>
		/// <returns>0 if successful.</returns>
		public static async Task<int> Main(string[] args)
		{
			var isDebug = new Option<bool>(
			name: "--debug",
			description: "Indicates the tool should write out debug logging.")
			{
				IsRequired = false,
			};

			isDebug.SetDefaultValue(false);

			var exeFilePath = new Option<string>(
				name: "--exe-file-path",
				description: "File path to the executable.")
			{
				IsRequired = true
			};

			var exeArguments = new Option<string>(
				name: "--exe-arguments",
				description: "Arguments to call during the exe run.")
			{
				IsRequired = false
			};

			var msiName = new Option<string>(
			name: "--msi-name",
			description: "Name of the msi file. Also the name of the installed app in windows.")
			{
				IsRequired = true
			};

			var msiVersion = new Option<string>(
				name: "--msi-version",
				description: "Version of the .msi. Should be format 'A.B.C.D'.")
			{
				IsRequired = true
			};

			var rootCommand = new RootCommand("Creates a .msi from a provided .exe installer.")
			{
				isDebug,
				exeFilePath,
				exeArguments,
				msiName,
				msiVersion,
			};

			rootCommand.SetHandler(Process, isDebug, exeFilePath, exeArguments, msiName, msiVersion);

			return await rootCommand.InvokeAsync(args);
		}

		public static async Task<int> Process(bool isDebug, string exePath, string exeArguments, string msiName, string msiVersion)
		{
			try
			{
				exeArguments = exeArguments.Trim('\"');
				var logConfig = new LoggerConfiguration().WriteTo.Console();
				logConfig.MinimumLevel.Is(isDebug ? Serilog.Events.LogEventLevel.Debug : Serilog.Events.LogEventLevel.Information);
				var seriLog = logConfig.CreateLogger();

				using var loggerFactory = LoggerFactory.Create(builder => builder.AddSerilog(seriLog));
				var logger = loggerFactory.CreateLogger("Skyline.DataMiner.CICD.Tools.ExeToMsi");
				try
				{
					//Main Code for program here
					string msiNameWithExtension;

					if (!msiName.EndsWith(".msi"))
					{
						msiNameWithExtension = msiName + ".msi";
					}
					else
					{
						msiNameWithExtension = msiName;
						msiName = msiName.Replace(".msi", "");
					}

					string installPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WiXToolset");

					if (!Directory.Exists(installPath))
					{
						Console.WriteLine("Installing WiX Toolset...");
						ExtractWixBuildTools(installPath);
						Console.WriteLine("WiX Toolset installed successfully.");
					}
					else
					{
						Console.WriteLine("WiX Toolset already installed.");
					}

					string exeName = Path.GetFileName(exePath);
					string workingDir = Path.GetDirectoryName(exePath);
					string wxsFile = Path.Combine(workingDir, "setup.wxs");
					string msiFile = Path.Combine(workingDir, $"{msiNameWithExtension}");

					string candlePath = $"{installPath}\\candle.exe";
					string lightPath = $"{installPath}\\light.exe";

					Guid upgradeCode = GenerateGuidFromString(msiNameWithExtension);

					// Step 1: Create a basic WXS template
					string wxsTemplate = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Wix xmlns=""http://schemas.microsoft.com/wix/2006/wi"">
  <Product Id=""*""
           Name=""{msiName} ExeToMsi""
           Language=""1033""
           Version=""{msiVersion}""
           Manufacturer=""YourCompany""
           UpgradeCode=""{upgradeCode}"">

    <Package InstallerVersion=""500"" Compressed=""yes"" InstallScope=""perMachine"" />

	<MajorUpgrade AllowSameVersionUpgrades=""yes"" DowngradeErrorMessage=""A newer version of [ProductName] is already installed."" />

    <Media Id=""1"" Cabinet=""product.cab"" EmbedCab=""yes"" />

    <Directory Id=""TARGETDIR"" Name=""SourceDir"">
      <Directory Id=""ProgramFilesFolder"">
        <Directory Id=""INSTALLFOLDER"" Name=""{msiName}"" />
      </Directory>
    </Directory>

    <Component Id=""MainExecutable"" Guid=""{Guid.NewGuid()}"" Directory=""INSTALLFOLDER"">
      <File Id=""{msiName}"" Source=""{exePath}"" KeyPath=""yes"" />
    </Component>

    <Feature Id=""ProductFeature"" Title=""{msiName}"" Level=""1"">
      <ComponentRef Id=""MainExecutable"" />
    </Feature>

    <!-- Deferred Custom Action: Associated with the Executable -->
    <CustomAction Id=""RunEXE""
                  Directory=""INSTALLFOLDER""
                  ExeCommand=""&quot;[INSTALLFOLDER]{exeName}&quot; {exeArguments}""
                  Execute=""deferred""
                  Return=""asyncNoWait""
                  Impersonate=""no"" />

    <!-- Schedule the Custom Action -->
    <InstallExecuteSequence>
      <Custom Action=""RunEXE"" Before=""InstallFinalize"">NOT Installed</Custom>
    </InstallExecuteSequence>
  </Product>
</Wix>";

					Console.WriteLine("Generating WXS file...");
					File.WriteAllText(wxsFile, wxsTemplate);

					// Step 2: Run Candle to compile WXS to WIXOBJ
					string wixObjFile = Path.Combine(workingDir, "setup.wixobj");
					RunProcess(candlePath, $"-out \"{wixObjFile}\" \"{wxsFile}\"", workingDir);

					// Step 3: Run Light to link and produce MSI
					Console.WriteLine("Creating MSI package...");
					RunProcess(lightPath, $"-o \"{msiFile}\" \"{wixObjFile}\"", workingDir);

					Console.WriteLine($"MSI package created successfully: {msiFile}");
				}
				catch (Exception e)
				{
					logger.LogError($"Exception during Process Run: {e}");
					return 1;
				}
			}
			catch (Exception e)
			{
				Console.WriteLine($"Exception on Logger Creation: {e}");
				return 1;
			}

			return 0;
		}

		private static void ExtractWixBuildTools(string destination)
		{
			Console.WriteLine("Extracting...");

			var allKnownResources = Assembly.GetExecutingAssembly().GetManifestResourceNames();
			string name = "Skyline.DataMiner.CICD.Tools.ExeToMsi.Included.wixBuildTools.zip";
			if (!allKnownResources.Contains(name))
			{
				throw new InvalidOperationException("Failed to find wix Build Tools.");
			}

			var byteStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name);

			ZipFile.ExtractToDirectory(byteStream, destination);
		}

		private static Guid GenerateGuidFromString(string input)
		{
			using (var md5 = MD5.Create())
			{
				byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));
				return new Guid(hash);
			}
		}

		private static void RunProcess(string exePath, string arguments, string workingDir)
		{
			ProcessStartInfo processInfo = new ProcessStartInfo
			{
				FileName = exePath,
				Arguments = arguments,
				WorkingDirectory = workingDir,
				RedirectStandardOutput = true,
				RedirectStandardError = true,
				UseShellExecute = false,
				CreateNoWindow = true
			};

			using (Process process = new Process { StartInfo = processInfo })
			{
				process.OutputDataReceived += (sender, args) => Console.WriteLine(args.Data);
				process.ErrorDataReceived += (sender, args) => Console.Error.WriteLine(args.Data);

				process.Start();
				process.BeginOutputReadLine();
				process.BeginErrorReadLine();
				process.WaitForExit();

				if (process.ExitCode != 0)
					throw new InvalidOperationException($"Error: Process failed with exit code {process.ExitCode}");
			}
		}
	}
}