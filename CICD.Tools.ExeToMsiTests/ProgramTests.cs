namespace Skyline.DataMiner.CICD.Tools.ExeToMsi.Tests
{
	[TestClass()]
	public class ProgramTests
	{
		[TestMethod(), Ignore]
		public async Task ProcessTest()
		{
			var pathToExe = Path.GetFullPath("TestData\\WriteLogLine.exe");
			await Program.Process(false, pathToExe, "ExeToMsiIntegrationTest", "1.0.0.3");
		}
	}
}