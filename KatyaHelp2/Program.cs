using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace KatyaHelp2
{
	class Program
	{
		static void Main(string[] args)
		{
			bool.TryParse(ConfigurationManager.AppSettings.Get("debug").Trim(),out bool debug);
			try
			{
				if (args.Any() || debug)
				{
					for (int a = 0; a < (args.Any() ? args.Count():1); a++)
					{
						string date = "", tender = "";
						string _SEPARATOR = "|";
						string _COMMAND_SEPARATOR = ";";
						string[] activeCommands = ConfigurationManager.AppSettings.Get("commandList").Trim().Split(new string[] { _COMMAND_SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);
						string timeValidator = "(?<=:|\\s|^)(\\d{1,2}:){2}\\d{1,2}.\\d{1,3}";
						string ipValid = "([0-9]{1,3}.){3}[0-9]{1,3}";

						StreamReader sr;
						List<Dictionary<string, string>> listDictForFile = new List<Dictionary<string, string>>();
						string fileName = string.Format("{0}\\{1}", Environment.CurrentDirectory, "test.txt");

						if (debug)
						{
							sr = new StreamReader(fileName);
							fileName = fileName.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries).Last().Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries).First();
						}
						else
						{
							sr = new StreamReader(args[a]);
							fileName = args[a].Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries).Last().Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries).First();
						}

						List<string> lineList = new List<string>();
						while (!sr.EndOfStream)
						{
							string line = sr.ReadLine().Trim();
							if (line.Length > 1)
							{
								lineList.Add(line);
							}
						}
						sr.Close();

						for (int i = 0; i < lineList.Count; i++)
						{
							string ip = "", command = "", commandName = "", proposition = "";
							List<string> filesName = new List<string>();
							Dictionary<DateTime, string> filesWithTime = new Dictionary<DateTime, string>();
							DateTime time = DateTime.Now;
							Dictionary<string, string> dictForFile = new Dictionary<string, string>();

							CommandDictionary<string, DateTime, DateTime, string> commandDictionary = new CommandDictionary<string, DateTime, DateTime, string>();

							var t = lineList[i];
							string[] temp = t.Split(new string[] { _SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);
							if (temp.Length == 1)
							{
								DateTime tDate;
								if (DateTime.TryParse(temp[0].Trim(), out tDate))
								{
									date = tDate.ToLongDateString();
								}
							}
							else
							{
								var u = Regex.Match(temp[0], timeValidator);
								if (u.Success)
								{
									DateTime.TryParse(u.Value, out time);
								}

								if (temp[1].Trim().Equals("In"))
								{
									//string ipValid = "([0-9]{1,3}.){3}[0-9]{1,3}";
									ip = Regex.Match(temp[2].Trim(), ipValid).Value;
									command = temp[3].Trim().Replace("CommandName:", "").Trim();
									if (activeCommands.Any(y => y.Equals(command)))
									{
										commandName = ConfigurationManager.AppSettings.Get(command).Trim(); 
										if (command.Equals("Upload"))
										{
											filesName = new List<string>();
											filesWithTime = new Dictionary<DateTime, string>();
											for (int ii = 1; ii < lineList.Count; ii++)
											{
												try
												{
													var temp2 = lineList[i + ii].Split(new string[] { _SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);
													if (temp2[1].Trim().Equals("Out"))
													{
														ii = lineList.Count;
													}
													else if (temp2[1].Trim().Equals("info"))
													{
														filesName.Add(temp2[2]);
														filesWithTime.Add(DateTime.Parse(temp2[0].Remove(temp2[0].IndexOf("->")).Trim()), temp2[2]);
													}
												}
												catch { };
											}
											string fN = string.Join(",", filesName.ToArray());
											commandName = ConfigurationManager.AppSettings.Get("Upload").Trim();
										}
										else if (command.Equals("LocalDownload"))
										{
											filesName = new List<string>();
											filesWithTime = new Dictionary<DateTime, string>();
											for (int ii = 1; ii < lineList.Count; ii++)
											{
												try
												{
													var temp2 = lineList[i + ii].Split(new string[] { _SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);
													if (temp2[1].Trim().Equals("Out"))
													{
														ii = lineList.Count;
													}
													else if (temp2[1].Trim().Equals("OK"))
													{
														filesName.Add(temp2[2].Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries).Last());
														filesWithTime.Add(DateTime.Parse(temp2[0].Remove(temp2[0].IndexOf("->")).Trim()), temp2[2].Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries).Last());
													}
												}
												catch { };
												string fN = string.Join(",", filesName.ToArray());
												commandName = ConfigurationManager.AppSettings.Get("LocalDownload").Trim();
											}
										}
										/*else if (command.Equals("jSetBid")){
										}*/
									}
									else
									{
										commandName = ConfigurationManager.AppSettings.Get("UnknownCommand").Trim();
									}
								}
								else if (temp[1].Trim().Equals("page"))
								{
									//string ipValid = "([0-9]{1,3}.){3}[0-9]{1,3}";
									ip = Regex.Match(temp[3].Trim(), ipValid).Value;
									if (temp[2].Trim().StartsWith("/PositionForm"))
									{
										tender = temp[2].Trim().Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
										commandName = string.Format(ConfigurationManager.AppSettings.Get("/PositionForm").Trim(), tender);

									}
									else if (temp[2].Trim().StartsWith("/BidForm"))
									{
										if (temp.Count() == 4)
										{
											if (temp[2].Trim().Substring(temp[2].Length - 5, temp.Length).Equals("lot="))
											{
												commandName = string.Format(ConfigurationManager.AppSettings.Get("/BidForm").Trim(), tender);
											}
											else if (temp[2].Trim().Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[2].StartsWith("web"))
											{
												commandName = ConfigurationManager.AppSettings.Get("CreateDraft").Trim();
											}
										}
									}

								}
							}
							if (filesWithTime.Any())
							{
								foreach (var fn in filesWithTime)
								{
									dictForFile = new Dictionary<string, string>();
									dictForFile.Add("date", date);
									dictForFile.Add("time", fn.Key.ToString(ConfigurationManager.AppSettings.Get("timeFormat").Trim()));
									dictForFile.Add("ip", ip);
									dictForFile.Add("commandName", string.Format(commandName, tender, fn.Value));
									if (!string.IsNullOrEmpty(dictForFile["commandName"]))
									{
										listDictForFile.Add(dictForFile);
									}
								}
							}
							else
							{
								dictForFile.Add("date", date);
								dictForFile.Add("time", time.ToString(ConfigurationManager.AppSettings.Get("timeFormat").Trim()));
								dictForFile.Add("ip", ip);
								dictForFile.Add("commandName", commandName);
								if (!string.IsNullOrEmpty(dictForFile["commandName"]))
								{
									listDictForFile.Add(dictForFile);
								}
							}					
						}

						#region xlsFileCreate
						string outFile = string.Format("{0}\\{1}.xlsx", Environment.CurrentDirectory, fileName);// a.Next(0, 4578));
						if (File.Exists(outFile))
						{
							File.Delete(outFile);
						}

						var file = new FileInfo(outFile);
						ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

						using (var package = new ExcelPackage(file))
						{
							ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Work");
							// --------- Data and styling goes here -------------- //
							int col = 1, row = 1;
							worksheet.DefaultColWidth = 25;
							worksheet.Cells[row, col++].Value = ConfigurationManager.AppSettings.Get("firstColumn").Trim();
							worksheet.Cells[row, col++].Value = ConfigurationManager.AppSettings.Get("secondColumn").Trim();
							worksheet.Cells[row, col++].Value = ConfigurationManager.AppSettings.Get("thirdColumn").Trim();
							worksheet.Cells[row, col++].Value = ConfigurationManager.AppSettings.Get("fourthColumn").Trim();
							int iii = 2;
							foreach (var entity in listDictForFile)
							{
								col = 1;
								try
								{
									worksheet.Cells[iii, col++].Value = entity["ip"];
								}
								catch { }
								try
								{
									worksheet.Cells[iii, col++].Value = entity["date"];
									}
								catch { }
								try
								{
									worksheet.Cells[iii, col++].Value = entity["time"];
								}
								catch { }
								try
								{
									worksheet.Cells[iii, col++].Value = entity["commandName"];
								}
								catch { }
								iii++;
							}
							var startRow = 2;
							var startColumn = 1;
							var endRow = 100;
							var endColumn = 10;
							///my
							int[] sortColumn = new int[] { 1, 2 };
							bool[] descending = new bool[] { false, false };
							///my
							using (ExcelRange excelRange = worksheet.Cells[startRow, startColumn, endRow, endColumn])
							{
								excelRange.Sort(sortColumn, descending, null, CompareOptions.IgnoreSymbols);
							}
							package.Save();
						}
						#endregion xlsFileCreate
						Console.WriteLine(string.Format("End{0}Press any key to continue", Environment.NewLine));
						Console.ReadKey();
					}
				}
				else
				{
					Console.WriteLine(ConfigurationManager.AppSettings.Get("noFile").Trim());
					Console.ReadKey();
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				Console.ReadKey();
			}
		}
	}
}
