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
			bool debug = true;
			try
			{
				if (args.Any() || debug)
				{
					string date = "", tender = "";
					string _SEPARATOR = "|";
					string _COMMAND_SEPARATOR = ";";
					string[] activeCommands = ConfigurationManager.AppSettings.Get("commandList").Trim().Split(new string[] { _COMMAND_SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);					
					string timeValidator = "(?<=:|\\s)(\\d{1,2}:){2}\\d{1,2}";

					StreamReader sr;
					List<Dictionary<string, string>> listDictForFile = new List<Dictionary<string, string>>();
					//string[] commandList = ConfigurationManager.AppSettings.Get("commandList").Trim().Split(command_SEPARATOR, StringSplitOptions.RemoveEmptyEntries);

					string fileName = string.Format("{0}\\{1}", Environment.CurrentDirectory, "test.txt");
					
					if (debug) { 
						sr = new StreamReader(fileName);
						fileName = fileName.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries).Last().Split(new string[] {"."}, StringSplitOptions.RemoveEmptyEntries).First();
					}
					else
					{
						 sr = new StreamReader(args[0]);
						fileName = args[0].Split(new string[] {"\\"}, StringSplitOptions.RemoveEmptyEntries).Last();
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
						DateTime time = DateTime.Now;
						Dictionary<string, string> dictForFile = new Dictionary<string, string>();
						var t = lineList[i];
						string[] temp = t.Split(new string[] { _SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);
						if (temp.Length == 1)
						{
							/*
							if (Regex.Match(temp[0], dateValidator).Success)
							{
								DateTime tDate;
								if (DateTime.TryParse(temp[0].Trim(), out tDate))
								{
									date = tDate.ToLongDateString();
								}
							}*/
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

							//string ttime = temp[0].Substring(temp[0].IndexOf(": ")).Remove(temp[0].IndexOf(" ->")).Remove(temp[0].IndexOf(".")).Replace(":", "");
							/*time = DateTime.ParseExact(
								temp[0].Substring(temp[0].IndexOf(": "))
									   .Remove(temp[0].IndexOf(" ->"))
									   .Remove(temp[0].IndexOf("."))
									   .Replace(":", ""),
									"HHmmss", System.Globalization.CultureInfo.InvariantCulture
									).ToString("HH:mm");*/
							if (temp[1].Trim().Equals("In"))
							{
								ip = temp[2].Trim().Replace("IP:", "").Trim();
								command = temp[3].Trim().Replace("CommandName:", "").Trim();
								if(activeCommands.Any(y => y.Equals(command))){ 
									commandName = ConfigurationManager.AppSettings.Get(command).Trim(); ////
									if (command.Equals("Upload"))
									{
										List<string> filesName = new List<string>();
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
												}
											}
											catch { };
										}
										string fN = string.Join(",", filesName);
										commandName = string.Format(ConfigurationManager.AppSettings.Get("Upload").Trim(), tender, fN);
									}
								}
								else
								{
									commandName = ConfigurationManager.AppSettings.Get("UnknownCommand").Trim();
								}

								
							}
							else if (temp[1].Trim().Equals("page"))
							{
								ip = temp[3].Trim().Replace("IP: ", "");
								if (temp[2].Trim().StartsWith("/PositionForm"))
								{
									tender = temp[2].Trim().Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
									commandName = string.Format(ConfigurationManager.AppSettings.Get("/PositionForm").Trim(), tender);

								}
								else if (temp[2].Trim().StartsWith("/BidForm"))
								{
									proposition = temp[2].Trim().Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
									commandName = string.Format(ConfigurationManager.AppSettings.Get("/BidForm").Trim(), tender);
								}

							}
						}
						//Console.ReadKey();
						dictForFile.Add("date", date);
						dictForFile.Add("time", time.ToString("HH:mm:ss"));
						dictForFile.Add("ip", ip);
						dictForFile.Add("commandName", commandName);
						//string h = dictForFile["date"];
						if (!string.IsNullOrEmpty(dictForFile["commandName"]))
						{
							listDictForFile.Add(dictForFile);
						}
					}

					#region xlsFileCreate
					Random a = new Random();
					string outFile = string.Format("{0}\\{1}.xlsx", Environment.CurrentDirectory, fileName);// a.Next(0, 4578));
					if (File.Exists(outFile))
					{
						File.Delete(outFile);
					}

					var file = new FileInfo(outFile);
					var package = new ExcelPackage(file);
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
							//worksheet.Cells[i, 2].Value = entity["rgc_value"].ToString().Remove(entity["rgc_value"].ToString().IndexOf(",") + 2);
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
					bool[] vs = new bool[] { false, false };
					///my
					using (ExcelRange excelRange = worksheet.Cells[startRow, startColumn, endRow, endColumn])
					{
						excelRange.Sort(sortColumn, vs, null, CompareOptions.None);
					}
					package.Save();

					/*if (File.Exists(correctionsFile))
					{
						File.Delete(correctionsFile);
					}*/
					#endregion xlsFileCreate
					Console.WriteLine(string.Format("End{0}Press any key to continue", Environment.NewLine));
					Console.ReadKey();
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
