// Copyright 2018 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using Google.Api.Ads.AdWords.Lib;
using Google.Api.Ads.AdWords.Util.Reports;
using Google.Api.Ads.AdWords.Util.Reports.v201809;
using Google.Api.Ads.AdWords.v201809;
using Google.Api.Ads.Common.Util.Reports;

using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace Google.Api.Ads.AdWords.Examples.CSharp.v201809
{
	/// <summary>
	/// This code example gets and downloads a criteria Ad Hoc report from an AWQL
	/// query. See https://developers.google.com/adwords/api/docs/guides/awql for
	/// AWQL documentation.
	/// </summary>
	public class DownloadCriteriaReportWithAwql : ExampleBase
	{
		/// <summary>
		/// Log File Name
		/// </summary>
		public static string logFineName = null;

		/// <summary>
		/// Error 
		/// </summary>
		public static string errorExisted = null;

		/// <summary>
		/// Main method, to run this code example as a standalone application.
		/// </summary>
		/// <param name="args">The command line arguments.</param>
		public static void Main(string[] args)
		{
			DownloadCriteriaReportWithAwql codeExample = new DownloadCriteriaReportWithAwql();
			Console.WriteLine(codeExample.Description);
			try
			{
				//string fileName = "INSERT_FILE_NAME_HERE";
				codeExample.Run(new AdWordsUser());
			}
			catch (Exception e)
			{
				Console.WriteLine("An exception occurred while running this code example. {0}",
					ExampleUtilities.FormatException(e));
			}
		}

		/// <summary>
		/// Returns a description about the code example.
		/// </summary>
		public override string Description
		{
			get
			{
				return "This code example gets and downloads a criteria Ad Hoc report from an " +
					"AWQL query. See " +
					"https://developers.google.com/adwords/api/docs/guides/awql for AWQL " +
					"documentation.";
			}
		}

		/// <summary>
		/// Runs the code example.
		/// </summary>
		/// <param name="user">The AdWords user.</param>

		public void Run(AdWordsUser user)
		{
			//ReportQuery query = new ReportQueryBuilder()
			//    .Select("CampaignId", "AdGroupId", "Id", "Criteria", "CriteriaType",
			//        "Impressions", "Clicks", "Cost")
			//    .From(ReportDefinitionReportType.CRITERIA_PERFORMANCE_REPORT)
			//    .Where("Status").In("ENABLED", "PAUSED")
			//    .During(ReportDefinitionDateRangeType.LAST_7_DAYS)
			//    .Build();

			//string filePath =
			//    ExampleUtilities.GetHomeDir() + Path.DirectorySeparatorChar + fileName;
			string DebugClientID = "";
			WriteLog("===========================Scheduler Started=========================================" + DateTime.Now);
			try
			{
				var connectionString = ConfigurationManager.ConnectionStrings["AMS"].ConnectionString; //"DATA SOURCE=10.108.135.231; UID=amslogin05; PWD=MoJ$t0$xAMSxps; INITIAL CATALOG=UAT_TRAINING_AMS;";//
				DataSet ds = new DataSet();

				SqlConnection conn = new SqlConnection(connectionString);
				if (conn.State.ToString() != "1")
					conn.Open();

				SqlCommand cmd = new SqlCommand("select distinct GoogleAdWordsClientID from branch where isnull(GoogleAdWordsClientID,'0') <> '0' and GoogleAdWordsClientID<>'362-801-4584' ", conn);
				cmd.CommandType = CommandType.Text;
				SqlDataAdapter sda = new SqlDataAdapter(cmd);
				sda.Fill(ds);

				foreach (DataRow row in ds.Tables[0].Rows)
				{
					string query = "SELECT CampaignId, AdGroupId, Id, Criteria, CriteriaType, Impressions, " +
					"Clicks, Cost FROM CRITERIA_PERFORMANCE_REPORT WHERE Status IN [ENABLED, PAUSED] " +
					"DURING YESTERDAY";

					((Google.Api.Ads.AdWords.Lib.AdWordsAppConfig)user.Config).ClientCustomerId = row["GoogleAdWordsClientID"].ToString();

					//string filePath = ExampleUtilities.GetHomeDir() + Path.DirectorySeparatorChar + fileName;
					string filePath = ConfigurationManager.AppSettings["FolderPath"].ToString() + row["GoogleAdWordsClientID"].ToString() + ".csv";
					DebugClientID = row["GoogleAdWordsClientID"].ToString();
					try
					{
						ReportUtilities utilities = new ReportUtilities(user, "v201809", query,
							DownloadFormat.CSV.ToString());
						using (ReportResponse response = utilities.GetResponse())
						{
							response.Save(filePath);
						}
						Console.WriteLine("Report was downloaded to '{0}'.", filePath);
					}
					catch (Exception e)
					{
						errorExisted = "yes";
						//throw new System.ApplicationException("Failed to download report.", e);
						WriteLog("*******************************Failed to download report for the Client ID=========================================" + DebugClientID);
						WriteLog(e.ToString());
					}
				}
			}
			catch (Exception ex)
			{
				errorExisted = "yes";
				WriteLog("*******************************ExCeption Starts for the Client ID=========================================" + DebugClientID);
				WriteLog(ex.ToString());
				WriteLog("*******************************ExCeption Ends=========================================");
				//throw;
			}
			finally
			{
				WriteLog("===========================Scheduler Ended=========================================" + DateTime.Now);

				string SMTPCLIENT = System.Configuration.ConfigurationManager.AppSettings["SMTPCLIENT"].ToString();
				string MAILUSERNAME = System.Configuration.ConfigurationManager.AppSettings["MAILUSERNAME"].ToString();
				string MAILPWD = System.Configuration.ConfigurationManager.AppSettings["MAILPWD"].ToString();
				int Port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Port"]);
				Boolean EnableSsl = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["EnableSsl"]);

				System.Net.Mail.MailMessage msg;
				

				msg = new System.Net.Mail.MailMessage();
				msg.Body = "Hi Team, <br><br> Please find the attached file for the Google AdWords log details. <br> <br> Thanks.";


				msg.From = new MailAddress("amsadmin@allmysons.com");

				string[] strMails = ConfigurationManager.AppSettings["ToMail"].ToString().Split(';');
				for (int i = 0; i < strMails.Length - 1; i++)
				{
					msg.To.Add(strMails[i].ToString());

				}


				if (errorExisted == "yes")
				{
					msg.Subject = "Google AdWords Scheduler Status on " + System.DateTime.Today.ToString("MM-dd-yyyy") + " ****** Error Occurred **********";
				}
				else
				{
					msg.Subject = "Google AdWords Scheduler Status on " + System.DateTime.Today.ToString("MM-dd-yyyy") + " ******* Run Successfully ******";
				}

				System.Net.Mail.SmtpClient smtpServer = null;
				System.Net.NetworkCredential credentials = null;
				smtpServer = new System.Net.Mail.SmtpClient(SMTPCLIENT);
				smtpServer.Port = Port;
				smtpServer.EnableSsl = EnableSsl;
				smtpServer.UseDefaultCredentials = false;
				smtpServer.Port = Port;
				smtpServer.EnableSsl = EnableSsl;
				smtpServer.UseDefaultCredentials = false;
				credentials = new System.Net.NetworkCredential(MAILUSERNAME, MAILPWD);
				smtpServer.Credentials = credentials;
				smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
				System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;

				MemoryStream ms1 = new MemoryStream(File.ReadAllBytes(logFineName));
				Attachment att1 = new System.Net.Mail.Attachment(ms1, "LogFile.txt");
				msg.Attachments.Add(att1);

				msg.IsBodyHtml = true;
				smtpServer.Credentials = credentials;
				smtpServer.Send(msg);


			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="strLog"></param>
		public static void WriteLog(string strLog)
		{
			StreamWriter log;
			FileStream fileStream = null;
			DirectoryInfo logDirInfo = null;
			FileInfo logFileInfo;

			string logFilePath = System.Configuration.ConfigurationManager.AppSettings["LogFolderPath"].ToString();
			logFilePath = logFilePath + "Log-File_Creation_" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
			logFineName = logFilePath;
			logFileInfo = new FileInfo(logFilePath);
			logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
			if (!logDirInfo.Exists) logDirInfo.Create();
			if (!logFileInfo.Exists)
			{
				fileStream = logFileInfo.Create();
			}
			else
			{
				fileStream = new FileStream(logFilePath, FileMode.Append);
			}
			log = new StreamWriter(fileStream);
			log.WriteLine(strLog);
			log.Close();
			fileStream.Close();
		}
	}
}
