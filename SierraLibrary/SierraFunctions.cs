using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.IO;
using ExcelDna.Integration;


using Newtonsoft.Json;

namespace SierraLibrary
{
	/// <summary>
	/// SierraFunctions
	/// This class implements excel functions to make SIERRA APIS (https://techdocs.iii.com/sierraapi/Content/titlePage.htm#) 
	/// call using Excel-DNA (https://excel-dna.net/)
	/// </summary>
	public static class SierraFunctions
    {
		public static string base_url = "https://risc01.sabanciuniv.edu/iii/sierra-api/v5/";
		public static string header_basic_auth = "Basic M1hNREd3aDdkT3FpY1JjVHdINzRsanJHVHRkcTpnb2theTE5Nzc=";
		public static string username = "your_username";
		public static string password = "your_password";

		/// <summary>
		/// Returns author information
		/// </summary>		
		[ExcelFunction(Category = "sierra", Description = "About author\r\nSample usage: =About()", Name = "About")]
        public static string About()
        {
            return "Gyokay Nurvet Mustafa\r\nhttps://gyokay.cloud/";
		}

		/// <summary>
		/// Returns bearer token
		/// </summary>
		/// <returns>
		/// "v0Qvd3EscNjMPF9zH606RebLuOaVrTuG6Bs9Vf1_cPFxRKCJPWSbTPOlTOi-bLF17Hcl-8-A2UdTvyMhZfIDATYKLgnh5y_02xNqYq9PGIQ"
		/// </returns>
		/// <example>
		/// {
		///   "access_token": "v0Qvd3EscNjMPF9zH606RebLuOaVrTuG6Bs9Vf1_cPFxRKCJPWSbTPOlTOi-bLF17Hcl-8-A2UdTvyMhZfIDATYKLgnh5y_02xNqYq9PGIQ",
		///   "token_type": "bearer",
		///   "expires_in": 3600
		/// }
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Get token\r\nSample usage: =GetToken()", Name = "GetToken")]
		public static string GetToken()
		{
			/*		
			string authInfo = username + ":" + password;
			authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
			*/
			string token_url = base_url + "token";
			WebRequest request = WebRequest.Create(token_url);
			request.Method = "POST";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers.Add("Authorization", "Basic M1hNREd3aDdkT3FpY1JjVHdINzRsanJHVHRkcTpnb2theTE5Nzc=");
			//request.Headers["Authorization"] = "Basic " + authInfo;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.access_token;
		}

		/// <summary>
		/// Find patron id by barcode
		/// </summary>
		/// <returns>
		/// Patron id
		/// </returns> 
		/// <example>
		/// <code>
		/// string patronId = Barcode2Id("00001845");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find patron id by barcode\r\nSample usage: = Barcode2Id(\"00001845\")", Name = "Barcode2Id")]
		public static string Barcode2Id(string barcode)
		{
			string token_url = base_url + $"patrons/find?barcode={barcode}&fields=id,names,patronType,emails,moneyOwed";
			string access_token = GetToken();

			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.id;
		}

		/// <summary>
		/// Find patron name by barcode
		/// </summary>
		/// <returns>
		/// Patron name
		/// </returns> 
		/// <example>
		/// <code>
		/// string patronName = Barcode2Name("00001845");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find patron name by barcode\r\nSample usage: =Barcode2Name(\"00001845\")", Name = "Barcode2Name")]
		public static string Barcode2Name(string barkod)
		{
			string token_url = base_url + $"patrons/find?barcode={barkod}&fields=id,names,patronType,emails,moneyOwed";
			string access_token = GetToken();
			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.names[0];
		}

		/// <summary>
		/// Find patron e-mail by barcode
		/// </summary>
		/// <returns>
		/// Patron e-mail
		/// </returns> 
		/// <example>
		/// <code>
		/// string patronName = Barcode2Email("00001845");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find patron e-mail by barcode\r\nSample usage: =Barcode2Email(\"00001845\")", Name = "Barcode2Email")]
		public static string Barcode2Email(string barcode)
		{
			string token_url = base_url + $"patrons/find?barcode={barcode}&fields=id,names,patronType,emails,moneyOwed";
			string access_token = GetToken();

			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.emails[0];
		}

		/// <summary>
		/// Find patron type by barcode
		/// </summary>
		/// <returns>
		/// Patron type
		/// </returns> 
		/// <example>
		/// <code>
		/// string patronType = Barcode2PatronType("00001845");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find patron type by barcode\r\nSample usage: =Barcode2PatronType(\"00001845\")", Name = "Barcode2PatronType")]
		public static string Barcode2PatronType(string barcode)
		{
			string token_url = base_url + $"patrons/find?barcode={barcode}&fields=id,names,patronType,emails,moneyOwed";
			string access_token = GetToken();

			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.patronType;
		}

		/// <summary>
		/// Find patron money owed
		/// </summary>
		/// <returns>
		/// money owed
		/// </returns> 
		/// <example>
		/// <code>
		/// string patronMoneyOwed = Barcode2MoneyOwed("00001845");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find patron money owed by barcode\r\nSample usage: =Barcode2MoneyOwed(\"00001845\")", Name = "Barcode2MoneyOwed")]
		public static string Barcode2MoneyOwed(string barcode)
		{
			string token_url = base_url + $"patrons/find?barcode={barcode}&fields=id,names,patronType,emails,moneyOwed";
			string access_token = GetToken();

			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.moneyOwed;
		}

		/// <summary>
		/// Find patron checkout items
		/// </summary>
		/// <returns>
		/// checkout items
		/// </returns> 
		/// <example>
		/// <code>
		/// string patronCheckoutItems = Barcode2CheckoutItems("00001845");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find checkout items by barcode\r\nSample usage: =Barcode2CheckoutItems(\"00001845\")", Name = "Barcode2CheckoutItems")]
		public static object Barcode2CheckoutItems(string barcode)
		{
			string id = Barcode2Id(barcode);
			string access_token = GetToken();

			string token_url = base_url + $"patrons/{id}/checkouts?fields=item";
			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			List<String> items = new List<string>();
			var stack = new Stack<char>();
			string item = "";
			string checkout_item = "";
			object[,] result = new string[sonuc.entries.Count,1];


			for (int i = 0; i < sonuc.entries.Count; i++)
			{
				item = (String)sonuc.entries[i].item;

				checkout_item = string.Concat(item.ToArray().Reverse().TakeWhile(char.IsNumber).Reverse());

				items.Add(checkout_item);
				result[i,0] = checkout_item;
			}
				
			var json = JsonConvert.SerializeObject(items);
			//return json;
			return result;
		}

		/// <summary>
		/// Find bib id by item
		/// </summary>
		/// <returns>
		/// checkout items
		/// </returns> 
		/// <example>
		/// <code>
		/// string bibId = Item2BibId("1136526");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find bib id by item\r\nSample usage: =Item2BibId(\"1136526\")", Name = "Item2BibId")]
		public static string Item2BibId(string item)
		{
			string token_url = base_url + $"items/{item}";
			string access_token = GetToken();

			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.bibIds[0];
		}

		/// <summary>
		/// Find title by bib id
		/// </summary>
		/// <returns>
		/// title
		/// </returns> 
		/// <example>
		/// <code>
		/// string title = BibId2Title("1159654");
		/// </code>
		/// </example>
		[ExcelFunction(Category = "sierra", Description = "Find title by bib id\r\nSample usage: =BibId2Title(\"1159654\")", Name = "BibId2Title")]
		public static string BibId2Title(string bibId)
		{
			string token_url = base_url + $"bibs/{bibId}";
			string access_token = GetToken();

			WebRequest request = WebRequest.Create(token_url);
			request.Method = "GET";
			request.ContentType = "application/json; charset=utf-8";
			request.Headers["Authorization"] = "Bearer " + access_token;

			var response = (HttpWebResponse)request.GetResponse();

			string strResponse = "";
			using (var sr = new StreamReader(response.GetResponseStream()))
			{
				strResponse = sr.ReadToEnd();
			}

			dynamic sonuc = JsonConvert.DeserializeObject(strResponse);
			return sonuc.title;
		}

		/// <summary>
		/// Makes sample array in excel
		/// </summary>
		/// <example>
		/// <code>
		/// = MakeArray(3,4)
		/// </code>
		/// </example>
		/*
		[ExcelFunction(Category = "sierra", Description = "Make array\r\nSample usage: =MakeArray(3,4)", Name = "MakeArray")]
		public static object MakeArray(int rows, int columns)
		{
			object[,] result = new string[rows, columns];
			for (int i = 0; i < rows; i++)
			{
				for (int j = 0; j < columns; j++)
				{
					result[i, j] = string.Format("({0},{1})", i, j);
				}
			}

			return result;
		}
		*/
	}
}





