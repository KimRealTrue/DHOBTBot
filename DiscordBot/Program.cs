using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.IO;

using Discord;
using Discord.Commands;
using Discord.WebSocket;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace DiscordBot
{

	class Program
	{

		static int FestTime = 2;
		static int DesTime = 3;
		static int FloodTime = 5;
		static int WarTime = 7;
		static int SupTime = 11;
		static int DevTime = 13;
		static int GoodTime = 17;
		static int RichTime = 19;

		DiscordSocketClient _client; //봇 클라이언트
		CommandService _commands;    //명령어 수신 클라이언트

		static GSheet _sheetConnect;
		static Dictionary<string, List<object>> _areaDictionary = new Dictionary<string, List<object>>();
		static Dictionary<string, List<object>> _trendDictionary = new Dictionary<string, List<object>>();
		static Dictionary<string, string> _sheetKeyDictionary = new Dictionary<string, string>();

		static ISocketMessageChannel _lastSocketChannel;

		static string trendMessage = "한 권역에 대유행이 중첩될 경우\n가장 텀이 긴 유행 두 개만 발생합니다.\n\n";


		static string[] northSeaArea = new string[] { "북해", "발트해" };
		static string[] southMidSeaArea = new string[] { "서지중해", "대서양" };
		static string[] eastMidSeaArea = new string[] { "동지중해", "흑해" };
		static string[] caribArea = new string[] { "동아메리카", "동미", "카리브" };
		static string[] southAmericaArea = new string[] { "남아메리카", "남미" };
		static string[] westAfricaArea = new string[] { "서아프리카", "서아프" };
		static string[] southAfricaArea = new string[] { "남아프리카", "남아프" };
		static string[] eastAfricaArea = new string[] { "동아프리카", "동아프" };
		static string[] arabiaArea = new string[] { "아라비아", "아랍", "서인도" };
		static string[] eastIndiaArea = new string[] { "동인도", "인도차이나" };
		static string[] southEastAsiaArea = new string[] { "남아시아", "동남아" };
		static string[] oceaniaArea = new string[] { "오세아니아", "호주" };
		static string[] chinaArea = new string[] { "동아시아", "명"};
		static string[] koreaArea = new string[] { "극동아시아", "극동", "조선", "일본"};
		static string[] bot_help_key = new string[] { "설명서", "help", "도움", "헬프", "봇" };
		static string[] log = new string[] { "로그", "업데이트" };


		/*
		 "사용법") || requestHead.Equals("설명서") || requestHead.Equals("help")
				|| requestHead.Equals("도움") || requestHead.Equals("헬프") || requestHead.Equals("봇"
		 */
		static void Main(string[] args)
		{
			_sheetConnect = new GSheet();

			SetAreaDictionary();
			SetTrendDictionary();
			SetSheetKey();

			string token = System.IO.File.ReadAllText("mainToken.txt");

			new Program()
				.MainAsync(token)
				.GetAwaiter()
				.GetResult();
		}

		static List<string> AreaList {
			get {
				return _areaDictionary.Keys.ToList();
			}
		}

		static void SetAreaDictionary()
		{
			var areaList = _sheetConnect.GetSheet("Area!A1:AE14");

			foreach (var row in areaList) {
				string key = "";
				List<object> valueList = new List<object>();
				for (int i = 0; i < row.Count; i++) {
					if (i == 0) {
						if (row[i] != null) {
							key = row[i].ToString();
						}
						else {
							key = "";
						}
					}
					else {
						if (row[i] != null) {
							valueList.Add(row[i]);
						}
						else {
							valueList.Add("");
						}
					}
				}
				_areaDictionary.Add(key, valueList);
			}

			Console.WriteLine(string.Format("Area Sheet Ready"));
		}

		static void SetTrendDictionary()
		{
			var typeList = _sheetConnect.GetSheet("Trend!A2:E9");

			foreach (var row in typeList) {
				string key = "";
				List<object> valueList = new List<object>();
				for (int i = 0; i < row.Count; i++) {
					if (i == 0) {
						if (row[i] != null) {
							key = row[i].ToString();
						}
						else {
							key = "";
						}
					}
					else {
						if (row[i] != null) {
							valueList.Add(row[i]);
						}
						else {
							valueList.Add("");
						}
					}
				}

				_trendDictionary.Add(key, valueList);
			}

			Console.WriteLine(string.Format("Trend Sheet Ready"));

		}

		static void SetSheetKey()
		{
			var sheetKeyList = _sheetConnect.GetSheet("SheetKey!A1:B22");

			foreach (var row in sheetKeyList) {
				string key = "";
				string value = "";
				List<object> valueList = new List<object>();
				for (int i = 0; i < row.Count; i++) {
					if (i == 0) {
						if (row[i] != null) {
							key = row[i].ToString();
						}
						else {
							key = "";
						}
					}
					else {
						if (row[i] != null) {
							value = row[i].ToString();
						}
						else {
							value = "";
						}
					}
				}
				_sheetKeyDictionary.Add(key, value);
			}

			Console.WriteLine(string.Format("SheetKey Ready"));
		}

		public static string GetAreaName(object cityName)
		{
			foreach (string key in _areaDictionary.Keys) {
				if (_areaDictionary[key].Contains(cityName)) {
					return key;
				}
			}
			return null;
		}

		public static List<object> GetTrendItems(object trendName)
		{
			foreach (string key in _trendDictionary.Keys) {
				if (key.Equals(trendName.ToString())) {
					return _trendDictionary[key];
				}
			}
			return null;
		}

		public static string GetTimeSheetKey(string cityName, string trendName)
		{
			string area = GetAreaName(cityName);
			string areaKey = "";
			string trendKey = "";

			if (area != null) {
				areaKey = _sheetKeyDictionary[area];
			}

			if (trendName != null) {
				if (_sheetKeyDictionary.ContainsKey(trendName) == true) {
					trendKey = _sheetKeyDictionary[trendName];
				}
				else {
					Console.WriteLine("InvalidKey: " + trendName);
					return null;
				}
			}

			return string.Format("Time!{0}{1}", areaKey, trendKey);
		}

		public Program()
		{
			// Config used by DiscordSocketClient
			// Define intents for the client
			var config = new DiscordSocketConfig {
				GatewayIntents = GatewayIntents.AllUnprivileged | GatewayIntents.MessageContent
			};

			// It is recommended to Dispose of a client when you are finished
			// using it, at the end of your app's lifetime.
			_client = new DiscordSocketClient(config);

			// Subscribing to client events, so that we may receive them whenever they're invoked.
			_client.Log += LogAsync;
			_client.Ready += ReadyAsync;
			_client.MessageReceived += MessageReceivedAsync;
		}

		public async Task MainAsync(string token)
		{
			// Tokens should be considered secret data, and never hard-coded.
			await _client.LoginAsync(TokenType.Bot, token);
			// Different approaches to making your token a secret is by putting them in local .json, .yaml, .xml or .txt files, then reading them on startup.

			await _client.StartAsync();

			// Block the program until it is closed.
			await Task.Delay(Timeout.Infinite);
		}

		Task LogAsync(LogMessage log)
		{
			Console.WriteLine(log.ToString());
			return Task.CompletedTask;
		}

		// The Ready event indicates that the client has opened a
		// connection and it is now safe to access the cache.
		Task ReadyAsync()
		{
			Console.WriteLine($"{_client.CurrentUser} is connected!");

			return Task.CompletedTask;
		}

		// This is not the recommended way to write a bot - consider
		// reading over the Commands Framework sample.
		async Task MessageReceivedAsync(SocketMessage socketMessage)
		{
			// The bot should never respond to itself.
			if (socketMessage.Author.Id == _client.CurrentUser.Id) {
				return;
			}

			var message = socketMessage as SocketUserMessage;
			string receiveMessage = socketMessage.Content;

			int pos = 0;
			if ((receiveMessage.StartsWith('!') == false
				|| message.HasMentionPrefix(_client.CurrentUser, ref pos))
				|| message.Author.IsBot) {
				return;
			}

			List<string> messageList = receiveMessage.Split(' ').ToList();

			string requestHead = messageList[0].Substring(1);

			_lastSocketChannel = socketMessage.Channel;

			//도시 권역 리퀘스트
			if (requestHead.Equals("권역")) {
				string areaName = GetAreaName(messageList[1]);
				if (areaName != null) {
					string returnMessage = string.Format("{0}의 대유행 권역: {1}", messageList[1], areaName);
					await SendMessage(socketMessage, returnMessage);
				}
				else {
					await SendMessage(socketMessage, string.Format("잘못된 입력입니다: {0}", receiveMessage));
				}
			}
			//대유행 품목 리퀘스트
			else if (requestHead.Equals("품목")) {
				List<object> itemList = GetTrendItems(messageList[1]);
				if (itemList != null) {

					string items = "";
					for(int i = 1; i < itemList.Count; i++) {
						items += itemList[i].ToString() + "\n";
					}
					string returnMessage = string.Format("{0} 대유행 품목:\n{1}", messageList[1], items);
					await SendMessage(socketMessage, returnMessage);
				}
				else {
					await SendMessage(socketMessage, string.Format("잘못된 입력입니다: {0}", receiveMessage));
				}
			}/*
			//대유행 발생 기록
			else if (requestHead.Equals("기록")) {
				string areaName = GetAreaName(messageList[1]);
				if (areaName != null) {
					string sheetKey = GetTimeSheetKey(messageList[1], messageList[2]);
					if (sheetKey != null) {
						if (messageList.Count == 4) {
							List<string> timeString = messageList[3].Split('.').ToList();
							DateTime dt = new DateTime(int.Parse(timeString[0]), int.Parse(timeString[1]), int.Parse(timeString[2]),
								int.Parse(timeString[3]), 0, 0);

							string logMessage = string.Format("{0}({1})에서 대유행:{2} {3}에 발생",
								messageList[1], areaName, messageList[2], DateTimeToString(dt));
							_sheetConnect.InsertData(message.Author.Username, GSheet.KEY.ROWS, sheetKey, new List<object> { DateTimeToString(dt) }, logMessage);

							await SendMessage(socketMessage, string.Format("기록되었습니다.\n{0}", logMessage));
						}
						else {
							DateTime dt = DateTime.Now;
							DateTime hDt = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, 0, 0);

							string logMessage = string.Format("{0}({1})에서 대유행:{2} {3}에 발생",
								messageList[1], areaName, messageList[2], DateTimeToString(hDt));
							_sheetConnect.InsertData(message.Author.Username, GSheet.KEY.ROWS, sheetKey, new List<object> { DateTimeToString(hDt) }, logMessage);

							await SendMessage(socketMessage, string.Format("기록되었습니다.\n{0}", logMessage));
						}
					}
					else {
						await SendMessage(socketMessage, string.Format("잘못된 입력입니다: {0}", receiveMessage));
					}
				}
				else {
					await SendMessage(socketMessage, string.Format("잘못된 입력입니다: {0}", receiveMessage));
				}
			}*/
			else if (requestHead.Equals("발생")) {
				string targetArea = messageList[1];

				if (targetArea.Equals("리스트")) {
					string prevMessage = string.Format("===== 입력 가능한 권역 리스트 =====\n" +
						"▷북해, 발트해\n" +
						"▷서지중해, 대서양\n" +
						"▷동지중해, 흑해\n" +
						"▷동아메리카, 동미, 카리브\n" +
						"▷남아메리카, 남미\n" +
						"▷서아프리카, 서아프\n" +
						"▷남아프리카, 남아프\n" +
						"▷동아프리카, 동아프\n" +
						"▷아라비아, 서인도, 아랍\n" +
						"▷동인도, 인도차이나\n" +
						"▷남아시아, 동남아\n" +
						"▷오세아니아, 호주\n" +
						"▷동아시아, 명\n" +
						"▷극동아시아, 극동, 조선, 일본\n");

					await SendMessage(socketMessage, prevMessage);
				}
				else if (northSeaArea.Contains(targetArea)) {
					string key = "B";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "북해, 발트해 권역==========\n" + msg);
				}
				else if (southMidSeaArea.Contains(targetArea)) {
					string key = "C";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "서지중해, 대서양 권역==========\n" + msg);
				}
				else if (eastMidSeaArea.Contains(targetArea)) {
					string key = "D";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "동지중해, 흑해 권역==========\n" + msg);
				}
				else if (caribArea.Contains(targetArea)) {
					string key = "E";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "카리브, 동미 권역==========\n" + msg);
				}
				else if (southAmericaArea.Contains(targetArea)) {
					string key = "F";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "남아메리카 권역==========\n" + msg);
				}
				else if (westAfricaArea.Contains(targetArea)) {
					string key = "G";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "서아프리카 권역==========\n" + msg);
				}
				else if (southAfricaArea.Contains(targetArea)) {
					string key = "H";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "남아프리카 권역==========\n" + msg);
				}
				else if (eastAfricaArea.Contains(targetArea)) {
					string key = "I";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "동아프리카 권역==========\n" + msg);
				}
				else if (arabiaArea.Contains(targetArea)) {
					string key = "J";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "아라비아, 서인도 권역==========\n" + msg);
				}
				else if (eastIndiaArea.Contains(targetArea)) {
					string key = "K";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "동인도, 인도차이나 권역==========\n" + msg);
				}
				else if (southEastAsiaArea.Contains(targetArea)) {
					string key = "L";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "동남아시아 권역==========\n" + msg);
				}
				else if (oceaniaArea.Contains(targetArea)) {
					string key = "M";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "호주 권역==========\n" + msg);
				}
				else if (chinaArea.Contains(targetArea)) {
					string key = "N";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "명나라 권역==========\n" + msg);
				}
				else if (koreaArea.Contains(targetArea)) {
					string key = "O";
					string msg = GetOccurMessage(socketMessage.Author.Username, string.Format("Time!{0}2:{0}9", key));
					await SendMessage(socketMessage, trendMessage + "조선 일본 권역==========\n" + msg);
				}
			}
			else if (bot_help_key.Contains(requestHead)) {
				string s =
				"안녕하세요 대유행 봇(v1.1)입니다!\n" +
				"아래는 봇 사용 설명서입니다.\n\n" +
				"==================== 정보 확인 ====================\n" +
				"▷	!권역 도시\n" +
				">> 해당 도시의 대유행 권역 반환\n" +
				"\n" +
				"▷	!품목 대유행\n" +
				">> 해당 대유행 품목 반환\n" +
				"\n\n" +
				"==================== 체크 ====================\n" +
				"▷	!발생 리스트\n" +
				" >> 발생 키워드 리스트 출력\n" +
				"\n" +
				"▷	!발생 권역\n" +
				">> 현재 시간 기준으로 해당 권역에 다음번 대유행 발생 시간 추정\n" +
				">> 추정 값으로 기록을 덮어씌움" +
				"\n";


				await SendMessage(socketMessage, s);

			}else if (log.Contains(requestHead)) {
				string s = "업데이트 기록\n" +
					"1.1===============\n" +
					"기록 멍령어 제거\n" +
					"유행 호출 시 권역 이름도 함께 나오도록 변경\n" +
					"대유행 발생 메세지 수정\n";

				await SendMessage(socketMessage, s);
			}
			else {
				await SendMessage(socketMessage, string.Format("알 수 없는 메세지: " + receiveMessage));
			}
		}

		async Task SendMessage(SocketMessage socketMessage, string sendMessage)
		{
			Console.WriteLine(string.Format("[{0}]For {1} Message Sending: {2}\n", DateTime.Now.ToShortTimeString(), socketMessage.Author.Username, sendMessage));
			await socketMessage.Channel.SendMessageAsync(sendMessage);
		}


		static string GetOccurMessage(string userName, string colKey)
		{
			string sheetKey = colKey;
			var list = _sheetConnect.GetSheet(sheetKey);
			List<object> timeList = new List<object>();
			foreach (var l in list) {
				for (int i = 0; i < l.Count; i++) {
					if(l[i].Equals("-") == false) {
						DateTime time = StringToDateTime(l[i].ToString());
						timeList.Add(time);
					}
					else {
						timeList.Add(null);
					}
				}
			}

			//최근 시간에 가깝게 시간 추가
			for(int i = 0; i < timeList.Count; i++) {
				if(i == 0) {//축제
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(FestTime);
						}
						timeList[i] = dt;
					}
				}
				else if(i == 1) {//전염병
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(DesTime);
						}
						timeList[i] = dt;
					}
				}
				else if (i == 2) {//홍수
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(FloodTime);
						}
						timeList[i] = dt;
					}
				}
				else if (i == 3) {//전쟁
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(WarTime);
						}
						timeList[i] = dt;
					}
				}
				else if (i == 4) {//후원
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(SupTime);
						}
						timeList[i] = dt;
					}
				}
				else if (i == 5) {//개발
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(DevTime);
						}
						timeList[i] = dt;
					}
				}
				else if (i == 6) {//호황
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(GoodTime);
						}
						timeList[i] = dt;
					}
				}
				else if (i == 7) {//사치
					if (timeList[i] != null) {
						DateTime dt = (DateTime)timeList[i];
						DateTime now = DateTime.Now;
						while (dt < now) {
							dt = dt.AddHours(RichTime);
						}
						timeList[i] = dt;
					}
				}
			}

			List<string> nextTimeMessageList = new List<string>();
			//DT를 텍스트로 변경
			for(int i = 0; i < timeList.Count; i++) {
				if (timeList[i] != null) {
					DateTime dt = ((DateTime)timeList[i]);
					if (i == 0) {//축제
						nextTimeMessageList.Add(string.Format("●축제({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, FestTime));
					}
					else if (i == 1) {//전염병
						nextTimeMessageList.Add(string.Format("○전염병({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, DesTime));
					}
					else if (i == 2) {//홍수
						nextTimeMessageList.Add(string.Format("●홍수({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, FloodTime));
					}
					else if (i == 3) {//전쟁
						nextTimeMessageList.Add(string.Format("○전쟁({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, WarTime));
					}
					else if (i == 4) {//후원
						nextTimeMessageList.Add(string.Format("●후원({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, SupTime));
					}
					else if (i == 5) {//개발
						nextTimeMessageList.Add(string.Format("○개발({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, DevTime));
					}
					else if (i == 6) {//호황
						nextTimeMessageList.Add(string.Format("●호황({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, GoodTime));
					}
					else if (i == 7) {//사치
						nextTimeMessageList.Add(string.Format("○사치({3}h): {0}월 {1}일 {2}시 예정", dt.Month, dt.Day, dt.Hour, RichTime));
					}
					timeList[i] = DateTimeToString(dt);
				}
				else {
					if (i == 0) {//축제
						nextTimeMessageList.Add("●축제: 이전 기록 없음");
					}
					else if (i == 1) {//전염병
						nextTimeMessageList.Add("○전염병: 이전 기록 없음");
					}
					else if (i == 2) {//홍수
						nextTimeMessageList.Add("●홍수: 이전 기록 없음");
					}
					else if (i == 3) {//전쟁
						nextTimeMessageList.Add("○전쟁: 이전 기록 없음");
					}
					else if (i == 4) {//후원
						nextTimeMessageList.Add("●후원: 이전 기록 없음");
					}
					else if (i == 5) {//개발
						nextTimeMessageList.Add("○개발: 이전 기록 없음");
					}
					else if (i == 6) {//호황
						nextTimeMessageList.Add("●호황: 이전 기록 없음");
					}
					else if (i == 7) {//사치
						nextTimeMessageList.Add("○사치: 이전 기록 없음");
					}
				}
			}

			//다음 변경 시간 미리 세팅
			_sheetConnect.InsertData(userName, GSheet.KEY.COLUMNS, sheetKey, timeList, "Trend Time Call");

			string res = "";
			for(int i = 0; i < nextTimeMessageList.Count; i++) {
				res += nextTimeMessageList[i] + "\n";
			}

			return res;
		}

		static string DateTimeToString(DateTime dt)
		{
			return string.Format("{0}:{1}:{2}:{3}", dt.Year, dt.Month, dt.Day, dt.Hour);
		}

		static DateTime StringToDateTime(string time)
		{
			List<int> timeList = time.Split(':').Select(a => int.Parse(a)).ToList();
			return new DateTime(timeList[0], timeList[1], timeList[2], timeList[3], 0, 0);
		}
	}

	public class GSheet
	{
		public enum KEY
		{
			ROWS,
			COLUMNS
		}

		static string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static string ApplicationName = "Google Sheets API .NET Quickstart";
		string sheetID = "";
		static string jsonPath = "client_secret.json";

		SheetsService? _service;
		bool _isCredentialed = false;

		public GSheet()
		{
			sheetID =  System.IO.File.ReadAllText("gsheetID.txt");
			Credentialize();
		}

		void Credentialize()
		{
			string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
			credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

			UserCredential credential;

			using (var stream = new FileStream(jsonPath, FileMode.Open, FileAccess.Read)) {
				credential = GetCredential(stream, credPath);
			}

			// Create Google Sheets API service.
			_service = new SheetsService(new BaseClientService.Initializer() {
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});

			_isCredentialed = _service != null;
		}

		UserCredential GetCredential(FileStream stream, string path)
		{
			return GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.FromStream(stream).Secrets, Scopes, "user",
					CancellationToken.None, new FileDataStore(path, true)).Result;
		}

		/// <summary>
		/// 시트 가져오기
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="col"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public IList<IList<object>> GetSheet(string sheetKey)
		{
			return SelectData(sheetKey);
		}

		IList<IList<object>> SelectData(string sheetKey)
		{
			if (_isCredentialed == false) {
				Credentialize();
				Console.WriteLine("Error: Recredentialize");
				return null;
			}
			else {
				var request = _service.Spreadsheets.Values.Get(sheetID, sheetKey);
				try {
					ValueRange response = request.Execute();
					var data = response.Values;
					return data;
				}
				catch (Exception ex) {
					Console.WriteLine(ex.Message);
					return null;
				}
			}
		}

		/// <summary>
		/// 값 입력
		/// </summary>
		/// <param name="rowCol">입력 방향(ROWS: 가로부터, COLUMNS: 세로부터)</param>
		/// <param name="sheetKey">입력 위치(시트!시작점:끝점)</param>
		/// <param name="list_data">입력 값</param>
		public void InsertData(string userName, KEY rowCol, string sheetKey, List<object> list_data, string logMessage = "")
		{
			if (_isCredentialed == false) {
				Credentialize();
				Console.WriteLine("Error: Recredentialize");
				return;
			}
			else {
				var valueRange = new ValueRange() {
					// ROWS or COLUMNS
					MajorDimension = rowCol.ToString(),
					// 추가할 데이터
					Values = new List<IList<object>> { list_data }
				};

				string s = "";
				foreach(var d in list_data) {
					s += d + "&";
				}

				Console.WriteLine(string.Format("[{0}]{1}'s Data Insert: {2}\n", DateTime.Now.ToShortTimeString(), userName, sheetKey, s));

				var update = _service.Spreadsheets.Values.Update(valueRange, sheetID, sheetKey);
				update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
				update.Execute();

				AppendLogData(userName, KEY.ROWS, "Log!A1", logMessage);
			}
		}

		public void AppendLogData(string userName, KEY rowCol, string sheetKey, string logMessage)
		{
			if (_isCredentialed == false) {
				Credentialize();
				Console.WriteLine("Error: Recredentialize");
				return;
			}
			else {
				string timeText = DateTime.Now.ToString("yy:MM:dd:HH:mm:ss");

				var valueRange = new ValueRange() {
					// ROWS or COLUMNS
					MajorDimension = rowCol.ToString(),
					// 추가할 데이터
					Values = new List<IList<object>> { new List<object> { string.Format("[{0}][{1}] LOG: {2}", timeText, userName, logMessage) } }
				};

				Console.WriteLine(string.Format("[{0}]{1}'s Data Append: {2}\n", timeText, userName, sheetKey, logMessage));

				var update = _service.Spreadsheets.Values.Append(valueRange, sheetID, sheetKey);
				update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
				update.Execute();
			}
		}
	}
}