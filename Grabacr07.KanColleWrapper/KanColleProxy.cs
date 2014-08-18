using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Fiddler;
using Grabacr07.KanColleWrapper.Win32;
using Livet;

namespace Grabacr07.KanColleWrapper
{
	public partial class KanColleProxy
	{
		private readonly IConnectableObservable<Session> connectableSessionSource;
		private readonly IConnectableObservable<Session> apiSource;
		private readonly LivetCompositeDisposable compositeDisposable;

		public IObservable<Session> SessionSource
		{
		    get { return this.connectableSessionSource.AsObservable(); }
		}

	    public IObservable<Session> ApiSessionSource
	    {
	        get { return this.apiSource.AsObservable(); }
	    }

	    public IProxySettings UpstreamProxySettings { get; set; }

        /// <summary>
        /// 通信結果を艦これ統計データベースに送信するかどうかを取得または設定します。
        /// </summary>
        public bool SendDb { get; set; }

        /// <summary>
        /// 艦これ統計データベースのアクセスキーを取得または設定します。
        /// </summary>
        public string DbAccessKey { get; set; }


		public KanColleProxy()
		{
			this.compositeDisposable = new LivetCompositeDisposable();

			this.connectableSessionSource = Observable
				.FromEvent<SessionStateHandler, Session>(
					action => new SessionStateHandler(action),
					h => FiddlerApplication.AfterSessionComplete += h,
					h => FiddlerApplication.AfterSessionComplete -= h)
				.Publish();

			this.apiSource = this.connectableSessionSource
				.Where(s => s.PathAndQuery.StartsWith("/kcsapi"))
				.Where(s => s.oResponse.MIMEType.Equals("text/plain"))
				#region .Do(debug)
#if DEBUG
.Do(session =>
				{
					Debug.WriteLine("==================================================");
					Debug.WriteLine("Fiddler session: ");
					Debug.WriteLine(session);
					Debug.WriteLine("");
				})
#endif
			#endregion
                #region .Do(send-database)
.Do(s =>
{
    // 艦これ統計データベースへ送信が有効で、かつアクセスキーが入力されていた場合は送信
    if (this.SendDb && !this.DbAccessKey.IsEmpty())
    {
        string[] urls = 
						{
							"api_port/port",
							"api_get_member/kdock",
							"api_get_member/ship2",
							"api_get_member/ship3",
							"api_req_hensei/change",
							"api_req_kousyou/createship",
							"api_req_kousyou/getship",
							"api_req_kousyou/createitem",
							"api_req_map/start",
							"api_req_map/next",
							"api_req_sortie/battle",
							"api_req_battle_midnight/battle",
							"api_req_battle_midnight/sp_midnight",
							"api_req_sortie/night_to_day",
							"api_req_sortie/battleresult",
							"api_req_practice/battle",
							"api_req_practice/battle_result",
                            "api_req_combined_battle/battle",
                            "api_req_combined_battle/airbattle",
                            "api_req_combined_battle/midnight_battle",
                            "api_req_combined_battle/battleresult",
						};
        foreach (var url in urls)
        {
            if (s.fullUrl.IndexOf(url) > 0)
            {
                using (System.Net.WebClient wc = new System.Net.WebClient())
                {
                    System.Collections.Specialized.NameValueCollection post = new System.Collections.Specialized.NameValueCollection();
                    post.Add("token", this.DbAccessKey);
                    post.Add("agent", "LZXNXVGPejgSnEXLH2ur");  // このクライアントのエージェントキー
                    post.Add("url", s.fullUrl);
                    string requestBody = System.Text.RegularExpressions.Regex.Replace(s.GetRequestBodyAsString(), @"&api(_|%5F)token=[0-9a-f]+|api(_|%5F)token=[0-9a-f]+&?", "");	// api_tokenを送信しないように削除
                    post.Add("requestbody", requestBody);
                    post.Add("responsebody", s.GetResponseBodyAsString());

                    wc.UploadValuesAsync(new Uri("http://api.kancolle-db.net/2/"), post);
#if DEBUG
									Debug.WriteLine("==================================================");
									Debug.WriteLine("Send to KanColle statistics database");
									Debug.WriteLine(s.fullUrl);
									Debug.WriteLine("==================================================");
#endif
                }
                break;
            }
        }
    }
})
            #endregion
				.Publish();
		}


		public void Startup(int proxy = 37564)
		{
			FiddlerApplication.Startup(proxy, false, true);
			FiddlerApplication.BeforeRequest += this.SetUpstreamProxyHandler;

			SetIESettings("localhost:" + proxy);

			this.compositeDisposable.Add(this.connectableSessionSource.Connect());
			this.compositeDisposable.Add(this.apiSource.Connect());
		}

		public void Shutdown()
		{
			this.compositeDisposable.Dispose();

			FiddlerApplication.BeforeRequest -= this.SetUpstreamProxyHandler;
			FiddlerApplication.Shutdown();
		}


		private static void SetIESettings(string proxyUri)
		{
			// ReSharper disable InconsistentNaming
			const int INTERNET_OPTION_PROXY = 38;
			const int INTERNET_OPEN_TYPE_PROXY = 3;
			// ReSharper restore InconsistentNaming

			INTERNET_PROXY_INFO proxyInfo;
			proxyInfo.dwAccessType = INTERNET_OPEN_TYPE_PROXY;
			proxyInfo.proxy = Marshal.StringToHGlobalAnsi(proxyUri);
			proxyInfo.proxyBypass = Marshal.StringToHGlobalAnsi("local");

			var proxyInfoSize = Marshal.SizeOf(proxyInfo);
			var proxyInfoPtr = Marshal.AllocCoTaskMem(proxyInfoSize);
			Marshal.StructureToPtr(proxyInfo, proxyInfoPtr, true);

			NativeMethods.InternetSetOption(IntPtr.Zero, INTERNET_OPTION_PROXY, proxyInfoPtr, proxyInfoSize);
		}

		/// <summary>
		/// Fiddler からのリクエスト発行時にプロキシを挟む設定を行います。
		/// </summary>
		/// <param name="requestingSession">通信を行おうとしているセッション。</param>
		private void SetUpstreamProxyHandler(Session requestingSession)
		{
			var settings = this.UpstreamProxySettings;
			if (settings == null) return;

			var useGateway = !string.IsNullOrEmpty(settings.Host) && settings.IsEnabled;
			if (!useGateway || (IsSessionSSL(requestingSession) && !settings.IsEnabledOnSSL)) return;

			var gateway = settings.Host.Contains(":")
				// IPv6 アドレスをプロキシホストにした場合はホストアドレス部分を [] で囲う形式にする。
				? string.Format("[{0}]:{1}", settings.Host, settings.Port)
				: string.Format("{0}:{1}", settings.Host, settings.Port);

			requestingSession["X-OverrideGateway"] = gateway;
		}

		/// <summary>
		/// セッションが SSL 接続を使用しているかどうかを返します。
		/// </summary>
		/// <param name="session">セッション。</param>
		/// <returns>セッションが SSL 接続を使用する場合は true、そうでない場合は false。</returns>
		internal static bool IsSessionSSL(Session session)
		{
			// 「http://www.dmm.com:433/」の場合もあり、これは Session.isHTTPS では判定できない
			return session.isHTTPS || session.fullUrl.StartsWith("https:") || session.fullUrl.Contains(":443");
		}
	}
}
