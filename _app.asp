<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.12.2023 11:39 @badursun
'**********************************************

' Set rsTempClass = New mngkargo_plugin 'MNGKargoServices
' 	rsTempClass.class_register()
' Set rsTempClass = Nothing

Class mngkargo_plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD

	Private MNGKARGO_CLIENTID, MNGKARGO_SECRET, MNGKARGO_API_USER, MNGKARGO_API_PASS, MNGKARGO_API_ID, MNGKARGO_API_HOST
	Private MNGKARGO_JWT, MNGKARGO_JWT_EXPIRE, MNGKARGO_REFRESH_TOKEN
	Private GENERATE_TOKEN_URL, REFRESH_TOKEN_URL, CREATE_ORDER, GET_ORDER, TRACK_ORDER
	Private DEBUG_MODE, MNGKARGO_ACTIVE
	Private smsPreference1, smsPreference2, smsPreference3
	Private DATA_REFERANS_ID, DATA_CONTENT, DATA_DESCRIPTION, DATA_IL, DATA_ILCE, DATA_ADRES, DATA_EPOSTA, DATA_ALICI, DATA_TELEFON, DATA_PAYATDOOR, DATA_PAYATDOOR_PRICE, DATA_DELIVERY_TYPE, DATA_PACK_TYPE, DATA_PAYMENT_TYPE, DATA_USER_ID
	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		
		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Check And Create Table
		'------------------------------
		Dim PluginTableName
			PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME

    	' If TableExist(PluginTableName) = False Then
    	' 	Conn.Execute("SET NAMES utf8mb4;") 
    	' 	Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
    	' 	Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"`")

    	' 	q=""
    	' 	q=q+"CREATE TABLE `"& PluginTableName &"` ( "
    	' 	q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
    	' 	q=q+"  `ORDER_ID` int(11) DEFAULT 0, "
    	' 	q=q+"  `ORDER_NO` bigint(20) DEFAULT 0, "
    	' 	q=q+"  `GUID` varchar(100) DEFAULT NULL, "
    	' 	q=q+"  `MSG` varchar(255) DEFAULT NULL, "
    	' 	q=q+"  `FATURA_TARIHI` datetime DEFAULT current_timestamp(), "
    	' 	q=q+"  `FATURA_URL` varchar(255) DEFAULT NULL, "
    	' 	q=q+"  PRIMARY KEY (`ID`), "
    	' 	q=q+"  KEY `IND1` (`ORDER_ID`) "
    	' 	q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
		' 	Conn.Execute(q)

    	' 	Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 

		' 	' Create Log
		' 	'------------------------------
    	' 	Call PanelLog(""& PLUGIN_CODE &" için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

		' 	' Register Settings
		' 	'------------------------------
		' 	DebugTimer ""& PLUGIN_CODE &" class_register() End"
    	' End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "mngkargo_plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "591")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "1")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)

		a=GetSettings(""&PLUGIN_CODE&"_API_KEY", "")
		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "AJAX:RemoveLog" Then
    		Query.PageContentType = "json"
    		
    		Call AdminSessionChecker()

    		REC_ID = Query.Data("RecID")

		    ' Conn.Execute("DELETE FROM tbl_plugin_bizimhesap WHERE ID="& REC_ID &"")
		    
		    Query.jsonResponse 200, "Güncellendi"
    		
    		Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write  		QuickSettings("checkbox", "MNGKARGO_ACTIVE", "Plugin Durumu", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-8 col-sm-12"">"
			.Write 			QuickSettings("input", "MNGKARGO_API_HOST", "API Host URL", "https://testapi.mngkargo.com.tr/", TO_DB)
			.Write "    </div>"

			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", "MNGKARGO_CLIENTID", "İstemci Tanımlayıcısı (Client ID)", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", "MNGKARGO_SECRET", "İstemci Güvenlik Dizgisi (Secret Key)", "", TO_DB)
			.Write "    </div>"

			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("input", "MNGKARGO_API_USER", "API Kullanıcı", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("input", "MNGKARGO_API_PASS", "API Parola", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("input", "MNGKARGO_API_ID", "API Kimlik (Identity)", "", TO_DB)
			.Write "    </div>"

			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("checkbox", "MNGKARGO_SMS_1", "Branch Kargo varış şubesine ulaştığında alıcıya SMS gitsin mi?", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("checkbox", "MNGKARGO_SMS_2", "Kargo ilk hazırlandığında alıcıya SMS gitsin mi?", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("checkbox", "MNGKARGO_SMS_3", "Kargo teslim edildiğinde göndericiye SMS gitsin mi?", "", TO_DB)
			.Write "    </div>"
			.Write "</div>"
			' .Write "<div class=""row"">"
			' .Write "    <div class=""col-lg-12 col-sm-12"">"
			' .Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=BizimHesapLog"" class=""btn btn-sm btn-primary"">"
			' .Write "        	Önbelleklenmiş Dosyaları Göster"
			' .Write "        </a>"
			' .Write "    </div>"
			' .Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	private sub class_initialize()
		DebugTimer "MNGKargoServices class_initialize() Start"
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_NAME 			= "MNG Kargo Plugin"
    	PLUGIN_CODE 			= "MNGKARGO"
    	PLUGIN_DB_NAME 			= "mngkargo"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/MNGKargo-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_FOLDER_NAME 		= "mngkargo-plugin"
    	PLUGIN_ICON 			= "zmdi-local-shipping"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= True
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

    	DEBUG_MODE 				= False
		DATA_REFERANS_ID		= ""
		DATA_PAYATDOOR			= 0
		DATA_PAYATDOOR_PRICE	= 0
		DATA_USER_ID			= 0
		DATA_DELIVERY_TYPE 		= 1 ' 1:STANDART_TESLİMAT, 7:GUNİCİ_TESLİMAT, 8:AKŞAM_TESLİMAT
		DATA_CONTENT 			= ""
		DATA_DESCRIPTION 		= ""
		DATA_IL 				= ""
		DATA_ILCE 				= ""
		DATA_ADRES 				= ""
		DATA_EPOSTA 			= ""
		DATA_ALICI 				= ""
		DATA_TELEFON 			= ""
		DATA_PACK_TYPE 			= 3 ' 1:DOSYA, 2:Mİ, 3:PAKET, 4:KOLİ Kargo Cinsi
		DATA_PAYMENT_TYPE 		= 2 ' 1:GONDERICI_ODER, 2:ALICI_ODER,3:PLATFORM_ODER.
		
		MNGKARGO_ACTIVE 		= Cint( GetSettings("MNGKARGO_ACTIVE", 0) )
    	MNGKARGO_CLIENTID 		= GetSettings("MNGKARGO_CLIENTID", "")
    	MNGKARGO_SECRET 		= GetSettings("MNGKARGO_SECRET", "")
    	MNGKARGO_API_USER 		= GetSettings("MNGKARGO_API_USER", "")
    	MNGKARGO_API_PASS 		= GetSettings("MNGKARGO_API_PASS", "")
    	MNGKARGO_API_ID 		= GetSettings("MNGKARGO_API_ID", "0")
    	MNGKARGO_API_HOST 		= GetSettings("MNGKARGO_API_HOST", "https://testapi.mngkargo.com.tr/")
    	smsPreference1 			= Cint( GetSettings("MNGKARGO_SMS_1", "1") )
    	smsPreference2 			= Cint( GetSettings("MNGKARGO_SMS_2", "1") )
    	smsPreference3 			= Cint( GetSettings("MNGKARGO_SMS_3", "1") )
    	MNGKARGO_JWT 			= GetSettings("MNGKARGO_JWT", "")
    	MNGKARGO_JWT_EXPIRE 	= GetSettings("MNGKARGO_JWT_EXPIRE", "")
    	MNGKARGO_REFRESH_TOKEN 	= GetSettings("MNGKARGO_REFRESH_TOKEN", "")
    	GENERATE_TOKEN_URL 		= MNGKARGO_API_HOST & "mngapi/api/token"
    	REFRESH_TOKEN_URL 		= MNGKARGO_API_HOST & "mngapi/api/refresh"
    	CREATE_ORDER 			= MNGKARGO_API_HOST & "mngapi/api/standardcmdapi/createOrder"
    	GET_ORDER 				= MNGKARGO_API_HOST & "mngapi/api/standardqueryapi/getorder/"
    	TRACK_ORDER 			= MNGKARGO_API_HOST & "mngapi/api/standardqueryapi/trackshipment/"
    	  
    	If MNGKARGO_JWT_EXPIRE="" Then 
    		' Response.Write "<h3>Class Init Non-Exist Token Refresh</h3>"
    		GetToken()
    	Else 
		    If IsDate(MNGKARGO_JWT_EXPIRE) Then
		        If DateDiff("h", Now(), CDate(MNGKARGO_JWT_EXPIRE)) < 2 Then 
    				' Response.Write "<h3>Class Init Expired Token Refresh</h3>"
					' RefreshToken()
					GetToken()
		        End If
		    End If
    	End If

    	
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()

    	'-------------------------------------------------------------------------------------
    	' Hook Auto Load Plugin
    	'-------------------------------------------------------------------------------------
    	If PLUGIN_AUTOLOAD_AT("WEB") = True Then 

    	End If
		DebugTimer ""& PLUGIN_CODE &" class_initialize() End"
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Property Get XMLHttp(Uri, xType, Data, AuthType)
		If MNGKARGO_ACTIVE = 0 Then Exit Property

		' ByPass Error
		'------------------------------------------------
		On Error Resume Next

		' Response.Write IIf(DEBUG_MODE=True, "<div><strong>MNGKARGO_CLIENTID :</strong> "& MNGKARGO_CLIENTID &"</div>", "")
		' Response.Write IIf(DEBUG_MODE=True, "<div><strong>MNGKARGO_SECRET :</strong> "& MNGKARGO_SECRET &"</div>", "")
		' Response.Write IIf(DEBUG_MODE=True, "<div><strong>MNGKARGO_API_USER :</strong> "& MNGKARGO_API_USER &"</div>", "")
		' Response.Write IIf(DEBUG_MODE=True, "<div><strong>MNGKARGO_API_PASS :</strong> "& MNGKARGO_API_PASS &"</div>", "")
		' Response.Write IIf(DEBUG_MODE=True, "<div><strong>MNGKARGO_API_ID :</strong> "& MNGKARGO_API_ID &"</div>", "")

		' Send Data
		'------------------------------------------------
	    'Set objXMLhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
	    Set objXMLhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0") 
			objXMLhttp.open xType, Uri, false
            objXMLhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
            objXMLhttp.setTimeouts 5000, 5000, 10000, 10000 'ms
			objXMLhttp.setRequestHeader "X-IBM-Client-Id" 		, MNGKARGO_CLIENTID
			objXMLhttp.setRequestHeader "X-IBM-Client-Secret"	, MNGKARGO_SECRET
			If AuthType = "BEARER" Then 
				objXMLhttp.setRequestHeader "Authorization"		, "Bearer " & MNGKARGO_JWT
			End If
			objXMLhttp.setRequestHeader "Accept" 				, "application/json"
			objXMLhttp.setRequestHeader "Content-type" 			, "application/json"
			objXMLhttp.setRequestHeader "Content-Length" 		, Len(Data)
			objXMLhttp.send TurkceKarakter2HTML( Data )

			If objXMLhttp.Status = 200 Then 
				' Response.Write IIf(DEBUG_MODE=True, "objXMLhttp.responseText: " & objXMLhttp.responseText & "<br>", "")
				' Response.Write IIf(DEBUG_MODE=True, "objXMLhttp.Status: " & objXMLhttp.Status & "<br>", "")
				API_REQUEST_HASH 				= Uguid()
				API_REQUEST_DATA 				= Data 
				
				' Conn.Execute("INSERT tbl_api_logs(HASH, START_TIME, END_TIME, APIKEY, IP, END_POINT, API_REQUEST, API_METHOD, HTTP_STATUS, API_RESPONSE) VALUES('"& API_REQUEST_HASH &"', UNIX_TIMESTAMP(), UNIX_TIMESTAMP(), 'MNG Kargo WS', '"& IPAdresi() &"', 'MNG.XMLHttp', '"& API_REQUEST_DATA &"', '"& xType &"', '"& objXMLhttp.Status &"','"& jSONReplace(TurkceKarakter2HTML(objXMLhttp.responseText)) &"')")

				CreateLog ""& PLUGIN_CODE &".XMLHttp", API_REQUEST_DATA, objXMLhttp.responseText, objXMLhttp.Status, "POST"
				
				XMLHttp = Array(objXMLhttp.Status, objXMLhttp.responseText)
			Else 
				On Error Resume Next
				Set parseJsonData = New aspJSON
					parseJsonData.loadJSON( objXMLhttp.responseText )
		    		Dim ERROR_TEXT 
		    			ERROR_TEXT = ""
					If IsNull(parseJsonData.data("error")) = True Then 
						ERROR_TEXT = "MNG Kargo Bağlantı Hatası"
					Else
						ERROR_TEXT = parseJsonData.data("error").item("Description")
					End IF
				Set parseJsonData = Nothing
				On Error Goto 0

				' Conn.Execute("INSERT tbl_api_logs(HASH, START_TIME, END_TIME, APIKEY, IP, END_POINT, API_REQUEST, API_METHOD, HTTP_STATUS, API_RESPONSE) VALUES('"& API_REQUEST_HASH &"', UNIX_TIMESTAMP(), UNIX_TIMESTAMP(), 'MNG Kargo WS', '"& IPAdresi() &"', 'MNG.XMLHttp', '"& API_REQUEST_DATA &"', '"& xType &"', '"& objXMLhttp.Status &"','"& jSONReplace(TurkceKarakter2HTML(objXMLhttp.responseText)) &"')")

				CreateLog ""& PLUGIN_CODE &".XMLHttp", "Status Error", ERROR_TEXT, objXMLhttp.Status, "POST"
				
				XMLHttp = Array(500, ERROR_TEXT)
			End If
	    Set objXMLhttp = Nothing
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get GetToken()
		If MNGKARGO_ACTIVE = 0 Then Exit Property

		If Len(MNGKARGO_CLIENTID) < 2 Then 
			Err.Raise 591, "MNGKargoServices Class", "Client ID Eksik"
			Exit Property
		End If

		If Len(MNGKARGO_SECRET) < 2 Then 
			Err.Raise 591, "MNGKargoServices Class", "Client Secret Eksik"
			Exit Property
		End If

	    Set oJSON = New aspJSON
	        With oJSON.data
	            .Add "customerNumber"       , ""& MNGKARGO_API_USER &""
	            .Add "password"          	, ""& MNGKARGO_API_PASS &""
	            .Add "identityType"         , CLng(MNGKARGO_API_ID)
	        End With
	        JSON_DATA = oJSON.JSONoutput()
	        ' Response.Write IIf(DEBUG_MODE=True, "<pre>"& JSON_DATA &"</pre>", "")
	    Set oJSON = Nothing

	    Dim TokenResult
	    	TokenResult = XMLHttp(GENERATE_TOKEN_URL, "POST", JSON_DATA, "")

	    If TokenResult(0) = 200 Then 
			Set parseJsonData = New aspJSON
				parseJsonData.loadJSON( TokenResult(1) )

	    		MNGKARGO_JWT 			= SetSettings("MNGKARGO_JWT", parseJsonData.data("jwt"))
	    		MNGKARGO_JWT_EXPIRE 	= SetSettings("MNGKARGO_JWT_EXPIRE", parseJsonData.data("jwtExpireDate"))
	    		MNGKARGO_REFRESH_TOKEN 	= SetSettings("MNGKARGO_REFRESH_TOKEN", parseJsonData.data("refreshToken"))

			Set parseJsonData = Nothing
	    End If

	    GetToken = TokenResult
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get GetOrder(referenceId)
	    Dim OrderResult
	    	OrderResult = XMLHttp(GET_ORDER & referenceId, "GET", "", "BEARER")
	    
	    If OrderResult(0) = 200 Then 
	    	Response.Write IIf(DEBUG_MODE=True, "<h4>GetOrder Success</h4>", "")
	    	Response.Write IIf(DEBUG_MODE=True, "<pre>"& OrderResult(1) &"</pre>", "")

	    	' Set parseJsonData = New aspJSON
	    	' 	parseJsonData.loadJSON( OrderResult(1) )
	    	' 	Response.Write IIf(DEBUG_MODE=True, "<h1>orderInvoiceId: "& parseJsonData.data(0).item("orderInvoiceId") &"</h1>", "")
	    	' 	Response.Write IIf(DEBUG_MODE=True, "<h1>orderInvoiceDetailId: "& parseJsonData.data(0).item("orderInvoiceDetailId") &"</h1>", "")
	    	' 	Response.Write IIf(DEBUG_MODE=True, "<h1>shipperBranchCode: "& parseJsonData.data(0).item("shipperBranchCode") &"</h1>", "")
	    	' Set parseJsonData = Nothing
		Else
	    	Response.Write IIf(DEBUG_MODE=True, "<h4>GetOrder Failed</h4>", "")
	    	Response.Write IIf(DEBUG_MODE=True, "<pre>"& OrderResult(1) &"</pre>", "")
	    End If
	    GetOrder = OrderResult
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get TrackOrder(referenceId)
	    Dim OrderResult
	    	OrderResult = XMLHttp(TRACK_ORDER & referenceId, "GET", "", "BEARER")
	    
	    If OrderResult(0) = 200 Then 
	    	Response.Write IIf(DEBUG_MODE=True, "<h4>TrackOrder Success</h4>", "")
	    	Response.Write IIf(DEBUG_MODE=True, "<pre>"& OrderResult(1) &"</pre>", "")

	    	' Set parseJsonData = New aspJSON
	    	' 	parseJsonData.loadJSON( OrderResult(1) )
	    	' 	Response.Write IIf(DEBUG_MODE=True, "<h1>orderInvoiceId: "& parseJsonData.data(0).item("orderInvoiceId") &"</h1>", "")
	    	' 	Response.Write IIf(DEBUG_MODE=True, "<h1>orderInvoiceDetailId: "& parseJsonData.data(0).item("orderInvoiceDetailId") &"</h1>", "")
	    	' 	Response.Write IIf(DEBUG_MODE=True, "<h1>shipperBranchCode: "& parseJsonData.data(0).item("shipperBranchCode") &"</h1>", "")
	    	' Set parseJsonData = Nothing
		Else
	    	Response.Write IIf(DEBUG_MODE=True, "<h4>TrackOrder Failed</h4>", "")
	    	Response.Write IIf(DEBUG_MODE=True, "<pre>"& OrderResult(1) &"</pre>", "")
	    End If
	    TrackOrder = OrderResult
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let SiparisID(val) 			: DATA_REFERANS_ID = val 		: End Property
	Public Property Let SiparisIcerik(val) 		: DATA_CONTENT = val 			: End Property
	Public Property Let SiparisAciklama(val) 	: DATA_DESCRIPTION = val 		: End Property
	Public Property Let SiparisUyeID(val) 		: DATA_USER_ID = val 			: End Property
	Public Property Let TeslimatTipi(val) 		: DATA_DELIVERY_TYPE = val 		: End Property
	Public Property Let PaketTipi(val) 			: DATA_PACK_TYPE = val 			: End Property
	Public Property Let OdemeyiYapan(val) 		: DATA_PAYMENT_TYPE = val 		: End Property
	Public Property Let KargoIl(val) 			: DATA_IL = val 				: End Property
	Public Property Let KargoIlce(val) 			: DATA_ILCE = val 				: End Property
	Public Property Let KargoAdres(val) 		: DATA_ADRES = val 				: End Property
	Public Property Let KargoEPosta(val) 		: DATA_EPOSTA = val 			: End Property
	Public Property Let KargoAlici(val) 		: DATA_ALICI = val 				: End Property
	Public Property Let KargoTelefon(val) 		: DATA_TELEFON = val 			: End Property
	Public Property Let KapidaOdeme(val) 		: DATA_PAYATDOOR = val 			: End Property
	Public Property Let KapidaOdemeTutar(val) 	: DATA_PAYATDOOR_PRICE = val 	: End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get CreateOrder()
	    i=0
	    Set oJSON = New aspJSON
	        With oJSON.data
	        	.Add "order" 					, oJSON.Collection()
	        	With oJSON.data("order")
	        		.Add "referenceId" 			, DATA_REFERANS_ID '"TESTORDER666"
	        		.Add "barcode" 				, DATA_REFERANS_ID '"TESTORDER666"
	        		.Add "billOfLandingId" 		, "İrsaliye 1"
	        		
	        		.Add "isCOD" 				, DATA_PAYATDOOR
	        		.Add "codAmount" 			, DATA_PAYATDOOR_PRICE
	        		
	        		.Add "shipmentServiceType" 	, DATA_DELIVERY_TYPE
	        		.Add "packagingType" 		, DATA_PACK_TYPE
	        		.Add "content" 				, ""& Left(DATA_CONTENT, 250) &"" '"Ürün İçeriği"
	        		.Add "smsPreference1" 		, smsPreference1 ' Branch Kargo varış şubesine ulaştığında alıcıya SMS gitsin mi?
	        		.Add "smsPreference2" 		, smsPreference2 ' Kargo ilk hazırlandığında alıcıya SMS gitsin mi?
	        		.Add "smsPreference3" 		, smsPreference3 ' Kargo teslim edildiğinde göndericiye SMS gitsin mi?
	        		.Add "paymentType" 			, DATA_PAYMENT_TYPE
	        		.Add "deliveryType" 		, 1
	        		.Add "description" 			, ""& Left(DATA_DESCRIPTION, 250) &"" '"Açıklama Metni"
	        		.Add "marketPlaceShortCode" , ""
	        		.Add "marketPlaceSaleCode" 	, ""
	        	End With

	        	.Add "orderPieceList" 			, oJSON.Collection()
	        	With oJSON.data("orderPieceList")
					.Add i, oJSON.Collection()
					With .item(i)
						.Add "barcode" 			, "URUNID_123"
						.Add "desi" 			, 2
						.Add "kg" 				, 2
						.Add "content" 			, "Satın Alınan Ürün"
	        		End With
	        		i=i+1
	        	End With

	        	.Add "recipient" 				, oJSON.Collection()
	        	With oJSON.data("recipient")
	            	.Add "customerId"          	, ""
	            	.Add "refCustomerId"        , ""
	            	.Add "cityCode"          	, 0
	            	.Add "cityName"          	, DATA_IL '"İstanbul"
	            	.Add "districtCode"         , 0
	            	.Add "districtName"         , DATA_ILCE '"Üsküdar"
	            	.Add "address"          	, DATA_ADRES '"Küçüksu Mah. Mustafa Düzgünman Cad. Tüylüoğlu Sk. No:1/B Çengelköy"
	            	.Add "email" 				, DATA_EPOSTA ' "badursun@gmail.com"
	            	.Add "taxOffice" 			, ""
	            	.Add "taxNumber" 			, ""
	            	.Add "fullName" 			, DATA_ALICI' "Anthony Burak DURSUN"
	            	.Add "homePhoneNumber" 		, ""
	            	.Add "mobilePhoneNumber" 	, PhoneNumberFixer(DATA_TELEFON) '"0507 342 23 45"
	        	End With

	        End With
	        JSON_DATA = oJSON.JSONoutput()
	    Set oJSON = Nothing

	    Dim OrderResult
	    	OrderResult = XMLHttp(CREATE_ORDER, "POST", JSON_DATA, "BEARER")

	    If OrderResult(0) = 200 Then 
	    	Response.Write IIf(DEBUG_MODE=True, "<h4>CreateOrder Success</h4>", "")
	    	Response.Write IIf(DEBUG_MODE=True, "<pre>"& OrderResult(1) &"</pre>", "")

	    	Set parseJsonData = New aspJSON
	    		parseJsonData.loadJSON( OrderResult(1) )
	    		Response.Write IIf(DEBUG_MODE=True, "<h1>orderInvoiceId: "& parseJsonData.data(0).item("orderInvoiceId") &"</h1>", "")
	    		Response.Write IIf(DEBUG_MODE=True, "<h1>orderInvoiceDetailId: "& parseJsonData.data(0).item("orderInvoiceDetailId") &"</h1>", "")
	    		Response.Write IIf(DEBUG_MODE=True, "<h1>shipperBranchCode: "& parseJsonData.data(0).item("shipperBranchCode") &"</h1>", "")
	    	Set parseJsonData = Nothing
		Else
	    	Response.Write IIf(DEBUG_MODE=True, "<h4>CreateOrder Failed</h4>", "")
	    	Response.Write IIf(DEBUG_MODE=True, "<pre>"& OrderResult(1) &"</pre>", "")
	    End If

	    CreateOrder = OrderResult
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function PhoneNumberFixer(Text)
		Dim Txt 
			Txt = Text & ""

		If IsNull(Txt) OR IsEmpty(Txt) OR Txt = "" OR Len(Txt) < 1 Then 
			PhoneNumberFixer = Txt
			Exit Function
		End If

		Txt = Replace(Txt,"+90" ,"0" ,1,-1,1)  
		Txt = Replace(Txt,"+9 0" ,"0" ,1,-1,1)  
		PhoneNumberFixer = Txt  
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function TurkceKarakter2HTML(Txt)
		Txt = Txt & ""
		If IsNull(Txt) OR IsEmpty(Txt) OR Txt = "" OR Len(Txt) < 1 Then 
			Txt = ""
			Exit Function
		End If

		Txt = Replace(Txt,"ğ" ,"\u011F" ,1,-1,1)  
		Txt = Replace(Txt,"Ğ" ,"\u011E" ,1,-1,1)
		Txt = Replace(Txt,"ı" ,"\u0131" ,1,-1,1)
		Txt = Replace(Txt,"İ" ,"\u0130" ,1,-1,1)
		Txt = Replace(Txt,"ö" ,"\u00F6" ,1,-1,1)
		Txt = Replace(Txt,"Ö" ,"\u00D6" ,1,-1,1)
		Txt = Replace(Txt,"ü" ,"\u00FC" ,1,-1,1)
		Txt = Replace(Txt,"Ü" ,"\u00DC" ,1,-1,1)
		Txt = Replace(Txt,"ş" ,"\u015F" ,1,-1,1)
		Txt = Replace(Txt,"Ş" ,"\u015E" ,1,-1,1)
		Txt = Replace(Txt,"ç" ,"\u00E7" ,1,-1,1)
		Txt = Replace(Txt,"Ç" ,"\u00C7" ,1,-1,1)
		TurkceKarakter2HTML = Txt  
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
End Class
%>
