InvestCard("變更投資組合", "目標型投組", 1)


' ====================================================================================================
' Purpose: 依需求搜索符合條件的投資組合
' Input: Purpose => 執行目的 ("變更投資組合" Or "提高定期定額投資金額")
'        InvestPortfolio => 投資卡片的需求或條件 ("目標型頭組" Or "退休型投組" Or "策略型投組")
'        CardStartIndex => 從第幾個頭組卡片開始 (若無特定輸入1即可)
' Return: N/A
' Example: InvestCard "提高定期定額投資金額", "目標型頭組", 1
' Creator: 宇森(Yusen) 2023/10
' Chang History: 
' ====================================================================================================
Function InvestCard(Purpose, InvestPortfolio, CardStartIndex)
	Select Case Purpose '<目的>
		Case "變更投資組合", "變更投組"
			' ____________________________________________________________________________________________________
			Select Case InvestPortfolio '<投組卡片條件>
				Case "目標型投組"
					ChangePortfolioType CardStartIndex, "智動ＧＯ目標" '<參數CardStartIndex為卡片起始值，若投組名稱有變更改String即可>

				Case "退休型投組"
					ChangePortfolioType CardStartIndex, "智動ＧＯ退休"

				Case "策略型投組"
					' Print("策略型投組")

				Case Else '<沒有特定>
					
			End Select	
		
			' ____________________________________________________________________________________________________
		Case "提高定期定額投資金額", "提高定期調額"
			Select Case InvestPortfolio '<投組卡片條件>
				Case "目標型投組"
					ChangePortfolioType CardStartIndex, "智動ＧＯ目標" '<參數CardStartIndex為卡片起始值，若投組名稱有變更改String即可>

				Case "退休型投組"
					ChangePortfolioType CardStartIndex, "智動ＧＯ退休"

				Case "策略型投組"
					' Print("策略型投組")

				Case Else '<沒有特定>
	End Select

End Function ' InvestCard



' ====================================================================================================
' Purpose: 依Call Action的執行目的的參數Parameter("Purpose")和投資組合的名稱尋找符合需求的投資組合
' Input: CardStartIndex => 投資卡片起始位置(Array Index)
'        InvestPortfolioName => 投資組合名稱(String)
' Return: N/A
' Example: ChangePortfolioType(1, "智動ＧＯ目標")
' Creator: 宇森(Yusen) 2023/10
' Chang History: 
' ====================================================================================================
Function ChangePortfolioType(CardStartIndex, InvestPortfolioName)
	Addstep "點選"& Parameter("InvestPortfolio") &"卡片 > 進入【投資組合】頁面"
	' ____________________________________________________________________________________________________
	' >>> 定義整張投資卡片物件(無法點選) <<<
	Set objInvestProtfolioCardDese = Description.Create()
		objInvestProtfolioCardDese("html tag").Value = "DIV"
		objInvestProtfolioCardDese("visible").Value = True
		objInvestProtfolioCardDese("outertext").RegularExpression = True
		objInvestProtfolioCardDese("outertext").Value = ".*"
		objInvestProtfolioCardDese("class").RegularExpression = True
		objInvestProtfolioCardDese("class").Value = ".*vi-card overflow-hidden.*"
	' ____________________________________________________________________________________________________
	' >>> 定義投資卡片物件(可點選) <<<
	Set LinkDescription = Description.Create()
		LinkDescription("html tag").Value = "DIV"
		LinkDescription("visible").Value = True
		LinkDescription("outertext").RegularExpression = True
		LinkDescription("outertext").Value = ".*"
		LinkDescription("class").RegularExpression = True
		LinkDescription("class").Value = ".*vi-select-link.*"
	Set objLinkDescription = Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資卡片").ChildObjects(LinkDescription)
	Set InvestPortfolioCard = Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資卡片").ChildObjects(objInvestProtfolioCardDese)
	InvestPortfolioCount = InvestPortfolioCard.Count -1
	InvestPortfolioCardMatching = False '<找到符合條件卡片改為 => True ，若未找到默認為 => False 退出腳本>
	' ____________________________________________________________________________________________________
	Select Case Parameter("Purpose")
		Case "變更投資組合", "變更投組"
			For i = CardStartIndex To InvestPortfolioCount '<檢查投組卡片是否符合條件，並點選進入>
				Result(0) = i : Result(1) = InvestPortfolioCount '<用於後續檢核卡片不符條件，Set For迴圈新卡片起始位>
				item = InvestPortfolioCard.GetROProperty("outertext")
				If InStr(item, InvestPortfolioName) > 0 And Not InStr(item, "全部贖回中") > 0 And Not InStr(item, "部分贖回中") > 0 And Not InStr(item, "智能調整中") > 0 And Not InStr(item, "尚未扣款") > 0 And Not InStr(item, "投組變更中") > 0 Then
					InvestPortfolioCardMatching = True '<初步找到符合條件卡片改為 => True ，若未找到默認為False 後續檢核退出腳本>
					Set objLinkDescription = Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資卡片").ChildObjects(LinkDescription)
					objLinkDescription(i-1).Click
					SyncLoading()
				End If
			Next
			' ____________________________________________________________________________________________________
			If InvestPortfolioCardMatching = False Then
				AddCheckFail "應點選進入【"& Left(Parameter("InvestPortfolio"), 3) &"投資組合卡片】", "成功匹配進入【"& Left(Parameter("InvestPortfolio"), 3) &"投資組合卡片】", "未找到符合條件投組卡片，請更換測資"
				ExitTest
			ElseIf CheckInvestPortfolioStastus = True Then
				Exit Function
			Else
				If Resule(0) < Result(1) Then
					InvestCard Parameter("Purpose"), Parameter("InvestPortfolio"), Result(0)+1
				End If
			End If
			' ____________________________________________________________________________________________________
		Case "提高定期定額投資金額", "提高定期定額"
			For i = CardStartIndex To InvestPortfolioCount '<檢查投組卡片是否符合條件，並點選進入>
				Result(0) = i : Result(1) = InvestPortfolioCount '<用於後續檢核卡片不符條件，Set For迴圈新卡片起始位>
				item = InvestPortfolioCard.GetROProperty("outertext")
				If InStr(item, InvestPortfolioName) > 0 And Not InStr(item, "全部贖回中") > 0 And Not InStr(item, "扣停中") > 0 And Not InStr(item, "尚未扣款") > 0 And Not InStr(item, "投組變更中") > 0 Then
					InvestPortfolioCardMatching = True '<初步找到符合條件卡片改為 => True ，若未找到默認為False 後續檢核退出腳本>
					Set objLinkDescription = Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資卡片").ChildObjects(LinkDescription)
					objLinkDescription(i-1).Click
					SyncLoading()
				End If
			Next
			' ____________________________________________________________________________________________________
			If InvestPortfolioCardMatching = False Then
				AddCheckFail "應點選進入【"& Left(Parameter("InvestPortfolio"), 3) &"投資組合卡片】", "成功匹配進入【"& Left(Parameter("InvestPortfolio"), 3) &"投資組合卡片】", "未找到符合條件投組卡片，請更換測資"
				ExitTest
			End If
			' ____________________________________________________________________________________________________
	End Select
End Function ' ChangePortfolioType



' ====================================================================================================
' Purpose: For變更投資組合用，找到符合條件卡片後，先檢查投組是否可以變更
' Input: N/A
' Return: Boolean
' Example: CheckInvestPortfolioStastus()
' Creator: 宇森(Yusen) 2023/10
' Chang History: 
' ====================================================================================================
Function CheckInvestPortfolioStastus()
	Addstep "執行變更投組前，檢查【投資組合】下拉選單狀態是否可異動"
	' ____________________________________________________________________________________________________
	Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("修改投資計畫[btn]").Click
	SyncLoading()
	' ____________________________________________________________________________________________________
	Status = Device("Device").App("Home Bank").MobileWebView("EBMW").Page("修改投資計畫").WebElement("投資組合[下拉選單]").GetROProperty("class")
	If InStr(Status, "disabled") Then
		CheckInvestPortfolioStastus = False
		Device("Device").App("Home Bank").MobileWebView("EBMW").Page("通用物件").WebElement("返回上一頁[左上角icon]").Click
		SyncLoading()
		Device("Device").App("Home Bank").MobileWebView("EBMW").Page("通用物件").WebElement("返回上一頁[左上角icon]").Click
		SyncLoading()
	Else
		Device("Device").App("Home Bank").MobileWebView("EBMW").Page("修改投資計畫").WebElement("投資組合[下拉選單]").Click
		Set objInvestPortfolioItemDesc = Description.Create()
			objInvestPortfolioItemDesc("html tag").Value = "LABLE"
			objInvestPortfolioItemDesc("visible").Value = True
			objInvestPortfolioItemDesc("outertext").RegularExpression = True
			objInvestPortfolioItemDesc("outertext").Value = ".*"
		Set InvestPortfolioItem = Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資卡片").ChildObjects(objInvestPortfolioItemDesc)
		InvestPortfolioItemCount = InvestPortfolioItem.Count 
		Device("Device").App("Home Bank").MobileWebView("EBMW").Page("通用物件").WebElement("X關閉").Click
		If InvestPortfolioItemCount > 1 Then
			CheckInvestPortfolioStastus = True
			Device("Device").App("Home Bank").MobileWebView("EBMW").Page("通用物件").WebElement("返回上一頁[左上角icon]").Click
			SyncLoading()
		Else 
			CheckInvestPortfolioStastus = False
		End If
	End If
End Function ' CheckInvestPortfolioStastus
