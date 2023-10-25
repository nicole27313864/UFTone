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
SyncLoading()
Dim objLinkDescription, InvestPortfolioCard, InvestPortfolioCardMatching, InvestPortfolio, NweInvestPortfolio, InvestPortfolioItem, CardStartIndex, Result(1), CardStartIndex, i, j
Dim Addstep, AddCheckPass, AddCheckFail, CheckDetails, wait,GetROProperty, SyncLoading 'UFT內建或額外定義

InvestCard Parameter("Purpose"), Parameter("InvestPortfolio"), 1    '改Call共用 Purpose, InvestPortfolio, CardStartIndex默認都先為 1


' ____________________________________________________________________________________________________
Addstep "檢核投資組合頁內【詳細資訊】展開/收合"
' __________檢核【詳細資訊】展開/收合
For i=0 To 2
	If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("詳細資訊_內容[展開/收合]").Exist(20) Then
		Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("詳細資訊_內容[展開/收合]").Click
    ElseIf Not Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("詳細資訊_內容[展開/收合]").Exist(20) Then
		If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("詳細資訊_內容[展開/收合]").Exist(5) Then
			AddCheckPass "投資組合頁內【詳細資訊】應可[展開/收合]", "【詳細資訊】成功展開收合"
		Else
			AddCheckFail "投資組合頁內【詳細資訊】應可[展開/收合]", "【詳細資訊】成功展開收合", "未能成功展開收合【詳細資訊】請檢查是否成功展開收合物件【詳細資訊[展開/收合]】是否異動。"
		End If
	End If
Next 'i

' __________檢核【詳細資訊】內容
Addstep "檢核投資組合頁內【詳細資訊】內容"
CheckDetails = Split(Trim(Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("詳細資訊_內容[展開/收合]").GetROProperty("outertext")), " ")
For i = 0 To UBound(CheckDetails)
	Select Case CheckDetails(i)
		Case "投資組合"
			AddCheckPass "應成功顯示【投資組合】", "成功顯示【投資組合: "& CheckDetails(i+1) &"】"
			InvestPortfolio = CheckDetails(i+1)

        Case "風險屬性"
			AddCheckPass "應成功顯示【風險屬性】", "成功顯示【風險屬性: "& CheckDetails(i+1) &"】"
            
		Case "持有資產"
			For j = i+1 To UBound(CheckDetails)
				Select Case CheckDetails(i)
					Case "參考市值"
						AddCheckPass "應成功顯示【參考市值】", "成功顯示【參考市值: "& CheckDetails(i+1) &"】"

					Case "原始投資金額"
						AddCheckPass "應成功顯示【原始投資金額】", "成功顯示【原始投資金額: "& CheckDetails(i+1) &"】"

					Case "參考損益"
						AddCheckPass "應成功顯示【參考損益】", "成功顯示【參考損益: "& CheckDetails(i+1) &"】"

					Case "參考報酬率"
						AddCheckPass "應成功顯示【參考報酬率】", "成功顯示【參考報酬率: "& CheckDetails(i+1) &"】"
						Exit For
				End Select
			Next 'j
	End Select
Next 'i

wait 1

' __________檢核【資產配置、投資績效、投資計畫、交易紀錄】Tab是否存在
Addstep "檢核投資組合頁內【資產配置、投資績效、投資計畫、交易紀錄】Tab"
If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("資產配置[Tab]").Exist(5) Then
	AddCheckPass "應成功顯示【資產配置】Tab", "成功顯示【資產配置】Tab"
Else
	AddCheckFail "應成功顯示【資產配置】Tab", "成功顯示【資產配置】Tab", "未找到【資產配置Tab】物件，請確認環境是否正常，或物件及規格異動。"
End If

If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("投資績效[Tab]").Exist(5) Then
	AddCheckPass "應成功顯示【投資績效】Tab", "成功顯示【投資績效】Tab"
Else
	AddCheckFail "應成功顯示【投資績效】Tab", "成功顯示【投資績效】Tab", "未找到【投資績效】物件，請確認環境是否正常，或物件及規格異動。"
End If

If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("投資計畫[Tab]").Exist(5) Then
	AddCheckPass "應成功顯示【投資計畫】Tab", "成功顯示【投資計畫】Tab"
Else
	AddCheckFail "應成功顯示【投資計畫】Tab", "成功顯示【投資計畫】Tab", "未找到【投資計畫Tab】物件，請確認環境是否正常，或物件及規格異動。"
End If

If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("交易紀錄[Tab]").Exist(5) Then
	AddCheckPass "應成功顯示【交易紀錄】Tab", "成功顯示【交易紀錄】Tab"
Else
	AddCheckFail "應成功顯示【交易紀錄】Tab", "成功顯示【交易紀錄】Tab", "未找到【交易紀錄Tab】物件，請確認環境是否正常，或物件及規格異動。"
End If

' __________防呆先回到【資產配置】Tab
Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("資產配置[Tab]").Click

' __________檢核【資產配置】下方顯示【資產配置圓餅圖】
Addstep "檢核投資組合頁內【資產配置】Tab下方顯示【資產配置圓餅圖】"
If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("資產配置圓餅圖").Exist(5) Then
	AddCheckPass "應成功顯示【資產配置圓餅圖】", "成功顯示【資產配置圓餅圖】"
Else
	AddCheckFail "應成功顯示【資產配置圓餅圖】", "成功顯示【資產配置圓餅圖】", "未找到【資產配置圓餅圖】物件，請確認環境是否正常，或物件及規格異動。"
End If

' __________檢核【單筆投資、定期定額投資/異動、贖回、修改投資計畫】按鈕
Addstep "檢核投資組合頁內【單筆投資、定期定額投資/異動、贖回、修改投資計畫】按鈕"
If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("單筆投資").Exist(5) Then
	AddCheckPass "應成功顯示【單筆投資】按鈕", "成功顯示【單筆投資】按鈕"
Else
	AddCheckFail "應成功顯示【單筆投資】按鈕", "成功顯示【單筆投資】按鈕", "未找到【單筆投資】按鈕物件，請確認環境是否正常，或物件及規格異動。"
End If

If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("定期定額投資/異動").Exist(5) Then
	AddCheckPass "應成功顯示【定期定額投資/異動】按鈕", "成功顯示【定期定額投資/異動】按鈕"
Else
	AddCheckFail "應成功顯示【定期定額投資", "成功顯示【定期定額投資", "未找到【定期定額投資/異動】按鈕物件，請確認環境是否正常，或物件及規格異動。"
End If

If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("贖回").Exist(5) Then
	AddCheckPass "應成功顯示【贖回】按鈕", "成功顯示【贖回】按鈕"
Else
	' AddCheckFail "應成功顯示【贖回】按鈕", "成功顯示【贖回】按鈕", "未找到【贖回】按鈕物件，請確認環境是否正常，或物件及規格異動。"
End If

If Device("Device").App("Home Bank").MobileWebView("EBMW").Page("投資組合頁").WebElement("修改投資計畫").Exist(5) Then
	AddCheckPass "應成功顯示【修改投資計畫】按鈕", "成功顯示【修改投資計畫】按鈕"
Else
	AddCheckFail "應成功顯示【修改投資計畫】按鈕", "成功顯示【修改投資計畫】按鈕", "未找到【修改投資計畫】按鈕物件，請確認環境是否正常，或物件及規格異動。"
End If


' ____________________________________________________________________________________________________





















































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

































