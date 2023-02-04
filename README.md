想在 Outlook 搜尋郵件、連絡人，需要一些背景知識，官方文章的 [Filtering Items ](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items?WT.mc_id=DOP-MVP-37580)是不錯的入門。
以下是我整理搜尋收件匣的原理：
* 先取得收件匣 Folder 物件，使用類似 SQL 的查詢語法對 Items 項目集合進行篩選，若筆數少可使用 [Find()](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.find?WT.mc_id=DOP-MVP-37580)/[Restrict()](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict?WT.mc_id=DOP-MVP-37580) 逐筆取回，Items.Restrict() 則可一次傳回符合條件的集合。
* 要找到特定信件，還有個笨方法是對 Items 跑迴圈一筆一筆撈出來讀屬性比對(若查詢邏輯很特殊，這是唯一解)，篩選功能可用類 SQL 語法查資料夾，更簡便且有效率。
* Find()/Restrict() 用的類 SQL 查詢語法有兩種格式：Jet Query Language 及 DAV Searching and Locating(DASL)，Jet 格式為 [Subject] = '...'、DASL 則為 @SQL=urn:schemas:httpmail:subject = '...'，兩種查詢都支援 AND/OR 及一些簡單運算，但二者不能混用，而且只有 DASL 才支援 LIKE 查詢，要做到關鍵字查詢，只能用 DASL。
* DASL 語法(@SQL=urn:schemas...) 會用到欄位對映 urn，可參考文件 [Exchage Store Schema](https://learn.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2007/aa581579(v=exchg.80)?WT.mc_id=DOP-MVP-37580) 取得。
* 找到 [MailItem 物件](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem?WT.mc_id=DOP-MVP-37580)後，可由 Subject、Body、Attachments 讀取主旨、內文及附件，也可呼叫 Delete()、Move()、Reply()、Forward() 進行刪除、搬移、回覆及轉寄等動作，做出各種花式應用。
