<?php
// 強制瀏覽器使用 UTF-8 顯示，避免看到亂碼
header('Content-Type: text/html; charset=utf-8');

// 1. 設定目標 URL
$url = "https://https://web.pcc.gov.tw/tps/atm/AtmAwardWithoutSso/QueryAtmAwardDetail?pkAtmMain=NzExNjgwMDM=";

// 2. 使用 cURL 抓取資料
$ch = curl_init();
curl_setopt($ch, CURLOPT_URL, $url);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false); // 略過 SSL 檢查
curl_setopt($ch, CURLOPT_USERAGENT, 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36');
$html = curl_exec($ch);
curl_close($ch);

if (!$html) {
    die("無法取得網頁內容，請檢查連線。");
}

// 3. 處理轉碼：政府網站大多是 Big5，轉換為 UTF-8 供 PHP 處理
$html = mb_convert_encoding($html, 'UTF-8', 'Big5');

// 4. 解析 HTML
$dom = new DOMDocument();
libxml_use_internal_errors(true);
// 這裡很關鍵：告訴 DOM 解析器這段 HTML 是 UTF-8
$dom->loadHTML('<?xml encoding="UTF-8">' . $html); 
$xpath = new DOMXPath($dom);

// 5. 使用 XPath 抓取指定路徑
// 這裡用「包含文字」的方式定位標籤，比寫死路徑更準確
$caseNoNode = $xpath->query("//th[contains(text(), '標案案號')]/following-sibling::td")->item(0);
$caseNameNode = $xpath->query("//th[contains(text(), '標案名稱')]/following-sibling::td")->item(0);

$caseNo = $caseNoNode ? trim($caseNoNode->nodeValue) : "抓不到資料";
$caseName = $caseNameNode ? trim($caseNameNode->nodeValue) : "抓不到資料";

// 6. 輸出結果
echo "<h2>抓取結果：</h2>";
echo "<b>標案案號：</b>" . htmlspecialchars($caseNo) . "<br>";
echo "<b>標案名稱：</b>" . htmlspecialchars($caseName) . "<br>";

// 如果你想看抓到哪一些 td，可以取消註解下面這行
// print_r($xpath->query("//table//td"));
?>