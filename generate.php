<?php
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

if ($_SERVER['REQUEST_METHOD'] === 'POST') {

    // 判斷使用者點擊的是哪一個按鈕
    $targetDoc = $_POST['target_doc'] ?? 'analysis';

    if ($targetDoc === 'review') {
        // 設定「檢視表」的範本名稱與下載檔名前綴
        $templateFile = __DIR__ . '/商情蒐集資料檢視表template.docx';
        $downloadPrefix = '商情蒐集資料檢視表_';
    } else {
        // 設定原本「分析表」的範本名稱
        $templateFile = __DIR__ . '/商情分析表template.docx';
        $downloadPrefix = '商情分析表_';
    }

    // 檢查檔案是否存在
    if (!file_exists($templateFile)) {
        die("錯誤：找不到模板檔案 " . basename($templateFile));
    }

    // 1. 接收「非陣列」的基本文字資料
    $projectCode = htmlspecialchars($_POST['project_code'] ?? '');
    $productName = htmlspecialchars($_POST['product_name'] ?? '');
    $fillDate = htmlspecialchars($_POST['fill_date'] ?? '');

    // --- 改良後的「二、商源資料」處理邏輯 ---
    $vNames = $_POST['v_name'] ?? [];
    $vAddresses = $_POST['v_address'] ?? [];
    $vContacts = $_POST['v_contact'] ?? [];
    $vPhones = $_POST['v_phone'] ?? [];

    $vendorStrings = [];
    for ($i = 0; $i < count($vNames); $i++) {
        // 只要有填廠商名稱就處理
        if (!empty(trim($vNames[$i]))) {
            $index = $i + 1;
            $name = htmlspecialchars($vNames[$i]);
            $addr = htmlspecialchars($vAddresses[$i] ?? '無');
            $contact = htmlspecialchars($vContacts[$i] ?? '無');
            $phone = htmlspecialchars($vPhones[$i] ?? '無');

            // 格式化字串
            $vendorStrings[] = "{$index}. {$name} (住址：{$addr}，聯絡人：{$contact}，電話：{$phone})";
        }
    }

    // 組合並轉為 Word 換行格式
    $sourceDataFormatted = implode("\n", $vendorStrings);
    $sourceData = str_replace("\n", '<w:br/>', $sourceDataFormatted);

    // 2. 接收多行文字資料 (加入 \n 轉 <w:br/> 處理)
    $purpose = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['purpose'] ?? ''));
    $sourceData = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['source_data'] ?? ''));
    $prevRecords = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['prev_records'] ?? ''));
    $marketSurvey = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['market_survey'] ?? ''));
    $conclusion = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['conclusion'] ?? ''));
    $otherMilitary = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['other_military_records'] ?? ''));
    $otherMarket = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['other_market_prices'] ?? ''));
    $discountRecords = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['discount_records'] ?? ''));
    $priceIndex = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['price_index'] ?? ''));
    $costAnalysis = str_replace("\n", '<w:br/>', htmlspecialchars($_POST['cost_analysis'] ?? ''));

    // 3. 建立暫存檔並處理 Zip 替換 (將 copy 移至此處)[cite: 1]
    $tempOutputFile = __DIR__ . '/temp_output_' . time() . '.docx';
    if (!copy($templateFile, $tempOutputFile)) die("錯誤：無法複製模板。");

    // 4. 處理 Zip 替換
    $zip = new ZipArchive;
    if ($zip->open($tempOutputFile) === TRUE) {
        $xmlFile = 'word/document.xml';
        $documentXml = $zip->getFromName($xmlFile);

        if ($documentXml === false) {
            $zip->close();
            die("錯誤：模板不是標準的 docx 格式。");
        }

        // 抓取前端傳來的陣列資料
        $units = $_POST['unit'] ?? [];
        $quantities = $_POST['quantity'] ?? [];
        $prevPrices = $_POST['prev_unit_price'] ?? [];
        $marketPrices = $_POST['market_lowest_price'] ?? [];
        $estPrices = $_POST['est_unit_price'] ?? [];

        $sumPrevTotal = 0;
        $sumMarketTotal = 0;
        $sumEstTotal = 0;
        $generatedRowsXml = '';

        // 尋找 Word 裡面包含 ${unit} 的那一列 (Row)
        $pos = strpos($documentXml, '${unit}');
        if ($pos !== false) {
            // 匹配 <w:tr> 或 <w:tr (帶屬性)，避免誤抓 <w:trPr> 導致 XML 斷裂
            $chunk = substr($documentXml, 0, $pos);
            $rowStart1 = strrpos($chunk, '<w:tr>');
            $rowStart2 = strrpos($chunk, '<w:tr ');
            $rowStart = max((int)$rowStart1, (int)$rowStart2);

            $rowEnd = strpos($documentXml, '</w:tr>', $pos) + 7; // 找列的結尾
            $templateRowXml = substr($documentXml, $rowStart, $rowEnd - $rowStart);

            // 移除 Word 自動生成的隱藏書籤，避免 ID 重複導致檔案損毀
            $templateRowXml = preg_replace('/<w:bookmarkStart[^>]*\/>/', '', $templateRowXml);
            $templateRowXml = preg_replace('/<w:bookmarkEnd[^>]*\/>/', '', $templateRowXml);

            // 根據前端傳來了幾個項次，就複製幾列
            for ($i = 0; $i < count($units); $i++) {
                $rowXml = $templateRowXml;

                // 數學計算
                $qty = is_numeric($quantities[$i]) ? floatval($quantities[$i]) : 1;
                $prevP = floatval($prevPrices[$i] ?? 0);
                $marketP = floatval($marketPrices[$i] ?? 0);
                $estP = floatval($estPrices[$i] ?? 0);

                $prevT = $prevP * $qty;
                $marketT = $marketP * $qty;
                $estT = $estP * $qty;

                $sumPrevTotal += $prevT;
                $sumMarketTotal += $marketT;
                $sumEstTotal += $estT;

                // 替換這一列裡面的標籤
                $rowXml = str_replace('${item_index}', ($i + 1), $rowXml);
                $rowXml = str_replace('${unit}', htmlspecialchars($units[$i] ?? '式'), $rowXml);
                $rowXml = str_replace('${quantity}', number_format($qty), $rowXml);

                $rowXml = str_replace('${prev_unit_price}', $prevP > 0 ? number_format($prevP) : '', $rowXml);
                $rowXml = str_replace('${market_lowest_price}', $marketP > 0 ? number_format($marketP) : '', $rowXml);
                $rowXml = str_replace('${est_unit_price}', $estP > 0 ? number_format($estP) : '', $rowXml);

                $rowXml = str_replace('${prev_total}', $prevT > 0 ? number_format($prevT) : '', $rowXml);
                $rowXml = str_replace('${market_total}', $marketT > 0 ? number_format($marketT) : '', $rowXml);
                $rowXml = str_replace('${est_total}', $estT > 0 ? number_format($estT) : '', $rowXml);

                $generatedRowsXml .= $rowXml; // 將產生好的列接起來
            }

            // 把原本單一的樣板列，替換成剛剛產生的一大串多筆資料列
            $documentXml = str_replace($templateRowXml, $generatedRowsXml, $documentXml);
        }

        // 接收前端獨立傳來的合計總價
        $manualSumPrev = $_POST['sum_prev_total_input'] ?? '';
        $manualSumMarket = $_POST['sum_market_total_input'] ?? '';
        $manualSumEst = $_POST['sum_est_total_input'] ?? '';

        // 替換 Word 最下方的總合計標籤 (自動加上千分位)
        $documentXml = str_replace('${sum_prev_total}', is_numeric($manualSumPrev) ? number_format((float)$manualSumPrev) : '', $documentXml);
        $documentXml = str_replace('${sum_market_total}', is_numeric($manualSumMarket) ? number_format((float)$manualSumMarket) : '', $documentXml);
        $documentXml = str_replace('${sum_est_total}', is_numeric($manualSumEst) ? number_format((float)$manualSumEst) : '', $documentXml);

        // 替換表頭基本標籤
        $documentXml = str_replace('${project_code}', $projectCode, $documentXml);
        $documentXml = str_replace('${product_name}', $productName, $documentXml);
        $documentXml = str_replace('${fill_date}', $fillDate, $documentXml);

        // 這就是你剛處理好的商源資料
        $documentXml = str_replace('${source_data}', $sourceData, $documentXml);

        // 替換文字區域標籤
        $documentXml = str_replace('${purpose}', $purpose, $documentXml);
        $documentXml = str_replace('${source_data}', $sourceData, $documentXml);
        $documentXml = str_replace('${prev_records}', $prevRecords, $documentXml);
        $documentXml = str_replace('${market_survey}', $marketSurvey, $documentXml);
        $documentXml = str_replace('${conclusion}', $conclusion, $documentXml);

        $documentXml = str_replace('${other_military_records}', $otherMilitary, $documentXml);
        $documentXml = str_replace('${other_market_prices}', $otherMarket, $documentXml);
        $documentXml = str_replace('${discount_records}', $discountRecords, $documentXml);
        $documentXml = str_replace('${price_index}', $priceIndex, $documentXml);
        $documentXml = str_replace('${cost_analysis}', $costAnalysis, $documentXml);

        $zip->addFromString($xmlFile, $documentXml);
        $zip->close();
    } else {
        die("錯誤：無法開啟暫存檔案。");
    }

    // 4. 下載設定：將檔名改為動態[cite: 1]
    while (ob_get_level()) ob_end_clean();
    $downloadFileName = $downloadPrefix . date('Ymd_His') . '.docx';

    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . rawurlencode($downloadFileName) . '"');

    readfile($tempOutputFile);
    unlink($tempOutputFile);
    exit;
}
