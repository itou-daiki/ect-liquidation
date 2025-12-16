document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const organizationInput = document.getElementById('organization');
    const positionInput = document.getElementById('position');
    const nameInput = document.getElementById('name');
    const highwayFromSelect = document.getElementById('highway_from');
    const highwayToSelect = document.getElementById('highway_to');
    const oneWayFeeInput = document.getElementById('one_way_fee');
    const csvFileInput = document.getElementById('csv_file');
    const generateButton = document.getElementById('generate_button');
    const statusMessagesDiv = document.getElementById('status_messages');
    const downloadAreaDiv = document.getElementById('download_area');
    const dataPreviewDiv = document.getElementById('data_preview');

    // --- Initial Setup ---

    function getHighwaySections() {
        // From app.py
        const sections = [
            "玖珠", "湯布院", "別府", "大分米良", "大分光吉", "大分宮河内",
            "大分松岡", "大分", "津久見", "津久見南", "佐伯", "佐伯堅田",
            "北浦", "蒲江", "日田", "竹田", "朝地", "宇佐", "院内", "安心院",
        ];
        return sections.sort();
    }

    function populateSelectOptions() {
        const sections = getHighwaySections();
        const oitaIndex = sections.indexOf("大分米良") > -1 ? sections.indexOf("大分米良") : 0;
        const hitaIndex = sections.indexOf("日田") > -1 ? sections.indexOf("日田") : 1;

        sections.forEach((section, index) => {
            const optionFrom = new Option(section, section);
            const optionTo = new Option(section, section);
            highwayFromSelect.add(optionFrom);
            highwayToSelect.add(optionTo);
        });

        highwayFromSelect.selectedIndex = oitaIndex;
        highwayToSelect.selectedIndex = hitaIndex;
    }
    
    populateSelectOptions();

    // --- Event Listeners ---
    generateButton.addEventListener('click', handleGeneration);

    // --- Helper Functions ---
    function logStatus(message, type = 'info') {
        const p = document.createElement('p');
        p.textContent = message;
        p.className = `status-${type}`;
        statusMessagesDiv.appendChild(p);
    }
    
    function clearLogs() {
        statusMessagesDiv.innerHTML = '';
        downloadAreaDiv.innerHTML = '';
        dataPreviewDiv.innerHTML = '';
    }

    function displayDataPreview(data, headers) {
        const previewData = data.slice(0, 10);
        let table = '<table><thead><tr>';
        headers.forEach(header => {
            table += `<th>${header}</th>`;
        });
        table += '</tr></thead><tbody>';
        previewData.forEach(row => {
            table += '<tr>';
            headers.forEach(header => {
                table += `<td>${row[header] || ''}</td>`;
            });
            table += '</tr>';
        });
        table += '</tbody></table>';
        dataPreviewDiv.innerHTML = table;
    }
    

    // --- Core Logic ---

    function handleGeneration() {
        clearLogs();
        const file = csvFileInput.files[0];
        if (!file) {
            logStatus('CSVファイルを選択してください。', 'error');
            return;
        }

        logStatus('CSVファイルの読み込みを開始します...', 'info');
        
        // Japanese CSVs often use Shift_JIS. FileReader can read with specified encoding.
        const reader = new FileReader();
        reader.readAsText(file, 'Shift_JIS'); 

        reader.onload = (event) => {
            logStatus('CSVファイルの解析を開始します...', 'info');
            const csvText = event.target.result;

            Papa.parse(csvText, {
                header: true,
                skipEmptyLines: true,
                complete: (results) => {
                    if (results.errors.length > 0) {
                        logStatus(`CSV解析エラー: ${results.errors[0].message}`, 'error');
                        console.error(results.errors);
                        return;
                    }
                    logStatus('CSVの解析が完了しました。', 'success');
                    displayDataPreview(results.data, results.meta.fields);
                    processAndGenerateExcel(results.data);
                }
            });
        };

        reader.onerror = () => {
            logStatus('ファイルの読み込みに失敗しました。エンコーディングがShift_JISでない可能性があります。', 'error');
        };
    }
    
    function extractYearMonth(data) {
        let latestYear = 0, latestMonth = 0;
        const dateColumn = '利用年月日（自）';

        for (const row of data) {
            const dateStr = row[dateColumn];
            if (dateStr && typeof dateStr === 'string' && dateStr.includes('/')) {
                const parts = dateStr.split('/');
                if (parts.length >= 2) {
                    try {
                        let year = parseInt(parts[0], 10);
                        let month = parseInt(parts[1], 10);
                        
                        if (year < 50) year += 2000;
                        else if (year < 100) year += 1900;
                        
                        if (year > latestYear || (year === latestYear && month > latestMonth)) {
                            latestYear = year;
                            latestMonth = month;
                        }
                    } catch(e) {
                        continue;
                    }
                }
            }
        }
        return { year: latestYear, month: latestMonth };
    }

    function calculateDailyUsage(df, targetDate) {
        let morningAmount = 0;
        let afternoonAmount = 0;
        let morningConfirmed = null;
        let afternoonConfirmed = null;

        const dateColumn = '利用年月日（自）';
        const timeColumn = '時分（自）';
        const feeColumn = '後納料金';

        const targetY = targetDate.getFullYear();
        const targetM = targetDate.getMonth() + 1;
        const targetD = targetDate.getDate();

        for (const row of df) {
            const dateStr = row[dateColumn];
            if (!dateStr || typeof dateStr !== 'string') continue;
            
            let dateMatch = false;
            const parts = dateStr.split('/');
            if (parts.length >= 3) {
                try {
                    let csvYear = parseInt(parts[0], 10);
                    const csvMonth = parseInt(parts[1], 10);
                    const csvDay = parseInt(parts[2], 10);
                    
                    if (csvYear < 50) csvYear += 2000;
                    else if (csvYear < 100) csvYear += 1900;

                    if (csvYear === targetY && csvMonth === targetM && csvDay === targetD) {
                        dateMatch = true;
                    }
                } catch(e) {
                    continue;
                }
            }

            if (dateMatch) {
                const timeStr = row[timeColumn] || '';
                const amount = parseFloat(row[feeColumn]) || 0;
                let isMorning = true;
                if (timeStr.includes(':')) {
                    const hour = parseInt(timeStr.split(':')[0], 10);
                    isMorning = hour < 12;
                }
                
                if (isMorning) {
                    morningAmount += amount;
                    morningConfirmed = '○';
                } else {
                    afternoonAmount += amount;
                    afternoonConfirmed = '○';
                }
            }
        }

        return { morningAmount, afternoonAmount, morningConfirmed, afternoonConfirmed };
    }

    async function processAndGenerateExcel(data) {
        const { year, month } = extractYearMonth(data);
        if (!year || !month) {
            logStatus('データから有効な年月を抽出できませんでした。', 'error');
            return;
        }
        logStatus(`データ期間: ${year}年${month}月`, 'success');

        const oneWayFee = parseFloat(oneWayFeeInput.value);

        // --- Calculate usage for the whole month ---
        const lastDay = new Date(year, month, 0).getDate();
        const usageAmounts = {};
        for (let day = 1; day <= lastDay; day++) {
            const targetDate = new Date(year, month - 1, day);
            usageAmounts[day] = calculateDailyUsage(data, targetDate);
        }
        
        // --- Load Template and Generate Excel ---
        try {
            logStatus('Excelテンプレートを読み込んでいます...', 'info');
            const templatePath = 'templates/2025_04_高速道路等利用実績簿（テンプレート）.xlsx';
            const response = await fetch(templatePath);
            if (!response.ok) {
                throw new Error(`テンプレートファイルが見つかりません: ${response.statusText}`);
            }
            const arrayBuffer = await response.arrayBuffer();
            const wb = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, bookVBA: true });
            const ws = wb.Sheets[wb.SheetNames[0]];

            logStatus('Excelファイルにデータを書き込んでいます...', 'info');
            
            // --- Write data to worksheet ---
            ws['C3'] = { v: organizationInput.value };
            ws['K3'] = { v: positionInput.value };
            ws['N3'] = { v: nameInput.value };
            
            ws['B5'] = { v: year - 2018 }; // 令和年
            ws['D5'] = { v: month };
            
            ws['M5'] = { v: highwayFromSelect.value };
            ws['P5'] = { v: highwayToSelect.value };
            ws['M6'] = { v: oneWayFee };

            // For formulas if needed, but direct value writing is safer here.
            ws['E56'] = { t: 'd', v: new Date(Date.UTC(year, month - 1, 1)) };
            ws['E57'] = { t: 'd', v: new Date(Date.UTC(year, month - 1, lastDay)) };

            for (let day = 1; day <= lastDay; day++) {
                const dayData = usageAmounts[day];
                if (dayData.morningConfirmed || dayData.afternoonConfirmed) {
                    if (day <= 15) {
                        const row = day + 13;
                        if (dayData.morningConfirmed) {
                            ws[`D${row}`] = { v: dayData.morningConfirmed };
                            ws[`E${row}`] = { v: dayData.morningAmount };
                        }
                        if (dayData.afternoonConfirmed) {
                            ws[`G${row}`] = { v: dayData.afternoonConfirmed };
                            ws[`H${row}`] = { v: dayData.afternoonAmount };
                        }
                    } else { // 16-31
                        const row = (day <= 30) ? (day - 15 + 13) : 29; // 16->14 .. 31->29
                        if (dayData.morningConfirmed) {
                            ws[`L${row}`] = { v: dayData.morningConfirmed };
                            ws[`M${row}`] = { v: dayData.morningAmount };
                        }
                        if (dayData.afternoonConfirmed) {
                            ws[`O${row}`] = { v: dayData.afternoonConfirmed };
                            ws[`P${row}`] = { v: dayData.afternoonAmount };
                        }
                    }
                }
            }
            
            // --- Create downloadable file ---
            logStatus('Excelファイルを生成しています...', 'info');
            const outputWb = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([outputWb], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);

            const downloadLink = document.createElement('a');
            downloadLink.href = url;
            downloadLink.download = `${year}_${month}_高速道路利用実績簿（${nameInput.value}）.xlsx`;
            downloadLink.textContent = '📥 Excelファイルをダウンロード';
            downloadLink.className = 'download-button';
            downloadAreaDiv.appendChild(downloadLink);
            
            logStatus('利用実績簿が正常に生成されました！', 'success');
            logStatus('生成されたExcelファイルは必ず確認し、必要に応じて手動で調整してください。', 'warning');

        } catch (error) {
            console.error(error);
            logStatus(`エラーが発生しました: ${error.message}`, 'error');
        }
    }
});
