document.addEventListener('DOMContentLoaded', function() {
    // 設定今日日期為預設值
    const loginDateInput = document.getElementById('loginDate');
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    loginDateInput.value = now.toISOString().slice(0, 16);
    
    // 表單提交處理
    const form = document.getElementById('businessForm');
    const resultDiv = document.getElementById('result');
    
    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        // 清除先前的結果
        resultDiv.className = 'result hidden';
        resultDiv.innerHTML = '';
        
        // 收集表單資料
        const formData = {
            loginDate: document.getElementById('loginDate').value,
            contactPerson: getSelectedValues('contactPerson'),
            workType: document.getElementById('workType').value,
            builder: document.getElementById('builder').value,
            firePermit: document.getElementById('firePermit').value,
            address: document.getElementById('address').value,
            client: document.getElementById('client').value,
            buildingType: document.getElementById('buildingType').value,
            floors: document.getElementById('floors').value,
            totalArea: document.getElementById('totalArea').value,
            certifications: getSelectedValues('certifications'),
            notes: document.getElementById('notes').value
        };
        
        // 驗證必填欄位
        if (!validateForm(formData)) {
            return;
        }
        
        // 送出資料
        try {
            resultDiv.innerHTML = '<p>🔄 正在送出資料，請稍候...</p>';
            resultDiv.className = 'result';
            
            const response = await fetch(API_URL, {
                method: 'POST',
                mode: 'no-cors', // Google Apps Script Web App 需要 no-cors
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(formData)
            });
            
            // 因為使用 no-cors，無法直接檢查 response status
            // 預設假设成功
            setTimeout(() => {
                resultDiv.className = 'result success';
                resultDiv.innerHTML = `
                    <h3>✅ 表單提交成功！</h3>
                    <p>感謝您的填寫，資料已記錄。</p>
                    <button onclick="location.reload()" class="btn btn-secondary" style="margin-top:10px;">新增下一筆</button>
                `;
                form.reset();
            }, 1000);
            
        } catch (error) {
            resultDiv.className = 'result error';
            resultDiv.innerHTML = `
                <h3>❌ 提交失敗</h3>
                <p>錯誤訊息：${error.message}</p>
                <p>請檢查網際網路連線或稍後再試。</p>
            `;
        }
    });
    
    // 取得選中的 checkbox 值
    function getSelectedValues(name) {
        const checkboxes = document.querySelectorAll(`input[name="${name}"]:checked`);
        return Array.from(checkboxes).map(cb => cb.value).join(', ');
    }
    
    // 驗證必填欄位
    function validateForm(formData) {
        const requiredFields = [
            { name: 'loginDate', label: '登入日期' },
            { name: 'contactPerson', label: '承辦人' },
            { name: 'workType', label: '工作分類' },
            { name: 'builder', label: '起造人' },
            { name: 'address', label: '地址' },
            { name: 'client', label: '委託人' },
            { name: 'buildingType', label: '建築物用途分類' },
            { name: 'floors', label: '樓層概要' },
            { name: 'totalArea', label: '總樓地板面積' }
        ];
        
        let hasError = false;
        let errorMessage = '❌ 請填寫以下必填欄位：\n\n';
        
        for (const field of requiredFields) {
            if (!formData[field.name] || (field.name === 'contactPerson' && field.value === '')) {
                // 特殊處理 contactPerson，至少選一個
                if (field.name === 'contactPerson') {
                    const checkboxes = document.querySelectorAll('input[name="contactPerson"]:checked');
                    if (checkboxes.length === 0) {
                        hasError = true;
                        errorMessage += `- ${field.label}\n`;
                    }
                } else {
                    hasError = true;
                    errorMessage += `- ${field.label}\n`;
                }
            }
        }
        
        if (hasError) {
            resultDiv.className = 'result error';
            resultDiv.innerHTML = `<p>${errorMessage}</p>`;
            return false;
        }
        
        return true;
    }
});
