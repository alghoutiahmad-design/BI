
# Faster Cash - Petty Cash Analyzer
$keyFile = "C:\Users\VIPCAR\Desktop\Claude\fluid-stratum-494915-u3-ae0c3730efa9.json"
$sheetId = "1buGXjNLN4aEW27YkOzRjGAHNberKTcp5xyID0sk7ckg"
$openssl = "C:\Program Files\Git\usr\bin\openssl.exe"
$outFile = "$PSScriptRoot\report.html"

Write-Host "Fetching data..." -ForegroundColor Cyan

# JWT
$creds   = Get-Content $keyFile -Raw | ConvertFrom-Json
$privKey = $creds.private_key -replace '\\n', "`n"
$tmpKey  = "$env:TEMP\sa_key.pem"
[System.IO.File]::WriteAllText($tmpKey, $privKey, [System.Text.Encoding]::ASCII)

function B64Url($data) {
    if ($data -is [string]) { $data = [System.Text.Encoding]::UTF8.GetBytes($data) }
    [Convert]::ToBase64String($data).TrimEnd('=').Replace('+','-').Replace('/','_')
}

$email   = $creds.client_email
$now     = [int][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$header  = B64Url '{"alg":"RS256","typ":"JWT"}'
$payload = B64Url "{`"iss`":`"$email`",`"scope`":`"https://www.googleapis.com/auth/spreadsheets.readonly`",`"aud`":`"https://oauth2.googleapis.com/token`",`"exp`":$($now+3600),`"iat`":$now}"
$unsigned = "$header.$payload"
$msgFile = "$env:TEMP\jwt_msg.txt"
$sigFile = "$env:TEMP\jwt_sig.bin"
[System.IO.File]::WriteAllText($msgFile, $unsigned, [System.Text.Encoding]::ASCII)
& $openssl dgst -sha256 -sign $tmpKey -out $sigFile $msgFile 2>&1 | Out-Null
$sig = B64Url ([System.IO.File]::ReadAllBytes($sigFile))
$jwt = "$unsigned.$sig"

$tokenResp = Invoke-RestMethod -Method Post -Uri "https://oauth2.googleapis.com/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{ grant_type="urn:ietf:params:oauth:grant-type:jwt-bearer"; assertion=$jwt }
$token = $tokenResp.access_token
Remove-Item $tmpKey,$msgFile,$sigFile -ErrorAction SilentlyContinue

# Fetch sheet
$enc  = [Uri]::EscapeDataString("Petty Cash Report!A1:G1000")
$resp = Invoke-RestMethod -Uri "https://sheets.googleapis.com/v4/spreadsheets/$sheetId/values/$enc" `
    -Headers @{ Authorization="Bearer $token" }
$rows = $resp.values
Write-Host "Fetched $($rows.Count) rows" -ForegroundColor Green

function ParseNum($s) {
    if ([string]::IsNullOrWhiteSpace($s) -or $s -eq '0' -or $s -eq '0.00') { return 0.0 }
    $v = 0.0
    if ([double]::TryParse(($s -replace ',','').Trim(), [ref]$v)) { return $v }
    return 0.0
}

# Data structure using English keys
$months = @('03','04','05')
$data = @{}
foreach ($m in $months) {
    $data[$m] = @{
        refill    = 0.0
        expenses  = 0.0
        c_salary  = 0.0
        c_fuel    = 0.0
        c_reg     = 0.0
        c_maint   = 0.0
        c_carbuy  = 0.0
        c_park    = 0.0
        c_ds      = 0.0
        c_refund  = 0.0
        c_office  = 0.0
        employees = @{}
        refills   = [System.Collections.ArrayList]@()
    }
}

$EMPS = @('WALEED','ABDULLAH','DIYAB','ANFAL','HUMAM','EMRAN','ADHAM',
          'MARUF','KABIR','MOHAMMAD','KHERAD','FAISAL','RONALD','SHYNOOJ',
          'JOSEPH','SAEED','ROSE','ZAKARIA','PETTER','ISMAEL')

foreach ($row in $rows[1..($rows.Count-1)]) {
    if ($row.Count -lt 5) { continue }
    $dateStr = $row[0]
    if ($dateStr -notmatch '/(\d{2})/2026') { continue }
    $m = $Matches[1]
    if ($m -notin $months) { continue }

    $notes  = if ($row.Count -gt 2) { $row[2] } else { "" }
    $dStr   = if ($row.Count -gt 3) { $row[3] } else { "0" }
    $cStr   = if ($row.Count -gt 4) { $row[4] } else { "0" }
    $debit  = ParseNum $dStr
    $credit = ParseNum $cStr
    $n      = $notes.ToUpper()

    if ($debit -gt 0) {
        $data[$m].refill += $debit
        if ($n -match 'RECEIVED|PETTY CASH FROM') {
            $null = $data[$m].refills.Add(@{ date=$dateStr; amount=$debit; notes=$notes })
        }
    }

    if ($credit -gt 0) {
        $data[$m].expenses += $credit

        if     ($n -match 'SALARY|ADVANCE|BONUS|PAYROLL|HOUSING') {
            $data[$m].c_salary += $credit
            $emp = 'Other'
            foreach ($e in $EMPS) { if ($n -match $e) { $emp = (Get-Culture).TextInfo.ToTitleCase($e.ToLower()); break } }
            if (-not $data[$m].employees.ContainsKey($emp)) { $data[$m].employees[$emp] = 0.0 }
            $data[$m].employees[$emp] += $credit
        } elseif ($n -match 'REFUEL|FUEL|PETROL') {
            $data[$m].c_fuel += $credit
        } elseif ($n -match 'REGISTR|INSUR|TRIPTICKET|TOURISM CERT|TASJEEL|PASSING TEST|EXPORT TEST|FAIL TEST|REG TEST|LIFE SPAN|RENEWAL') {
            $data[$m].c_reg += $credit
        } elseif ($n -match 'REPAIR|FIXING|MAINTEN|CHANGE DASH|PARTS|LABOUR|MASTERCLASS|TYRE|TIRE') {
            $data[$m].c_maint += $credit
        } elseif ($n -match 'PAYMENT FOR BUYING|BUYING NEW') {
            $data[$m].c_carbuy += $credit
        } elseif ($n -match 'PARKING|PARK RENTAL|ZAHART|LABAN|SKYLARK') {
            $data[$m].c_park += $credit
        } elseif ($n -match 'DRIVING SERV|TAXI|IDL DELIVERY|RECOVERY|PICK UP') {
            $data[$m].c_ds += $credit
        } elseif ($n -match 'REFUND|BALANCE OF CUSTOMER|TRANSFER|DEPOSIT MAIN|DEPOSIT IN BACK') {
            $data[$m].c_refund += $credit
        } else {
            $data[$m].c_office += $credit
        }
    }
}

# May projection
$today      = Get-Date
$daysInMay  = [DateTime]::DaysInMonth($today.Year, 5)
$mayPct     = if ($today.Month -eq 5) { [math]::Round($today.Day / $daysInMay * 100) } else { 100 }
$mayProj    = if ($today.Month -eq 5 -and $today.Day -gt 0) {
    [math]::Round($data['05'].expenses / $today.Day * $daysInMay)
} else { [math]::Round($data['05'].expenses) }

# Build JSON
function RoundVal($v) { [math]::Round($v, 2) }

$jsonObj = [ordered]@{
    generated = (Get-Date).ToString("dd/MM/yyyy HH:mm")
    months = [ordered]@{
        "03" = [ordered]@{
            name="March"; refill=RoundVal($data['03'].refill); expenses=RoundVal($data['03'].expenses)
            completion=100; projected=RoundVal($data['03'].expenses)
            cats=[ordered]@{ salary=RoundVal($data['03'].c_salary); fuel=RoundVal($data['03'].c_fuel); reg=RoundVal($data['03'].c_reg); maint=RoundVal($data['03'].c_maint); carbuy=RoundVal($data['03'].c_carbuy); park=RoundVal($data['03'].c_park); ds=RoundVal($data['03'].c_ds); refund=RoundVal($data['03'].c_refund); office=RoundVal($data['03'].c_office) }
            employees=$data['03'].employees
            refills=$data['03'].refills
        }
        "04" = [ordered]@{
            name="April"; refill=RoundVal($data['04'].refill); expenses=RoundVal($data['04'].expenses)
            completion=100; projected=RoundVal($data['04'].expenses)
            cats=[ordered]@{ salary=RoundVal($data['04'].c_salary); fuel=RoundVal($data['04'].c_fuel); reg=RoundVal($data['04'].c_reg); maint=RoundVal($data['04'].c_maint); carbuy=RoundVal($data['04'].c_carbuy); park=RoundVal($data['04'].c_park); ds=RoundVal($data['04'].c_ds); refund=RoundVal($data['04'].c_refund); office=RoundVal($data['04'].c_office) }
            employees=$data['04'].employees
            refills=$data['04'].refills
        }
        "05" = [ordered]@{
            name="May"; refill=RoundVal($data['05'].refill); expenses=RoundVal($data['05'].expenses)
            completion=$mayPct; projected=$mayProj
            cats=[ordered]@{ salary=RoundVal($data['05'].c_salary); fuel=RoundVal($data['05'].c_fuel); reg=RoundVal($data['05'].c_reg); maint=RoundVal($data['05'].c_maint); carbuy=RoundVal($data['05'].c_carbuy); park=RoundVal($data['05'].c_park); ds=RoundVal($data['05'].c_ds); refund=RoundVal($data['05'].c_refund); office=RoundVal($data['05'].c_office) }
            employees=$data['05'].employees
            refills=$data['05'].refills
        }
    }
}

$jsonStr = $jsonObj | ConvertTo-Json -Depth 10 -Compress
Write-Host "Generating HTML..." -ForegroundColor Cyan

# HTML template (single-quoted to avoid $ expansion)
$htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Faster Cash - Petty Cash Analysis</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  * { font-family: 'Inter', sans-serif; }
  .card { background:white; border-radius:14px; box-shadow:0 2px 16px rgba(0,0,0,0.07); padding:1.5rem; }
  .badge { display:inline-flex; align-items:center; gap:4px; padding:3px 10px; border-radius:999px; font-size:0.72rem; font-weight:700; }
  ::-webkit-scrollbar { width:6px; height:6px; }
  ::-webkit-scrollbar-track { background:#f1f5f9; }
  ::-webkit-scrollbar-thumb { background:#cbd5e1; border-radius:3px; }
</style>
</head>
<body class="bg-slate-50 min-h-screen">

<div class="bg-gradient-to-r from-blue-900 to-blue-700 text-white px-6 py-5 shadow-lg">
  <div class="max-w-7xl mx-auto flex flex-wrap justify-between items-center gap-3">
    <div>
      <h1 class="text-2xl font-extrabold tracking-tight">Faster Cash &mdash; Petty Cash Analysis</h1>
      <p class="text-blue-200 text-sm mt-0.5">March &middot; April &middot; May 2026</p>
    </div>
    <div class="text-xs text-blue-300" id="gen-time"></div>
  </div>
</div>

<div class="max-w-7xl mx-auto px-4 py-6 space-y-6">

  <!-- Summary Cards -->
  <div class="grid grid-cols-1 md:grid-cols-3 gap-4" id="cards"></div>

  <!-- Charts -->
  <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
    <div class="card">
      <h3 class="font-semibold text-gray-700 mb-4">Expenses by Category</h3>
      <canvas id="catChart"></canvas>
    </div>
    <div class="card">
      <h3 class="font-semibold text-gray-700 mb-4">Monthly Refill vs Expenses</h3>
      <canvas id="monthChart"></canvas>
    </div>
  </div>

  <!-- Employee Advances -->
  <div class="card">
    <h3 class="font-semibold text-gray-700 mb-4">Salary &amp; Advances by Employee</h3>
    <div class="overflow-x-auto">
      <table class="w-full text-sm">
        <thead>
          <tr class="text-gray-400 border-b border-gray-100">
            <th class="text-left pb-3 pl-2 font-semibold">Employee</th>
            <th class="text-center pb-3 font-semibold text-blue-600">March</th>
            <th class="text-center pb-3 font-semibold text-emerald-600">April</th>
            <th class="text-center pb-3 font-semibold text-amber-600">May</th>
            <th class="text-center pb-3 font-semibold text-gray-700">Total</th>
          </tr>
        </thead>
        <tbody id="empTable"></tbody>
      </table>
    </div>
  </div>

  <!-- Refill History -->
  <div class="card">
    <h3 class="font-semibold text-gray-700 mb-4">EIB Refill History</h3>
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6" id="refillSection"></div>
  </div>

</div>

<script>
const RAW = __DATA_PLACEHOLDER__;
const M = RAW.months;
const KEYS = ['03','04','05'];
const NAMES = {'03':'March','04':'April','05':'May'};
const CAT_NAMES = {
  salary:'Salary & Advances',
  fuel:'Fuel',
  reg:'Registration & Insurance',
  maint:'Repairs & Maintenance',
  carbuy:'Car Purchase',
  park:'Parking',
  ds:'Driving Service',
  refund:'Refunds & Transfers',
  office:'Office & Misc'
};
const COLORS = {
  '03':{main:'#3b82f6',light:'#eff6ff',text:'#1d4ed8',alpha:'rgba(59,130,246,0.75)'},
  '04':{main:'#10b981',light:'#ecfdf5',text:'#065f46',alpha:'rgba(16,185,129,0.75)'},
  '05':{main:'#f59e0b',light:'#fffbeb',text:'#92400e',alpha:'rgba(245,158,11,0.75)'},
};

const fmt = n => new Intl.NumberFormat('en-US',{maximumFractionDigits:0}).format(Math.round(n));
const aed = n => fmt(n) + ' AED';

document.getElementById('gen-time').textContent = 'Last updated: ' + RAW.generated;

// Cards
const cardsEl = document.getElementById('cards');
KEYS.forEach(k => {
  const d = M[k]; const c = COLORS[k];
  const net = d.refill - d.expenses;
  const incomplete = d.completion < 100;
  cardsEl.innerHTML += `
  <div class="card border-t-4" style="border-color:${c.main}">
    <div class="flex justify-between items-start mb-4">
      <div>
        <p class="text-xs text-gray-400 uppercase tracking-widest font-medium">2026</p>
        <h2 class="text-2xl font-extrabold" style="color:${c.main}">${NAMES[k]}</h2>
      </div>
      ${incomplete
        ? `<span class="badge" style="background:${c.light};color:${c.text}">&#9203; ${d.completion}% complete</span>`
        : '<span class="badge bg-green-50 text-green-700">&#10003; Complete</span>'}
    </div>
    <div class="space-y-3">
      <div class="flex justify-between items-center">
        <span class="text-gray-400 text-sm">Total Refill</span>
        <span class="font-bold text-green-600">${aed(d.refill)}</span>
      </div>
      <div class="flex justify-between items-center">
        <span class="text-gray-400 text-sm">Total Expenses</span>
        <span class="font-bold text-red-500">${aed(d.expenses)}</span>
      </div>
      <div class="h-px bg-gray-100"></div>
      <div class="flex justify-between items-center">
        <span class="text-gray-600 text-sm font-medium">Net</span>
        <span class="font-extrabold ${net>=0?'text-green-600':'text-red-500'}">${aed(net)}</span>
      </div>
      ${incomplete ? `
      <div class="rounded-lg p-2.5 text-xs" style="background:${c.light};color:${c.text}">
        &#128202; Projected month-end: <strong>${aed(d.projected)}</strong>
      </div>` : ''}
    </div>
  </div>`;
});

// Category Chart
const catKeys = Object.keys(CAT_NAMES);
new Chart(document.getElementById('catChart'), {
  type: 'bar',
  data: {
    labels: catKeys.map(k => CAT_NAMES[k]),
    datasets: KEYS.map(k => ({
      label: NAMES[k],
      data: catKeys.map(c => M[k].cats[c] || 0),
      backgroundColor: COLORS[k].alpha,
      borderColor: COLORS[k].main,
      borderWidth: 1,
      borderRadius: 5
    }))
  },
  options: {
    responsive: true,
    plugins: { legend: { position:'top', labels:{ font:{family:'Inter',size:12} } } },
    scales: {
      x: { ticks:{ font:{family:'Inter',size:10}, maxRotation:30 } },
      y: { ticks:{ callback: v => fmt(v), font:{size:10} } }
    }
  }
});

// Monthly Chart
new Chart(document.getElementById('monthChart'), {
  type: 'bar',
  data: {
    labels: KEYS.map(k => NAMES[k]),
    datasets: [
      { label:'Refill',    data:KEYS.map(k=>M[k].refill),   backgroundColor:'rgba(16,185,129,0.75)', borderRadius:6 },
      { label:'Expenses',  data:KEYS.map(k=>M[k].expenses), backgroundColor:'rgba(239,68,68,0.75)',  borderRadius:6 }
    ]
  },
  options: {
    responsive: true,
    plugins: { legend: { position:'top', labels:{ font:{family:'Inter',size:12} } } },
    scales: { y: { ticks:{ callback: v => fmt(v), font:{size:10} } } }
  }
});

// Employee Table
const allEmps = [...new Set(KEYS.flatMap(k => Object.keys(M[k].employees||{})))];
allEmps.sort((a,b) => {
  const ta = KEYS.reduce((s,k) => s+(M[k].employees[a]||0), 0);
  const tb = KEYS.reduce((s,k) => s+(M[k].employees[b]||0), 0);
  return tb - ta;
});
const tbody = document.getElementById('empTable');
const tots = {'03':0,'04':0,'05':0};
allEmps.forEach((emp, i) => {
  const vals  = KEYS.map(k => M[k].employees[emp] || 0);
  const total = vals.reduce((a,b) => a+b, 0);
  KEYS.forEach((k,i) => tots[k] += vals[i]);
  tbody.innerHTML += `
  <tr class="${i%2===0 ? 'bg-gray-50/60' : ''}">
    <td class="py-2.5 pl-2 font-medium text-gray-700">${emp}</td>
    ${vals.map((v,i) => `<td class="py-2.5 text-center ${['text-blue-700','text-emerald-700','text-amber-700'][i]}">${v>0 ? fmt(v) : '&mdash;'}</td>`).join('')}
    <td class="py-2.5 text-center font-bold text-gray-800">${fmt(total)}</td>
  </tr>`;
});
const grand = Object.values(tots).reduce((a,b) => a+b, 0);
tbody.innerHTML += `
<tr class="border-t-2 border-gray-200 font-bold bg-slate-100">
  <td class="py-3 pl-2 text-gray-700">Total</td>
  <td class="py-3 text-center text-blue-700">${fmt(tots['03'])}</td>
  <td class="py-3 text-center text-emerald-700">${fmt(tots['04'])}</td>
  <td class="py-3 text-center text-amber-700">${fmt(tots['05'])}</td>
  <td class="py-3 text-center text-red-600 text-base">${fmt(grand)}</td>
</tr>`;

// Refill History
const refillEl = document.getElementById('refillSection');
KEYS.forEach(k => {
  const c = COLORS[k];
  const refills = M[k].refills || [];
  const total = refills.reduce((s,r) => s+r.amount, 0);
  refillEl.innerHTML += `
  <div>
    <div class="flex justify-between items-center mb-3">
      <h4 class="font-semibold" style="color:${c.text}">${NAMES[k]}</h4>
      <span class="badge" style="background:${c.light};color:${c.text}">${refills.length} transfers &middot; ${fmt(total)} AED</span>
    </div>
    <div class="space-y-2">
      ${refills.length === 0
        ? '<p class="text-gray-400 text-sm">No data</p>'
        : refills.map(r => `
      <div class="flex justify-between items-center px-3 py-2 rounded-lg text-sm" style="background:${c.light}">
        <span class="text-gray-500">${r.date}</span>
        <span class="font-bold" style="color:${c.text}">${fmt(r.amount)} AED</span>
      </div>`).join('')}
    </div>
  </div>`;
});
</script>
</body>
</html>
'@

$finalHtml = $htmlTemplate.Replace('__DATA_PLACEHOLDER__', $jsonStr)
[System.IO.File]::WriteAllText($outFile, $finalHtml, [System.Text.Encoding]::UTF8)
Write-Host "Report generated: $outFile" -ForegroundColor Green
Start-Process $outFile
Write-Host "Opened in browser!" -ForegroundColor Green
