const TEMPLATE_SLIDES = [
  {
    id: 'template3',
    html: `<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>[プロジェクト名] | [サブタイトル]</title>
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            width: 100vw;
            height: 100vh;
            display: flex;
            flex-direction: column;
            background: white;
            color-scheme: light;
        }
        .slide-container {
            width: 100%;
            height: 100%;
            padding: 25px 40px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
        }
        .slide-title {
            font-size: 34px;
            font-weight: bold;
            color: #000000;
            margin-bottom: 5px;
        }
        .slide-subtitle {
            font-size: 18px;
            color: #333333;
            margin-bottom: 15px;
        }
        .roadmap-wrapper {
            flex: 1;
            display: flex;
            flex-direction: column;
        }
        .header-grid {
            display: grid;
            grid-template-columns: 120px repeat(10, 1fr);
            margin-bottom: 5px;
        }
        .month-label {
            text-align: center;
            font-weight: bold;
            padding: 3px;
            border-bottom: 2px solid #ddd;
            font-size: 14px;
        }
        .date-label {
            text-align: center;
            font-size: 11px;
            padding: 3px;
            color: #666;
        }
        .tasks-container {
            flex: 1;
        }
        .task-group {
            display: grid;
            grid-template-columns: 120px repeat(10, 1fr);
            margin-bottom: 3px;
            min-height: 22px;
            align-items: center;
        }
        .dept-label {
            padding: 0 10px;
            font-weight: bold;
            font-size: 12px;
        }
        .task-area {
            grid-column: 2 / -1;
            position: relative;
            height: 20px;
        }
        .task {
            position: absolute;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 10px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            height: 18px;
            line-height: 18px;
            text-align: center;
            border: 1px solid white;
        }
        .task-mail { background-color: #0052cc; }
        .task-cms { background-color: #00875a; }
        .task-code { background-color: #B22222; }
        .task-fusion { background-color: #6B46C1; }
        .task-hold { background-color: #d3d3d3; }
        .task-sub-mail {
            background-color: rgba(0, 82, 204, 0.2);
            color: #0052cc;
            border: 1px dashed #0052cc;
        }
        .task-sub-cms {
            background-color: rgba(0, 135, 90, 0.2);
            color: #00875a;
            border: 1px dashed #00875a;
        }
        .task-sub-code {
            background-color: rgba(178, 34, 34, 0.2);
            color: #B22222;
            border: 1px dashed #B22222;
        }
        .task-sub-fusion {
            background-color: rgba(107, 70, 193, 0.2);
            color: #6B46C1;
            border: 1px dashed #6B46C1;
        }
        .legend {
            margin-top: 10px;
            font-size: 11px;
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 5px;
        }
        .legend-dot {
            width: 12px;
            height: 12px;
            border-radius: 2px;
            border: 1px solid white;
        }
        .w-10 { width: 10%; }
        .w-20 { width: 20%; }
        .w-30 { width: 30%; }
        .w-40 { width: 40%; }
        .w-50 { width: 50%; }
        .w-60 { width: 60%; }
        .w-70 { width: 70%; }
        .w-100 { width: 100%; }
        .l-0 { left: 0%; }
        .l-10 { left: 10%; }
        .l-20 { left: 20%; }
        .l-30 { left: 30%; }
        .l-40 { left: 40%; }
        .l-50 { left: 50%; }
        .l-60 { left: 60%; }
        .l-70 { left: 70%; }
    </style>
</head>
<body>
    <div class="slide-container">
        <h1 class="slide-title">[プロジェクト名] | [ロードマップタイトル]</h1>
        <p class="slide-subtitle">[期間] [サブタイトル詳細]</p>
        <div class="roadmap-wrapper">
            <div class="header-grid">
                <div></div>
                <div class="month-label" style="grid-column: 2/5;">[月1]</div>
                <div class="month-label" style="grid-column: 5/9;">[月2]</div>
                <div class="month-label" style="grid-column: 9/12;">[月3]</div>
                <div></div>
                <div class="date-label">[日付1]</div>
                <div class="date-label">[日付2]</div>
                <div class="date-label">[日付3]</div>
                <div class="date-label">[日付4]</div>
                <div class="date-label">[日付5]</div>
                <div class="date-label">[日付6]</div>
                <div class="date-label">[日付7]</div>
                <div class="date-label">[日付8]</div>
                <div class="date-label">[日付9]</div>
                <div class="date-label">[日付10]</div>
            </div>
            <div class="tasks-container">
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名1]</div>
                    <div class="task-area">
                        <div class="task task-code l-0 w-70">メインタスク（7週間）</div>
                        <div class="task task-code l-70 w-30">フォローアップ（3週間）</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名2]</div>
                    <div class="task-area">
                        <div class="task task-mail l-0 w-30">初期フェーズ</div>
                        <div class="task task-sub-mail l-30 w-10">評価</div>
                        <div class="task task-mail l-40 w-30">展開フェーズ</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名3]</div>
                    <div class="task-area">
                        <div class="task task-cms l-10 w-40">CMS統合作業</div>
                        <div class="task task-sub-cms l-50 w-20">テスト期間</div>
                        <div class="task task-cms l-70 w-30">本番運用</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名4]</div>
                    <div class="task-area">
                        <div class="task task-fusion l-0 w-50">Fusion開発</div>
                        <div class="task task-sub-fusion l-50 w-10">検証</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名5]</div>
                    <div class="task-area">
                        <div class="task task-hold l-0 w-60">保留中（他部署待ち）</div>
                        <div class="task task-code l-60 w-20">PoC開始</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名6]</div>
                    <div class="task-area">
                        <div class="task task-code l-20 w-30">実装フェーズ</div>
                        <div class="task task-sub-code l-50 w-10">レビュー</div>
                        <div class="task task-code l-60 w-40">本番展開</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名7]</div>
                    <div class="task-area">
                        <div class="task task-mail l-0 w-20">メール配信テスト</div>
                        <div class="task task-cms l-30 w-40">CMS移行作業</div>
                    </div>
                </div>
                <div class="task-group">
                    <div class="dept-label">[部署/チーム名8]</div>
                    <div class="task-area">
                        <div class="task task-hold l-0 w-100">対応不要 / 状況確認中</div>
                    </div>
                </div>
            </div>
        </div>
        <div class="legend">
            <span><strong>凡例：</strong></span>
            <div class="legend-item">
                <div class="legend-dot task-mail"></div>
                <span>[凡例項目1]</span>
            </div>
            <div class="legend-item">
                <div class="legend-dot task-cms"></div>
                <span>[凡例項目2]</span>
            </div>
            <div class="legend-item">
                <div class="legend-dot task-code"></div>
                <span>[凡例項目3]</span>
            </div>
            <div class="legend-item">
                <div class="legend-dot task-fusion"></div>
                <span>[凡例項目4]</span>
            </div>
            <div class="legend-item">
                <div class="legend-dot task-hold"></div>
                <span>[凡例項目5]</span>
            </div>
        </div>
    </div>
</body>
</html>`
  },
  {
    id: 'template4',
    html: `<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>[スライドタイトル]</title>
    <link rel="stylesheet" href="\${safeGetURL('other/tailwind.min.css')}">
    <style>
        @keyframes fadeInUp { from { opacity: 0; transform: translateY(30px); } to { opacity: 1; transform: translateY(0); } }
        .fade-in-up { animation: fadeInUp 0.6s ease-out forwards; }
        .fade-in-up-delay-1 { animation: fadeInUp 0.6s ease-out 0.2s forwards; opacity: 0; }
        .fade-in-up-delay-2 { animation: fadeInUp 0.6s ease-out 0.4s forwards; opacity: 0; }
        .rakuten-red { color: #bf0000; }
        html, body {
            margin: 0;
            padding: 0;
            color-scheme: light;
        }
        .slide-wrapper { width: 100vw; height: 100vh; display: flex; align-items: center; justify-content: center; padding: 0; }
        .slide-container { width: 100%; max-width: 1422px; aspect-ratio: 16/9; background: white; border-radius: 8px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); padding: 30px 50px; display: flex; flex-direction: column; overflow: hidden; max-height: 100vh; }
        .arrow-box { position: relative; background: #f9fafb; padding: 12px 8px; margin-right: 50px; border-radius: 12px; min-height: 60px; display: flex; align-items: center; justify-content: center; flex: 1; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .arrow-box::after { content: ''; position: absolute; right: -20px; top: 50%; transform: translateY(-50%); width: 0; height: 0; border-left: 20px solid #f9fafb; border-top: 22px solid transparent; border-bottom: 22px solid transparent; }
        .arrow-box.red { background: #fff; border: 2px solid #bf0000; box-shadow: 0 4px 6px rgba(191,0,0,0.1); }
        .arrow-box.red::after { border-left-color: #fff; right: -25px; }
        .arrow-box:last-child { margin-right: 0; }
        .arrow-box:last-child::after { display: none; }
        .process-container { display: flex; align-items: center; justify-content: space-between; width: 100%; gap: 0; }
        .process-section { flex: 1; display: flex; flex-direction: column; gap: 20px; }
        @media (max-width: 1200px) { .slide-container { padding: 24px 40px; } .arrow-box { padding: 8px 6px; min-height: 50px; } }
        @media (max-width: 768px) { .slide-container { padding: 20px 30px; } .process-container { flex-direction: column; gap: 20px; } .arrow-box { margin-right: 0; width: 100%; min-height: 50px; } .arrow-box::after { display: none; } }
    </style>
</head>
<body class="bg-gray-100">
    <div class="slide-wrapper">
        <div class="slide-container">
            <div class="mb-8">
                <h1 class="text-black font-bold mb-3 fade-in-up" style="font-size: 34px; text-align: left;">Before・After</h1>
                <p class="text-gray-700 fade-in-up" style="font-size: 18px; text-align: left;">[このスライドで伝えたいことを150文字以内でまとめる]</p>
            </div>
            <div class="flex-1 flex flex-col justify-center gap-8">
                <div class="process-section fade-in-up-delay-1">
                    <div class="flex items-center mb-4">
                        <div class="bg-gray-100 rounded-full p-2 mr-3"><i class="fas fa-times text-gray-500"></i></div>
                        <h2 class="text-2xl font-bold text-gray-700">Before: [従来の方法・課題の名称]</h2>
                    </div>
                    <div class="process-container">
                        <div class="arrow-box"><div class="text-center"><i class="fas [Beforeプロセス1のアイコンクラス] text-4xl text-gray-400 mb-3"></i><p class="font-bold text-gray-700 text-lg">[Beforeプロセス1のタイトル]</p></div></div>
                        <div class="arrow-box"><div class="text-center"><i class="fas [Beforeプロセス2のアイコンクラス] text-4xl text-gray-400 mb-3"></i><p class="font-bold text-gray-700 text-lg">[Beforeプロセス2のタイトル]</p></div></div>
                        <div class="arrow-box"><div class="text-center"><i class="fas [Beforeプロセス3のアイコンクラス] text-4xl text-gray-400 mb-3"></i><p class="font-bold text-gray-700 text-lg">[Beforeプロセス3のタイトル]</p></div></div>
                    </div>
                </div>
                <div class="process-section fade-in-up-delay-2">
                    <div class="flex items-center mb-4">
                        <div class="bg-red-100 rounded-full p-2 mr-3"><i class="fas fa-check rakuten-red"></i></div>
                        <h2 class="text-2xl font-bold rakuten-red">After: [新しい方法・解決策の名称]</h2>
                    </div>
                    <div class="process-container">
                        <div class="arrow-box red"><div class="text-center"><i class="fas [Afterプロセス1のアイコンクラス] text-4xl rakuten-red mb-3"></i><p class="font-bold text-lg">[Afterプロセス1のタイトル]</p><p class="text-sm text-gray-600 mt-1">[Afterプロセス1の説明]</p></div></div>
                        <div class="arrow-box red"><div class="text-center"><i class="fas [Afterプロセス2のアイコンクラス] text-4xl rakuten-red mb-3"></i><p class="font-bold text-lg">[Afterプロセス2のタイトル]</p><p class="text-sm text-gray-600 mt-1">[Afterプロセス2の説明]</p></div></div>
                        <div class="arrow-box red"><div class="text-center"><i class="fas [Afterプロセス3のアイコンクラス] text-4xl rakuten-red mb-3"></i><p class="font-bold text-lg">[Afterプロセス3のタイトル]</p><p class="text-sm text-gray-600 mt-1">[Afterプロセス3の説明]</p></div></div>
                        <div class="arrow-box red"><div class="text-center"><i class="fas [Afterプロセス4のアイコンクラス] text-4xl rakuten-red mb-3"></i><p class="font-bold text-lg">[Afterプロセス4のタイトル]</p><p class="text-sm text-gray-600 mt-1">[Afterプロセス4の説明]</p></div></div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>`
  },
  {
    id: 'template5',
    html: `<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>表と集計結果</title>
    <link rel="stylesheet" href="\${safeGetURL('other/tailwind.min.css')}">
    <style>
        @keyframes fadeInUp { from { opacity: 0; transform: translateY(30px); } to { opacity: 1; transform: translateY(0); } }
        .fade-in-up { animation: fadeInUp 0.6s ease-out forwards; }
        .delay-1 { animation-delay: 0.1s; opacity: 0; }
        .delay-2 { animation-delay: 0.2s; opacity: 0; }
        .delay-3 { animation-delay: 0.3s; opacity: 0; }
        .rakuten-red { color: #bf0000; }
        html, body {
            margin: 0;
            padding: 0;
            color-scheme: light;
        }
        .slide-container { aspect-ratio: 16/9; max-width: 100vw; max-height: 100vh; }
        .table-header { background-color: #f8f8f8; border-bottom: 2px solid #bf0000; }
        .company-header { background-color: #f0f0f0; font-weight: bold; }
    </style>
</head>
<body class="bg-gray-100 flex items-center justify-center min-h-screen">
    <div class="slide-container w-full bg-white rounded-lg shadow-lg p-6 flex flex-col">
        <h1 class="text-black font-bold mb-4 fade-in-up" style="font-size: 34px; text-align: left;">[スライドタイトル（20文字以内）]</h1>
        <div class="flex-1 fade-in-up delay-1">
            <div class="overflow-hidden rounded-lg border border-gray-200">
                <table class="w-full text-sm">
                    <thead>
                        <tr class="table-header">
                            <th class="text-left p-3 font-semibold">[列1のヘッダー]</th>
                            <th class="text-left p-3 font-semibold">[列2のヘッダー]</th>
                            <th class="text-center p-3 font-semibold">[列3のヘッダー]</th>
                            <th class="text-center p-3 font-semibold">[列4のヘッダー]</th>
                            <th class="text-center p-3 font-semibold">[列5のヘッダー]</th>
                            <th class="text-left p-3 font-semibold">[列6のヘッダー]</th>
                        </tr>
                    </thead>
                    <tbody class="bg-white">
                        <tr class="border-b border-gray-200">
                            <td class="p-3 company-header" rowspan="[結合する行数]">[カテゴリー名1]</td>
                            <td class="p-3"><div class="flex items-center"><i class="fas fa-[アイコン名] rakuten-red mr-2"></i><span>[サブカテゴリー名1]</span></div></td>
                            <td class="text-center p-3"><span class="rakuten-red font-bold text-lg">[強調する数値1]</span></td>
                            <td class="text-center p-3"><span class="font-semibold">[データ1]</span></td>
                            <td class="text-center p-3"><span class="font-semibold">[データ2]</span></td>
                            <td class="p-3 text-xs">[詳細情報1]<br><span class="text-red-600 font-semibold">[重要な注記1]</span></td>
                        </tr>
                        <tr class="border-b border-gray-200">
                            <td class="p-3 company-header" rowspan="[結合する行数]">[カテゴリー名2]</td>
                            <td class="p-3"><div class="flex items-center"><i class="fas fa-[アイコン名] rakuten-red mr-2"></i><span>[サブカテゴリー名2]</span></div></td>
                            <td class="text-center p-3"><span class="rakuten-red font-bold text-lg">[強調する数値2]</span></td>
                            <td class="text-center p-3"><span class="font-semibold">[データ3]</span></td>
                            <td class="text-center p-3"><span class="font-semibold">[データ4]</span></td>
                            <td class="p-3 text-xs">[詳細情報2]<br><span class="text-red-600 font-semibold">[重要な注記2]</span></td>
                        </tr>
                        <tr class="border-b border-gray-200">
                            <td class="p-3 company-header">[カテゴリー名3]</td>
                            <td class="p-3"><div class="flex items-center"><i class="fas fa-[アイコン名] rakuten-red mr-2"></i><span>[サブカテゴリー名3]</span></div></td>
                            <td class="text-center p-3"><span class="rakuten-red font-bold text-lg">[強調する数値3]</span></td>
                            <td class="text-center p-3"><span class="font-semibold">[データ5]</span></td>
                            <td class="text-center p-3"><span class="font-semibold">[データ6]</span></td>
                            <td class="p-3 text-xs">[詳細情報3]<br><span class="text-red-600 font-semibold">[重要な注記3]</span></td>
                        </tr>
                        <tr class="border-b border-gray-200">
                            <td class="p-3 company-header">[カテゴリー名]</td>
                            <td class="p-3">[サブカテゴリー名またはハイフン]</td>
                            <td class="text-center p-3">-</td>
                            <td class="text-center p-3">-</td>
                            <td class="text-center p-3">-</td>
                            <td class="p-3 text-xs text-gray-500">[ステータス（例：調査中）]</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div class="mt-4 grid grid-cols-3 gap-4 fade-in-up delay-2">
                <div class="bg-white rounded-lg p-4 border border-gray-200 text-center">
                    <p class="text-sm text-gray-600 mb-2">[指標名1]</p>
                    <p class="text-3xl font-bold">[値1]</p>
                    <p class="text-xs text-gray-500 mt-1">[計算式または補足説明]</p>
                </div>
                <div class="bg-white rounded-lg p-4 border border-gray-200 text-center">
                    <p class="text-sm text-gray-600 mb-2">[指標名2]</p>
                    <p class="text-3xl font-bold">[値2]</p>
                    <p class="text-xs text-gray-500 mt-1">[計算式または補足説明]</p>
                </div>
                <div class="bg-white rounded-lg p-4 border border-gray-200 text-center">
                    <p class="text-sm text-gray-600 mb-2">[指標名3]</p>
                    <p class="text-3xl font-bold rakuten-red">[強調したい値3]</p>
                    <p class="text-xs text-gray-500 mt-1">[計算式または補足説明]</p>
                </div>
            </div>
        </div>
    </div>
</body>
</html>`
  },
  {
    id: 'template6',
    html: `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>レーダーチャート｜サンプル</title>
    <link rel="stylesheet" href="\${safeGetURL('other/tailwind.min.css')}">
    <script src="\${safeGetURL('other/chart.js')}"></script>
    <style>
        body {
            font-family: 'Noto Sans JP', sans-serif;
            background-color: #fff;
            margin: 0;
            padding: 0;
            color-scheme: light;
        }
        .slide {
            width: 100%;
            max-width: 1280px;
            aspect-ratio: 16/9;
            background-color: white;
            margin: 0 auto;
            padding: 30px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            position: relative;
            overflow: hidden;
        }
        .fade-in {
            animation: fadeInUp 0.8s ease-out;
        }
        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(30px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        .section-border {
            width: 4px;
            height: 20px;
            background-color: #EC4899;
            position: absolute;
            left: 0;
        }
        .table-custom {
            border-collapse: collapse;
            width: 100%;
        }
        .table-custom th {
            background-color: #F3F4F6;
            font-weight: bold;
            font-size: 14px;
            padding: 8px;
            border: 1px solid #E5E7EB;
            text-align: center;
        }
        .table-custom td {
            font-size: 14px;
            padding: 10px;
            border: 1px solid #E5E7EB;
            text-align: center;
        }
        .highlight-text {
            color: #EC4899;
            font-weight: bold;
        }
        .info-box {
            background-color: #F9FAFB;
            border-radius: 8px;
            padding: 16px;
        }
        @media (max-width: 768px) {
            .slide {
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="slide fade-in">
        <!-- Title -->
        <h1 style="font-size: 34px; font-weight: bold; color: #000000; text-align: left; margin-bottom: 12px;">
            レーダーチャート｜サンプル
        </h1>
        <!-- Subtitle -->
        <p style="font-size: 18px; color: #4B5563; text-align: left; margin-bottom: 20px; line-height: 1.4;">
            このスライドで伝えたい主要メッセージや、データ分析から得られた結論・インサイトを記載
        </p>
        <!-- Data Overview Section -->
        <div style="position: relative; margin-bottom: 16px;">
            <div class="section-border" style="top: 2px;"></div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding-left: 16px;">
                <h2 style="font-size: 18px; font-weight: bold; color: #000000;">表データ</h2>
                <span style="font-size: 12px; color: #6B7280;">データ更新日</span>
            </div>
        </div>
        <!-- Table -->
        <div style="margin-bottom: 20px; overflow-x: auto;">
            <table class="table-custom">
                <thead>
                    <tr>
                        <th style="text-align: left;">見出し</th>
                        <th>項目1</th>
                        <th>項目2</th>
                        <th>項目3</th>
                        <th>項目4</th>
                        <th>項目5</th>
                        <th>項目6</th>
                        <th>項目7</th>
                        <th>項目8</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="text-align: left; font-weight: bold; font-size: 14px;">行見出し</td>
                        <td class="highlight-text" style="font-size: 14px;">数値1</td>
                        <td class="highlight-text" style="font-size: 14px;">数値2</td>
                        <td class="highlight-text" style="font-size: 14px;">数値3</td>
                        <td class="highlight-text" style="font-size: 14px;">数値4</td>
                        <td style="font-size: 14px;">数値5</td>
                        <td style="font-size: 14px;">数値6</td>
                        <td style="font-size: 14px;">数値7</td>
                        <td style="font-size: 14px;">数値8</td>
                    </tr>
                </tbody>
            </table>
        </div>
        <!-- Bottom Section Container -->
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px;">
            <!-- Left: Radar Chart -->
            <div>
                <div style="position: relative; margin-bottom: 12px;">
                    <div class="section-border"></div>
                    <h2 style="font-size: 18px; font-weight: bold; color: #000000; padding-left: 16px;">レーダーチャート</h2>
                </div>
                <div style="position: relative; height: 220px;">
                    <canvas id="radarChart"></canvas>
                </div>
                <p style="font-size: 11px; color: #6B7280; text-align: center; margin-top: 4px;">
                    グラフの補足説明や注釈
                </p>
            </div>
            <!-- Right: Key Insights -->
            <div>
                <div style="position: relative; margin-bottom: 12px;">
                    <div class="section-border"></div>
                    <h2 style="font-size: 18px; font-weight: bold; color: #000000; padding-left: 16px;">分析結果・アクション</h2>
                </div>
                <table class="table-custom" style="margin-bottom: 12px;">
                    <thead>
                        <tr>
                            <th style="width: 35%; font-size: 14px;">分類</th>
                            <th style="width: 65%; font-size: 14px;">詳細内容</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="text-align: left; font-size: 13px;">強み・良い点</td>
                            <td style="text-align: left; color: #4B5563; font-size: 13px;">データから読み取れる強みや優れている点</td>
                        </tr>
                        <tr>
                            <td style="text-align: left; font-size: 13px;">課題・改善点</td>
                            <td style="text-align: left; color: #4B5563; font-size: 13px;">データから読み取れる課題や改善が必要な点</td>
                        </tr>
                    </tbody>
                </table>
                <div class="info-box">
                    <p style="font-size: 16px; color: #4B5564; line-height: 1.5;">
                        データ分析から導き出された重要な結論や、次のアクションにつながる提言を記載
                    </p>
                </div>
            </div>
        </div>
    </div>
    <script>
        // Radar Chart
        const ctx = document.getElementById('radarChart').getContext('2d');
        const radarChart = new Chart(ctx, {
            type: 'radar',
            data: {
                labels: ['評価軸1', '評価軸2', '評価軸3', '評価軸4', '評価軸5', '評価軸6'],
                datasets: [
                    {
                        label: '比較対象1',
                        data: [80, 80, 80, 80, 80, 80],
                        backgroundColor: 'rgba(59, 130, 246, 0.3)',
                        borderColor: '#3B82F6',
                        borderWidth: 2,
                        pointBackgroundColor: '#3B82F6',
                        pointBorderColor: '#FFFFFF',
                        pointBorderWidth: 2,
                        pointRadius: 6
                    },
                    {
                        label: '比較対象2',
                        data: [92, 88, 95, 78, 85, 90],
                        backgroundColor: 'rgba(236, 72, 153, 0.3)',
                        borderColor: '#EC4899',
                        borderWidth: 2,
                        pointBackgroundColor: '#EC4899',
                        pointBorderColor: '#FFFFFF',
                        pointBorderWidth: 2,
                        pointRadius: 6
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            font: {
                                size: 14,
                                family: "'Noto Sans JP', sans-serif"
                            },
                            color: '#000000',
                            padding: 10
                        }
                    }
                },
                scales: {
                    r: {
                        min: 0,
                        max: 100,
                        ticks: {
                            stepSize: 20,
                            font: {
                                size: 13,
                                family: "'Noto Sans JP', sans-serif"
                            },
                            color: '#6B7280'
                        },
                        grid: {
                            color: '#E5E7EB',
                            lineWidth: 0.5
                        },
                        pointLabels: {
                            font: {
                                size: 14,
                                family: "'Noto Sans JP', sans-serif",
                                weight: 'bold'
                            },
                            color: '#1F2937'
                        }
                    }
                }
            }
        });
    </script>
</body>
</html>
`
  }
];

window.TEMPLATE_SLIDES = TEMPLATE_SLIDES;
