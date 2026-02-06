/**
 * 市区町村別建築統計マップ
 * 
 * 機能:
 * - 国土地理院地図上に市区町村別建築統計をコロプレスマップで表示
 * - 着工件数（推計）、床面積、見込み工事額の3指標を切り替え可能
 * - インタラクティブな操作（ホバー、クリック、凡例）
 * 
 * データソース:
 * - data/municipality_stats.csv: 市区町村別統計データ
 * - data/municipalities.geojson: N03行政区域データ（25.85MB、9,296ポリゴン）
 * - 埋め込み: 都道府県別平均工事単価
 */

// ============================================================================
// 定数定義
// ============================================================================

const CONFIG = {
    // 地図設定
    MAP: {
        INITIAL_CENTER: [36.0, 138.0],
        INITIAL_ZOOM: 6,
        MAX_ZOOM: 18,
        TILE_URL: 'https://cyberjapandata.gsi.go.jp/xyz/pale/{z}/{x}/{y}.png',
        ATTRIBUTION: '<a href="https://maps.gsi.go.jp/development/ichiran.html">国土地理院</a>'
    },
    
    // データファイルパス
    DATA: {
        MUNICIPALITY_STATS: 'data/stats.csv',
        GEOJSON: 'data/municipalities.geojson'
    },
    
    // 色設定（5段階コロプレス）
    COLORS: {
        SCALE: ['#ffffcc', '#c2e699', '#78c679', '#31a354', '#006837'],
        NO_DATA: '#cccccc',
        BORDER: 'white',
        HIGHLIGHT: '#666'
    },
    
    // スタイル設定
    STYLE: {
        DEFAULT_WEIGHT: 1,
        HIGHLIGHT_WEIGHT: 3,
        DEFAULT_OPACITY: 1,
        DEFAULT_FILL_OPACITY: 0.7,
        NO_DATA_FILL_OPACITY: 0.1,
        HIGHLIGHT_FILL_OPACITY: 0.9
    },
    
    // 指標設定
    INDICATORS: {
        buildingCount: {
            field: 'buildingCount',
            label: '着工件数',
            unit: '棟'
        },
        floorAreaTotal: {
            field: 'floorAreaTotal',
            label: '全建物床面積',
            unit: '㎡'
        },
        estimatedAmount: {
            field: 'estimatedAmount',
            label: '見込み工事額',
            unit: '円'
        }
    }
};

// 都道府県別平均工事単価（埋め込みデータ）
const PREF_UNIT_COST_DATA = `pref_name,unit_cost_per_m2
千葉県,280000
東京都,340000
埼玉県,270000`;

// ============================================================================
// グローバル状態
// ============================================================================

const state = {
    // データストア
    prefUnitCostMap: new Map(),
    municipalityStatsMap: new Map(),
    municipalitiesGeoJSON: null,
    
    // Leafletオブジェクト
    map: null,
    geoJSONLayer: null,
    
    // UI状態
    currentYear: '2025',
    currentMonth: '9',
    currentIndicator: 'buildingCount',
    
    // 現在の階級区分
    currentBreaks: []
};

// ============================================================================
// ユーティリティ関数
// ============================================================================

/**
 * 文字列の正規化（空白除去）
 */
function normalizeName(str) {
    if (!str) return '';
    return str.trim().replace(/\s+/g, '').replace(/　/g, '');
}

/**
 * 市区町村データのキー生成
 */
function buildMunicipalityKey(prefName, cityName, year, month) {
    const p = normalizeName(prefName);
    const c = normalizeName(cityName);
    return `${p}__${c}__${year}__${month}`;
}

/**
 * 数値フォーマット（K/M表記）
 */
function formatNumber(num) {
    if (num >= 1000000) {
        return (num / 1000000).toFixed(1) + 'M';
    } else if (num >= 1000) {
        return (num / 1000).toFixed(1) + 'K';
    }
    return Math.round(num).toLocaleString();
}

/**
 * ローディング表示制御
 */
function setLoadingVisible(visible) {
    const loading = document.getElementById('loading');
    loading.classList.toggle('visible', visible);
}

// ============================================================================
// データ読み込み
// ============================================================================

/**
 * 都道府県別平均工事単価の読み込み
 */
function loadPrefUnitCost() {
    return new Promise((resolve, reject) => {
        Papa.parse(PREF_UNIT_COST_DATA, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                results.data.forEach(row => {
                    const prefName = row.pref_name;
                    const unitCost = Number(row.unit_cost_per_m2);
                    
                    if (prefName && !isNaN(unitCost)) {
                        const key = normalizeName(prefName);
                        state.prefUnitCostMap.set(key, unitCost);
                    }
                });
                console.log(`✓ 都道府県別単価読み込み完了: ${state.prefUnitCostMap.size}件`);
                resolve();
            },
            error: (error) => {
                reject(new Error('都道府県別単価の解析に失敗: ' + error.message));
            }
        });
    });
}

/**
 * 市区町村別統計データの読み込み
 */
async function loadMunicipalityStats() {
    try {
        const response = await fetch(CONFIG.DATA.MUNICIPALITY_STATS);
        if (!response.ok) {
            throw new Error(`HTTPエラー: ${response.status} ${response.statusText}`);
        }
        
        const csvText = await response.text();
        
        return new Promise((resolve, reject) => {
            Papa.parse(csvText, {
                header: true,
                skipEmptyLines: true,
                complete: (results) => {
                    results.data.forEach(row => {
                        const prefName = row.pref_name;
                        const cityName = row.city_name;
                        const year = row.year;
                        const month = row.month;
                        const buildingCount = Number(row.building_count_A_Residence);
                        const floorAreaTotal = Number(row.floor_area_total);
                        const aResidenceArea = Number(row.A_Residence_Area);
                        
                        if (!prefName || !cityName || !year || !month) return;
                        
                        // 都道府県の平均単価を取得
                        const prefKey = normalizeName(prefName);
                        const unitCost = state.prefUnitCostMap.get(prefKey) || 0;
                        
                        // 見込み工事額を計算（居住専用住宅の床面積 × 単価）
                        const estimatedAmount = aResidenceArea * unitCost;
                        
                        // データを格納
                        const key = buildMunicipalityKey(prefName, cityName, year, month);
                        state.municipalityStatsMap.set(key, {
                            prefName,
                            cityName,
                            year,
                            month,
                            buildingCount,
                            floorAreaTotal,
                            aResidenceArea,
                            estimatedAmount
                        });
                    });
                    
                    console.log(`✓ 市区町村統計読み込み完了: ${state.municipalityStatsMap.size}件`);
                    resolve();
                },
                error: (error) => {
                    reject(new Error('統計データの解析に失敗: ' + error.message));
                }
            });
        });
    } catch (error) {
        throw new Error('統計データの取得に失敗: ' + error.message);
    }
}

/**
 * GeoJSON（行政区域データ）の読み込み
 */
async function loadGeoJSON() {
    try {
        const response = await fetch(CONFIG.DATA.GEOJSON);
        if (!response.ok) {
            throw new Error(`HTTPエラー: ${response.status} ${response.statusText}`);
        }
        
        state.municipalitiesGeoJSON = await response.json();
        console.log(`✓ GeoJSON読み込み完了: ${state.municipalitiesGeoJSON.features.length}フィーチャー`);
    } catch (error) {
        throw new Error('行政区域データの取得に失敗: ' + error.message);
    }
}

/**
 * 全データの一括読み込み
 */
async function loadAllData() {
    try {
        // 都道府県単価を先に読み込み（統計データの計算に必要）
        await loadPrefUnitCost();
        
        // 残りのデータを並行読み込み
        await Promise.all([
            loadMunicipalityStats(),
            loadGeoJSON()
        ]);
        
        console.log('✓ 全データ読み込み完了');
    } catch (error) {
        throw error;
    }
}

// ============================================================================
// 地図初期化・描画
// ============================================================================

/**
 * Leaflet地図の初期化
 */
function initMap() {
    state.map = L.map('map').setView(
        CONFIG.MAP.INITIAL_CENTER,
        CONFIG.MAP.INITIAL_ZOOM
    );
    
    // 国土地理院タイルレイヤー
    L.tileLayer(CONFIG.MAP.TILE_URL, {
        attribution: CONFIG.MAP.ATTRIBUTION,
        maxZoom: CONFIG.MAP.MAX_ZOOM
    }).addTo(state.map);
    
    console.log('✓ 地図初期化完了');
}

/**
 * GeoJSONフィーチャーにデータを付与
 */
function attachDataToFeatures() {
    if (!state.municipalitiesGeoJSON) return;
    
    state.municipalitiesGeoJSON.features.forEach(feature => {
        const prefName = feature.properties.pref_name;
        const cityName = feature.properties.city_name;
        
        if (!prefName || !cityName) return;
        
        const key = buildMunicipalityKey(
            prefName,
            cityName,
            state.currentYear,
            state.currentMonth
        );
        
        const stats = state.municipalityStatsMap.get(key);
        
        if (stats) {
            feature.properties.buildingCount = stats.buildingCount;
            feature.properties.floorAreaTotal = stats.floorAreaTotal;
            feature.properties.estimatedAmount = stats.estimatedAmount;
            feature.properties.year = stats.year;
            feature.properties.month = stats.month;
        } else {
            // データなし
            feature.properties.buildingCount = null;
            feature.properties.floorAreaTotal = null;
            feature.properties.estimatedAmount = null;
        }
    });
}

/**
 * 階級区分の計算（5段階）
 */
function calculateBreaks(indicator) {
    const values = [];
    
    state.municipalitiesGeoJSON.features.forEach(feature => {
        const value = feature.properties[indicator];
        if (value !== null && value !== undefined && !isNaN(value)) {
            values.push(value);
        }
    });
    
    if (values.length === 0) {
        return [0, 0, 0, 0, 0, 0];
    }
    
    values.sort((a, b) => a - b);
    
    const min = values[0];
    const max = values[values.length - 1];
    const range = max - min;
    
    // 等間隔5段階
    return [
        min,
        min + range * 0.2,
        min + range * 0.4,
        min + range * 0.6,
        min + range * 0.8,
        max
    ];
}

/**
 * 値から色を取得
 */
function getColor(value, breaks) {
    if (value === null || value === undefined || isNaN(value)) {
        return CONFIG.COLORS.NO_DATA;
    }
    
    const colors = CONFIG.COLORS.SCALE;
    
    for (let i = breaks.length - 1; i >= 0; i--) {
        if (value >= breaks[i]) {
            return colors[Math.min(i, colors.length - 1)];
        }
    }
    
    return colors[0];
}

/**
 * フィーチャーのスタイル取得
 */
function getFeatureStyle(feature) {
    const value = feature.properties[state.currentIndicator];
    const hasData = value !== null && value !== undefined;
    
    return {
        fillColor: getColor(value, state.currentBreaks),
        weight: CONFIG.STYLE.DEFAULT_WEIGHT,
        opacity: CONFIG.STYLE.DEFAULT_OPACITY,
        color: CONFIG.COLORS.BORDER,
        fillOpacity: hasData 
            ? CONFIG.STYLE.DEFAULT_FILL_OPACITY 
            : CONFIG.STYLE.NO_DATA_FILL_OPACITY
    };
}

/**
 * 地図レイヤーの更新
 */
function updateMapLayer() {
    if (!state.municipalitiesGeoJSON) {
        console.warn('GeoJSONデータが読み込まれていません');
        return;
    }
    
    setLoadingVisible(true);
    
    // 既存レイヤーを削除
    if (state.geoJSONLayer) {
        state.map.removeLayer(state.geoJSONLayer);
    }
    
    // データ付与
    attachDataToFeatures();
    
    // 階級区分を計算
    state.currentBreaks = calculateBreaks(state.currentIndicator);
    
    // GeoJSONレイヤーを作成
    state.geoJSONLayer = L.geoJSON(state.municipalitiesGeoJSON, {
        style: getFeatureStyle,
        onEachFeature: (feature, layer) => {
            layer.on({
                mouseover: onFeatureMouseOver,
                mouseout: onFeatureMouseOut,
                click: onFeatureClick
            });
        }
    }).addTo(state.map);
    
    // 凡例を更新
    updateLegend();
    
    // ポリゴンレイヤーを最前面に移動
    bringDrawnItemsToFront();
    
    setLoadingVisible(false);
    
    console.log(`✓ 地図更新完了: ${state.currentYear}年${state.currentMonth}月 - ${CONFIG.INDICATORS[state.currentIndicator].label}`);
}

// ============================================================================
// インタラクション
// ============================================================================

/**
 * フィーチャーマウスオーバー
 */
function onFeatureMouseOver(e) {
    const layer = e.target;
    
    layer.setStyle({
        weight: CONFIG.STYLE.HIGHLIGHT_WEIGHT,
        color: CONFIG.COLORS.HIGHLIGHT,
        fillOpacity: CONFIG.STYLE.HIGHLIGHT_FILL_OPACITY
    });
    
    layer.bringToFront();
}

/**
 * フィーチャーマウスアウト
 */
function onFeatureMouseOut(e) {
    if (state.geoJSONLayer) {
        state.geoJSONLayer.resetStyle(e.target);
    }
}

/**
 * フィーチャークリック（ポップアップ表示）
 */
function onFeatureClick(e) {
    const layer = e.target;
    const props = layer.feature.properties;
    
    const prefName = props.pref_name || '';
    const cityName = props.city_name || '';
    
    // ポップアップコンテンツ生成
    let content = `<div class="popup-title">${prefName} ${cityName}</div>`;
    content += `<div class="popup-info">`;
    content += `<div class="popup-info-item">`;
    content += `<span class="popup-info-label">年月:</span>`;
    content += `<span class="popup-info-value">${state.currentYear}年${state.currentMonth}月</span>`;
    content += `</div>`;
    
    if (props.buildingCount !== null && props.buildingCount !== undefined) {
        content += `<div class="popup-info-item">`;
        content += `<span class="popup-info-label">着工件数（推計）:</span>`;
        content += `<span class="popup-info-value">${props.buildingCount.toLocaleString()} 棟</span>`;
        content += `</div>`;
        
        content += `<div class="popup-info-item">`;
        content += `<span class="popup-info-label">全建物床面積:</span>`;
        content += `<span class="popup-info-value">${props.floorAreaTotal.toLocaleString()} ㎡</span>`;
        content += `</div>`;
        
        content += `<div class="popup-info-item">`;
        content += `<span class="popup-info-label">見込み工事額:</span>`;
        content += `<span class="popup-info-value">${Math.round(props.estimatedAmount).toLocaleString()} 円</span>`;
        content += `</div>`;
    } else {
        content += `<div class="popup-no-data">データなし</div>`;
    }
    
    content += `</div>`;
    
    // ポリゴンの中心座標を取得
    const center = layer.getBounds().getCenter();
    
    L.popup()
        .setLatLng(center)
        .setContent(content)
        .openOn(state.map);
}

// ============================================================================
// 凡例
// ============================================================================

/**
 * 凡例の更新
 */
function updateLegend() {
    const legendContent = document.getElementById('legend-content');
    legendContent.innerHTML = '';
    
    const colors = CONFIG.COLORS.SCALE;
    const indicator = CONFIG.INDICATORS[state.currentIndicator];
    const unit = indicator.unit;
    
    // 階級ごとの凡例アイテム
    for (let i = colors.length - 1; i >= 0; i--) {
        const item = createLegendItem(
            colors[i],
            `${formatNumber(state.currentBreaks[i])} ${unit} 以上`
        );
        legendContent.appendChild(item);
    }
    
    // データなしの凡例アイテム
    const noDataItem = createLegendItem(CONFIG.COLORS.NO_DATA, 'データなし');
    legendContent.appendChild(noDataItem);
}

/**
 * 凡例アイテムの生成
 */
function createLegendItem(color, label) {
    const item = document.createElement('div');
    item.className = 'legend-item';
    
    const colorBox = document.createElement('div');
    colorBox.className = 'legend-color';
    colorBox.style.backgroundColor = color;
    
    const labelSpan = document.createElement('span');
    labelSpan.className = 'legend-label';
    labelSpan.textContent = label;
    
    item.appendChild(colorBox);
    item.appendChild(labelSpan);
    
    return item;
}

// ============================================================================
// イベントハンドラー
// ============================================================================

/**
 * 更新ボタンクリック
 */
function onUpdateButtonClick() {
    state.currentYear = document.getElementById('year-select').value;
    state.currentMonth = document.getElementById('month-select').value;
    state.currentIndicator = document.getElementById('indicator-select').value;
    
    updateMapLayer();
}

/**
 * 統計レイヤーの表示切替
 */
function toggleStatsLayer(show) {
    if (!state.geoJSONLayer) return;
    
    if (show) {
        if (!state.map.hasLayer(state.geoJSONLayer)) {
            state.map.addLayer(state.geoJSONLayer);
        }
        // 統計レイヤー表示後、エリア分析レイヤーを最前面に
        if (areaAnalysisState.drawnItems) {
            areaAnalysisState.drawnItems.bringToFront();
        }
    } else {
        if (state.map.hasLayer(state.geoJSONLayer)) {
            state.map.removeLayer(state.geoJSONLayer);
        }
    }
}

/**
 * イベントリスナーの設定
 */
function setupEventListeners() {
    document.getElementById('update-btn').addEventListener('click', onUpdateButtonClick);
    
    // 統計レイヤー表示切替
    const statsLayerCheckbox = document.getElementById('show-stats-layer-checkbox');
    if (statsLayerCheckbox) {
        statsLayerCheckbox.addEventListener('change', (e) => {
            toggleStatsLayer(e.target.checked);
        });
    }
}

// ============================================================================
// 初期化・起動
// ============================================================================

/**
 * アプリケーション初期化
 */
async function initialize() {
    console.log('アプリケーション起動中...');
    setLoadingVisible(true);
    
    try {
        // 地図初期化
        initMap();
        
        // データ読み込み
        await loadAllData();
        
        // イベントリスナー設定
        setupEventListeners();
        
        // 初回描画
        updateMapLayer();
        
        console.log('✓ アプリケーション起動完了');
    } catch (error) {
        console.error('初期化エラー:', error);
        alert(`データの読み込みに失敗しました。\n\n${error.message}\n\nサーバーが起動しているか、ファイルパスが正しいか確認してください。`);
    } finally {
        setLoadingVisible(false);
    }
}

// ============================================================================
// 建築計画データ機能
// ============================================================================

// 建築計画データの状態
const constructionState = {
    data: null,
    markers: [],
    currentMonth: null,
    isVisible: false
};

// 建築計画データを読み込む
async function loadConstructionData() {
    console.log('建築計画データ読み込み開始...');
    try {
        // 全件座標付きデータを使用（2025年1月～12月完全版）
        const response = await fetch('data/construction_projects.json');
        console.log('Response status:', response.status);
        
        if (!response.ok) {
            throw new Error(`建築計画データの読み込みに失敗しました: ${response.status}`);
        }
        
        constructionState.data = await response.json();
        
        // プロジェクトデータから完成年月・着工年月のリストを動的に生成
        const completionMonths = new Set();
        const startMonths = new Set();
        
        constructionState.data.projects.forEach(project => {
            // 完成日から年月を抽出 (YYYY/MM/DD → YYYY/MM)
            if (project.completion_date && project.completion_date !== 'nan') {
                const match = project.completion_date.match(/(\d{4})\/(\d{2})/);
                if (match) {
                    project.完成年月 = `${match[1]}/${match[2]}`;
                    completionMonths.add(project.完成年月);
                }
            }
            // 着工日から年月を抽出 (YYYY/MM/DD → YYYY/MM)
            if (project.start_date && project.start_date !== 'nan') {
                const match = project.start_date.match(/(\d{4})\/(\d{2})/);
                if (match) {
                    project.着工年月 = `${match[1]}/${match[2]}`;
                    startMonths.add(project.着工年月);
                }
            }
        });
        
        constructionState.data.completion_months = Array.from(completionMonths).sort();
        constructionState.data.start_months = Array.from(startMonths).sort();
        
        console.log(`✓ 建築計画データ読み込み完了`);
        console.log(`  - プロジェクト件数: ${constructionState.data.projects.length}件`);
        console.log(`  - 完成年月の種類: ${constructionState.data.completion_months.length}ヶ月`);
        console.log(`  - 着工年月の種類: ${constructionState.data.start_months.length}ヶ月`);
        console.log(`  - 完成年月リスト:`, constructionState.data.completion_months);
        console.log(`  - 着工年月リスト:`, constructionState.data.start_months);
        
        // 完成年月・着工年月のプルダウンを初期化
        initCompletionMonthSelect();
        initStartMonthSelect();
    } catch (error) {
        console.error('❌ 建築計画データ読み込みエラー:', error);
        alert(`建築計画データの読み込みに失敗しました。\n\nエラー: ${error.message}\n\nサーバーが起動しているか確認してください。`);
    }
}

// 完成年月プルダウンを初期化
function initCompletionMonthSelect() {
    console.log('完成年月プルダウン初期化開始...');
    const selectFrom = document.getElementById('completion-month-from');
    const selectTo = document.getElementById('completion-month-to');
    
    if (!selectFrom || !selectTo) {
        console.error('❌ completion-month-from または completion-month-to 要素が見つかりません');
        return;
    }
    
    if (!constructionState.data) {
        console.error('❌ constructionState.dataが未定義です');
        return;
    }
    
    console.log(`  - セレクト要素: 見つかりました`);
    console.log(`  - データ: ${constructionState.data.completion_months.length}ヶ月分`);
    
    // 既存のオプションをクリア（最初のオプションは残す）
    while (selectFrom.options.length > 1) {
        selectFrom.remove(1);
    }
    while (selectTo.options.length > 1) {
        selectTo.remove(1);
    }
    
    // 完成年月のオプションを両方のセレクトに追加
    constructionState.data.completion_months.forEach((month, index) => {
        const optionFrom = document.createElement('option');
        optionFrom.value = month;
        optionFrom.textContent = month;
        selectFrom.appendChild(optionFrom);
        
        const optionTo = document.createElement('option');
        optionTo.value = month;
        optionTo.textContent = month;
        selectTo.appendChild(optionTo);
        
        if (index < 5) {
            console.log(`  - オプション追加: ${month}`);
        }
    });
    
    // デフォルト値を設定（最初の月～最後の月）
    if (constructionState.data.completion_months.length > 0) {
        selectFrom.value = constructionState.data.completion_months[0];
        selectTo.value = constructionState.data.completion_months[constructionState.data.completion_months.length - 1];
        console.log(`  - デフォルト期間設定: ${selectFrom.value} ～ ${selectTo.value}`);
    }
    
    console.log(`✓ 完成年月プルダウン初期化完了: ${selectFrom.options.length - 1}個のオプション追加`);
}

// 着工年月プルダウンを初期化
function initStartMonthSelect() {
    console.log('着工年月プルダウン初期化開始...');
    const selectFrom = document.getElementById('start-month-from');
    const selectTo = document.getElementById('start-month-to');
    
    if (!selectFrom || !selectTo) {
        console.error('❌ start-month-from または start-month-to 要素が見つかりません');
        return;
    }
    
    if (!constructionState.data) {
        console.error('❌ constructionState.dataが未定義です');
        return;
    }
    
    if (!constructionState.data.start_months) {
        console.error('❌ constructionState.data.start_monthsが未定義です');
        return;
    }
    
    console.log(`  - セレクト要素: 見つかりました`);
    console.log(`  - データ: ${constructionState.data.start_months.length}ヶ月分`);
    
    // 既存のオプションをクリア（最初のオプションは残す）
    while (selectFrom.options.length > 1) {
        selectFrom.remove(1);
    }
    while (selectTo.options.length > 1) {
        selectTo.remove(1);
    }
    
    // 着工年月のオプションを両方のセレクトに追加
    constructionState.data.start_months.forEach((month, index) => {
        const optionFrom = document.createElement('option');
        optionFrom.value = month;
        optionFrom.textContent = month;
        selectFrom.appendChild(optionFrom);
        
        const optionTo = document.createElement('option');
        optionTo.value = month;
        optionTo.textContent = month;
        selectTo.appendChild(optionTo);
        
        if (index < 5) {
            console.log(`  - オプション追加: ${month}`);
        }
    });
    
    // デフォルト値を設定（最初の月～最後の月）
    if (constructionState.data.start_months.length > 0) {
        selectFrom.value = constructionState.data.start_months[0];
        selectTo.value = constructionState.data.start_months[constructionState.data.start_months.length - 1];
        console.log(`  - デフォルト期間設定: ${selectFrom.value} ～ ${selectTo.value}`);
    }
    
    console.log(`✓ 着工年月プルダウン初期化完了: ${selectFrom.options.length - 1}個のオプション追加`);
}

// 建築計画マーカーを表示
function showConstructionMarkers(completionMonthFrom, completionMonthTo, startMonthFrom, startMonthTo) {
    // 既存のマーカーをクリア
    clearConstructionMarkers();
    
    if (!constructionState.data) return;
    if (!completionMonthFrom && !completionMonthTo && !startMonthFrom && !startMonthTo) return;
    
    // 指定された条件でプロジェクトをフィルタ
    let projects = constructionState.data.projects;
    
    // 完成年月でフィルタ（期間指定）
    if (completionMonthFrom || completionMonthTo) {
        projects = projects.filter(p => {
            const completionMonth = p.完成年月;
            if (!completionMonth) return false;
            
            // 開始月が指定されている場合、それ以降であることを確認
            if (completionMonthFrom && completionMonth < completionMonthFrom) {
                return false;
            }
            
            // 終了月が指定されている場合、それ以前であることを確認
            if (completionMonthTo && completionMonth > completionMonthTo) {
                return false;
            }
            
            return true;
        });
    }
    
    // 着工年月でフィルタ（期間指定）
    if (startMonthFrom || startMonthTo) {
        projects = projects.filter(p => {
            const startMonth = p.着工年月;
            if (!startMonth) return false;
            
            // 開始月が指定されている場合、それ以降であることを確認
            if (startMonthFrom && startMonth < startMonthFrom) {
                return false;
            }
            
            // 終了月が指定されている場合、それ以前であることを確認
            if (startMonthTo && startMonth > startMonthTo) {
                return false;
            }
            
            return true;
        });
    }
    
    const completionPeriodText = completionMonthFrom && completionMonthTo 
        ? `${completionMonthFrom}～${completionMonthTo}` 
        : completionMonthFrom 
        ? `${completionMonthFrom}以降` 
        : completionMonthTo 
        ? `${completionMonthTo}以前` 
        : 'すべて';
    
    const startPeriodText = startMonthFrom && startMonthTo 
        ? `${startMonthFrom}～${startMonthTo}` 
        : startMonthFrom 
        ? `${startMonthFrom}以降` 
        : startMonthTo 
        ? `${startMonthTo}以前` 
        : 'すべて';
    
    console.log(`完成年月: ${completionPeriodText}, 着工年月: ${startPeriodText}: ${projects.length}件の建築計画を表示`);
    
    // マーカーを作成
    projects.forEach(project => {
        // カスタムアイコン
        const icon = L.divIcon({
            className: 'construction-marker',
            html: '<div class="construction-marker-inner"></div>',
            iconSize: [12, 12],
            iconAnchor: [6, 6],
            popupAnchor: [0, -6]
        });
        
        // マーカー作成
        const marker = L.marker([project.latitude, project.longitude], { icon: icon });
        
        // プロジェクトデータを保存
        marker.project = project;
        
        // ポップアップ内容
        const popupContent = `
            <div class="construction-popup">
                <h3>${project.name}</h3>
                <table>
                    <tr><th>住所</th><td>${project.address || '-'}</td></tr>
                    <tr><th>用途</th><td>${project.usage || '-'}</td></tr>
                    <tr><th>工事種別</th><td>${project.structure || '-'}</td></tr>
                    <tr><th>地上階</th><td>${project.floors || '-'}</td></tr>
                    <tr><th>延床面積</th><td>${project.area || '-'}</td></tr>
                    <tr><th>建築主</th><td>${project.owner || '-'}</td></tr>
                    <tr><th>設計者</th><td>${project.designer || '-'}</td></tr>
                    <tr><th>施工者</th><td>${project.constructor || '-'}</td></tr>
                    <tr><th>着工日</th><td>${project.start_date || '-'}</td></tr>
                    <tr><th>完成日</th><td>${project.completion_date || '-'}</td></tr>
                </table>
            </div>
        `;
        
        marker.bindPopup(popupContent);
        marker.addTo(state.map);
        
        constructionState.markers.push(marker);
    });
    
    // ポリゴンレイヤーを最前面に移動
    bringDrawnItemsToFront();
}

// 建築計画マーカーをクリア
function clearConstructionMarkers() {
    constructionState.markers.forEach(marker => {
        state.map.removeLayer(marker);
    });
    constructionState.markers = [];
}

// 建築計画表示を切り替え
function toggleConstructionDisplay() {
    const checkbox = document.getElementById('show-construction-checkbox');
    const completionMonthFrom = document.getElementById('completion-month-from');
    const completionMonthTo = document.getElementById('completion-month-to');
    const startMonthFrom = document.getElementById('start-month-from');
    const startMonthTo = document.getElementById('start-month-to');
    
    console.log('建築計画表示切り替え:');
    console.log(`  - チェックボックス: ${checkbox.checked}`);
    console.log(`  - 選択された完成年月（開始）: ${completionMonthFrom.value}`);
    console.log(`  - 選択された完成年月（終了）: ${completionMonthTo.value}`);
    console.log(`  - 選択された着工年月（開始）: ${startMonthFrom.value}`);
    console.log(`  - 選択された着工年月（終了）: ${startMonthTo.value}`);
    
    if (checkbox.checked && (completionMonthFrom.value || completionMonthTo.value || startMonthFrom.value || startMonthTo.value)) {
        showConstructionMarkers(completionMonthFrom.value, completionMonthTo.value, startMonthFrom.value, startMonthTo.value);
        constructionState.isVisible = true;
    } else {
        clearConstructionMarkers();
        constructionState.isVisible = false;
    }
}

// 建築計画機能のイベントリスナー設定
function setupConstructionEventListeners() {
    const checkbox = document.getElementById('show-construction-checkbox');
    const completionMonthFrom = document.getElementById('completion-month-from');
    const completionMonthTo = document.getElementById('completion-month-to');
    const startMonthFrom = document.getElementById('start-month-from');
    const startMonthTo = document.getElementById('start-month-to');
    
    if (checkbox) {
        checkbox.addEventListener('change', toggleConstructionDisplay);
    }
    
    if (completionMonthFrom) {
        completionMonthFrom.addEventListener('change', () => {
            if (checkbox.checked) {
                toggleConstructionDisplay();
            }
        });
    }
    
    if (completionMonthTo) {
        completionMonthTo.addEventListener('change', () => {
            if (checkbox.checked) {
                toggleConstructionDisplay();
            }
        });
    }
    
    if (startMonthFrom) {
        startMonthFrom.addEventListener('change', () => {
            if (checkbox.checked) {
                toggleConstructionDisplay();
            }
        });
    }
    
    if (startMonthTo) {
        startMonthTo.addEventListener('change', () => {
            if (checkbox.checked) {
                toggleConstructionDisplay();
            }
        });
    }
}

// ============================================================================
// 任意ポイント登録機能
// ============================================================================

const customPointsState = {
    points: [], // 登録済みポイント配列 { id, name, address, category, salesPerson, craftsman, lat, lng, marker, isVisible }
    nextId: 1,
    isVisible: true,
    editingId: null, // 編集中ID
    clickModeActive: false, // 地図クリック追加モードが有効かどうか
    tempMarker: null // 一時的なマーカー（クリック位置を示す）
};

/**
 * 住所から座標を取得（国土地理院ジオコーディング）
 */
async function geocodeAddress(address) {
    try {
        // 国土地理院ジオコーディングAPI
        const url = `https://msearch.gsi.go.jp/address-search/AddressSearch?q=${encodeURIComponent(address)}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error('ジオコーディングAPIエラー');
        }
        
        const data = await response.json();
        
        if (!data || data.length === 0) {
            throw new Error('住所が見つかりませんでした');
        }
        
        // 最初の結果を使用
        const result = data[0];
        const lat = parseFloat(result.geometry.coordinates[1]);
        const lng = parseFloat(result.geometry.coordinates[0]);
        
        return { lat, lng, fullAddress: result.properties.title };
    } catch (error) {
        console.error('ジオコーディングエラー:', error);
        throw error;
    }
}

/**
 * 座標から住所を取得（逆ジオコーディング）
 */
async function reverseGeocode(lat, lng) {
    try {
        // 国土地理院の逆ジオコーディングAPIは公式にはないため、
        // 近似的な住所を生成（緯度経度を表示）
        // より正確な住所が必要な場合は、別のサービス（Google Maps APIなど）の利用を検討
        return `緯度: ${lat.toFixed(6)}, 経度: ${lng.toFixed(6)}`;
    } catch (error) {
        console.error('逆ジオコーディングエラー:', error);
        return `緯度: ${lat.toFixed(6)}, 経度: ${lng.toFixed(6)}`;
    }
}

/**
 * 任意ポイントマーカーを作成
 */
function createCustomPointMarker(point) {
    const marker = L.marker([point.lat, point.lng], {
        icon: L.divIcon({
            className: 'custom-point-marker',
            html: '<div class="custom-point-marker-inner"></div>',
            iconSize: [14, 14],
            iconAnchor: [7, 7]
        })
    });
    
    // ポップアップ内容
    const title = point.name || '登録ポイント';
    let popupContent = `
        <div class="custom-point-popup">
            <h3>${title}</h3>
    `;
    
    if (point.category) {
        popupContent += `<div class="popup-detail"><span class="popup-category">${point.category}</span></div>`;
    }
    
    popupContent += `<div class="popup-detail"><span class="popup-label">住所:</span><br><span class="popup-value">${point.address}</span></div>`;
    
    if (point.salesPerson) {
        popupContent += `<div class="popup-detail"><span class="popup-label">営業担当者:</span> <span class="popup-value">${point.salesPerson}</span></div>`;
    }
    
    if (point.craftsman) {
        popupContent += `<div class="popup-detail"><span class="popup-label">担当職人:</span> <span class="popup-value">${point.craftsman}</span></div>`;
    }
    
    popupContent += `</div>`;
    
    marker.bindPopup(popupContent);
    
    return marker;
}

/**
 * 任意ポイントを地図に追加
 */
function addCustomPointToMap(point) {
    // 個別の表示設定も確認
    if (point.isVisible === false) return;
    
    const marker = createCustomPointMarker(point);
    marker.addTo(state.map);
    point.marker = marker;
}

/**
 * 任意ポイントを地図から削除
 */
function removeCustomPointFromMap(point) {
    if (point.marker) {
        state.map.removeLayer(point.marker);
        point.marker = null;
    }
}

/**
 * すべての任意ポイントマーカーを表示/非表示
 */
function toggleCustomPointsVisibility() {
    customPointsState.points.forEach(point => {
        if (customPointsState.isVisible) {
            if (!point.marker && point.isVisible !== false) {
                addCustomPointToMap(point);
            }
        } else {
            removeCustomPointFromMap(point);
        }
    });
}

/**
 * 個別ポイントの表示/非表示を切り替え
 */
function togglePointVisibility(id) {
    const point = customPointsState.points.find(p => p.id === id);
    if (!point) return;
    
    point.isVisible = !point.isVisible;
    
    if (point.isVisible && customPointsState.isVisible) {
        // 表示ONで全体も表示中の場合、マーカーを追加
        if (!point.marker) {
            addCustomPointToMap(point);
        }
    } else {
        // 表示OFFの場合、マーカーを削除
        removeCustomPointFromMap(point);
    }
    
    // LocalStorageに保存
    saveCustomPointsToStorage();
    
    // UI更新
    updateCustomPointsList();
}

/**
 * 任意ポイントを追加
 */
async function addCustomPoint(address) {
    try {
        console.log('ジオコーディング開始:', address);
        
        // 住所から座標を取得
        const { lat, lng, fullAddress } = await geocodeAddress(address);
        
        console.log('座標取得成功:', lat, lng);
        
        // ポイントオブジェクト作成
        const point = {
            id: customPointsState.nextId++,
            name: '',
            address: fullAddress || address,
            category: '',
            salesPerson: '',
            craftsman: '',
            lat,
            lng,
            marker: null,
            isVisible: true
        };
        
        // 配列に追加
        customPointsState.points.push(point);
        
        // 地図に表示
        if (customPointsState.isVisible) {
            addCustomPointToMap(point);
        }
        
        // LocalStorageに保存
        saveCustomPointsToStorage();
        
        // UI更新
        updateCustomPointsList();
        
        // 地図をポイントに移動
        state.map.setView([lat, lng], 15);
        
        // 属性編集モーダルを開く
        openEditModal(point.id);
        
        return point;
    } catch (error) {
        console.error('ポイント追加エラー:', error);
        alert(`ポイントの追加に失敗しました:\n${error.message}`);
        throw error;
    }
}

/**
 * 任意ポイントを削除
 */
function deleteCustomPoint(id) {
    const index = customPointsState.points.findIndex(p => p.id === id);
    if (index === -1) return;
    
    const point = customPointsState.points[index];
    
    // 地図から削除
    removeCustomPointFromMap(point);
    
    // 配列から削除
    customPointsState.points.splice(index, 1);
    
    // LocalStorageに保存
    saveCustomPointsToStorage();
    
    // UI更新
    updateCustomPointsList();
}

/**
 * 地図クリック追加モードを切り替え
 */
function toggleClickMode() {
    customPointsState.clickModeActive = !customPointsState.clickModeActive;
    
    const button = document.getElementById('add-point-by-click-btn');
    if (!button) return;
    
    if (customPointsState.clickModeActive) {
        button.textContent = 'クリックモード終了';
        button.classList.add('active');
        state.map.getContainer().style.cursor = 'crosshair';
        alert('地図上の任意の位置をクリックしてポイントを追加してください');
    } else {
        button.textContent = '地図クリックで追加';
        button.classList.remove('active');
        state.map.getContainer().style.cursor = '';
        
        // 一時マーカーを削除
        if (customPointsState.tempMarker) {
            state.map.removeLayer(customPointsState.tempMarker);
            customPointsState.tempMarker = null;
        }
    }
}

/**
 * 地図クリック時にポイントを追加
 */
async function handleMapClickForPoint(e) {
    if (!customPointsState.clickModeActive) return;
    
    const lat = e.latlng.lat;
    const lng = e.latlng.lng;
    
    try {
        // 一時マーカーがあれば削除
        if (customPointsState.tempMarker) {
            state.map.removeLayer(customPointsState.tempMarker);
        }
        
        // 逆ジオコーディングで住所を取得
        const address = await reverseGeocode(lat, lng);
        
        // 新しいポイントを作成
        const point = {
            id: customPointsState.nextId++,
            name: '',
            address: address,
            category: '',
            salesPerson: '',
            craftsman: '',
            lat: lat,
            lng: lng,
            marker: null,
            isVisible: true
        };
        
        customPointsState.points.push(point);
        
        // 地図に表示
        if (customPointsState.isVisible) {
            addCustomPointToMap(point);
        }
        
        // LocalStorageに保存
        saveCustomPointsToStorage();
        
        // UI更新
        updateCustomPointsList();
        
        // クリックモードを終了
        toggleClickMode();
        
        // 属性編集モーダルを開く
        openEditModal(point.id);
        
    } catch (error) {
        console.error('地図クリックポイント追加エラー:', error);
        alert(`ポイントの追加に失敗しました:\n${error.message}`);
    }
}

/**
 * 地図クリックイベントリスナーを設定
 */
function setupMapClickListener() {
    if (state.map) {
        state.map.on('click', handleMapClickForPoint);
    }
}

/**
 * ポイントに地図を移動
 */
function locateCustomPoint(id) {
    const point = customPointsState.points.find(p => p.id === id);
    if (!point) return;
    
    state.map.setView([point.lat, point.lng], 15);
    
    // マーカーが表示されていればポップアップを開く
    if (point.marker) {
        point.marker.openPopup();
    }
}

/**
 * LocalStorageに保存
 */
function saveCustomPointsToStorage() {
    try {
        const data = customPointsState.points.map(p => ({
            id: p.id,
            name: p.name,
            address: p.address,
            category: p.category,
            salesPerson: p.salesPerson,
            craftsman: p.craftsman,
            lat: p.lat,
            lng: p.lng
        }));
        
        localStorage.setItem('customPoints', JSON.stringify(data));
        localStorage.setItem('customPointsNextId', customPointsState.nextId.toString());
    } catch (error) {
        console.error('LocalStorage保存エラー:', error);
    }
}

/**
 * LocalStorageから読み込み
 */
function loadCustomPointsFromStorage() {
    try {
        const data = localStorage.getItem('customPoints');
        const nextId = localStorage.getItem('customPointsNextId');
        
        if (data) {
            const points = JSON.parse(data);
            customPointsState.points = points.map(p => ({ ...p, marker: null }));
            
            // 地図に表示
            if (customPointsState.isVisible) {
                customPointsState.points.forEach(point => {
                    addCustomPointToMap(point);
                });
            }
        }
        
        if (nextId) {
            customPointsState.nextId = parseInt(nextId, 10);
        }
        
        // UI更新
        updateCustomPointsList();
    } catch (error) {
        console.error('LocalStorage読み込みエラー:', error);
    }
}

/**
 * ポイント一覧UIを更新
 */
function updateCustomPointsList() {
    const countElement = document.getElementById('custom-points-count');
    const itemsElement = document.getElementById('custom-points-items');
    
    if (!countElement || !itemsElement) return;
    
    // 件数更新
    countElement.textContent = customPointsState.points.length;
    
    // リストクリア
    itemsElement.innerHTML = '';
    
    // ポイントがない場合
    if (customPointsState.points.length === 0) {
        itemsElement.innerHTML = '<p style="color: #999; font-size: 14px; padding: 10px;">登録されたポイントはありません</p>';
        return;
    }
    
    // 各ポイントのアイテムを作成
    customPointsState.points.forEach(point => {
        const item = document.createElement('div');
        item.className = 'custom-point-item';
        
        const addressDiv = document.createElement('div');
        addressDiv.className = 'custom-point-item-address';
        
        let displayText = point.name || point.address;
        if (point.name && point.category) {
            displayText = `${point.name} [${point.category}]`;
        } else if (point.category) {
            displayText = `${point.address} [${point.category}]`;
        }
        
        addressDiv.textContent = displayText;
        
        const actionsDiv = document.createElement('div');
        actionsDiv.className = 'custom-point-item-actions';
        
        // 編集ボタン
        const editBtn = document.createElement('button');
        editBtn.className = 'custom-point-item-btn custom-point-edit-btn';
        editBtn.textContent = '編集';
        editBtn.onclick = () => openEditModal(point.id);
        
        // 表示/非表示ボタン
        const toggleBtn = document.createElement('button');
        toggleBtn.className = 'custom-point-item-btn custom-point-locate-btn';
        const isVisible = point.isVisible !== false;
        toggleBtn.textContent = isVisible ? '非表示' : '表示';
        toggleBtn.onclick = () => togglePointVisibility(point.id);
        
        // ズームボタン
        const zoomBtn = document.createElement('button');
        zoomBtn.className = 'custom-point-item-btn custom-point-edit-btn';
        zoomBtn.textContent = 'ズーム';
        zoomBtn.onclick = () => locateCustomPoint(point.id);
        
        // 削除ボタン
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'custom-point-item-btn custom-point-delete-btn';
        deleteBtn.textContent = '削除';
        deleteBtn.onclick = () => {
            if (confirm(`このポイントを削除しますか?\n${point.name || point.address}`)) {
                deleteCustomPoint(point.id);
            }
        };
        
        actionsDiv.appendChild(editBtn);
        actionsDiv.appendChild(toggleBtn);
        actionsDiv.appendChild(zoomBtn);
        actionsDiv.appendChild(deleteBtn);
        
        item.appendChild(addressDiv);
        item.appendChild(actionsDiv);
        
        itemsElement.appendChild(item);
    });
}

/**
 * ポイント属性編集モーダルを開く
 */
function openEditModal(pointId) {
    const point = customPointsState.points.find(p => p.id === pointId);
    if (!point) return;
    
    customPointsState.editingId = pointId;
    
    // フォームに現在の値を設定
    document.getElementById('edit-point-name').value = point.name || '';
    document.getElementById('edit-point-address').value = point.address || '';
    document.getElementById('edit-point-category').value = point.category || '';
    document.getElementById('edit-sales-person').value = point.salesPerson || '';
    document.getElementById('edit-craftsman').value = point.craftsman || '';
    
    // モーダル表示
    const modal = document.getElementById('point-edit-modal');
    if (modal) {
        modal.classList.add('show');
    }
}

/**
 * ポイント属性編集モーダルを閉じる
 */
function closeEditModal() {
    const modal = document.getElementById('point-edit-modal');
    if (modal) {
        modal.classList.remove('show');
    }
    customPointsState.editingId = null;
}

/**
 * ポイント属性を保存
 */
async function savePointAttributes() {
    const pointId = customPointsState.editingId;
    if (!pointId) return;
    
    const point = customPointsState.points.find(p => p.id === pointId);
    if (!point) return;
    
    // フォームから値を取得
    const name = document.getElementById('edit-point-name').value.trim();
    const newAddress = document.getElementById('edit-point-address').value.trim();
    const category = document.getElementById('edit-point-category').value;
    const salesPerson = document.getElementById('edit-sales-person').value.trim();
    const craftsman = document.getElementById('edit-craftsman').value.trim();
    
    // 必須項目チェック
    if (!name) {
        alert('ポイント名を入力してください');
        return;
    }
    
    if (!newAddress) {
        alert('住所を入力してください');
        return;
    }
    
    // 住所が変更されている場合は再ジオコーディング
    if (newAddress !== point.address) {
        try {
            const result = await geocodeAddress(newAddress);
            point.lat = result.lat;
            point.lng = result.lng;
            point.address = result.fullAddress || newAddress;
        } catch (error) {
            alert(`住所のジオコーディングに失敗しました:\n${error.message}\n元の位置を保持します。`);
            // エラーの場合は元の座標を保持して、住所だけ更新
            point.address = newAddress;
        }
    }
    
    // 属性更新
    point.name = name;
    point.category = category;
    point.salesPerson = salesPerson;
    point.craftsman = craftsman;
    
    // マーカーを再作成
    if (point.marker) {
        removeCustomPointFromMap(point);
        addCustomPointToMap(point);
        
        // 新しい位置に地図を移動
        state.map.setView([point.lat, point.lng], state.map.getZoom());
    }
    
    // LocalStorageに保存
    saveCustomPointsToStorage();
    
    // UI更新
    updateCustomPointsList();
    
    // モーダルを閉じる
    closeEditModal();
    
    alert('ポイント属性を保存しました');
}

/**
 * モーダルイベントリスナー設定
 */
function setupModalEventListeners() {
    const modal = document.getElementById('point-edit-modal');
    const closeBtn = document.getElementById('modal-close-btn');
    const cancelBtn = document.getElementById('modal-cancel-btn');
    const saveBtn = document.getElementById('modal-save-btn');
    
    if (!modal || !closeBtn || !cancelBtn || !saveBtn) return;
    
    // 閉じるボタン
    closeBtn.addEventListener('click', closeEditModal);
    cancelBtn.addEventListener('click', closeEditModal);
    
    // 保存ボタン
    saveBtn.addEventListener('click', savePointAttributes);
    
    // 背景クリックで閉じる
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            closeEditModal();
        }
    });
    
    // ESCキーで閉じる
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && modal.classList.contains('show')) {
            closeEditModal();
        }
    });
}

// ============================================================================
// CSV エクスポート・インポート機能
// ============================================================================

/**
 * ポイントデータをCSV形式でエクスポート
 */
function exportPointsToCSV() {
    if (customPointsState.points.length === 0) {
        alert('エクスポートするポイントがありません');
        return;
    }
    
    try {
        // CSVヘッダー
        const headers = ['ID', 'ポイント名', '住所', 'カテゴリ', '営業担当者名', '担当職人名', '緯度', '経度'];
        
        // CSVデータ行を作成
        const rows = customPointsState.points.map(point => {
            return [
                point.id,
                point.name || '',
                point.address || '',
                point.category || '',
                point.salesPerson || '',
                point.craftsman || '',
                point.lat,
                point.lng
            ].map(value => {
                // カンマや改行を含む場合はダブルクォートで囲む
                const strValue = String(value);
                if (strValue.includes(',') || strValue.includes('\n') || strValue.includes('"')) {
                    return `"${strValue.replace(/"/g, '""')}"`;
                }
                return strValue;
            }).join(',');
        });
        
        // CSVテキストを生成
        const csvContent = [headers.join(','), ...rows].join('\n');
        
        // BOM付きUTF-8でエンコード（Excelで文字化けしないように）
        const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
        const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
        
        // ダウンロード用のリンクを作成
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        
        // ファイル名（日時付き）
        const now = new Date();
        const dateStr = now.toISOString().slice(0, 19).replace(/:/g, '-').replace('T', '_');
        link.download = `登録ポイント_${dateStr}.csv`;
        
        // ダウンロード実行
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        alert(`${customPointsState.points.length}件のポイントをエクスポートしました`);
    } catch (error) {
        console.error('CSVエクスポートエラー:', error);
        alert(`エクスポートに失敗しました:\n${error.message}`);
    }
}

/**
 * CSV形式のテキストをパースしてポイント配列に変換
 */
function parseCSV(csvText) {
    const lines = csvText.split('\n').filter(line => line.trim());
    if (lines.length < 2) {
        throw new Error('CSVデータが空です');
    }
    
    // ヘッダー行をスキップ
    const dataLines = lines.slice(1);
    
    const points = [];
    let lineNumber = 2; // ヘッダーが1行目なのでデータは2行目から
    
    for (const line of dataLines) {
        try {
            // 簡易CSVパーサー（ダブルクォート対応）
            const values = [];
            let currentValue = '';
            let insideQuotes = false;
            
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                const nextChar = line[i + 1];
                
                if (char === '"') {
                    if (insideQuotes && nextChar === '"') {
                        // エスケープされたダブルクォート
                        currentValue += '"';
                        i++; // 次の文字をスキップ
                    } else {
                        // クォートの開始/終了
                        insideQuotes = !insideQuotes;
                    }
                } else if (char === ',' && !insideQuotes) {
                    // フィールド区切り
                    values.push(currentValue.trim());
                    currentValue = '';
                } else {
                    currentValue += char;
                }
            }
            values.push(currentValue.trim()); // 最後のフィールド
            
            if (values.length < 8) {
                console.warn(`${lineNumber}行目: フィールド数が不足しています（スキップ）`);
                lineNumber++;
                continue;
            }
            
            const [id, name, address, category, salesPerson, craftsman, lat, lng] = values;
            
            // 必須項目チェック
            if (!address) {
                console.warn(`${lineNumber}行目: 住所が空です（スキップ）`);
                lineNumber++;
                continue;
            }
            
            const latNum = parseFloat(lat);
            const lngNum = parseFloat(lng);
            
            if (isNaN(latNum) || isNaN(lngNum)) {
                console.warn(`${lineNumber}行目: 座標が不正です（スキップ）`);
                lineNumber++;
                continue;
            }
            
            points.push({
                id: parseInt(id, 10) || customPointsState.nextId++,
                name: name || '',
                address: address || '',
                category: category || '',
                salesPerson: salesPerson || '',
                craftsman: craftsman || '',
                lat: latNum,
                lng: lngNum,
                marker: null
            });
            
            lineNumber++;
        } catch (error) {
            console.error(`${lineNumber}行目のパースエラー:`, error);
            lineNumber++;
        }
    }
    
    return points;
}

/**
 * CSVファイルからポイントデータをインポート
 */
function importPointsFromCSV(file) {
    if (!file) {
        alert('ファイルが選択されていません');
        return;
    }
    
    if (!file.name.endsWith('.csv')) {
        alert('CSVファイルを選択してください');
        return;
    }
    
    const reader = new FileReader();
    
    reader.onload = (e) => {
        try {
            const csvText = e.target.result;
            const importedPoints = parseCSV(csvText);
            
            if (importedPoints.length === 0) {
                alert('インポート可能なデータがありませんでした');
                return;
            }
            
            // 確認ダイアログ
            const message = customPointsState.points.length > 0
                ? `${importedPoints.length}件のポイントをインポートします。\n既存の${customPointsState.points.length}件のポイントは削除されます。\n\n続行しますか？`
                : `${importedPoints.length}件のポイントをインポートします。\n\n続行しますか？`;
            
            if (!confirm(message)) {
                return;
            }
            
            // 既存のマーカーをすべて削除
            customPointsState.points.forEach(point => {
                removeCustomPointFromMap(point);
            });
            
            // ポイントを置き換え
            customPointsState.points = importedPoints;
            
            // nextIdを更新
            const maxId = Math.max(...importedPoints.map(p => p.id), 0);
            customPointsState.nextId = maxId + 1;
            
            // 地図に表示
            if (customPointsState.isVisible) {
                customPointsState.points.forEach(point => {
                    addCustomPointToMap(point);
                });
            }
            
            // LocalStorageに保存
            saveCustomPointsToStorage();
            
            // UI更新
            updateCustomPointsList();
            
            alert(`${importedPoints.length}件のポイントをインポートしました`);
            
            // 最初のポイントに地図を移動
            if (importedPoints.length > 0) {
                const firstPoint = importedPoints[0];
                state.map.setView([firstPoint.lat, firstPoint.lng], 12);
            }
            
        } catch (error) {
            console.error('CSVインポートエラー:', error);
            alert(`インポートに失敗しました:\n${error.message}`);
        }
    };
    
    reader.onerror = () => {
        alert('ファイルの読み込みに失敗しました');
    };
    
    reader.readAsText(file, 'UTF-8');
}

/**
 * CSV機能のイベントリスナー設定
 */
function setupCSVEventListeners() {
    const exportBtn = document.getElementById('export-csv-btn');
    const importBtn = document.getElementById('import-csv-btn');
    const fileInput = document.getElementById('csv-file-input');
    
    if (!exportBtn || !importBtn || !fileInput) {
        console.warn('CSV機能の要素が見つかりません');
        return;
    }
    
    // エクスポートボタン
    exportBtn.addEventListener('click', exportPointsToCSV);
    
    // インポートボタン（ファイル選択ダイアログを開く）
    importBtn.addEventListener('click', () => {
        fileInput.click();
    });
    
    // ファイル選択時
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            importPointsFromCSV(file);
        }
        // 同じファイルを再度選択できるようにリセット
        fileInput.value = '';
    });
}

/**
 * 任意ポイント機能のイベントリスナー設定
 */
function setupCustomPointsEventListeners() {
    const addressInput = document.getElementById('custom-point-address');
    const addButton = document.getElementById('add-custom-point-btn');
    const addByClickButton = document.getElementById('add-point-by-click-btn');
    const showCheckbox = document.getElementById('show-custom-points-checkbox');
    
    if (!addressInput || !addButton || !showCheckbox) {
        console.warn('任意ポイント登録要素が見つかりません');
        return;
    }
    
    // 住所から追加ボタン
    addButton.addEventListener('click', async () => {
        const address = addressInput.value.trim();
        
        if (!address) {
            alert('住所を入力してください');
            return;
        }
        
        try {
            await addCustomPoint(address);
            addressInput.value = ''; // 成功したらクリア
        } catch (error) {
            // エラーは addCustomPoint 内で処理済み
        }
    });
    
    // 地図クリックで追加ボタン
    if (addByClickButton) {
        addByClickButton.addEventListener('click', () => {
            toggleClickMode();
        });
    }
    
    // Enterキーで追加
    addressInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            addButton.click();
        }
    });
    
    // 表示切り替え
    showCheckbox.addEventListener('change', () => {
        customPointsState.isVisible = showCheckbox.checked;
        toggleCustomPointsVisibility();
    });
}

// ============================================================================
// 折りたたみ機能
// ============================================================================

/**
 * セクションの折りたたみ状態を管理
 */
const collapsibleState = {
    stats: false,           // 統計データ
    construction: false,    // 建築計画データ
    customPoints: false,    // 任意ポイント登録
    areaAnalysis: false,    // エリア分析
    pointsList: false,      // 登録済みポイント一覧
    polygonsList: false     // 保存済みポリゴン一覧
};

/**
 * 折りたたみ状態をLocalStorageに保存
 */
function saveCollapsibleState() {
    try {
        localStorage.setItem('collapsibleState', JSON.stringify(collapsibleState));
    } catch (error) {
        console.error('折りたたみ状態の保存エラー:', error);
    }
}

/**
 * 折りたたみ状態をLocalStorageから読み込み
 */
function loadCollapsibleState() {
    try {
        const saved = localStorage.getItem('collapsibleState');
        if (saved) {
            const state = JSON.parse(saved);
            Object.assign(collapsibleState, state);
        }
    } catch (error) {
        console.error('折りたたみ状態の読み込みエラー:', error);
    }
}

/**
 * セクションを折りたたむ/展開する
 */
function toggleSection(sectionElement, sectionKey) {
    if (!sectionElement) return;
    
    const isCollapsed = sectionElement.classList.contains('collapsed');
    
    if (isCollapsed) {
        // 展開
        sectionElement.classList.remove('collapsed');
        collapsibleState[sectionKey] = false;
    } else {
        // 折りたたむ
        sectionElement.classList.add('collapsed');
        collapsibleState[sectionKey] = true;
    }
    
    saveCollapsibleState();
}

/**
 * 折りたたみ機能のイベントリスナー設定
 */
function setupCollapsibleListeners() {
    // 各セクションのヘッダーにクリックイベントを設定
    const sections = [
        { selector: '[data-section="stats"]', key: 'stats' },
        { selector: '[data-section="construction"]', key: 'construction' },
        { selector: '[data-section="custom-points"]', key: 'customPoints' },
        { selector: '[data-section="points-list"]', key: 'pointsList' },
        { selector: '[data-section="area-analysis"]', key: 'areaAnalysis' },
        { selector: '[data-section="polygons-list"]', key: 'polygonsList' }
    ];
    
    sections.forEach(({ selector, key }) => {
        const section = document.querySelector(selector);
        if (!section) return;
        
        const header = section.querySelector('.section-header');
        if (!header) return;
        
        // 保存された状態を復元
        if (collapsibleState[key]) {
            section.classList.add('collapsed');
        }
        
        // クリックイベント
        header.addEventListener('click', (e) => {
            // ボタン以外の場所をクリックした場合も動作
            toggleSection(section, key);
        });
        
        // ボタンのクリックイベント（伝播を止める必要はない）
        const btn = header.querySelector('.collapse-btn');
        if (btn) {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                toggleSection(section, key);
            });
        }
    });
}

// ============================================================================
// サイドバー制御機能
// ============================================================================

/**
 * サイドバーの状態管理
 */
const sidebarState = {
    isOpen: true,
    width: 320 // デフォルト幅（px）
};

/**
 * サイドバーの状態をLocalStorageに保存
 */
function saveSidebarState() {
    try {
        localStorage.setItem('sidebarState', JSON.stringify(sidebarState));
    } catch (error) {
        console.error('サイドバー状態の保存エラー:', error);
    }
}

/**
 * サイドバーの状態をLocalStorageから読み込み
 */
function loadSidebarState() {
    try {
        const saved = localStorage.getItem('sidebarState');
        if (saved) {
            const state = JSON.parse(saved);
            Object.assign(sidebarState, state);
        }
    } catch (error) {
        console.error('サイドバー状態の読み込みエラー:', error);
    }
}

/**
 * サイドバーを開閉
 */
function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    
    sidebarState.isOpen = !sidebarState.isOpen;
    
    if (sidebarState.isOpen) {
        sidebar.classList.remove('collapsed');
    } else {
        sidebar.classList.add('collapsed');
    }
    
    saveSidebarState();
    
    // 地図のサイズを更新
    setTimeout(() => {
        if (state.map) {
            state.map.invalidateSize();
        }
    }, 300); // アニメーション完了後
}

/**
 * サイドバーの幅を設定
 */
function setSidebarWidth(width) {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    
    // 幅の制限
    const minWidth = 280;
    const maxWidth = 500;
    const constrainedWidth = Math.max(minWidth, Math.min(maxWidth, width));
    
    sidebar.style.width = `${constrainedWidth}px`;
    sidebarState.width = constrainedWidth;
    
    saveSidebarState();
    
    // 地図のサイズを更新
    if (state.map) {
        state.map.invalidateSize();
    }
}

/**
 * サイドバーリサイザーの設定
 */
function setupSidebarResizer() {
    const resizer = document.getElementById('sidebar-resizer');
    const sidebar = document.getElementById('sidebar');
    
    if (!resizer || !sidebar) return;
    
    let isResizing = false;
    let startX = 0;
    let startWidth = 0;
    
    resizer.addEventListener('mousedown', (e) => {
        isResizing = true;
        startX = e.clientX;
        startWidth = sidebar.offsetWidth;
        
        resizer.classList.add('resizing');
        document.body.style.cursor = 'ew-resize';
        document.body.style.userSelect = 'none';
        
        e.preventDefault();
    });
    
    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        
        const deltaX = e.clientX - startX;
        const newWidth = startWidth + deltaX;
        
        setSidebarWidth(newWidth);
    });
    
    document.addEventListener('mouseup', () => {
        if (isResizing) {
            isResizing = false;
            resizer.classList.remove('resizing');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// ============================================================================
// エリア分析機能
// ============================================================================

/**
 * ポリゴンレイヤーを最前面に移動するヘルパー関数
 */
function bringDrawnItemsToFront() {
    if (areaAnalysisState.drawnItems) {
        areaAnalysisState.drawnItems.bringToFront();
    }
}

// エリア分析の状態管理
const areaAnalysisState = {
    drawnItems: null,          // Leaflet.draw レイヤーグループ
    drawControl: null,         // 描画コントロール
    currentPolygon: null,      // 現在描画中/編集中のポリゴン
    projectsInPolygon: [],     // ポリゴン内のプロジェクト
    savedPolygons: [],         // 保存済みポリゴン配列
    nextId: 1,                 // 次のID
    visiblePolygonIds: new Set() // 表示中のポリゴンID
};

/**
 * 主要用途をカテゴリに分類
 */
function categorizeUsage(usage) {
    if (!usage) return 'その他';
    
    const u = usage.toLowerCase();
    
    // 住宅系
    if (u.includes('共同住宅') || u.includes('長屋') || u.includes('一戸建') || 
        u.includes('寄宿舎') || u.includes('下宿') || u.includes('アパート') || 
        u.includes('マンション')) {
        return '住宅系';
    }
    
    // 商業系
    if (u.includes('店舗') || u.includes('ホテル') || u.includes('飲食') || 
        u.includes('物販') || u.includes('百貨店') || u.includes('スーパー') ||
        u.includes('旅館') || u.includes('料理店')) {
        return '商業系';
    }
    
    // 業務系
    if (u.includes('事務所') || u.includes('銀行') || u.includes('オフィス')) {
        return '業務系';
    }
    
    // 工業系
    if (u.includes('工場') || u.includes('倉庫') || u.includes('物流') || 
        u.includes('作業所')) {
        return '工業系';
    }
    
    // 公共系
    if (u.includes('学校') || u.includes('病院') || u.includes('診療所') || 
        u.includes('福祉') || u.includes('保育') || u.includes('幼稚園') ||
        u.includes('図書館') || u.includes('美術館') || u.includes('体育館') ||
        u.includes('公民館') || u.includes('庁舎')) {
        return '公共系';
    }
    
    return 'その他';
}

/**
 * Point-in-Polygon判定（Ray Casting Algorithm）
 */
function isPointInPolygon(lat, lng, polygon) {
    const latlngs = polygon.getLatLngs()[0]; // 外側のリング
    let inside = false;
    
    for (let i = 0, j = latlngs.length - 1; i < latlngs.length; j = i++) {
        const xi = latlngs[i].lat, yi = latlngs[i].lng;
        const xj = latlngs[j].lat, yj = latlngs[j].lng;
        
        const intersect = ((yi > lng) !== (yj > lng)) &&
            (lat < (xj - xi) * (lng - yi) / (yj - yi) + xi);
        
        if (intersect) inside = !inside;
    }
    
    return inside;
}

/**
 * 延べ床面積を数値に変換（"1234.56 ㎡" → 1234.56）
 */
function parseFloorArea(areaStr) {
    if (!areaStr) return 0;
    const num = parseFloat(areaStr.toString().replace(/[^\d.]/g, ''));
    return isNaN(num) ? 0 : num;
}

/**
 * 保存済みポリゴンをLocalStorageに保存
 */
function savePolygonsToStorage() {
    try {
        const data = areaAnalysisState.savedPolygons.map(p => ({
            id: p.id,
            name: p.name,
            latlngs: p.latlngs,
            createdAt: p.createdAt
        }));
        
        localStorage.setItem('savedPolygons', JSON.stringify(data));
        localStorage.setItem('savedPolygonsNextId', areaAnalysisState.nextId.toString());
        console.log(`✓ ポリゴンを保存しました (${data.length}件)`);
    } catch (error) {
        console.error('ポリゴン保存エラー:', error);
    }
}

/**
 * 保存済みポリゴンをLocalStorageから読み込み
 */
function loadPolygonsFromStorage() {
    try {
        const data = localStorage.getItem('savedPolygons');
        const nextId = localStorage.getItem('savedPolygonsNextId');
        
        if (data) {
            areaAnalysisState.savedPolygons = JSON.parse(data).map(p => ({ ...p, layer: null }));
            
            // 初期表示状態：全て表示
            areaAnalysisState.savedPolygons.forEach(polygon => {
                areaAnalysisState.visiblePolygonIds.add(polygon.id);
                addPolygonToMap(polygon);
            });
            
            console.log(`✓ 保存済みポリゴンを復元しました (${areaAnalysisState.savedPolygons.length}件)`);
        }
        
        if (nextId) {
            areaAnalysisState.nextId = parseInt(nextId, 10);
        }
        
        // UI更新
        updateSavedPolygonsList();
    } catch (error) {
        console.error('ポリゴン読み込みエラー:', error);
    }
}

/**
 * ポリゴンを地図に追加
 */
function addPolygonToMap(polygonData) {
    if (!polygonData || !areaAnalysisState.drawnItems) return;
    
    // 既に追加済みの場合はスキップ
    if (polygonData.layer) return;
    
    const latlngs = polygonData.latlngs.map(ll => [ll.lat, ll.lng]);
    const polygon = L.polygon(latlngs, {
        color: '#0078ff',
        fillColor: '#0078ff',
        fillOpacity: 0.2,
        weight: 2
    });
    
    // ポリゴンにデータを付与
    polygon.polygonId = polygonData.id;
    
    // ポップアップを追加
    polygon.bindPopup(`<b>${polygonData.name || 'ポリゴン #' + polygonData.id}</b><br>登録日: ${new Date(polygonData.createdAt).toLocaleString('ja-JP')}<br><small>クリックで集計表示</small>`);
    
    // クリックイベントを追加
    polygon.on('click', function(e) {
        // ポップアップを開く
        polygon.openPopup();
        
        // 分析実行
        const analysisData = analyzeProjectsInPolygon(polygon);
        displayAnalysisResults(analysisData);
        
        console.log(`✓ ポリゴン「${polygonData.name}」をクリック - 集計を表示しました`);
        
        // イベントの伝播を止める（地図のクリックイベントが発火しないように）
        L.DomEvent.stopPropagation(e);
    });
    
    // マウスオーバーで強調表示
    polygon.on('mouseover', function(e) {
        polygon.setStyle({
            fillOpacity: 0.4,
            weight: 3
        });
    });
    
    // マウスアウトで元に戻す
    polygon.on('mouseout', function(e) {
        polygon.setStyle({
            fillOpacity: 0.2,
            weight: 2
        });
    });
    
    areaAnalysisState.drawnItems.addLayer(polygon);
    polygonData.layer = polygon;
    
    // 最前面に移動
    bringDrawnItemsToFront();
}

/**
 * ポリゴンを地図から削除
 */
function removePolygonFromMap(polygonData) {
    if (!polygonData || !polygonData.layer) return;
    
    areaAnalysisState.drawnItems.removeLayer(polygonData.layer);
    polygonData.layer = null;
}

/**
 * 新しいポリゴンを保存
 */
function saveNewPolygon() {
    if (!areaAnalysisState.currentPolygon) {
        alert('保存するポリゴンがありません。まずポリゴンを描画してください。');
        return;
    }
    
    const nameInput = document.getElementById('polygon-name');
    const name = nameInput.value.trim() || `エリア ${areaAnalysisState.nextId}`;
    
    const latlngs = areaAnalysisState.currentPolygon.getLatLngs()[0];
    const polygonData = {
        id: areaAnalysisState.nextId++,
        name: name,
        latlngs: latlngs.map(ll => ({ lat: ll.lat, lng: ll.lng })),
        createdAt: new Date().toISOString(),
        layer: null
    };
    
    // 保存済みリストに追加
    areaAnalysisState.savedPolygons.push(polygonData);
    areaAnalysisState.visiblePolygonIds.add(polygonData.id);
    
    // 描画中のポリゴンを削除
    areaAnalysisState.drawnItems.removeLayer(areaAnalysisState.currentPolygon);
    areaAnalysisState.currentPolygon = null;
    
    // 地図に追加
    addPolygonToMap(polygonData);
    
    // 保存
    savePolygonsToStorage();
    
    // UI更新
    updateSavedPolygonsList();
    nameInput.value = '';
    
    // ボタンを戻す
    document.getElementById('save-polygon-btn').style.display = 'none';
    document.getElementById('start-polygon-btn').style.display = 'inline-block';
    
    // 分析結果を非表示
    document.getElementById('analysis-results').style.display = 'none';
    
    alert(`ポリゴン「${name}」を保存しました！`);
}

/**
 * 保存済みポリゴンを削除
 */
function deleteSavedPolygon(id) {
    const index = areaAnalysisState.savedPolygons.findIndex(p => p.id === id);
    if (index === -1) return;
    
    const polygon = areaAnalysisState.savedPolygons[index];
    
    // 地図から削除
    removePolygonFromMap(polygon);
    
    // 配列から削除
    areaAnalysisState.savedPolygons.splice(index, 1);
    areaAnalysisState.visiblePolygonIds.delete(id);
    
    // 保存
    savePolygonsToStorage();
    
    // UI更新
    updateSavedPolygonsList();
}

/**
 * ポリゴンの表示/非表示を切り替え
 */
function togglePolygonVisibility(id) {
    const polygon = areaAnalysisState.savedPolygons.find(p => p.id === id);
    if (!polygon) return;
    
    if (areaAnalysisState.visiblePolygonIds.has(id)) {
        // 非表示に
        removePolygonFromMap(polygon);
        areaAnalysisState.visiblePolygonIds.delete(id);
    } else {
        // 表示に
        addPolygonToMap(polygon);
        areaAnalysisState.visiblePolygonIds.add(id);
    }
    
    // UI更新
    updateSavedPolygonsList();
}

/**
 * ポリゴンにズーム
 */
function zoomToPolygon(id) {
    const polygon = areaAnalysisState.savedPolygons.find(p => p.id === id);
    if (!polygon) return;
    
    // 表示されていない場合は表示
    if (!areaAnalysisState.visiblePolygonIds.has(id)) {
        addPolygonToMap(polygon);
        areaAnalysisState.visiblePolygonIds.add(id);
        updateSavedPolygonsList();
    }
    
    if (polygon.layer) {
        state.map.fitBounds(polygon.layer.getBounds());
        polygon.layer.openPopup();
        
        // 分析実行
        const analysisData = analyzeProjectsInPolygon(polygon.layer);
        displayAnalysisResults(analysisData);
    }
}

/**
 * ポリゴン名編集モーダルを開く
 */
function openPolygonEditModal(id) {
    const polygon = areaAnalysisState.savedPolygons.find(p => p.id === id);
    if (!polygon) return;
    
    const modal = document.getElementById('polygon-edit-modal');
    const nameInput = document.getElementById('edit-polygon-name');
    
    if (!modal || !nameInput) {
        console.error('モーダル要素が見つかりません', { modal, nameInput });
        return;
    }
    
    // 現在の値を設定
    nameInput.value = polygon.name;
    nameInput.dataset.polygonId = id;
    
    // モーダルを表示
    modal.style.display = 'block';
    
    // 少し遅延させてからフォーカス
    setTimeout(() => {
        nameInput.focus();
        nameInput.select(); // テキストを選択状態にする
        
        // IMEを明示的に有効化（Chromeなど）
        nameInput.style.imeMode = 'active';
        nameInput.setAttribute('lang', 'ja');
    }, 100);
    
    console.log('✓ ポリゴン編集モーダルを開きました', { id, name: polygon.name });
}

/**
 * ポリゴン名編集モーダルを閉じる
 */
function closePolygonEditModal() {
    const modal = document.getElementById('polygon-edit-modal');
    if (modal) {
        modal.style.display = 'none';
    }
}

/**
 * ポリゴン名を保存
 */
function savePolygonName() {
    const nameInput = document.getElementById('edit-polygon-name');
    if (!nameInput) {
        console.error('入力フィールドが見つかりません');
        return;
    }
    
    const id = parseInt(nameInput.dataset.polygonId, 10);
    const newName = nameInput.value.trim();
    
    console.log('保存処理開始', { id, newName, rawValue: nameInput.value });
    
    if (!newName) {
        alert('ポリゴン名を入力してください');
        return;
    }
    
    const polygon = areaAnalysisState.savedPolygons.find(p => p.id === id);
    if (!polygon) {
        console.error('ポリゴンが見つかりません', id);
        return;
    }
    
    // 名前を更新
    polygon.name = newName;
    
    // レイヤーのポップアップも更新
    if (polygon.layer) {
        polygon.layer.setPopupContent(`<b>${polygon.name}</b><br>登録日: ${new Date(polygon.createdAt).toLocaleString('ja-JP')}`);
    }
    
    // 保存
    savePolygonsToStorage();
    
    // UI更新
    updateSavedPolygonsList();
    
    // モーダルを閉じる
    closePolygonEditModal();
    
    console.log('✓ ポリゴン名を保存しました', { id, newName });
}

/**
 * ポリゴン名編集モーダルのイベントリスナー設定
 */
function setupPolygonEditModalListeners() {
    const modal = document.getElementById('polygon-edit-modal');
    const closeBtn = document.getElementById('polygon-modal-close-btn');
    const cancelBtn = document.getElementById('polygon-modal-cancel-btn');
    const saveBtn = document.getElementById('polygon-modal-save-btn');
    const nameInput = document.getElementById('edit-polygon-name');
    
    if (closeBtn) {
        closeBtn.addEventListener('click', closePolygonEditModal);
    }
    
    if (cancelBtn) {
        cancelBtn.addEventListener('click', closePolygonEditModal);
    }
    
    if (saveBtn) {
        saveBtn.addEventListener('click', savePolygonName);
    }
    
    // モーダル外クリックで閉じる
    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                closePolygonEditModal();
            }
        });
    }
}

/**
 * ポリゴンにズーム
 */
function zoomToPolygonOld(id) {
    const polygon = areaAnalysisState.savedPolygons.find(p => p.id === id);
    if (!polygon) return;
    
    // 表示されていない場合は表示
    if (!areaAnalysisState.visiblePolygonIds.has(id)) {
        addPolygonToMap(polygon);
        areaAnalysisState.visiblePolygonIds.add(id);
        updateSavedPolygonsList();
    }
    
    if (polygon.layer) {
        state.map.fitBounds(polygon.layer.getBounds());
        polygon.layer.openPopup();
        
        // 分析実行
        const analysisData = analyzeProjectsInPolygon(polygon.layer);
        displayAnalysisResults(analysisData);
    }
}

/**
 * 保存済みポリゴン一覧UIを更新
 */
function updateSavedPolygonsList() {
    const countElement = document.getElementById('saved-polygons-count');
    const itemsElement = document.getElementById('saved-polygons-items');
    
    if (!countElement || !itemsElement) return;
    
    // 件数更新
    countElement.textContent = areaAnalysisState.savedPolygons.length;
    
    // リストクリア
    itemsElement.innerHTML = '';
    
    // ポリゴンがない場合
    if (areaAnalysisState.savedPolygons.length === 0) {
        itemsElement.innerHTML = '<p style="color: #999; font-size: 14px; padding: 10px;">保存されたポリゴンはありません</p>';
        return;
    }
    
    // 各ポリゴンのアイテムを作成
    areaAnalysisState.savedPolygons.forEach(polygon => {
        const item = document.createElement('div');
        item.className = 'custom-point-item'; // 同じスタイルを再利用
        
        const nameDiv = document.createElement('div');
        nameDiv.className = 'custom-point-item-address';
        nameDiv.textContent = polygon.name;
        
        const actionsDiv = document.createElement('div');
        actionsDiv.className = 'custom-point-item-actions';
        
        // 編集ボタン
        const editBtn = document.createElement('button');
        editBtn.className = 'custom-point-item-btn custom-point-edit-btn';
        editBtn.textContent = '編集';
        editBtn.onclick = () => openPolygonEditModal(polygon.id);
        
        // 表示/非表示ボタン
        const toggleBtn = document.createElement('button');
        toggleBtn.className = 'custom-point-item-btn custom-point-locate-btn';
        const isVisible = areaAnalysisState.visiblePolygonIds.has(polygon.id);
        toggleBtn.textContent = isVisible ? '非表示' : '表示';
        toggleBtn.onclick = () => togglePolygonVisibility(polygon.id);
        
        // ズームボタン
        const zoomBtn = document.createElement('button');
        zoomBtn.className = 'custom-point-item-btn custom-point-edit-btn';
        zoomBtn.textContent = 'ズーム';
        zoomBtn.onclick = () => zoomToPolygon(polygon.id);
        
        // 削除ボタン
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'custom-point-item-btn custom-point-delete-btn';
        deleteBtn.textContent = '削除';
        deleteBtn.onclick = () => {
            if (confirm(`ポリゴン「${polygon.name}」を削除しますか？`)) {
                deleteSavedPolygon(polygon.id);
            }
        };
        
        actionsDiv.appendChild(editBtn);
        actionsDiv.appendChild(toggleBtn);
        actionsDiv.appendChild(zoomBtn);
        actionsDiv.appendChild(deleteBtn);
        
        item.appendChild(nameDiv);
        item.appendChild(actionsDiv);
        
        itemsElement.appendChild(item);
    });
}

/**
 * ポリゴンをLocalStorageに保存（旧版、下位互換用に残す）
 */
function savePolygonToStorage() {
    // 何もしない（新しいsavePolygonsToStorageを使用）
}

/**
 * ポリゴンをLocalStorageから読み込み（旧版、下位互換用に残す）
 */
function loadPolygonFromStorage() {
    // 何もしない（新しいloadPolygonsFromStorageを使用）
}

/**
 * ポリゴン内のプロジェクトを分析
 */
function analyzeProjectsInPolygon(polygon) {
    // 表示中のマーカーがない場合は、全データを対象にする（後方互換性のため）
    if (!constructionState.markers || constructionState.markers.length === 0) {
        // マーカーが表示されていない場合は全データを対象
        if (!constructionState.data || !constructionState.data.projects) {
            return {
                count: 0,
                totalArea: 0,
                avgArea: 0,
                usageBreakdown: {},
                constructionTypeBreakdown: {}
            };
        }
        
        const projects = constructionState.data.projects;
        const projectsInside = [];
        
        // ポリゴン内の点を抽出
        for (const project of projects) {
            if (project.緯度 && project.経度) {
                if (isPointInPolygon(project.緯度, project.経度, polygon)) {
                    projectsInside.push(project);
                }
            }
        }
        
        return analyzeProjects(projectsInside);
    }
    
    // 表示中のマーカーに対応するプロジェクトのみを対象に集計
    const projectsInside = [];
    
    constructionState.markers.forEach(marker => {
        const project = marker.project;
        if (!project || !project.緯度 || !project.経度) return;
        
        if (isPointInPolygon(project.緯度, project.経度, polygon)) {
            projectsInside.push(project);
        }
    });
    
    console.log(`ポリゴン内分析: 表示中${constructionState.markers.length}件のうち${projectsInside.length}件がポリゴン内`);
    
    return analyzeProjects(projectsInside);
}

/**
 * プロジェクトデータを集計
 */
function analyzeProjects(projectsInside) {
    const totalCount = projectsInside.length;
    
    if (totalCount === 0) {
        return {
            count: 0,
            totalArea: 0,
            avgArea: 0,
            usageBreakdown: {},
            constructionTypeBreakdown: {}
        };
    }
    
    // 延べ床面積の集計
    const areas = projectsInside.map(p => parseFloorArea(p.延床面積));
    const totalArea = areas.reduce((sum, a) => sum + a, 0);
    const avgArea = totalArea / totalCount;
    
    // 用途別集計（カテゴリ分類）
    const usageBreakdown = {};
    projectsInside.forEach(p => {
        const category = categorizeUsage(p.主要用途);
        usageBreakdown[category] = (usageBreakdown[category] || 0) + 1;
    });
    
    // 工事種別集計
    const constructionTypeBreakdown = {};
    projectsInside.forEach(p => {
        const type = p.工事種別 || 'その他';
        constructionTypeBreakdown[type] = (constructionTypeBreakdown[type] || 0) + 1;
    });
    
    return {
        count: totalCount,
        totalArea,
        avgArea,
        usageBreakdown,
        constructionTypeBreakdown,
        projects: projectsInside
    };
}

/**
 * 分析結果をUIに表示
 */
function displayAnalysisResults(analysisData) {
    const resultsDiv = document.getElementById('analysis-results');
    
    if (!analysisData || analysisData.count === 0) {
        // データがない場合でも表示は保持し、「-」を表示
        document.getElementById('total-count').textContent = '-';
        document.getElementById('total-area').textContent = '-';
        document.getElementById('avg-area').textContent = '-';
        document.getElementById('usage-breakdown').textContent = '-';
        document.getElementById('construction-type-breakdown').textContent = '-';
        areaAnalysisState.projectsInPolygon = [];
        return;
    }
    
    // 基本情報
    document.getElementById('total-count').textContent = `${analysisData.count.toLocaleString()}件`;
    document.getElementById('total-area').textContent = `${analysisData.totalArea.toLocaleString()}㎡`;
    document.getElementById('avg-area').textContent = `${analysisData.avgArea.toFixed(1)}㎡`;
    
    // 用途別内訳
    const usageDiv = document.getElementById('usage-breakdown');
    const sortedUsage = Object.entries(analysisData.usageBreakdown)
        .sort((a, b) => b[1] - a[1]);
    
    usageDiv.innerHTML = sortedUsage.map(([category, count]) => {
        const percentage = (count / analysisData.count * 100).toFixed(1);
        return `
            <div class="breakdown-item">
                <span class="breakdown-label">${category}</span>
                <span class="breakdown-value">${count}件 (${percentage}%)</span>
            </div>
        `;
    }).join('');
    
    // 工事種別
    const ctDiv = document.getElementById('construction-type-breakdown');
    const sortedCT = Object.entries(analysisData.constructionTypeBreakdown)
        .sort((a, b) => b[1] - a[1]);
    
    ctDiv.innerHTML = sortedCT.map(([type, count]) => {
        const percentage = (count / analysisData.count * 100).toFixed(1);
        return `
            <div class="breakdown-item">
                <span class="breakdown-label">${type}</span>
                <span class="breakdown-value">${count}件 (${percentage}%)</span>
            </div>
        `;
    }).join('');
    
    // ポリゴン内プロジェクトを保存
    areaAnalysisState.projectsInPolygon = analysisData.projects;
}

/**
 * ポリゴン描画時のイベントハンドラ
 */
function onPolygonCreated(e) {
    const layer = e.layer;
    
    // 既存の描画中ポリゴンを削除
    if (areaAnalysisState.currentPolygon) {
        areaAnalysisState.drawnItems.removeLayer(areaAnalysisState.currentPolygon);
    }
    
    // 新しいポリゴンを追加
    areaAnalysisState.drawnItems.addLayer(layer);
    areaAnalysisState.currentPolygon = layer;
    
    // ポリゴンレイヤーを最前面に持ってくる
    areaAnalysisState.drawnItems.bringToFront();
    
    // 分析実行
    const analysisData = analyzeProjectsInPolygon(layer);
    displayAnalysisResults(analysisData);
    
    // 範囲内マーカー強調（チェックボックスがONの場合）
    const highlightCheckbox = document.getElementById('highlight-in-polygon-checkbox');
    if (highlightCheckbox && highlightCheckbox.checked) {
        highlightMarkersInPolygon();
    }
    
    // 保存ボタンを表示
    const startBtn = document.getElementById('start-polygon-btn');
    const saveBtn = document.getElementById('save-polygon-btn');
    if (startBtn) startBtn.style.display = 'none';
    if (saveBtn) saveBtn.style.display = '';
    
    console.log('✓ ポリゴン描画完了 - 保存ボタンを表示しました');
}

/**
 * ポリゴン編集時のイベントハンドラ
 */
function onPolygonEdited(e) {
    if (areaAnalysisState.currentPolygon) {
        const analysisData = analyzeProjectsInPolygon(areaAnalysisState.currentPolygon);
        displayAnalysisResults(analysisData);
        
        const highlightCheckbox = document.getElementById('highlight-in-polygon-checkbox');
        if (highlightCheckbox && highlightCheckbox.checked) {
            highlightMarkersInPolygon();
        }
    }
}

/**
 * ポリゴン削除時のイベントハンドラ
 */
function onPolygonDeleted(e) {
    areaAnalysisState.currentPolygon = null;
    areaAnalysisState.projectsInPolygon = [];
    document.getElementById('analysis-results').style.display = 'none';
    
    // マーカー強調を解除
    resetMarkerHighlight();
    
    // ボタンを戻す
    const saveBtn = document.getElementById('save-polygon-btn');
    const startBtn = document.getElementById('start-polygon-btn');
    if (saveBtn) saveBtn.style.display = 'none';
    if (startBtn) startBtn.style.display = '';
}

/**
 * 範囲内マーカーを強調表示
 */
function highlightMarkersInPolygon() {
    if (!constructionState.markers || constructionState.markers.length === 0) {
        return;
    }
    
    const projectsInside = new Set(areaAnalysisState.projectsInPolygon.map(p => p.件名));
    
    constructionState.markers.forEach(marker => {
        const project = marker.project;
        
        // projectプロパティが存在しない場合はスキップ
        if (!project) {
            console.warn('マーカーにprojectプロパティがありません', marker);
            return;
        }
        
        if (projectsInside.has(project.件名)) {
            // 範囲内: 緑色のマーカー
            marker.setIcon(L.icon({
                iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-green.png',
                shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-shadow.png',
                iconSize: [25, 41],
                iconAnchor: [12, 41],
                popupAnchor: [1, -34],
                shadowSize: [41, 41]
            }));
        } else {
            // 範囲外: デフォルト（青）の半透明
            marker.setOpacity(0.3);
        }
    });
}

/**
 * マーカー強調をリセット
 */
function resetMarkerHighlight() {
    if (!constructionState.markers || constructionState.markers.length === 0) {
        return;
    }
    
    constructionState.markers.forEach(marker => {
        marker.setOpacity(1.0);
        // 元の赤い円形マーカーに戻す
        marker.setIcon(L.divIcon({
            className: 'construction-marker',
            html: '<div class="construction-marker-inner"></div>',
            iconSize: [12, 12],
            iconAnchor: [6, 6],
            popupAnchor: [0, -6]
        }));
    });
}

/**
 * エリア分析機能のセットアップ
 */
function setupAreaAnalysis() {
    // Leaflet.drawレイヤーグループを作成
    areaAnalysisState.drawnItems = new L.FeatureGroup();
    state.map.addLayer(areaAnalysisState.drawnItems);
    
    // ポリゴンレイヤーを常に最前面に表示
    areaAnalysisState.drawnItems.bringToFront();
    
    // 描画コントロールを初期化（デフォルトは非表示）
    areaAnalysisState.drawControl = new L.Control.Draw({
        draw: {
            polygon: {
                allowIntersection: false,
                shapeOptions: {
                    color: '#0078ff',
                    fillColor: '#0078ff',
                    fillOpacity: 0.2,
                    weight: 2
                }
            },
            polyline: false,
            rectangle: false,
            circle: false,
            marker: false,
            circlemarker: false
        },
        edit: {
            featureGroup: areaAnalysisState.drawnItems,
            remove: true
        }
    });
    
    // イベントリスナー
    state.map.on(L.Draw.Event.CREATED, onPolygonCreated);
    state.map.on(L.Draw.Event.EDITED, onPolygonEdited);
    state.map.on(L.Draw.Event.DELETED, onPolygonDeleted);
    
    // 地図操作後にポリゴンレイヤーを最前面に保つ
    state.map.on('zoomend moveend', bringDrawnItemsToFront);
    
    // 保存されたポリゴンを復元
    loadPolygonsFromStorage();
    
    // ボタンイベント
    const startBtn = document.getElementById('start-polygon-btn');
    const saveBtn = document.getElementById('save-polygon-btn');
    const highlightCheckbox = document.getElementById('highlight-in-polygon-checkbox');
    
    let controlAdded = false; // コントロール追加フラグ
    
    if (startBtn) {
        startBtn.addEventListener('click', () => {
            // 描画コントロールを地図に追加（初回のみ）
            if (!controlAdded) {
                state.map.addControl(areaAnalysisState.drawControl);
                controlAdded = true;
            }
            
            // 統計レイヤーを一時的に非表示（描画を邪魔しないように）
            const statsCheckbox = document.getElementById('show-stats-layer-checkbox');
            if (statsCheckbox && statsCheckbox.checked) {
                statsCheckbox.checked = false;
                toggleStatsLayer(false);
                
                // ヒント表示
                console.log('ℹ️ ポリゴン描画のため統計レイヤーを一時的に非表示にしました');
            }
            
            // ポリゴン描画を開始
            new L.Draw.Polygon(state.map, areaAnalysisState.drawControl.options.draw.polygon).enable();
            
            startBtn.textContent = '🖊️ 描画中...';
            startBtn.disabled = true;
            
            // 描画完了またはキャンセル後にボタンを戻す
            const resetButton = () => {
                startBtn.textContent = '🖊️ ポリゴン描画開始';
                startBtn.disabled = false;
            };
            
            state.map.once(L.Draw.Event.CREATED, resetButton);
            state.map.once(L.Draw.Event.DRAWSTOP, resetButton);
        });
    }
    
    if (saveBtn) {
        saveBtn.addEventListener('click', saveNewPolygon);
    }
    
    if (highlightCheckbox) {
        highlightCheckbox.addEventListener('change', (e) => {
            if (e.target.checked && areaAnalysisState.currentPolygon) {
                highlightMarkersInPolygon();
            } else {
                resetMarkerHighlight();
            }
        });
    }
}

/**
 * サイドバー制御のセットアップ
 */
function setupSidebarControls() {
    // 状態を読み込み
    loadSidebarState();
    
    // サイドバーの初期状態を設定
    const sidebar = document.getElementById('sidebar');
    if (sidebar) {
        setSidebarWidth(sidebarState.width);
        
        if (!sidebarState.isOpen) {
            sidebar.classList.add('collapsed');
        }
    }
    
    // 切替ボタン
    const toggleBtn = document.getElementById('sidebar-toggle-btn');
    if (toggleBtn) {
        toggleBtn.addEventListener('click', toggleSidebar);
    }
    
    // リサイザー
    setupSidebarResizer();
}

// ============================================================================
// エントリーポイント
// ============================================================================

window.addEventListener('DOMContentLoaded', async () => {
    await initialize();
    
    // 建築計画データを読み込み
    await loadConstructionData();
    
    // 建築計画機能のイベントリスナー設定
    setupConstructionEventListeners();
    
    // 任意ポイント機能の初期化
    loadCustomPointsFromStorage();
    setupCustomPointsEventListeners();
    setupModalEventListeners();
    setupCSVEventListeners();
    setupMapClickListener(); // 地図クリック追加機能
    
    // エリア分析機能の初期化
    setupAreaAnalysis();
    
    // ポリゴン編集モーダルの初期化
    setupPolygonEditModalListeners();
    
    // 折りたたみ機能の初期化
    loadCollapsibleState();
    setupCollapsibleListeners();
    
    // サイドバー機能の初期化
    setupSidebarControls();
});
