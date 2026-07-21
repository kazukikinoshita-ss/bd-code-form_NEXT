/**
 * @fileoverview 資産化比率設定ツール（UI検証用モック） バックエンド処理
 * 既存3ツールの統合デザインシステムおよび通信仕様に完全準拠したモックロジック
 * 
 * @version 1.3.1
 */

/**
 * グローバル設定オブジェクト
 * ハードコードを排除し、保守性を高める設定定義
 */
const ASSETIZATION_CONFIG = {
  APP_TITLE: '資産化比率設定ツール',
  DEFAULT_START_MONTH: '2026-07',
  DEFAULT_END_MONTH: '2027-01',
  TIMEZONE: 'JST',
  DATE_FORMAT: 'yyyy/MM/dd HH:mm:ss'
};

/**
 * モック用の初期表示データ（BU・MPコード・PJマスタおよび初期比率データ）を取得します。
 * 
 * @returns {Object} 初期データオブジェクト
 */
function getMockInitialData() {
  try {
    // 1. BUマスタ（モック）
    const buList = ['すべて', '第一開発部', '第二開発部', '基盤運用部', 'デジタル推進部'];

    // 2. MPコードマスタ（モック）
    const mpList = [
      { code: 'MP-10023', name: '基幹系システムクラウド移行', bu: '第一開発部' },
      { code: 'MP-10024', name: '全社Web UI刷新プロジェクト', bu: '第一開発部' },
      { code: 'MP-20051', name: '次世代ECプラットフォーム構築', bu: '第二開発部' },
      { code: 'MP-30012', name: '社内インフラセキュリティ強化', bu: '基盤運用部' },
      { code: 'MP-40089', name: 'データ分析基盤（DWH）構築', bu: 'デジタル推進部' }
    ];

    // 3. プロジェクトマスタ（モック）
    const projectList = [
      { code: 'PJ-001', name: '共通基盤開発' },
      { code: 'PJ-002', name: 'フロントエンド実装' },
      { code: 'PJ-003', name: 'バックエンドAPI作成' },
      { code: 'PJ-004', name: 'データベース設計・移行' },
      { code: 'PJ-005', name: 'インフラ構築・CI/CD' }
    ];

    // 4. 初期サンプル比率設定（マトリクスデータ構造）
    const initialRatios = {
      'MP-10023': [
        {
          projectCode: 'PJ-001',
          projectName: '共通基盤開発',
          monthlyRatios: {
            '2026/07': 60,
            '2026/08': 50,
            '2026/09': 40,
            '2026/10': 30,
            '2026/11': 20,
            '2026/12': 10,
            '2027/01': 0
          }
        },
        {
          projectCode: 'PJ-002',
          projectName: 'フロントエンド実装',
          monthlyRatios: {
            '2026/07': 40,
            '2026/08': 50,
            '2026/09': 60,
            '2026/10': 70,
            '2026/11': 80,
            '2026/12': 90,
            '2027/01': 100
          }
        }
      ]
    };

    return {
      success: true,
      config: ASSETIZATION_CONFIG,
      buList: buList,
      mpList: mpList,
      projectList: projectList,
      initialRatios: initialRatios
    };
  } catch (error) {
    console.error('getMockInitialData Error:', error);
    return {
      success: false,
      message: '初期設定・マスタデータの読み込みに失敗しました: ' + error.toString()
    };
  }
}

/**
 * 画面で編集されたマトリクスデータを専用シート保存用の「縦持ちレコード構造（1行1レコード）」に
 * 自動分解・変換する処理の検証シミュレーションを行います。
 * 
 * @param {Object} data 画面から送信されたカード形式のマトリクスデータ
 * @returns {Object} 処理結果および分解後の縦持ちレコード一覧
 */
function saveAssetizationRatioMock(data) {
  try {
    const activeUser = Session.getActiveUser().getEmail() || 'kazuki.kinoshita@example.com';
    const nowStr = Utilities.formatDate(new Date(), ASSETIZATION_CONFIG.TIMEZONE, ASSETIZATION_CONFIG.DATE_FORMAT);
    const verticalRecords = [];

    if (!data || !data.cards || data.cards.length === 0) {
      throw new Error('保存対象のデータが存在しません。');
    }

    data.cards.forEach(card => {
      const mpCode = card.mpCode;
      const mpName = card.mpName;
      const bu = card.bu;

      card.projects.forEach(pj => {
        Object.keys(pj.monthlyRatios).forEach(yearMonth => {
          const ratio = Number(pj.monthlyRatios[yearMonth]) || 0;
          
          // 将来的に専用スプレッドシート（縦持ち形式）へ書き込まれる1レコードの構造
          verticalRecords.push({
            yearMonth: yearMonth,          // 対象年月 (YYYY/MM)
            mpCode: mpCode,                // MPコード
            mpName: mpName,                // MPコード名称
            bu: bu,                        // BU (事業部)
            projectCode: pj.projectCode,   // プロジェクトコード
            projectName: pj.projectName,   // プロジェクト名
            ratio: ratio,                  // 配分比率 (%)
            updatedBy: activeUser,         // 最終更新者
            updatedAt: nowStr              // 最終更新日時
          });
        });
      });
    });

    return {
      success: true,
      message: `【一括更新シミュレーション完了】\n合計 ${data.cards.length} 件のMPコード、全 ${verticalRecords.length} 件の月別縦持ちレコードに正常変換・構造化されました。`,
      recordCount: verticalRecords.length,
      sampleRecords: verticalRecords
    };
  } catch (error) {
    console.error('saveAssetizationRatioMock Error:', error);
    return {
      success: false,
      message: '一括更新処理中にエラーが発生しました: ' + error.toString()
    };
  }
}
