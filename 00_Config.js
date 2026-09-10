/**
 * =============================================================================
 * 近海_請求書：作業リスト連動
 * =============================================================================
 *
 * 【ファイル分担】
 * 00_Config.js            … シート名・行・ヘッダー別名だけ。仕様変更はまずここ。
 * 01_Utils.js             … 文字列正規化・重複排除・順番ソート・内部書き込みガード。
 * 02_WorkListRepository.js… 作業リストのヘッダー解決とレコード化。
 * 05_Triggers.js          … onOpen / onEdit / メニュー。
 * 07_WebApp.js            … 入力 Web アプリと印刷シート作成の呼び出し。
 * 08_ServiceInfo.js       … 整備情報マスタ。
 * 09_InvoiceTemplate.js   … 明細テンプレートシートの読込とサンプル作成。
 * 10_ListRefresh.js       … リストの並べ替え（図形ボタン refreshAllMasterLists）。
 * 12_InvoiceSave.js       … 請求書の保存・呼び出しと印刷PDF。
 * 13_InvoiceSearch.js     … シート「請求書検索」（K-No・一覧・表示）。
 *
 * 【中項目】
 * 入力アプリが作業リストから候補を出し、選ぶと作業内容・部品を展開する。
 * シート「車検_入力」「中項目候補」「中項目_参照」は使わない。
 *
 * 【作業リストの列】
 * A 大項目 / B 中項目 / C 作業内容 / D 技術料 / E 作業コード / F 部品_中項目 / G 単価 / H 数量 / （他） / 最後が順番。
 * 作業コードは作業者リストと同じ数値コード。氏名と作業内容が一致すると空欄だけ自動入力。
 * E と部品_中項目はリスト選択＋手入力可（既存値は消さない）。
 *
 * 【順番列】
 * 「順番」または「表示順」。大項目の塊 → 中項目の塊。
 * 中項目の中は 3 レーン（部品セットが増えても番号がぶつからない）:
 *   0 先頭行（1 行だけ。作業内容が空。無ければ順番 1 の空行）
 *   1 作業内容（順番昇順）
 *   2 追加の部品セット（作業内容が空で部品あり。順番昇順）
 * 空の順番は更新ボタンのときだけシートへ書く。更新忘れでも入力アプリは同じ規則で仮の順番を使う。
 * 手入力は上書きしない。更新は図形ボタン（refreshAllMasterLists）またはメニュー。
 * 技術の中項目を選ぶと、同じ中項目の部品（作業内容の有無を問わず）を順番昇順で部品側へ出す。
 * =============================================================================
 */

const CONFIG = {
  /**
   * false にすると Logger と Cache ガードを省いて応答を速くする。
   * インストール型 onEdit を使ってループする場合だけ useWriteGuard を true。
   */
  verboseLog: false,
  useWriteGuard: false,

  /** マスタ（縦持ち 1 行 = 1 レコード） */
  workList: {
    sheetName: '作業リスト',
    headerRow: 1,
    /**
     * 論理名 → 実際のヘッダー候補（左から優先）
     * 実シートは「大項目」「中項目」。仕様書表記の「作業_大項目」なども許容する。
     */
    headers: {
      major: ['作業_大項目', '大項目', '作業内容_大項目'],
      mid: ['作業_中項目', '中項目'],
      content: ['作業内容'],
      fee: ['技術料'],
      workerCode: ['作業コード', '作業者コード'],
      order: ['順番', '表示順'],
      partMajor: ['部品_大項目'],
      partMid: ['部品_中項目', '部品名'],
      qty: ['数量'],
      unitPrice: ['単価']
    },
    /** E=作業コード, F=部品_中項目, G=単価, H=数量。順番は末尾。 */
    layout: {
      workerCode: 5,
      partMid: 6,
      unitPrice: 7,
      qty: 8
    }
  },

  /** 入力アプリの明細行数（表内スクロール） */
  app: {
    lineCount: 120
  },

  parts: {
    sheetName: '部品リスト',
    headerRow: 1,
    headers: {
      major: ['部品_大項目', '大項目'],
      mid: ['部品_中項目', '中項目', 'セット名'],
      name: ['部品名', '品名', '作業内容', '名称'],
      qty: ['数量'],
      unitPrice: ['単価', '価格'],
      order: ['順番', '表示順']
    }
  },

  workers: {
    sheetName: '作業者リスト',
    headerRow: 1,
    headers: {
      code: ['コード', '作業者コード', 'ID'],
      name: ['名前', '作業者', '氏名']
    }
  },

  /**
   * 作業リスト／部品リスト。図形に refreshAllMasterLists を割り当てて更新する。
   * 空の順番は更新時だけまとめて付ける。手で入れた順番は残す。
   */
  listRefresh: {
    sheets: ['作業リスト', '部品リスト'],
    buttonLabel: '更新',
    orderStep: 10
  },

  serviceInfo: {
    /** A＝全て（未選択時の種別順）。B〜E＝種別1〜4（大型／小型／BP板金／部品販売）。F＝受付担当。 */
    sheetName: '整備情報'
  },

  /**
   * 入力アプリの明細テンプレート。
   * 同じ「テンプレート名」の行が、選んだときの明細になる。
   */
  invoiceTemplate: {
    sheetName: '明細テンプレート',
    headerRow: 1,
    headers: {
      name: ['テンプレート名', 'テンプレ名', 'テンプレート'],
      major: ['作業_大項目', '大項目', '作業内容_大項目'],
      mid: ['作業_中項目', '中項目', '作業内容'],
      fee: ['技術料'],
      workerCode: ['作業者コード', '作業者', '担当コード'],
      partMajor: ['部品_大項目'],
      partMid: ['部品_中項目', '部品名'],
      unitPrice: ['単価'],
      qty: ['数量'],
      discYen: ['値引額', '値引']
    }
  },

  /**
   * 請求書保存。一覧は 1 件 1 行、明細は年ごとの別シート（年間2000件想定）。
   * K-No は 4 桁（K- は付けない）。
   */
  invoiceSave: {
    sheetName: '請求書保存',
    detailPrefix: '請求書保存明細_',
    headerRow: 1
  },

  /**
   * K-No で保存（CSV相当）を探し、印刷_表示へ描画する。
   */
  invoiceSearch: {
    sheetName: '請求書検索'
  },

  /**
   * 印刷は固定2シート。正本は請求書保存（CSV相当）。タブは増やさない。
   * sheetName＝作成ボタン。viewSheetName＝検索からの表示。
   */
  print: {
    sheetName: '印刷',
    viewSheetName: '印刷_表示',
    sampleSheetName: '印刷原本',
    sheetNamePrefix: '印刷_',
    samplePrefix: '印刷原本_',
    title: '近海請求書',
    linesPerPage: 30,
    pageRows: 42,
    colCount: 8,
    layout: {
      title: 0,
      metaL1: 1,
      metaV1: 2,
      metaL2: 3,
      metaV2: 4,
      metaL3: 5,
      metaV3: 6,
      spacer: 7,
      colHead: 8,
      firstLine: 9,
      footerStart: 39
    }
  },

  /** インストール型 onEdit の再入防止用キャッシュキー */
  internalWriteCacheKey: 'KINKAI_INTERNAL_WRITE',

  logPrefix: '[請求書入力]'
};
