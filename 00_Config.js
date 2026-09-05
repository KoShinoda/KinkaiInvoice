/**
 * =============================================================================
 * 近海_請求書：作業リスト連動
 * =============================================================================
 *
 * 【ファイル分担】
 * 00_Config.js            … シート名・行・ヘッダー別名だけ。仕様変更はまずここ。
 * 01_Utils.js             … 文字列正規化・重複排除・順番ソート・内部書き込みガード。
 * 02_WorkListRepository.js… 作業リスト／入力シートのヘッダー解決とレコード化。
 * 03_InputSheetWriter.js  … プルダウン設定と、中項目直下への作業内容・技術料の書き込み。
 * 04_SelectionService.js  … 大項目変更時のクリア／中項目選択時の明細展開。
 * 05_Triggers.js          … onOpen / onEdit / メニュー。
 * 06_CandidateSheet.js    … 中項目候補マスタ。
 * 07_WebApp.js            … 入力 Web アプリと印刷シート作成の呼び出し。
 * 08_ServiceInfo.js       … 整備情報マスタ。
 * 09_InvoiceTemplate.js   … 明細テンプレートシートの読込とサンプル作成。
 * 10_ListRefresh.js       … リストの並べ替え（図形ボタン refreshAllMasterLists）。
 * 11_PrintLayout.js       … A4 印刷原本（1シート・ページ区切り）。
 *
 * 【中項目プルダウン】
 * 大項目を選ぶたびに GAS で候補を付け替えない。
 * シート「中項目候補」＋「中項目_参照」の数式で即時に切り替わる。
 * 大項目が空のときは列「（全て）」を参照し、全中項目から選べる。
 *
 * 【中項目選択時】
 * 空白以外の作業内容を、その行のすぐ下の空行へ上書きする（行の挿入はしない）。
 * 技術料は D 列。作業内容が空のマスタ行の技術料は、中項目行の D 列に載せる。
 *
 * 【作業リストの列】
 * A 大項目 / B 中項目 / C 作業内容 / D 技術料 / E 部品_中項目 / F 単価 / G 数量 / （他） / 最後が順番。
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
      order: ['順番', '表示順'],
      partMajor: ['部品_大項目'],
      partMid: ['部品_中項目', '部品名'],
      qty: ['数量'],
      unitPrice: ['単価']
    },
    /** E=部品_中項目, F=単価, G=数量。順番は末尾。 */
    layout: {
      partMid: 5,
      unitPrice: 6,
      qty: 7
    }
  },

  /** 入力シート（請求明細の入力側） */
  input: {
    sheetName: '車検_入力',
    headerRow: 8,
    /**
     * 大項目（B）・中項目（C）の入力開始行。
     * ここから maxSelectRows 行まで、行ごとに中項目プルダウンを付け替える。
     * 候補はシートに書き出さず、各セルの入力規則だけを更新する（行は増えない）。
     */
    dataStartRow: 9,
    /** 入力アプリの明細行数（表内スクロール） */
    appRows: 120,
    /** 大項目を並べられる行数（B9 なら 208 行目まで） */
    maxSelectRows: 200,
    /** 単一行処理の既定行（dataStartRow と同じでよい） */
    selectRow: 9,
    /**
     * 列番号を固定すると、編集のたびに見出し行を読まない（高速化）。
     * 車検_入力：B=大項目, C=中項目, D=技術料。
     */
    fixedCols: {
      major: 2,
      mid: 3,
      fee: 4
    },
    /** 下行を何行まで見て旧明細を消すか（200 行全読みはしない） */
    detailPeekRows: 80,
    headers: {
      serial: ['連番'],
      major: ['作業_大項目', '作業内容_大項目', '大項目'],
      /** C9 は中項目選択。C10 以降は明細名。同じ列を使う */
      mid: ['作業_中項目', '作業内容', '中項目'],
      fee: ['技術料'],
      workerCode: ['作業者コード', '作業者', '担当コード'],
      partMajor: ['部品_大項目'],
      partName: ['部品名', '部品_中項目'],
      qty: ['数量'],
      unitPrice: ['単価']
    }
  },

  /**
   * 大項目×中項目の重複なしマスタ（横持ち）。
   * A列「（全て）」＝全中項目。B列以降＝各大項目の中項目。
   */
  candidates: {
    sheetName: '中項目候補',
    allHeader: '（全て）',
    maxRows: 100
  },

  /**
   * 入力行ごとの FILTER 結果。C 列の入力規則がここを参照する。
   * 列1＝入力9行目、列2＝10行目…（非表示）
   */
  lookup: {
    sheetName: '中項目_参照'
  },

  parts: {
    sheetName: '部品リスト',
    headerRow: 1,
    headers: {
      major: ['部品_大項目', '大項目'],
      mid: ['部品_中項目', '部品名', '中項目'],
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
   * 車検_入力のヘッダー付近にある小計・値引（無ければ書き込みスキップ）。
   */
  summary: {
    techSub: 'E1',
    techDisc: 'E2',
    techTotal: 'E3',
    partSub: 'H1',
    partDisc: 'H2',
    grand: 'H5'
  },

  /**
   * 印刷用。車検原紙は使わず、A4 縦 1 シートにページを縦積みする。
   * ヘッダーは 1 枚目だけ、フッターは最終枚だけ。右上に No.n／m。
   * 中間ページは列見出し・明細・No. のみ。印刷設定は A4・幅1ページ。
   */
  print: {
    sheetName: '印刷',
    sampleSheetName: '印刷原本',
    sheetNamePrefix: '印刷_',
    samplePrefix: '印刷原本_',
    title: '近海請求書',
    linesPerPage: 20,
    pageRows: 36,
    colCount: 8,
    layout: {
      title: 0,
      metaL1: 1,
      metaV1: 2,
      metaL2: 3,
      metaV2: 4,
      spacer: 5,
      colHead: 6,
      firstLine: 7,
      footerStart: 28
    }
  },

  /** インストール型 onEdit の再入防止用キャッシュキー */
  internalWriteCacheKey: 'KINKAI_INTERNAL_WRITE',

  logPrefix: '[請求書入力]'
};
