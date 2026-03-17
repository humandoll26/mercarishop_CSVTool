const STORAGE_KEY = "mercari-shops-draft-multi-record";
const ENCODING_STORAGE_KEY = "mercari-shops-draft-encoding";
const MASTER_STORAGE_KEY = "mercari-shops-draft-master-cache";
const BUNDLED_MASTER_FILES = {
  brand: "brand_master_sjis.csv",
  category: "category_master_updated_sjis.csv"
};
const MASTER_HINT_DEFAULT = {
  brand: "ブランドCSVを開くと検索できます。",
  category: "カテゴリCSVを開くと検索できます。"
};
const MASTER_LABELS = {
  brand: "ブランド",
  category: "カテゴリ"
};
const SHARED_FIELD_KEYS = ["condition", "shippingMethod", "prefecture", "shipDays", "status", "shippingPayer"];

const CSV_HEADERS = [
  "商品画像名_1","商品画像名_2","商品画像名_3","商品画像名_4","商品画像名_5",
  "商品画像名_6","商品画像名_7","商品画像名_8","商品画像名_9","商品画像名_10",
  "商品画像名_11","商品画像名_12","商品画像名_13","商品画像名_14","商品画像名_15",
  "商品画像名_16","商品画像名_17","商品画像名_18","商品画像名_19","商品画像名_20",
  "商品名","商品説明",
  "SKU1_種類","SKU1_在庫数","SKU1_商品管理コード","SKU1_JANコード",
  "SKU2_種類","SKU2_在庫数","SKU2_商品管理コード","SKU2_JANコード",
  "SKU3_種類","SKU3_在庫数","SKU3_商品管理コード","SKU3_JANコード",
  "SKU4_種類","SKU4_在庫数","SKU4_商品管理コード","SKU4_JANコード",
  "SKU5_種類","SKU5_在庫数","SKU5_商品管理コード","SKU5_JANコード",
  "SKU6_種類","SKU6_在庫数","SKU6_商品管理コード","SKU6_JANコード",
  "SKU7_種類","SKU7_在庫数","SKU7_商品管理コード","SKU7_JANコード",
  "SKU8_種類","SKU8_在庫数","SKU8_商品管理コード","SKU8_JANコード",
  "SKU9_種類","SKU9_在庫数","SKU9_商品管理コード","SKU9_JANコード",
  "SKU10_種類","SKU10_在庫数","SKU10_商品管理コード","SKU10_JANコード",
  "ブランドID","販売価格","カテゴリID","商品の状態","配送方法","発送元の地域",
  "発送までの日数","商品ステータス","配送料の負担","送料ID","メルカリBiz配送_クール区分"
];

const OPTION_MASTERS = {
  condition: [
    { value: "", label: "選択してください" },
    { value: "1", label: "1: 新品、未使用" },
    { value: "2", label: "2: 未使用に近い" },
    { value: "3", label: "3: 目立った傷や汚れなし" },
    { value: "4", label: "4: やや傷や汚れあり" },
    { value: "5", label: "5: 傷や汚れあり" }
  ],
  shippingMethod: [
    { value: "", label: "選択してください" },
    { value: "1", label: "1: らくらくメルカリ便" },
    { value: "2", label: "2: ゆうゆうメルカリ便" },
    { value: "3", label: "3: メルカリBiz配送" },
    { value: "9", label: "9: その他" }
  ],
  prefecture: [
    { value: "", label: "選択してください" },
    ...[
      ["jp01","北海道"],["jp02","青森県"],["jp03","岩手県"],["jp04","宮城県"],["jp05","秋田県"],["jp06","山形県"],
      ["jp07","福島県"],["jp08","茨城県"],["jp09","栃木県"],["jp10","群馬県"],["jp11","埼玉県"],["jp12","千葉県"],
      ["jp13","東京都"],["jp14","神奈川県"],["jp15","新潟県"],["jp16","富山県"],["jp17","石川県"],["jp18","福井県"],
      ["jp19","山梨県"],["jp20","長野県"],["jp21","岐阜県"],["jp22","静岡県"],["jp23","愛知県"],["jp24","三重県"],
      ["jp25","滋賀県"],["jp26","京都府"],["jp27","大阪府"],["jp28","兵庫県"],["jp29","奈良県"],["jp30","和歌山県"],
      ["jp31","鳥取県"],["jp32","島根県"],["jp33","岡山県"],["jp34","広島県"],["jp35","山口県"],["jp36","徳島県"],
      ["jp37","香川県"],["jp38","愛媛県"],["jp39","高知県"],["jp40","福岡県"],["jp41","佐賀県"],["jp42","長崎県"],
      ["jp43","熊本県"],["jp44","大分県"],["jp45","宮崎県"],["jp46","鹿児島県"],["jp47","沖縄県"]
    ].map(([value, label]) => ({ value, label: `${value}: ${label}` }))
  ],
  shipDays: [
    { value: "", label: "選択してください" },
    { value: "1", label: "1: 1〜2日で発送" },
    { value: "2", label: "2: 2〜3日で発送" },
    { value: "3", label: "3: 4〜7日で発送" }
  ],
  status: [
    { value: "1", label: "1: 非公開" },
    { value: "2", label: "2: 公開" }
  ],
  shippingPayer: [
    { value: "", label: "選択してください" },
    { value: "1", label: "1: 送料込み(出品者負担)" },
    { value: "2", label: "2: 着払い(購入者負担)" }
  ],
  coolType: [
    { value: "", label: "未設定" },
    { value: "1", label: "1: 通常" },
    { value: "2", label: "2: 冷蔵" },
    { value: "3", label: "3: 冷凍" }
  ]
};

const MASTER_ENCODINGS = { brand: "shift-jis", category: "utf-8" };

const ui = {
  brandFileInput: document.getElementById("brandFileInput"),
  categoryFileInput: document.getElementById("categoryFileInput"),
  brandMasterStatus: document.getElementById("brandMasterStatus"),
  categoryMasterStatus: document.getElementById("categoryMasterStatus"),
  brandSearchHint: document.getElementById("brandSearchHint"),
  categorySearchHint: document.getElementById("categorySearchHint"),
  recordList: document.getElementById("recordList"),
  addRecordButton: document.getElementById("addRecordButton"),
  duplicateRecordButton: document.getElementById("duplicateRecordButton"),
  deleteRecordButton: document.getElementById("deleteRecordButton"),
  clearRecordsButton: document.getElementById("clearRecordsButton"),
  currentRecordLabel: document.getElementById("currentRecordLabel"),
  productName: document.getElementById("productName"),
  description: document.getElementById("description"),
  price: document.getElementById("price"),
  brandSearch: document.getElementById("brandSearch"),
  brandResults: document.getElementById("brandResults"),
  brandSelection: document.getElementById("brandSelection"),
  categorySearch: document.getElementById("categorySearch"),
  categoryHierarchy: document.getElementById("categoryHierarchy"),
  categoryQuickResults: document.getElementById("categoryQuickResults"),
  clearCategoryButton: document.getElementById("clearCategoryButton"),
  categorySelection: document.getElementById("categorySelection"),
  condition: document.getElementById("condition"),
  shippingMethod: document.getElementById("shippingMethod"),
  prefecture: document.getElementById("prefecture"),
  shipDays: document.getElementById("shipDays"),
  status: document.getElementById("status"),
  shippingPayer: document.getElementById("shippingPayer"),
  shippingFeeId: document.getElementById("shippingFeeId"),
  shippingFeeHint: document.getElementById("shippingFeeHint"),
  coolType: document.getElementById("coolType"),
  coolTypeHint: document.getElementById("coolTypeHint"),
  sku1Stock: document.getElementById("sku1Stock"),
  sku1Code: document.getElementById("sku1Code"),
  sku1Jan: document.getElementById("sku1Jan"),
  sampleButton: document.getElementById("sampleButton"),
  resetButton: document.getElementById("resetButton"),
  validateButton: document.getElementById("validateButton"),
  validationSummary: document.getElementById("validationSummary"),
  errorList: document.getElementById("errorList"),
  encodingStatus: document.getElementById("encodingStatus"),
  encodingNote: document.getElementById("encodingNote"),
  downloadButton: document.getElementById("downloadButton"),
  recordCountView: document.getElementById("recordCountView"),
  categoryCodeView: document.getElementById("categoryCodeView"),
  brandCodeView: document.getElementById("brandCodeView"),
  encodingModeView: document.getElementById("encodingModeView"),
  csvPreview: document.getElementById("csvPreview"),
  jsonPreview: document.getElementById("jsonPreview"),
  clearStorageButton: document.getElementById("clearStorageButton"),
  encodingRadios: Array.from(document.querySelectorAll('input[name="encodingMode"]'))
};

const state = {
  brandMaster: [],
  categoryMaster: [],
  categoryTree: [],
  sharedDefaults: createSharedDefaults(),
  products: [],
  selectedIndex: 0
};

initialize();

function initialize() {
  renderSelects();
  bindEvents();
  hydrateFromStorage();
  hydrateMasterCache();
  hydrateEncodingMode();
  resetMasterHint("brand");
  resetMasterHint("category");
  setInitialMasterStatus();
  renderAll();
  void loadBundledMasters();
}

function renderSelects() {
  fillSelect(ui.condition, OPTION_MASTERS.condition);
  fillSelect(ui.shippingMethod, OPTION_MASTERS.shippingMethod);
  fillSelect(ui.prefecture, OPTION_MASTERS.prefecture);
  fillSelect(ui.shipDays, OPTION_MASTERS.shipDays);
  fillSelect(ui.status, OPTION_MASTERS.status);
  fillSelect(ui.shippingPayer, OPTION_MASTERS.shippingPayer);
  fillSelect(ui.coolType, OPTION_MASTERS.coolType);
}

function fillSelect(select, options) {
  select.innerHTML = "";
  options.forEach((option) => {
    const el = document.createElement("option");
    el.value = option.value;
    el.textContent = option.label;
    select.appendChild(el);
  });
}

function createEmptyProduct() {
  return {
    images: Array(20).fill(""),
    productName: "",
    description: "",
    skus: Array.from({ length: 10 }, () => ({ type: "", stock: "", code: "", jan: "" })),
    brandId: "",
    brandLabel: "",
    price: "",
    categoryId: "",
    categoryLabel: "",
    categoryPathDraft: [],
    condition: "",
    shippingMethod: "",
    prefecture: "",
    shipDays: "",
    status: "1",
    shippingPayer: "",
    shippingFeeId: "",
    coolType: ""
  };
}

function createSharedDefaults() {
  return {
    condition: "",
    shippingMethod: "",
    prefecture: "",
    shipDays: "",
    status: "1",
    shippingPayer: ""
  };
}

function applySharedDefaultsToProduct(product) {
  SHARED_FIELD_KEYS.forEach((key) => {
    product[key] = state.sharedDefaults[key];
  });
  return product;
}

function cloneProduct(product) {
  return {
    ...product,
    images: [...product.images],
    skus: product.skus.map((sku) => ({ ...sku }))
  };
}

function normalizeStoredProduct(product) {
  const empty = createEmptyProduct();
  return {
    ...empty,
    ...product,
    images: empty.images.map((_, index) => product.images?.[index] || ""),
    skus: empty.skus.map((sku, index) => ({ ...sku, ...(product.skus?.[index] || {}) })),
    categoryPathDraft: Array.isArray(product.categoryPathDraft) ? product.categoryPathDraft.filter(Boolean) : []
  };
}

function hydrateFromStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    state.sharedDefaults = createSharedDefaults();
    state.products = [applySharedDefaultsToProduct(createEmptyProduct())];
    return;
  }

  try {
    const parsed = JSON.parse(raw);
    state.sharedDefaults = { ...createSharedDefaults(), ...(parsed.sharedDefaults || {}) };
    state.products = Array.isArray(parsed.products) && parsed.products.length > 0
      ? parsed.products.map(normalizeStoredProduct)
      : [applySharedDefaultsToProduct(createEmptyProduct())];
    state.selectedIndex = Math.min(parsed.selectedIndex || 0, state.products.length - 1);
  } catch {
    state.sharedDefaults = createSharedDefaults();
    state.products = [applySharedDefaultsToProduct(createEmptyProduct())];
    state.selectedIndex = 0;
  }
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    sharedDefaults: state.sharedDefaults,
    products: state.products,
    selectedIndex: state.selectedIndex
  }));
}

function hydrateMasterCache() {
  const raw = localStorage.getItem(MASTER_STORAGE_KEY);
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    state.brandMaster = Array.isArray(parsed.brandMaster) ? parsed.brandMaster : [];
    state.categoryMaster = Array.isArray(parsed.categoryMaster) ? parsed.categoryMaster : [];
    state.categoryTree = buildCategoryTree(state.categoryMaster);
  } catch {
    state.brandMaster = [];
    state.categoryMaster = [];
    state.categoryTree = [];
  }
}

function persistMasterCache() {
  const fullCache = JSON.stringify({
    brandMaster: state.brandMaster,
    categoryMaster: state.categoryMaster
  });

  try {
    localStorage.setItem(MASTER_STORAGE_KEY, fullCache);
    return;
  } catch {}

  try {
    localStorage.setItem(MASTER_STORAGE_KEY, JSON.stringify({
      brandMaster: [],
      categoryMaster: state.categoryMaster
    }));
  } catch {}
}

function setInitialMasterStatus() {
  setMasterStatus("brand", state.brandMaster.length > 0);
  setMasterStatus("category", state.categoryMaster.length > 0);

  if (state.brandMaster.length > 0) {
    setMasterHint("brand", "前回読み込んだブランドCSVを復元しました。");
  }
  if (state.categoryMaster.length > 0) {
    setMasterHint("category", "前回読み込んだカテゴリCSVを復元しました。");
  }
}

function bindEvents() {
  ui.brandFileInput.addEventListener("change", (event) => handleMasterFile(event, "brand"));
  ui.categoryFileInput.addEventListener("change", (event) => handleMasterFile(event, "category"));
  document.getElementById("brandLoadButton").addEventListener("click", () => ui.brandFileInput.click());
  document.getElementById("categoryLoadButton").addEventListener("click", () => ui.categoryFileInput.click());

  ui.addRecordButton.addEventListener("click", addRecord);
  ui.duplicateRecordButton.addEventListener("click", duplicateRecord);
  ui.deleteRecordButton.addEventListener("click", deleteRecord);
  ui.clearRecordsButton.addEventListener("click", clearAllRecords);
  ui.sampleButton.addEventListener("click", applySample);
  ui.resetButton.addEventListener("click", resetCurrentRecord);
  ui.validateButton.addEventListener("click", () => renderValidation(validateAllProducts()));
  ui.downloadButton.addEventListener("click", downloadCsv);
  ui.clearCategoryButton.addEventListener("click", clearCategorySelection);
  ui.clearStorageButton.addEventListener("click", clearAllSavedData);

  [
    ui.productName, ui.description, ui.price, ui.condition, ui.shippingMethod, ui.prefecture,
    ui.shipDays, ui.status, ui.shippingPayer, ui.shippingFeeId, ui.coolType,
    ui.sku1Stock, ui.sku1Code, ui.sku1Jan
  ].forEach((element) => {
    element.addEventListener("input", handleFormChange);
    element.addEventListener("change", handleFormChange);
  });

  ui.brandSearch.addEventListener("input", () => {
    const product = currentProduct();
    if (product.brandId && ui.brandSearch.value.trim() !== product.brandLabel) {
      product.brandId = "";
      product.brandLabel = "";
      ui.brandSelection.textContent = "未選択";
      persistState();
    }
    renderBrandResults();
  });

  ui.categorySearch.addEventListener("input", () => {
    const product = currentProduct();
    const inputValue = ui.categorySearch.value.trim();
    const selectedCategory = state.categoryMaster.find((item) => item.id === product.categoryId);
    const selectedValue = selectedCategory?.fullPath || product.categoryLabel;
    const draftValue = product.categoryPathDraft.join(" > ");
    const shouldResetConfirmed = product.categoryId && inputValue !== selectedValue;
    const shouldResetDraft = !product.categoryId && product.categoryPathDraft.length > 0 && inputValue !== draftValue;

    if (shouldResetConfirmed || shouldResetDraft) {
      product.categoryId = "";
      product.categoryLabel = "";
      product.categoryPathDraft = [];
      ui.categorySelection.textContent = "未選択";
      persistState();
    }
    renderCategoryHierarchy();
    renderCategoryQuickResults();
  });

  ui.encodingRadios.forEach((radio) => {
    radio.addEventListener("change", () => {
      localStorage.setItem(ENCODING_STORAGE_KEY, getEncodingMode());
      updateOutputs();
    });
  });
}

async function loadBundledMasters() {
  const results = await Promise.allSettled(
    Object.keys(BUNDLED_MASTER_FILES).map((kind) => loadBundledMaster(kind))
  );

  if (results.some((result) => result.status === "fulfilled" && result.value)) {
    renderAll();
    persistState();
  }
}

async function loadBundledMaster(kind) {
  try {
    const response = await fetch(BUNDLED_MASTER_FILES[kind], { cache: "no-store" });
    if (!response.ok) throw new Error(`failed to load ${kind} master`);
    const buffer = await response.arrayBuffer();
    return applyMasterBuffer(kind, buffer, "bundled");
  } catch {
    resetMasterHint(kind);
    return false;
  }
}

function handleFormChange() {
  syncSharedDefaultsFromForm();
  const product = currentProduct();
  product.productName = ui.productName.value.trim();
  product.description = ui.description.value.trim();
  product.price = ui.price.value.trim();
  product.condition = ui.condition.value;
  product.shippingMethod = ui.shippingMethod.value;
  product.prefecture = ui.prefecture.value;
  product.shipDays = ui.shipDays.value;
  product.status = ui.status.value;
  product.shippingPayer = ui.shippingPayer.value;
  product.shippingFeeId = ui.shippingFeeId.value.trim();
  product.coolType = ui.coolType.value;
  product.skus[0].stock = ui.sku1Stock.value.trim();
  product.skus[0].code = ui.sku1Code.value.trim();
  product.skus[0].jan = ui.sku1Jan.value.trim();
  applyFieldDependencies(product);
  persistState();
  renderRecordList();
  updateOutputs();
}

function syncSharedDefaultsFromForm() {
  state.sharedDefaults.condition = ui.condition.value;
  state.sharedDefaults.shippingMethod = ui.shippingMethod.value;
  state.sharedDefaults.prefecture = ui.prefecture.value;
  state.sharedDefaults.shipDays = ui.shipDays.value;
  state.sharedDefaults.status = ui.status.value;
  state.sharedDefaults.shippingPayer = ui.shippingPayer.value;
}

function currentProduct() {
  return state.products[state.selectedIndex];
}

function renderAll() {
  if (state.products.length === 0) {
    state.products = [applySharedDefaultsToProduct(createEmptyProduct())];
    state.selectedIndex = 0;
  }
  populateForm(currentProduct());
  renderRecordList();
  renderBrandResults();
  renderCategoryHierarchy();
  renderCategoryQuickResults();
  renderValidation(validateAllProducts());
  updateOutputs();
}

function populateForm(product) {
  ui.currentRecordLabel.textContent = `商品 ${state.selectedIndex + 1} を編集中です。`;
  ui.productName.value = product.productName;
  ui.description.value = product.description;
  ui.price.value = product.price;
  ui.brandSearch.value = product.brandLabel;
  ui.categorySearch.value = product.categoryLabel;
  const selectedCategory = state.categoryMaster.find((item) => item.id === product.categoryId);
  if (selectedCategory?.fullPath) {
    ui.categorySearch.value = selectedCategory.fullPath;
  } else if (product.categoryPathDraft.length > 0) {
    ui.categorySearch.value = product.categoryPathDraft.join(" > ");
  }
  ui.brandSelection.textContent = product.brandId ? `${product.brandId} / ${product.brandLabel || product.brandId}` : "未選択";
  ui.categorySelection.textContent = product.categoryId
    ? `${product.categoryId} / ${(selectedCategory?.fullPath || product.categoryLabel || product.categoryId)}`
    : "未選択";
  ui.condition.value = product.condition;
  ui.shippingMethod.value = product.shippingMethod;
  ui.prefecture.value = product.prefecture;
  ui.shipDays.value = product.shipDays;
  ui.status.value = product.status;
  ui.shippingPayer.value = product.shippingPayer;
  ui.shippingFeeId.value = product.shippingFeeId;
  ui.coolType.value = product.coolType;
  ui.sku1Stock.value = product.skus[0].stock || "";
  ui.sku1Code.value = product.skus[0].code || "";
  ui.sku1Jan.value = product.skus[0].jan || "";
  applyFieldDependencies(product);
}

function renderRecordList() {
  ui.recordList.innerHTML = "";
  state.products.forEach((product, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `record-item${index === state.selectedIndex ? " is-active" : ""}`;
    button.innerHTML = `<strong>${escapeHtml(product.productName || `商品 ${index + 1}`)}</strong><small>${escapeHtml(product.categoryId || "カテゴリ未選択")} / ${escapeHtml(product.price || "価格未入力")}</small>`;
    button.addEventListener("click", () => {
      state.selectedIndex = index;
      renderAll();
      persistState();
    });
    ui.recordList.appendChild(button);
  });
}

function addRecord() {
  state.products.push(applySharedDefaultsToProduct(createEmptyProduct()));
  state.selectedIndex = state.products.length - 1;
  renderAll();
  persistState();
}

function duplicateRecord() {
  state.products.splice(state.selectedIndex + 1, 0, cloneProduct(currentProduct()));
  state.selectedIndex += 1;
  renderAll();
  persistState();
}

function deleteRecord() {
  if (state.products.length === 1) {
    state.products[0] = applySharedDefaultsToProduct(createEmptyProduct());
    state.selectedIndex = 0;
  } else {
    state.products.splice(state.selectedIndex, 1);
    state.selectedIndex = Math.max(0, state.selectedIndex - 1);
  }
  renderAll();
  persistState();
}

function clearAllRecords() {
  const confirmed = window.confirm("商品レコードをすべて削除します。よろしいですか。");
  if (!confirmed) return;

  state.products = [applySharedDefaultsToProduct(createEmptyProduct())];
  state.selectedIndex = 0;
  renderAll();
  persistState();
}

function applySample() {
  const product = applySharedDefaultsToProduct(createEmptyProduct());
  product.productName = "昭和レトロ ガラス皿 青";
  product.description = "昭和期のガラス皿です。小キズあり。下書き登録用サンプルです。";
  product.price = "1800";
  product.condition = "4";
  product.shippingMethod = "1";
  product.prefecture = "jp13";
  product.shipDays = "2";
  product.status = "1";
  product.shippingPayer = "1";
  product.skus[0].stock = "1";
  if (state.categoryMaster[0]) {
    product.categoryId = state.categoryMaster[0].id;
    product.categoryLabel = state.categoryMaster[0].name;
  }
  state.products[state.selectedIndex] = product;
  renderAll();
  persistState();
}

function resetCurrentRecord() {
  state.products[state.selectedIndex] = applySharedDefaultsToProduct(createEmptyProduct());
  renderAll();
  persistState();
}

function applyFieldDependencies(product) {
  const shippingFeeEnabled = product.shippingPayer !== "1" && product.shippingPayer !== "";
  const coolTypeEnabled = product.shippingMethod === "3";

  if (!shippingFeeEnabled) {
    product.shippingFeeId = "";
    ui.shippingFeeId.value = "";
  }
  if (!coolTypeEnabled) {
    product.coolType = "";
    ui.coolType.value = "";
  }

  ui.shippingFeeId.disabled = !shippingFeeEnabled;
  ui.coolType.disabled = !coolTypeEnabled;
  ui.shippingFeeHint.textContent = shippingFeeEnabled ? "現在の設定では送料IDを使用できます。" : "現在の設定では送料IDは使用されません。";
  ui.coolTypeHint.textContent = coolTypeEnabled ? "現在の設定ではクール区分を使用できます。" : "メルカリBiz配送のときだけクール区分を使用できます。";
}

function validateProduct(product) {
  const errors = [];
  const warnings = [];

  if (!product.productName) errors.push("商品名が空です。");
  if (!product.description) errors.push("商品説明が空です。");
  if (!product.price) errors.push("販売価格が空です。");
  else if (!/^\d+$/.test(product.price) || Number(product.price) <= 0) errors.push("販売価格は正整数で入力してください。");
  if (!product.categoryId) errors.push("カテゴリIDが未選択です。");
  if (!product.condition) errors.push("商品の状態が未選択です。");
  if (!product.shippingMethod) errors.push("配送方法が未選択です。");
  if (!product.prefecture) errors.push("発送元の地域が未選択です。");
  if (!product.shipDays) errors.push("発送までの日数が未選択です。");
  if (!product.status || !["1", "2"].includes(product.status)) errors.push("商品ステータスが不正です。");
  if (!product.shippingPayer) errors.push("配送料の負担が未選択です。");
  if (!["", "1", "2", "3"].includes(product.coolType)) errors.push("クール区分の値が不正です。");
  if (product.brandId && !state.brandMaster.some((item) => item.id === product.brandId)) errors.push("ブランドIDがマスタ未一致です。");
  if (product.categoryId && !state.categoryMaster.some((item) => item.id === product.categoryId)) errors.push("カテゴリIDがマスタ未一致です。");
  if (product.skus[0].stock && !/^\d+$/.test(product.skus[0].stock)) errors.push("SKU1_在庫数は整数で入力してください。");

  if (!product.brandId) warnings.push("ブランド未設定です。");
  if (!product.skus[0].stock) errors.push("SKU1_在庫数が空です。");
  if (product.skus[0].jan && !/^\d+$/.test(product.skus[0].jan)) warnings.push("SKU1_JANコードは数字のみで入力してください。");
  if (product.shippingPayer === "1" && product.shippingFeeId) warnings.push("出品者負担では送料IDは出力時に空欄へ補正されます。");
  if (product.shippingMethod !== "3" && product.coolType) warnings.push("メルカリBiz配送以外ではクール区分は出力時に空欄へ補正されます。");
  return { errors, warnings };
}

function validateAllProducts() {
  const errors = [];
  const warnings = [];
  state.products.forEach((product, index) => {
    const result = validateProduct(product);
    result.errors.forEach((message) => errors.push(`商品 ${index + 1}: ${message}`));
    result.warnings.forEach((message) => warnings.push(`商品 ${index + 1}: ${message}`));
  });
  return { errors, warnings };
}

function renderValidation(result) {
  renderList(ui.errorList, result.errors, "エラーはありません。");
  if (result.errors.length > 0) {
    ui.validationSummary.textContent = `エラー ${result.errors.length} 件。`;
    ui.validationSummary.className = "summary summary-error";
  } else {
    ui.validationSummary.textContent = "エラーはありません。CSVを出力できます。";
    ui.validationSummary.className = "summary summary-success";
  }
}

function renderList(container, items, emptyMessage) {
  container.innerHTML = "";
  (items.length > 0 ? items : [emptyMessage]).forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    container.appendChild(li);
  });
}

function normalizeBeforeExport(product) {
  const normalized = cloneProduct(product);
  if (normalized.shippingPayer === "1") normalized.shippingFeeId = "";
  if (normalized.shippingMethod !== "3") normalized.coolType = "";
  return normalized;
}

function productToCsvRow(product) {
  const normalized = normalizeBeforeExport(product);
  const row = {};
  normalized.images.forEach((value, index) => { row[`商品画像名_${index + 1}`] = value; });
  row["商品名"] = normalized.productName;
  row["商品説明"] = normalized.description;
  normalized.skus.forEach((sku, index) => {
    row[`SKU${index + 1}_種類`] = sku.type;
    row[`SKU${index + 1}_在庫数`] = sku.stock;
    row[`SKU${index + 1}_商品管理コード`] = sku.code;
    row[`SKU${index + 1}_JANコード`] = sku.jan;
  });
  row["ブランドID"] = normalized.brandId;
  row["販売価格"] = normalized.price;
  row["カテゴリID"] = normalized.categoryId;
  row["商品の状態"] = normalized.condition;
  row["配送方法"] = normalized.shippingMethod;
  row["発送元の地域"] = normalized.prefecture;
  row["発送までの日数"] = normalized.shipDays;
  row["商品ステータス"] = normalized.status;
  row["配送料の負担"] = normalized.shippingPayer;
  row["送料ID"] = normalized.shippingFeeId;
  row["メルカリBiz配送_クール区分"] = normalized.coolType;
  return row;
}

function buildCsv(products) {
  const header = CSV_HEADERS.map(escapeCsvCell).join(",");
  const rows = products.map((product) => {
    const row = productToCsvRow(product);
    return CSV_HEADERS.map((column) => escapeCsvCell(row[column] || "")).join(",");
  });
  return [header, ...rows].join("\n");
}

function escapeCsvCell(value) {
  return `"${String(value ?? "").replace(/"/g, "\"\"")}"`;
}

function updateOutputs() {
  const selected = currentProduct();
  ui.jsonPreview.value = JSON.stringify({
    selectedIndex: state.selectedIndex,
    products: state.products
  }, null, 2);
  ui.csvPreview.value = buildCsv(state.products);
  ui.recordCountView.textContent = `${state.products.length}件`;
  ui.categoryCodeView.textContent = selected.categoryId || "空欄";
  ui.brandCodeView.textContent = selected.brandId || "空欄";
  ui.encodingModeView.textContent = getEncodingMode() === "sjis" ? "S-JIS" : "UTF-8 BOM付き";
  updateEncodingSummary();
}

function getEncodingMode() {
  return ui.encodingRadios.find((radio) => radio.checked)?.value || "sjis";
}

function hydrateEncodingMode() {
  const mode = localStorage.getItem(ENCODING_STORAGE_KEY) === "utf8bom" ? "utf8bom" : "sjis";
  ui.encodingRadios.forEach((radio) => { radio.checked = radio.value === mode; });
}

function updateEncodingSummary() {
  if (getEncodingMode() === "sjis") {
    ui.encodingStatus.textContent = "文字コード: S-JIS（Windows向け推奨）";
    ui.encodingNote.textContent = "WindowsでメルカリShopsへアップロードする場合は、S-JISが推奨です。";
  } else {
    ui.encodingStatus.textContent = "文字コード: UTF-8 BOM付き";
    ui.encodingNote.textContent = "Windows環境では S-JIS の方が安全な場合があります。";
  }
}

function downloadCsv() {
  const validation = validateAllProducts();
  renderValidation(validation);
  persistState();
  if (validation.errors.length > 0) {
    ui.validationSummary.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  const csvText = buildCsv(state.products);
  const blob = buildCsvBlob(csvText, getEncodingMode());
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "mercari_shops_draft_records.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function buildCsvBlob(csvText, mode) {
  if (mode === "sjis") {
    const unicodeArray = Encoding.stringToCode(csvText);
    const sjisArray = Encoding.convert(unicodeArray, { to: "SJIS", from: "UNICODE" });
    return new Blob([new Uint8Array(sjisArray)], { type: "text/csv" });
  }
  return new Blob([`\uFEFF${csvText}`], { type: "text/csv;charset=utf-8" });
}

function handleMasterFile(event, kind) {
  const file = event.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      applyMasterBuffer(kind, reader.result, "manual");
      renderAll();
      persistState();
    } catch {
      setMasterStatus(kind, false);
      setMasterHint(kind, `${MASTER_LABELS[kind]}CSVを解析できませんでした。`);
    } finally {
      event.target.value = "";
    }
  };
  reader.readAsArrayBuffer(file);
}

function applyMasterBuffer(kind, buffer, source) {
  const text = decodeMasterText(buffer, kind);
  const rows = parseCsv(text);
  const items = kind === "brand" ? normalizeBrands(rows) : normalizeCategories(rows);

  if (kind === "brand") {
    state.brandMaster = items;
  } else {
    state.categoryMaster = items;
    state.categoryTree = buildCategoryTree(items);
  }

  persistMasterCache();

  setMasterStatus(kind, items.length > 0);
  if (items.length > 0) {
    const prefix = source === "bundled" ? "同梱の" : "";
    setMasterHint(kind, `${prefix}${MASTER_LABELS[kind]}CSVを読み込みました。`);
    return true;
  }

  setMasterHint(kind, `${MASTER_LABELS[kind]}CSVを解析できませんでした。`);
  return false;
}

function setMasterStatus(kind, isOk) {
  const el = kind === "brand" ? ui.brandMasterStatus : ui.categoryMasterStatus;
  el.textContent = isOk ? "OK" : "NG";
  el.className = `master-status-badge ${isOk ? "is-ok" : "is-ng"}`;
}

function setMasterHint(kind, message) {
  if (kind === "brand") {
    ui.brandSearchHint.textContent = message;
  } else {
    ui.categorySearchHint.textContent = message;
  }
}

function resetMasterHint(kind) {
  setMasterHint(kind, MASTER_HINT_DEFAULT[kind]);
}

function decodeMasterText(buffer, kind) {
  const primary = MASTER_ENCODINGS[kind] || "utf-8";
  const candidates = primary === "shift-jis" ? ["shift-jis", "utf-8"] : ["utf-8", "shift-jis"];

  for (const encoding of candidates) {
    try {
      const text = new TextDecoder(encoding).decode(buffer);
      const rows = parseCsv(text);
      if (rows.length > 0 && rows[0].length > 1) {
        return text;
      }
    } catch {}
  }

  return new TextDecoder(primary).decode(buffer);
}

function normalizeBrands(rows) {
  if (!rows.length) return [];
  const header = rows[0].map((value) => value.trim());
  const idIndex = findHeaderIndex(header, [/brand.?id/i, /^id$/i, /ブランドid/i]);
  const nameIndex = findHeaderIndex(header, [/brand.?name/i, /name/i, /ブランド名/i]);
  const kanaIndex = findHeaderIndex(header, [/kana/i, /カナ/i]);
  const englishIndex = findHeaderIndex(header, [/english/i, /alphabet/i, /英語/i]);
  const resolvedIdIndex = idIndex === -1 ? 0 : idIndex;
  const resolvedNameIndex = nameIndex === -1 ? 1 : nameIndex;
  const resolvedKanaIndex = kanaIndex === -1 ? 2 : kanaIndex;
  const resolvedEnglishIndex = englishIndex === -1 ? 3 : englishIndex;
  return rows.slice(1).map((row) => ({
    id: (row[resolvedIdIndex] || "").trim(),
    name: (row[resolvedNameIndex] || "").trim(),
    kana: (row[resolvedKanaIndex] || "").trim(),
    english: (row[resolvedEnglishIndex] || "").trim()
  })).filter((item) => item.id && item.name);
}

function normalizeCategories(rows) {
  if (!rows.length) return [];
  const header = rows[0].map((value) => value.trim());
  const idIndex = findHeaderIndex(header, [/category.?id/i, /^id$/i, /カテゴリid/i]);
  const nameIndex = findHeaderIndex(header, [/category.?name/i, /name/i, /カテゴリ名/i]);
  const fullPathIndex = findHeaderIndex(header, [/full/i, /path/i, /フル/i]);
  const resolvedIdIndex = idIndex === -1 ? 0 : idIndex;
  const resolvedNameIndex = nameIndex === -1 ? 1 : nameIndex;
  const resolvedFullPathIndex = fullPathIndex === -1 ? 2 : fullPathIndex;
  return rows.slice(1).map((row) => ({
    id: (row[resolvedIdIndex] || "").trim(),
    name: (row[resolvedNameIndex] || "").trim(),
    fullPath: (row[resolvedFullPathIndex] || row[resolvedNameIndex] || "").trim(),
    pathParts: (row[resolvedFullPathIndex] || row[resolvedNameIndex] || "")
      .split(">")
      .map((value) => value.trim())
      .filter(Boolean)
  })).filter((item) => item.id && item.name);
}

function buildCategoryTree(categories) {
  const root = [];

  categories.forEach((category) => {
    const parts = category.pathParts?.length ? category.pathParts : [category.name];
    let level = root;

    parts.forEach((part, index) => {
      let node = level.find((item) => item.label === part);
      if (!node) {
        node = { label: part, children: [], item: null };
        level.push(node);
      }
      if (index === parts.length - 1) {
        node.item = category;
      }
      level = node.children;
    });
  });

  const sortNodes = (nodes) => {
    nodes.sort((a, b) => a.label.localeCompare(b.label, "ja"));
    nodes.forEach((node) => sortNodes(node.children));
  };

  sortNodes(root);
  return root;
}

function renderCategoryHierarchy() {
  ui.categoryHierarchy.innerHTML = "";

  if (!state.categoryTree.length) {
    ui.categoryHierarchy.innerHTML = "<div class=\"category-hierarchy-empty\">カテゴリCSVを読み込むと階層選択できます。</div>";
    return;
  }

  const product = currentProduct();
  const selectedCategory = state.categoryMaster.find((item) => item.id === product.categoryId);
  const selectedParts = selectedCategory?.pathParts?.length
    ? selectedCategory.pathParts
    : product.categoryPathDraft;
  let nodes = state.categoryTree;
  let level = 0;

  while (nodes.length > 0) {
    const currentLevel = level;
    const field = document.createElement("label");
    field.className = "field category-level-field";

    const title = document.createElement("span");
    title.textContent = `階層 ${currentLevel + 1}`;
    field.appendChild(title);

    const select = document.createElement("select");
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "選択してください";
    select.appendChild(placeholder);

    nodes.forEach((node) => {
      const option = document.createElement("option");
      option.value = node.label;
      option.textContent = node.label;
      select.appendChild(option);
    });

    const selectedLabel = selectedParts[currentLevel] || "";
    select.value = selectedLabel;
    select.addEventListener("change", () => handleCategoryLevelChange(currentLevel, select.value));
    field.appendChild(select);
    ui.categoryHierarchy.appendChild(field);

    if (!selectedLabel) break;

    const nextNode = nodes.find((node) => node.label === selectedLabel);
    if (!nextNode) break;
    nodes = nextNode.children;
    level += 1;
  }
}

function handleCategoryLevelChange(level, value) {
  const product = currentProduct();
  const selectedCategory = state.categoryMaster.find((item) => item.id === product.categoryId);
  const currentParts = selectedCategory?.pathParts?.length
    ? selectedCategory.pathParts
    : product.categoryPathDraft;
  const nextParts = currentParts.slice(0, level);

  if (value) {
    nextParts[level] = value;
  }

  const matchedCategory = state.categoryMaster.find((item) =>
    item.pathParts.length === nextParts.length &&
    item.pathParts.every((part, index) => part === nextParts[index])
  );

  if (matchedCategory) {
    applyCategorySelection(matchedCategory, true);
    return;
  }

  product.categoryId = "";
  product.categoryLabel = "";
  product.categoryPathDraft = nextParts;
  ui.categorySearch.value = nextParts.join(" > ");
  ui.categorySelection.textContent = nextParts.length > 0 ? `${nextParts.join(" > ")} / 未確定` : "未選択";
  persistState();
  renderCategoryHierarchy();
  renderCategoryQuickResults();
  renderRecordList();
  updateOutputs();
}

function clearCategorySelection() {
  const product = currentProduct();
  product.categoryId = "";
  product.categoryLabel = "";
  product.categoryPathDraft = [];
  ui.categorySearch.value = "";
  ui.categorySelection.textContent = "未選択";
  persistState();
  renderCategoryHierarchy();
  renderCategoryQuickResults();
  renderRecordList();
  updateOutputs();
}

function clearAllSavedData() {
  const confirmed = window.confirm("保存済みの商品データ、設定、読込済みマスターを削除します。よろしいですか。");
  if (!confirmed) return;

  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(ENCODING_STORAGE_KEY);
  localStorage.removeItem(MASTER_STORAGE_KEY);

  state.brandMaster = [];
  state.categoryMaster = [];
  state.categoryTree = [];
  state.sharedDefaults = createSharedDefaults();
  state.products = [applySharedDefaultsToProduct(createEmptyProduct())];
  state.selectedIndex = 0;

  resetMasterHint("brand");
  resetMasterHint("category");
  setMasterStatus("brand", false);
  setMasterStatus("category", false);
  hydrateEncodingMode();
  renderAll();
}

function findHeaderIndex(header, patterns) {
  return header.findIndex((cell) => patterns.some((pattern) => pattern.test(cell)));
}

function renderBrandResults() {
  renderSearchResults({
    source: state.brandMaster,
    query: ui.brandSearch.value.trim().toLowerCase(),
    selectedId: currentProduct().brandId,
    container: ui.brandResults,
    matcher: (item, query) => [item.id, item.name, item.kana, item.english].some((value) => (value || "").toLowerCase().includes(query)),
    renderer: (item) => `<strong>${escapeHtml(item.name)}</strong><small>ID: ${escapeHtml(item.id)}</small>`,
    onSelect: (item) => {
      const product = currentProduct();
      product.brandId = item.id;
      product.brandLabel = item.name;
      ui.brandSearch.value = item.name;
      ui.brandSelection.textContent = `${item.id} / ${item.name}`;
      persistState();
      updateOutputs();
      renderBrandResults();
      renderRecordList();
    }
  });
}

function applyCategorySelection(item, syncSearchInput = false) {
  const product = currentProduct();
  product.categoryId = item.id;
  product.categoryLabel = item.name;
  product.categoryPathDraft = [...(item.pathParts || [])];
  if (syncSearchInput) {
    ui.categorySearch.value = item.fullPath || item.name;
  }
  ui.categorySelection.textContent = `${item.id} / ${item.fullPath || item.name}`;
  persistState();
  renderCategoryHierarchy();
  renderCategoryQuickResults();
  renderRecordList();
  updateOutputs();
}

function renderCategoryQuickResults() {
  const query = ui.categorySearch.value.trim().toLowerCase();
  const container = ui.categoryQuickResults;

  if (!query || state.categoryMaster.length === 0) {
    container.innerHTML = "";
    return;
  }

  const results = state.categoryMaster
    .filter((item) => [item.id, item.name, item.fullPath].some((value) => (value || "").toLowerCase().includes(query)))
    .slice(0, 8);

  if (results.length === 0) {
    container.innerHTML = "<div class=\"result-item\"><small>候補はありません。</small></div>";
    return;
  }

  container.innerHTML = "";
  results.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `result-item${item.id === currentProduct().categoryId ? " is-selected" : ""}`;
    button.innerHTML = `<strong>${escapeHtml(item.name)}</strong><small>${escapeHtml(item.fullPath || item.name)}</small>`;
    button.addEventListener("click", () => applyCategorySelection(item, true));
    container.appendChild(button);
  });
}

function renderSearchResults({ source, query, selectedId, container, matcher, renderer, onSelect }) {
  const results = source.filter((item) => !query || matcher(item, query)).slice(0, 20);
  if (!results.length) {
    container.innerHTML = "<div class=\"result-item\"><small>候補はありません。</small></div>";
    return;
  }
  container.innerHTML = "";
  results.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `result-item${item.id === selectedId ? " is-selected" : ""}`;
    button.innerHTML = renderer(item);
    button.addEventListener("click", () => onSelect(item));
    container.appendChild(button);
  });
}

function parseCsv(text) {
  const rows = [];
  let cell = "";
  let row = [];
  let inQuotes = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === "\"") {
      if (inQuotes && next === "\"") {
        cell += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  if (cell !== "" || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows.filter((line) => line.some((value) => value.trim() !== ""));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
