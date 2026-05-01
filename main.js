const { app, BrowserWindow, Menu, globalShortcut, ipcMain, dialog, Tray, shell } = require('electron');
const remoteMain = require('@electron/remote/main');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const chokidar = require('chokidar');
const sharp = require('sharp');
const XLSX = require('xlsx');

remoteMain.initialize();

// --- Tray ---
let tray = null;
let mainWindow = null;
let autoUpdaterEventsBound = false;
let autoUpdaterFeedUrl = '';
let updateState = {
    checking: false,
    available: false,
    downloading: false,
    downloaded: false,
    progress: 0,
    status: '未检查更新',
    latestVersion: '',
    releaseDate: '',
    releaseName: '',
    error: '',
    lastCheckedAt: ''
};

// --- Config persistence ---
const CONFIG_FILE = path.join(app.getPath('userData'), 'compress-config.json');
const CLASSIFY_CONFIG_FILE = path.join(app.getPath('userData'), 'classify-config.json');
const SLICE_CONFIG_FILE = path.join(app.getPath('userData'), 'slice-config.json');
const TEMPLATE_CONFIG_FILE = path.join(app.getPath('userData'), 'template-config.json');
const PRODUCT_PUBLISH_CONFIG_FILE = path.join(app.getPath('userData'), 'product-publish-config.json');
const PRODUCT_PUBLISH_DATA_FILE = path.join(app.getPath('userData'), 'product-publish-data.json');
const PRINT_AI_CONFIG_FILE = path.join(app.getPath('userData'), 'print-ai-config.json');
const PRINT_AI_DATA_FILE = path.join(app.getPath('userData'), 'print-ai-data.json');
const PRINT_AI_STORAGE_DIR = path.join(app.getPath('userData'), 'print-ai');
const WORKFLOW_CONFIG_FILE = path.join(app.getPath('userData'), 'workflow-config.json');
const WORKFLOW_DATA_FILE = path.join(app.getPath('userData'), 'workflow-data.json');
const WORKFLOW_STORAGE_DIR = path.join(app.getPath('userData'), 'workflow');
const UPDATE_CONFIG_FILE = path.join(app.getPath('userData'), 'update-config.json');
const PACKAGE_JSON_FILE = path.join(__dirname, 'package.json');
const LEGACY_WATERMARK_PRESETS_FILE = path.join(app.getPath('userData'), 'watermark-presets.json');
const TEMPLATE_PARAMETER_PRESETS_FILE = path.join(app.getPath('userData'), 'template-parameter-presets.json');
const PRODUCT_PUBLISH_IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.bmp']);
const PRINT_AI_IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.avif']);
const WORKFLOW_IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.avif']);
const PRODUCT_PUBLISH_TEMU_TEMPLATE_NAME = '妙手Temu导入模板-非服饰类模板.xlsx';

function createDefaultProductPublishTypeMappings() {
    return [
        { id: 'type-rug', name: '地垫', keywords: ['地垫', '门垫', '浴室垫', 'floor mat', 'doormat', 'door mat', 'bath mat', 'bathroom rug', 'entryway mat', 'accent rug', 'rug'] },
        { id: 'type-mousepad', name: '鼠标垫', keywords: ['鼠标垫', 'mouse pad', 'mousepad', 'computer mousepad', 'gaming mousepad'] },
        { id: 'type-coffee-machine', name: '咖啡机垫', keywords: ['咖啡机垫', 'espresso machine pad', 'coffee maker underpad', 'coffee machine pad'] }
    ];
}

function normalizeProductPublishTypeMappings(mappings) {
    const defaults = createDefaultProductPublishTypeMappings();
    const fallbackMap = new Map(defaults.map((item) => [item.id, item]));
    const source = Array.isArray(mappings) && mappings.length ? mappings : defaults;
    const normalized = source
        .map((item, index) => {
            const rawId = String(item?.id || '').trim();
            const rawName = String(item?.name || '').trim();
            if (!rawName || rawName === '咖啡垫') return null;

            return {
                id: rawId || `product-type-${Date.now()}-${index + 1}`,
                name: rawName,
                // 关键词不再由用户维护，内部按类型名称和内置别名生成。
                keywords: []
            };
        })
        .filter((item) => item && item.name);

    const hasCoffeeMachine = normalized.some((item) => item.id === 'type-coffee-machine' || item.name === '咖啡机垫');
    if (!hasCoffeeMachine) {
        normalized.push({ ...fallbackMap.get('type-coffee-machine') });
    }
    return normalized;
}

function getProductPublishTypeKeywords(mapping) {
    const name = String(mapping?.name || '').trim();
    const keywords = new Set([name.toLowerCase()]);
    if (name === '地垫') {
        ['门垫', '浴室垫', 'floor mat', 'doormat', 'door mat', 'bath mat', 'bathroom rug', 'entryway mat', 'accent rug', 'rug'].forEach((item) => keywords.add(item.toLowerCase()));
    } else if (name === '鼠标垫') {
        ['mouse pad', 'mousepad', 'computer mousepad', 'gaming mousepad'].forEach((item) => keywords.add(item.toLowerCase()));
    } else if (name === '咖啡机垫') {
        ['espresso machine pad', 'coffee maker underpad', 'coffee machine pad'].forEach((item) => keywords.add(item.toLowerCase()));
    }
    return Array.from(keywords);
}

function loadConfig() {
    try {
        const cfg = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf-8'));
        return {
            directory: '',
            thresholdMB: 20,
            autoStart: false,
            ...(cfg || {})
        };
    } catch {
        return { directory: '', thresholdMB: 20, autoStart: false };
    }
}

function saveConfig(cfg) {
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2), 'utf-8');
}

function loadClassifyConfig() {
    try {
        const cfg = JSON.parse(fs.readFileSync(CLASSIFY_CONFIG_FILE, 'utf-8'));
        return {
            sourceDir: '',
            targetDir: '',
            userName: '',
            autoStart: false,
            ...(cfg || {})
        };
    } catch {
        return { sourceDir: '', targetDir: '', userName: '', autoStart: false };
    }
}

function saveClassifyConfig(cfg) {
    fs.writeFileSync(CLASSIFY_CONFIG_FILE, JSON.stringify(cfg, null, 2), 'utf-8');
}

function getDefaultSliceOutputDir() {
    return path.join(app.getPath('pictures'), 'ImageFlow切片结果');
}

function loadSliceConfig() {
    try {
        return JSON.parse(fs.readFileSync(SLICE_CONFIG_FILE, 'utf-8'));
    } catch {
        return { outputDir: getDefaultSliceOutputDir() };
    }
}

function saveSliceConfig(cfg) {
    fs.writeFileSync(SLICE_CONFIG_FILE, JSON.stringify(cfg, null, 2), 'utf-8');
}

function readJsonFile(filePath, fallback) {
    try {
        return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    } catch {
        return fallback;
    }
}

function writeJsonFile(filePath, data) {
    ensureDir(path.dirname(filePath));
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8');
}

function ensureDir(dirPath) {
    fs.mkdirSync(dirPath, { recursive: true });
    return dirPath;
}

function getExternalAssetsBaseDir() {
    return app.isPackaged ? path.dirname(process.execPath) : __dirname;
}

function getBundledTemplateRootDir() {
    if (app.isPackaged) {
        const externalTemplateRoot = path.join(path.dirname(process.execPath), 'templates');
        if (fs.existsSync(externalTemplateRoot)) {
            return externalTemplateRoot;
        }
    }
    return path.join(__dirname, 'templates');
}

function getBundledWatermarkRootDir() {
    if (app.isPackaged) {
        const externalWatermarkRoot = path.join(path.dirname(process.execPath), 'watermarks');
        if (fs.existsSync(externalWatermarkRoot)) {
            return externalWatermarkRoot;
        }
    }
    return path.join(__dirname, 'watermarks');
}

function getTemplateRendererScriptPath() {
    const externalPath = path.join(getExternalAssetsBaseDir(), 'template_renderer.py');
    if (app.isPackaged && fs.existsSync(externalPath)) {
        return externalPath;
    }
    return path.join(__dirname, 'template_renderer.py');
}

function getPythonRuntimeCandidates() {
    const externalBaseDir = getExternalAssetsBaseDir();
    const candidatePaths = [];
    if (app.isPackaged) {
        candidatePaths.push(
            path.join(externalBaseDir, 'python-runtime', 'python.exe'),
            path.join(externalBaseDir, 'python-runtime', 'Scripts', 'python.exe')
        );
    }
    candidatePaths.push(
        path.join(__dirname, 'python-runtime', 'python.exe'),
        path.join(__dirname, 'python-runtime', 'Scripts', 'python.exe')
    );
    return Array.from(new Set(candidatePaths));
}

function getDefaultTemplateRootDir() {
    return app.isPackaged
        ? path.join(app.getPath('userData'), 'templates')
        : path.join(getExternalAssetsBaseDir(), 'templates');
}

function getDefaultWatermarkDir() {
    return app.isPackaged
        ? path.join(app.getPath('userData'), 'watermarks')
        : path.join(getExternalAssetsBaseDir(), 'watermarks');
}

function normalizeDirectoryPath(dirPath, fallbackDir) {
    const trimmed = String(dirPath || '').trim();
    return path.resolve(trimmed || fallbackDir);
}

function hasSavedTemplateConfig() {
    return fs.existsSync(TEMPLATE_CONFIG_FILE);
}

function copyMissingDirectoryContents(sourceDir, targetDir) {
    if (!fs.existsSync(sourceDir)) {
        return;
    }
    ensureDir(targetDir);
    for (const entry of fs.readdirSync(sourceDir, { withFileTypes: true })) {
        if (entry.name === '.gitkeep') continue;
        const sourcePath = path.join(sourceDir, entry.name);
        const targetPath = path.join(targetDir, entry.name);
        if (fs.existsSync(targetPath)) {
            continue;
        }
        fs.cpSync(sourcePath, targetPath, {
            recursive: true,
            force: false,
            errorOnExist: false
        });
    }
}

function isLegacyPackagedTemplateRoot(dirPath) {
    if (!app.isPackaged) {
        return false;
    }
    return path.resolve(dirPath) === path.resolve(path.join(path.dirname(process.execPath), 'templates'));
}

function isLegacyPackagedWatermarkRoot(dirPath) {
    if (!app.isPackaged) {
        return false;
    }
    return path.resolve(dirPath) === path.resolve(path.join(path.dirname(process.execPath), 'watermarks'));
}

function persistTemplateConfigMigration(partialCfg) {
    try {
        saveTemplateConfig({
            ...loadTemplateConfig(),
            ...(partialCfg || {})
        });
    } catch {
        // 启动时迁移不应因配置写入失败阻塞程序打开
    }
}

function seedTemplateRootDir(targetDir) {
    const sourceDir = path.resolve(getBundledTemplateRootDir());
    const resolvedTargetDir = path.resolve(targetDir);
    if (resolvedTargetDir === sourceDir || !fs.existsSync(sourceDir)) {
        return;
    }
    const existingEntries = fs.readdirSync(resolvedTargetDir, { withFileTypes: true })
        .filter((entry) => entry.name !== '.gitkeep');
    if (existingEntries.length > 0) {
        return;
    }
    if (app.isPackaged && hasSavedTemplateConfig()) {
        return;
    }
    copyMissingDirectoryContents(sourceDir, resolvedTargetDir);
}

function seedWatermarkDir(targetDir) {
    const sourceDir = path.resolve(getBundledWatermarkRootDir());
    const resolvedTargetDir = path.resolve(targetDir);
    if (resolvedTargetDir === sourceDir || !fs.existsSync(sourceDir)) {
        return;
    }
    const existingEntries = fs.readdirSync(resolvedTargetDir, { withFileTypes: true })
        .filter((entry) => entry.name !== '.gitkeep');
    if (existingEntries.length > 0) {
        return;
    }
    if (app.isPackaged && hasSavedTemplateConfig()) {
        return;
    }
    copyMissingDirectoryContents(sourceDir, resolvedTargetDir);
}

function getDefaultTemplateOutputDir() {
    return path.join(app.getPath('pictures'), 'ImageFlow智能模板结果');
}

function getDefaultProductPublishOutputDir() {
    return path.join(app.getPath('documents'), 'ImageFlow产品发布');
}

function loadTemplateConfig() {
    const defaults = {
        outputDir: getDefaultTemplateOutputDir(),
        selectedTemplates: [],
        templateOrder: [],
        watermarkPresetId: '',
        parameterPresetId: '',
        defaultPreviewPath: '',
        defaultPreviewName: '',
        templateRootDir: getDefaultTemplateRootDir(),
        watermarkDir: getDefaultWatermarkDir()
    };
    try {
        const parsed = JSON.parse(fs.readFileSync(TEMPLATE_CONFIG_FILE, 'utf-8'));
        return {
            ...defaults,
            ...(parsed || {}),
            templateOrder: Array.isArray(parsed && parsed.templateOrder)
                ? Array.from(new Set(parsed.templateOrder.map((item) => String(item || '').trim()).filter(Boolean)))
                : defaults.templateOrder,
            templateRootDir: normalizeDirectoryPath(parsed && parsed.templateRootDir, defaults.templateRootDir),
            watermarkDir: normalizeDirectoryPath(parsed && parsed.watermarkDir, defaults.watermarkDir)
        };
    } catch {
        return defaults;
    }
}

function saveTemplateConfig(cfg) {
    const defaults = loadTemplateConfig();
    const nextCfg = {
        ...defaults,
        ...(cfg || {}),
        templateOrder: Array.isArray(cfg && cfg.templateOrder)
            ? Array.from(new Set(cfg.templateOrder.map((item) => String(item || '').trim()).filter(Boolean)))
            : defaults.templateOrder,
        templateRootDir: normalizeDirectoryPath(cfg && cfg.templateRootDir, defaults.templateRootDir),
        watermarkDir: normalizeDirectoryPath(cfg && cfg.watermarkDir, defaults.watermarkDir)
    };
    fs.writeFileSync(TEMPLATE_CONFIG_FILE, JSON.stringify(nextCfg, null, 2), 'utf-8');
    return nextCfg;
}

function getTemplateRootDir(cfg = loadTemplateConfig()) {
    const savedTemplateRootDir = normalizeDirectoryPath(cfg && cfg.templateRootDir, getDefaultTemplateRootDir());
    let templateRootDir = savedTemplateRootDir;
    if (isLegacyPackagedTemplateRoot(savedTemplateRootDir)) {
        templateRootDir = getDefaultTemplateRootDir();
        persistTemplateConfigMigration({ templateRootDir });
    }
    ensureDir(templateRootDir);
    seedTemplateRootDir(templateRootDir);
    return templateRootDir;
}

function getWatermarkDir(cfg = loadTemplateConfig()) {
    const savedWatermarkDir = normalizeDirectoryPath(cfg && cfg.watermarkDir, getDefaultWatermarkDir());
    let watermarkDir = savedWatermarkDir;
    if (isLegacyPackagedWatermarkRoot(savedWatermarkDir)) {
        watermarkDir = getDefaultWatermarkDir();
        persistTemplateConfigMigration({ watermarkDir });
    }
    ensureDir(watermarkDir);
    seedWatermarkDir(watermarkDir);
    return watermarkDir;
}

function getWatermarkPresetsFile(cfg = loadTemplateConfig()) {
    const watermarkDir = getWatermarkDir(cfg);
    const presetFile = path.join(watermarkDir, 'watermark-presets.json');
    if (!fs.existsSync(presetFile) && fs.existsSync(LEGACY_WATERMARK_PRESETS_FILE)) {
        fs.copyFileSync(LEGACY_WATERMARK_PRESETS_FILE, presetFile);
    }
    return presetFile;
}

function loadWatermarkPresets() {
    try {
        const presets = JSON.parse(fs.readFileSync(getWatermarkPresetsFile(), 'utf-8'));
        return Array.isArray(presets) ? presets : [];
    } catch {
        return [];
    }
}

function saveWatermarkPresets(presets) {
    fs.writeFileSync(getWatermarkPresetsFile(), JSON.stringify(Array.isArray(presets) ? presets : [], null, 2), 'utf-8');
}

function loadTemplateParameterPresets() {
    try {
        const presets = JSON.parse(fs.readFileSync(TEMPLATE_PARAMETER_PRESETS_FILE, 'utf-8'));
        return Array.isArray(presets) ? presets : [];
    } catch {
        return [];
    }
}

function saveTemplateParameterPresets(presets) {
    fs.writeFileSync(TEMPLATE_PARAMETER_PRESETS_FILE, JSON.stringify(Array.isArray(presets) ? presets : [], null, 2), 'utf-8');
}

const DEFAULT_PRODUCT_PUBLISH_PROMPT_DOC = `# Role

跨境电商中英标题生成助手（Vision + SEO + Compliance）

# Task

根据当前上传图片与已锁定的产品类型，生成一组可直接用于商品发布的英文标题和中文标题。

# Vision Rules

1. 识别优先级：
   - 产品类型已经由文件夹名称或图片名称锁定，必须严格服从，不要重新识别产品类型。
   - 重点识别图片中的图案、风格、配色、文字内容。
2. 图片理解范围：
   - 只能基于当前上传图片识别，不要编造图片里没有的元素。
   - 如果上传多张图片，请综合当前上传图片内容，为同一条产品记录生成一组统一标题，不要逐图编号输出。
3. 稳定识别要求：
   - 先在脑中提取稳定的识别摘要：产品类型、主图案、风格、颜色、文字内容。
   - 标题只能使用这份稳定摘要里的信息，不要为了差异化随意替换主题词。
   - 对地垫类产品，优先稳定识别主图案和核心风格，不要频繁改变产品名和场景词。
4. 图案文字识别：
   - 如果图片中有可辨认文字、标语、短句，必须提取。
   - 英文标题里用双引号包裹文字内容。
   - 如果图片文字是中文，英文标题中必须翻译成英文，不得保留中文字符。

# Critical Rules

1. 长度严格限制（英文标题，包含空格和标点，必须严格命中，不可超出）：
   - 鼠标垫（Mouse Pad）：150 - 200 字符
   - 其他产品（如地垫、咖啡机垫）：150 - 250 字符
   - 如果初稿超长，必须主动压缩到合规长度后再输出。

2. 强制前缀：
   - 地垫或用户明确要求 2D 的产品：英文必须以 \`[2D Flat Print]1pc \` 开头
   - 其他默认产品：英文必须以 \`1pc \` 开头

3. 合规过滤：
   - 绝对禁止出现毒品、武器、色情、暴力、政治敏感、涉及人类未成年等违禁内容及其同义表达。
   - 如果图片中存在高危元素，不要停止生成，直接净化为中性、抽象、艺术化表达。
   - 不要输出额外警告说明，只输出最终标题。

4. 输出格式：
   - 只输出两行纯文本
   - 第一行：英文标题
   - 第二行：中文标题
   - 不要解释，不要编号，不要代码块，不要字段标签

5. 语言与标点要求：
   - 英文标题必须是纯英文，不得混入中文
   - 中文标题必须是自然中文，不要变成营销文案
   - 英文标题必须使用标准英文标点组织结构，至少使用 3 个英文逗号分隔主要语义块
   - 中文标题必须使用自然中文标点，至少使用 3 个中文逗号分隔主要语义块
   - 不允许输出没有标点的一整串文本
   - 如果初稿没有标点，必须先补全标点后再输出最终结果

# Title Logic

1. 英文标题目标：
   - 先服从已锁定的产品类型
   - 再准确描述图案、风格、文字内容
   - 在准确基础上做 SEO 组织，不要为了 SEO 牺牲识别准确性
   - 相同图片多次生成时，应尽量保持产品名、主图案、核心风格一致

2. 负向过滤：
   - 移除材质词：Rubber, Polyester 等，统一替换为 \`Non-Slip Backing\`
   - 移除尺寸词：XL, XXL, 任何具体尺寸数字
   - 移除封边词：Edge, Stitched, Locked 等
   - 移除营销虚词：Outdoor, Washable, Super, Best Gift 等

3. 英文标题结构：
   - 前缀 + 同类通用产品名 + 图案描述/文字内容 + Non-Slip Backing + 动态长尾词
   - 不要限制长尾词数量，必须在识别准确的前提下自然补足标题长度
   - 当英文标题明显短于目标字符范围时，应继续补充与图片内容强相关的长尾词、用途词、风格词与场景词，直到接近目标字符范围
   - 如果英文标题低于 150 字符，视为不合格，必须继续扩充到 150 字符以上再输出
   - 只允许补充与当前图片真实内容、真实产品用途相关的词，不要为了凑长度堆砌无关词
   - 标题必须有清晰逗号分段，不能写成一整串无标点文本
   - 不要把产品名固定成唯一写法，同类产品可自然变化，但必须保持产品类别正确
   - 地垫类可自然使用：Floor Mat, Doormat, Bathroom Rug, Accent Rug, Entryway Rug, Decorative Rug 等
   - 鼠标垫类可自然使用：Mouse Pad, Desk Mat, Mousepad, Desktop Mat, Office Desk Mat 等
   - 咖啡机垫类可自然使用：Coffee Machine Mat, Coffee Bar Mat, Counter Mat, Espresso Machine Pad, Appliance Mat 等

4. 动态词库：
   - 鼠标垫：Office Desk Mat, Gaming Accessories, Desktop Protector, Computer Mousepad, Workstation Decor, PC Keyboard Mat, Laptop Pad, Gamer Setup, Workspace Decoration, Typing Mat, PC Table Cover, Home Office Supply, Workroom Essential, Gamer Gear
   - 咖啡机垫：Coffee Bar Mat, Kitchen Countertop Protector, Cafe Station, Espresso Machine Pad, Table Mat, Barista Station Accessory, Kitchen Counter Decor, Espresso Bar Setup, Coffee Maker Underpad, Tea Corner Mat, Dining Table Saver
   - 地垫：Entryway Mat, Bathroom Rug, Kitchen Floor Mat, Welcome Mat, Home Decor Carpet, Area Rug, Porch Carpet, Hallway Rug, Shower Floor Pad, Living Room Accent, Indoor Entrance Mat, Vanity Rug, Bedside Carpet

# Translation Rules

1. 中文标题必须以英文标题为唯一依据，逐段直译英文标题，不得自行补充、删减或改写信息。
2. 英文标题确定后，中文标题必须严格按英文标题的逗号分段顺序一一对应翻译。
3. 中文标题必须与英文标题语义完全一致，不允许英文写一种内容、中文写另一种内容。
4. 中文标题必须保留清晰分段，不能输出成没有任何标点的一整句。
5. 前缀强映射：
   - \`[2D Flat Print]1pc \` 对应 \`【2D平面打印】一件\`
   - \`1pc \` 对应 \`一件\`
6. 产品词必须准确：
   - 鼠标垫类必须体现鼠标垫/桌垫语义，但不要固定成唯一叫法
   - 地垫类必须体现门垫/地垫/浴室垫/装饰地毯语义，但不要固定成唯一叫法
   - 咖啡机垫类必须体现咖啡机垫/咖啡机台垫/咖啡吧台垫语义，但不要固定成唯一叫法

# Reference Examples

以下示例只用于学习格式、标点和节奏，不可照抄内容：

1pc Desk Mat, Watercolor Cat Floral Pattern, "Hello Summer", Non-Slip Backing, Office Mousepad, Desktop Protector, Gamer Setup, Workspace Decoration, Computer Mouse Pad, Home Office Supply, Typing Mat
一件 桌垫，水彩猫花卉图案，“Hello Summer”字样，防滑底，办公鼠标垫，桌面保护垫，游戏桌搭配，工作区装饰，电脑鼠标垫，居家办公用品，打字桌垫

[2D Flat Print]1pc Bathroom Rug, Vintage Botanical Leaves Pattern, Non-Slip Backing, Entryway Mat, Decorative Accent Rug, Kitchen Floor Mat, Indoor Entrance Rug, Hallway Carpet, Living Room Accent
【2D平面打印】一件 浴室垫，复古植物叶片图案，防滑底，门垫，装饰地毯，厨房地垫，室内入口地毯，走廊地毯，客厅点缀地毯

# Final Instruction

请严格基于当前上传图片与产品类型约束输出最终结果：
- 第一行：英文标题
- 第二行：中文标题
- 不要解释
- 不要编号
- 不要额外提醒
- 不要输出字段标签
- 英文标题必须带标准英文标点
- 中文标题必须带中文标点
- 必须先写英文标题，再把英文标题逐段翻译成中文标题
- 中文标题不得新增英文标题中没有的信息
- 中文标题不得省略英文标题中已有的信息
- 先稳定识别，再生成标题，不要为了变化而变化
- 如果结果没有标点或结构混乱，先自行修正后再输出`;

function createDefaultProductPublishPromptPresets() {
    return [
        {
            id: 'default-general',
            name: '通用标题',
            doc: DEFAULT_PRODUCT_PUBLISH_PROMPT_DOC
        }
    ];
}

function createDefaultProductPublishSettingsPresets() {
    return [
        {
            id: 'default-publish',
            name: '默认发布',
            aiApiUrl: '',
            aiApiKey: '',
            aiModel: '',
            urlPrefix: '',
            ossBucket: '',
            ossRegion: '',
            ossAccessKeyId: '',
            ossAccessKeySecret: '',
            ossObjectPrefix: 'products'
        }
    ];
}

function createDefaultProductPublishAiPresets() {
    return [
        {
            id: 'default-ai',
            name: '默认AI配置',
            aiApiUrl: 'https://www.vivaapi.cn',
            aiApiKey: '',
            aiModel: 'gpt-5.4-nano-2026-03-17'
        }
    ];
}

function createDefaultProductPublishOssPresets() {
    return [
        {
            id: 'default-oss',
            name: '默认OSS配置',
            urlPrefix: 'https://imageflow.oss-cn-hangzhou.aliyuncs.com',
            ossBucket: 'imageflow',
            ossRegion: 'oss-cn-hangzhou',
            ossAccessKeyId: '',
            ossAccessKeySecret: '',
            ossObjectPrefix: 'products'
        }
    ];
}

function normalizeProductPublishPromptPresets(presets, selectedId = '') {
    const defaultPresets = createDefaultProductPublishPromptPresets();
    const normalized = [];
    const seenIds = new Set();
    const hasExplicitPresets = Array.isArray(presets);
    (Array.isArray(presets) ? presets : []).forEach((preset, index) => {
        const id = String(preset?.id || `preset-${Date.now()}-${index + 1}`).trim();
        const name = String(preset?.name || '').trim();
        const doc = String(preset?.doc || '').trim();
        if (!id || !name || !doc || seenIds.has(id)) {
            return;
        }
        seenIds.add(id);
        normalized.push({ id, name, doc });
    });
    if (!hasExplicitPresets) {
        defaultPresets.forEach((preset) => {
            if (seenIds.has(preset.id)) {
                return;
            }
            seenIds.add(preset.id);
            normalized.unshift({ ...preset });
        });
    }
    const activeId = normalized.some((preset) => preset.id === selectedId)
        ? selectedId
        : (normalized[0]?.id || '');
    const activePreset = normalized.find((preset) => preset.id === activeId) || normalized[0] || null;
    return {
        presets: normalized,
        activeId,
        activePreset
    };
}

function normalizeProductPublishSettingsPresets(presets, selectedId = '', currentValues = {}) {
    const defaults = createDefaultProductPublishSettingsPresets();
    const normalized = (Array.isArray(presets) ? presets : [])
        .map((preset, index) => ({
            id: String(preset?.id || `publish-preset-${Date.now()}-${index + 1}`).trim(),
            name: String(preset?.name || '未命名发布预设').trim() || '未命名发布预设',
            aiApiUrl: String(preset?.aiApiUrl || '').trim(),
            aiApiKey: String(preset?.aiApiKey || '').trim(),
            aiModel: String(preset?.aiModel || '').trim(),
            urlPrefix: String(preset?.urlPrefix || '').trim(),
            ossBucket: String(preset?.ossBucket || '').trim(),
            ossRegion: String(preset?.ossRegion || '').trim(),
            ossAccessKeyId: String(preset?.ossAccessKeyId || '').trim(),
            ossAccessKeySecret: String(preset?.ossAccessKeySecret || '').trim(),
            ossObjectPrefix: String(preset?.ossObjectPrefix || 'products').trim() || 'products'
        }))
        .filter((preset) => preset.id && preset.name);
    if (!normalized.length) {
        normalized.push(...defaults);
    }
    let activeId = String(selectedId || normalized[0]?.id || defaults[0].id).trim();
    let activePreset = normalized.find((preset) => preset.id === activeId);
    if (!activePreset) {
        activeId = normalized[0]?.id || defaults[0].id;
        activePreset = normalized.find((preset) => preset.id === activeId) || normalized[0] || defaults[0];
    }
    const current = {
        aiApiUrl: String(currentValues?.aiApiUrl || '').trim(),
        aiApiKey: String(currentValues?.aiApiKey || '').trim(),
        aiModel: String(currentValues?.aiModel || '').trim(),
        urlPrefix: String(currentValues?.urlPrefix || '').trim(),
        ossBucket: String(currentValues?.ossBucket || '').trim(),
        ossRegion: String(currentValues?.ossRegion || '').trim(),
        ossAccessKeyId: String(currentValues?.ossAccessKeyId || '').trim(),
        ossAccessKeySecret: String(currentValues?.ossAccessKeySecret || '').trim(),
        ossObjectPrefix: String(currentValues?.ossObjectPrefix || 'products').trim() || 'products'
    };
    const activeIndex = normalized.findIndex((preset) => preset.id === activeId);
    if (activeIndex >= 0) {
        normalized[activeIndex] = {
            ...normalized[activeIndex],
            ...current
        };
        activePreset = normalized[activeIndex];
    }
    return {
        presets: normalized,
        activeId,
        activePreset
    };
}

function normalizeProductPublishAiPresets(presets, selectedId = '', currentValues = {}) {
    const defaults = createDefaultProductPublishAiPresets();
    const normalized = (Array.isArray(presets) ? presets : [])
        .map((preset, index) => ({
            id: String(preset?.id || `ai-preset-${Date.now()}-${index + 1}`).trim(),
            name: String(preset?.name || '未命名AI配置').trim() || '未命名AI配置',
            aiApiUrl: String(preset?.aiApiUrl || '').trim(),
            aiApiKey: String(preset?.aiApiKey || '').trim(),
            aiModel: String(preset?.aiModel || '').trim()
        }))
        .filter((preset) => preset.id && preset.name);
    if (!normalized.length) normalized.push(...defaults);
    let activeId = String(selectedId || normalized[0]?.id || defaults[0].id).trim();
    let activePreset = normalized.find((preset) => preset.id === activeId);
    if (!activePreset) {
        activeId = normalized[0]?.id || defaults[0].id;
        activePreset = normalized.find((preset) => preset.id === activeId) || normalized[0] || defaults[0];
    }
    const current = {
        aiApiUrl: String(currentValues?.aiApiUrl || '').trim(),
        aiApiKey: String(currentValues?.aiApiKey || '').trim(),
        aiModel: String(currentValues?.aiModel || '').trim()
    };
    const activeIndex = normalized.findIndex((preset) => preset.id === activeId);
    if (activeIndex >= 0) {
        normalized[activeIndex] = { ...normalized[activeIndex], ...current };
        activePreset = normalized[activeIndex];
    }
    return { presets: normalized, activeId, activePreset };
}

function normalizeProductPublishOssPresets(presets, selectedId = '', currentValues = {}) {
    const defaults = createDefaultProductPublishOssPresets();
    const normalized = (Array.isArray(presets) ? presets : [])
        .map((preset, index) => ({
            id: String(preset?.id || `oss-preset-${Date.now()}-${index + 1}`).trim(),
            name: String(preset?.name || '未命名OSS配置').trim() || '未命名OSS配置',
            urlPrefix: String(preset?.urlPrefix || '').trim(),
            ossBucket: String(preset?.ossBucket || '').trim(),
            ossRegion: String(preset?.ossRegion || '').trim(),
            ossAccessKeyId: String(preset?.ossAccessKeyId || '').trim(),
            ossAccessKeySecret: String(preset?.ossAccessKeySecret || '').trim(),
            ossObjectPrefix: String(preset?.ossObjectPrefix || 'products').trim() || 'products'
        }))
        .filter((preset) => preset.id && preset.name);
    if (!normalized.length) normalized.push(...defaults);
    let activeId = String(selectedId || normalized[0]?.id || defaults[0].id).trim();
    let activePreset = normalized.find((preset) => preset.id === activeId);
    if (!activePreset) {
        activeId = normalized[0]?.id || defaults[0].id;
        activePreset = normalized.find((preset) => preset.id === activeId) || normalized[0] || defaults[0];
    }
    const current = {
        urlPrefix: String(currentValues?.urlPrefix || '').trim(),
        ossBucket: String(currentValues?.ossBucket || '').trim(),
        ossRegion: String(currentValues?.ossRegion || '').trim(),
        ossAccessKeyId: String(currentValues?.ossAccessKeyId || '').trim(),
        ossAccessKeySecret: String(currentValues?.ossAccessKeySecret || '').trim(),
        ossObjectPrefix: String(currentValues?.ossObjectPrefix || 'products').trim() || 'products'
    };
    const activeIndex = normalized.findIndex((preset) => preset.id === activeId);
    if (activeIndex >= 0) {
        normalized[activeIndex] = { ...normalized[activeIndex], ...current };
        activePreset = normalized[activeIndex];
    }
    return { presets: normalized, activeId, activePreset };
}

function createDefaultProductPublishConfig() {
    return {
        aiProvider: 'auto',
        aiApiUrl: 'https://www.vivaapi.cn',
        aiApiKey: '',
        aiModel: 'gpt-5.4-nano-2026-03-17',
        aiModelHistory: [
            'gpt-5.4-nano-2026-03-17',
            'gpt-5.4',
            'gpt-5.3-codex'
        ],
        aiPresetId: 'default-ai',
        aiPresets: createDefaultProductPublishAiPresets(),
        ossPresetId: 'default-oss',
        ossPresets: createDefaultProductPublishOssPresets(),
        settingsPresetId: 'default-publish',
        settingsPresets: createDefaultProductPublishSettingsPresets(),
        titlePromptPresetId: 'default-general',
        titlePromptPresets: createDefaultProductPublishPromptPresets(),
        titlePromptDoc: DEFAULT_PRODUCT_PUBLISH_PROMPT_DOC,
        productTypeMappings: createDefaultProductPublishTypeMappings(),
        exportTemplateDefaults: {
            mainCodePrefix: 'A',
            categoryId: '124300',
            outputDir: getDefaultProductPublishOutputDir(),
            urlPrefix: 'https://imageflow.oss-cn-hangzhou.aliyuncs.com',
            ossBucket: 'imageflow',
            ossRegion: 'oss-cn-hangzhou',
            ossAccessKeyId: '',
            ossAccessKeySecret: '',
            ossObjectPrefix: 'products',
            shipLeadTime: '2',
            originPlace: '中国-浙江省',
            customized: '否',
            specName1: '颜色',
            specName2: '尺寸',
            specValue1: '白色',
            specValue2: 'xl',
            declaredPrice: '10',
            suggestedPrice: '10',
            lengthCm: '10',
            widthCm: '10',
            heightCm: '10',
            weightG: '10',
            inventory: '10',
            sensitive: '否'
        },
        exportTemplateProfiles: [
            {
                id: 'default-flannel-rug',
                name: '法兰绒地垫',
                fields: {
                    mainCodePrefix: 'A',
                    categoryId: '124300',
                    outputDir: '',
                    urlPrefix: 'https://imageflow.oss-cn-hangzhou.aliyuncs.com',
                    ossBucket: 'imageflow',
                    ossRegion: 'oss-cn-hangzhou',
                    ossAccessKeyId: '',
                    ossAccessKeySecret: '',
                    ossObjectPrefix: 'products',
                    shipLeadTime: '2',
                    originPlace: '中国-浙江省',
                    customized: '否',
                    specName1: '颜色',
                    specName2: '尺寸',
                    specValue1: '白色',
                    specValue2: 'xl',
                    declaredPrice: '10',
                    suggestedPrice: '10',
                    lengthCm: '10',
                    widthCm: '10',
                    heightCm: '10',
                    weightG: '10',
                    inventory: '10',
                    sensitive: '否'
                }
            }
        ]
    };
}

function normalizeProductPublishAiProvider(value) {
    const provider = String(value || '').trim().toLowerCase();
    if (provider === 'openai') return 'openai';
    if (provider === 'gemini') return 'gemini';
    if (provider === 'claude') return 'claude';
    if (provider === 'text') return 'text';
    return 'auto';
}

const PRODUCT_PUBLISH_GEMINI_MODELS = [
    'gemini-1.5-flash',
    'gemini-1.5-pro',
    'gemini-2.0-flash',
    'gemini-2.0-flash-exp',
    'gemini-2.0-flash-exp-image-generation',
    'gemini-2.0-flash-thinking-exp',
    'gemini-2.5-flash',
    'gemini-2.5-flash-image',
    'gemini-2.5-flash-image-preview',
    'gemini-2.5-pro',
    'gemini-3-pro-preview',
    'gemini-3-pro-image-preview',
    'gemini-3.1-flash-image-preview',
    'gemini-embedding-001',
    'gemini-image-editing'
];

const PRODUCT_PUBLISH_CLAUDE_MODELS = [
    'claude-3-haiku',
    'claude-3-sonnet',
    'claude-3-opus',
    'claude-3-5-haiku',
    'claude-3-5-haiku-latest',
    'claude-3-5-sonnet',
    'claude-3-5-sonnet-latest',
    'claude-3-7-sonnet',
    'claude-3-7-sonnet-latest',
    'claude-4-sonnet',
    'claude-4-opus',
    'claude-sonnet-4-20250514',
    'claude-opus-4-20250514'
];

const PRODUCT_PUBLISH_GATEWAY_EXTRA_MODELS = [
    'chatgpt-4o-latest',
    'gpt-4.1',
    'gpt-4.1-mini',
    'gpt-4o',
    'gpt-4o-mini',
    'gpt-5',
    'gpt-5.4',
    'o3',
    'o3-pro',
    'o4-mini',
    'deepseek-chat',
    'deepseek-reasoner',
    'grok-3',
    'glm4',
    'glm-4',
    'glm-4-plus',
    'qwen-max',
    'qwen-plus',
    'qwen-turbo',
    'qwen2.5-72b-instruct',
    'kimi-k2',
    'doubao-seedance-1-5-pro-251215'
];

function loadProductPublishConfig() {
    const defaults = createDefaultProductPublishConfig();
    try {
        const parsed = JSON.parse(fs.readFileSync(PRODUCT_PUBLISH_CONFIG_FILE, 'utf-8'));
        const cfg = {
            ...defaults,
            ...(parsed || {})
        };
        cfg.productTypeMappings = normalizeProductPublishTypeMappings(cfg.productTypeMappings);
        return cfg;
    } catch {
        return defaults;
    }
}

function saveProductPublishConfig(cfg) {
    const nextCfg = {
        ...createDefaultProductPublishConfig(),
        ...(cfg || {})
    };
    nextCfg.productTypeMappings = normalizeProductPublishTypeMappings(nextCfg.productTypeMappings);
    nextCfg.aiModelHistory = Array.from(new Set((Array.isArray(nextCfg.aiModelHistory) ? nextCfg.aiModelHistory : [])
        .map((item) => String(item || '').trim())
        .filter(Boolean))).slice(0, 30);
    nextCfg.exportTemplateDefaults = {
        ...createDefaultProductPublishConfig().exportTemplateDefaults,
        ...(nextCfg.exportTemplateDefaults || {})
    };
    nextCfg.exportTemplateDefaults.outputDir = normalizeDirectoryPath(
        nextCfg.exportTemplateDefaults.outputDir,
        getDefaultProductPublishOutputDir()
    );
    const aiPresetState = normalizeProductPublishAiPresets(
        nextCfg.aiPresets,
        nextCfg.aiPresetId,
        {
            aiApiUrl: nextCfg.aiApiUrl,
            aiApiKey: nextCfg.aiApiKey,
            aiModel: nextCfg.aiModel
        }
    );
    nextCfg.aiPresets = aiPresetState.presets;
    nextCfg.aiPresetId = aiPresetState.activeId;
    nextCfg.aiApiUrl = aiPresetState.activePreset?.aiApiUrl || '';
    nextCfg.aiApiKey = aiPresetState.activePreset?.aiApiKey || '';
    nextCfg.aiModel = aiPresetState.activePreset?.aiModel || '';
    const ossPresetState = normalizeProductPublishOssPresets(
        nextCfg.ossPresets,
        nextCfg.ossPresetId,
        {
            urlPrefix: nextCfg.exportTemplateDefaults?.urlPrefix,
            ossBucket: nextCfg.exportTemplateDefaults?.ossBucket,
            ossRegion: nextCfg.exportTemplateDefaults?.ossRegion,
            ossAccessKeyId: nextCfg.exportTemplateDefaults?.ossAccessKeyId,
            ossAccessKeySecret: nextCfg.exportTemplateDefaults?.ossAccessKeySecret,
            ossObjectPrefix: nextCfg.exportTemplateDefaults?.ossObjectPrefix
        }
    );
    nextCfg.ossPresets = ossPresetState.presets;
    nextCfg.ossPresetId = ossPresetState.activeId;
    nextCfg.exportTemplateDefaults = {
        ...nextCfg.exportTemplateDefaults,
        urlPrefix: ossPresetState.activePreset?.urlPrefix || '',
        ossBucket: ossPresetState.activePreset?.ossBucket || '',
        ossRegion: ossPresetState.activePreset?.ossRegion || '',
        ossAccessKeyId: ossPresetState.activePreset?.ossAccessKeyId || '',
        ossAccessKeySecret: ossPresetState.activePreset?.ossAccessKeySecret || '',
        ossObjectPrefix: ossPresetState.activePreset?.ossObjectPrefix || 'products'
    };
    const promptPresetState = normalizeProductPublishPromptPresets(nextCfg.titlePromptPresets, nextCfg.titlePromptPresetId);
    nextCfg.titlePromptPresets = promptPresetState.presets;
    nextCfg.titlePromptPresetId = promptPresetState.activeId;
    nextCfg.titlePromptDoc = String(
        nextCfg.titlePromptDoc || promptPresetState.activePreset?.doc || DEFAULT_PRODUCT_PUBLISH_PROMPT_DOC
    ).trim() || DEFAULT_PRODUCT_PUBLISH_PROMPT_DOC;
    const selectedPresetIndex = nextCfg.titlePromptPresets.findIndex((preset) => preset.id === nextCfg.titlePromptPresetId);
    if (selectedPresetIndex >= 0) {
        nextCfg.titlePromptPresets[selectedPresetIndex] = {
            ...nextCfg.titlePromptPresets[selectedPresetIndex],
            doc: nextCfg.titlePromptDoc
        };
    }
    fs.writeFileSync(PRODUCT_PUBLISH_CONFIG_FILE, JSON.stringify(nextCfg, null, 2), 'utf-8');
    return nextCfg;
}

function createDefaultPrintAiConfig() {
    let productCfg = {};
    try {
        productCfg = loadProductPublishConfig();
    } catch {}
    return {
        baseUrl: String(productCfg.aiApiUrl || '').trim(),
        apiKey: String(productCfg.aiApiKey || '').trim(),
        extractModel: String(productCfg.aiModel || '').trim(),
        variationModel: String(productCfg.aiModel || '').trim(),
        timeoutMs: 300000,
        concurrency: 1,
        variationCount: 1,
        aspectRatio: '3:2',
        outputDir: path.join(app.getPath('pictures'), 'ImageFlow印花裂变'),
        extractionPrompt: '完整提取商品表面的印花图案，去除背景、透视、阴影和材质干扰，输出独立正视角图案图片。保持主体完整、画面干净、无产品场景。',
        variationPrompts: [
            {
                id: 'print-ai-variation-default',
                name: '通用裂变',
                prompt: '基于提取出的印花图案生成新的同风格图案，保持电商印花可用、构图完整、背景干净，不要生成产品场景。'
            }
        ]
    };
}

function normalizePrintAiConfig(cfg) {
    const defaults = createDefaultPrintAiConfig();
    const variationPrompts = Array.isArray(cfg?.variationPrompts)
        ? cfg.variationPrompts.map((item, index) => ({
            id: String(item?.id || `print-ai-variation-${Date.now()}-${index + 1}`).trim(),
            name: String(item?.name || `裂变方式 ${index + 1}`).trim(),
            prompt: String(item?.prompt || '').trim()
        })).filter((item) => item.id && item.name && item.prompt)
        : defaults.variationPrompts;
    return {
        ...defaults,
        ...(cfg || {}),
        baseUrl: String(cfg?.baseUrl ?? defaults.baseUrl).trim(),
        apiKey: String(cfg?.apiKey ?? defaults.apiKey).trim(),
        extractModel: String(cfg?.extractModel ?? defaults.extractModel).trim(),
        variationModel: String(cfg?.variationModel ?? defaults.variationModel).trim(),
        timeoutMs: Math.max(1000, Math.min(900000, Number(cfg?.timeoutMs || defaults.timeoutMs))),
        concurrency: Math.max(1, Math.min(6, Math.round(Number(cfg?.concurrency || defaults.concurrency)))),
        variationCount: Math.max(1, Math.min(8, Math.round(Number(cfg?.variationCount || defaults.variationCount)))),
        aspectRatio: ['3:2', '21:9'].includes(cfg?.aspectRatio) ? cfg.aspectRatio : defaults.aspectRatio,
        outputDir: normalizeDirectoryPath(cfg?.outputDir, defaults.outputDir),
        extractionPrompt: String(cfg?.extractionPrompt || defaults.extractionPrompt).trim(),
        variationPrompts: variationPrompts.length ? variationPrompts : defaults.variationPrompts
    };
}

function loadPrintAiConfig() {
    return normalizePrintAiConfig(readJsonFile(PRINT_AI_CONFIG_FILE, null));
}

function savePrintAiConfig(cfg) {
    const nextCfg = normalizePrintAiConfig(cfg || {});
    writeJsonFile(PRINT_AI_CONFIG_FILE, nextCfg);
    return nextCfg;
}

function createDefaultPrintAiData() {
    return { tasks: [] };
}

function normalizePrintAiTask(task, index = 0) {
    const sourcePath = String(task?.sourcePath || '').trim();
    const sourceName = String(task?.sourceName || path.basename(sourcePath || `印花任务${index + 1}`)).trim();
    const variants = Array.isArray(task?.variants)
        ? task.variants.map((item, itemIndex) => ({
            id: String(item?.id || `variant-${itemIndex + 1}`).trim(),
            promptId: String(item?.promptId || '').trim(),
            promptName: String(item?.promptName || '').trim(),
            status: ['pending', 'running', 'done', 'failed'].includes(item?.status) ? item.status : 'pending',
            imagePath: String(item?.imagePath || '').trim(),
            error: String(item?.error || '').trim()
        }))
        : [];
    let status = ['pending', 'running', 'extracted', 'done', 'failed'].includes(task?.status) ? task.status : 'pending';
    const hasVariants = variants.length > 0;
    const variantsTerminal = hasVariants && variants.every((item) => item.status === 'done' || item.status === 'failed');
    if ((status === 'running' || status === 'extracted') && variantsTerminal) {
        status = variants.some((item) => item.status === 'done') ? 'done' : 'failed';
    }
    return {
        id: String(task?.id || `print-ai-task-${Date.now()}-${index + 1}`).trim(),
        sourceName,
        sourcePath,
        status,
        extractedPath: String(task?.extractedPath || '').trim(),
        description: task?.description || null,
        variants,
        error: String(task?.error || '').trim(),
        createdAt: String(task?.createdAt || new Date().toISOString()),
        updatedAt: String(task?.updatedAt || task?.createdAt || new Date().toISOString())
    };
}

function loadPrintAiData() {
    const data = readJsonFile(PRINT_AI_DATA_FILE, createDefaultPrintAiData());
    return {
        tasks: Array.isArray(data?.tasks)
            ? data.tasks.map((task, index) => normalizePrintAiTask(task, index))
            : []
    };
}

function savePrintAiData(data) {
    const nextData = {
        tasks: Array.isArray(data?.tasks)
            ? data.tasks.map((task, index) => normalizePrintAiTask(task, index))
            : []
    };
    writeJsonFile(PRINT_AI_DATA_FILE, nextData);
    return nextData;
}

function broadcastPrintAiTasks(sender, data = loadPrintAiData()) {
    const target = sender || mainWindow?.webContents;
    if (target) {
        safeSend(target, 'print-ai:tasks', data);
    }
}

function resolvePrintAiApiRoot(rawUrl) {
    return String(rawUrl || '').trim().replace(/\/+$/, '').replace(/\/(?:images\/edits|images\/generations|chat\/completions|models)\/?$/i, '');
}

function resolvePrintAiApiUrl(baseUrl, suffix) {
    const root = resolvePrintAiApiRoot(baseUrl);
    if (!root) return '';
    return `${root}/${String(suffix || '').replace(/^\/+/, '')}`;
}

async function detectPrintAiModels(config) {
    const modelsUrl = resolvePrintAiApiUrl(config.baseUrl, '/models');
    if (!modelsUrl) {
        throw new Error('请先填写 API 地址');
    }
    const headers = {};
    if (config.apiKey) {
        headers.Authorization = `Bearer ${config.apiKey}`;
    }
    const response = await fetch(modelsUrl, {
        method: 'GET',
        headers,
        signal: AbortSignal.timeout(config.timeoutMs)
    });
    const rawText = await response.text().catch(() => '');
    if (!response.ok) {
        throw new Error(`模型识别失败：${response.status}${rawText ? ` ${rawText.slice(0, 500)}` : ''}`);
    }
    let payload = {};
    try {
        payload = rawText ? JSON.parse(rawText) : {};
    } catch {
        throw new Error(`模型识别失败：接口返回的不是合法 JSON${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
    }
    const source = Array.isArray(payload)
        ? payload
        : (Array.isArray(payload?.data) ? payload.data : (Array.isArray(payload?.models) ? payload.models : []));
    const models = source
        .map((item) => {
            if (typeof item === 'string') return item;
            return String(item?.id || item?.name || item?.model || '').trim();
        })
        .filter(Boolean);
    const uniqueModels = Array.from(new Set(models));
    if (!uniqueModels.length) {
        throw new Error('模型识别失败：接口没有返回可用模型');
    }
    return {
        models: uniqueModels,
        preferredModel: uniqueModels.includes('gpt-image-2') ? 'gpt-image-2' : uniqueModels[0],
        provider: 'openai'
    };
}

async function testPrintAiModelAvailability(config, model) {
    const result = await detectPrintAiModels(config);
    if (!result.models.includes(model)) {
        throw new Error(`当前 API Key 分组未返回该模型：${model}`);
    }
    return {
        ok: true,
        model,
        message: `模型可用：${model}`,
        total: result.models.length
    };
}

function getPrintAiImageSize(aspectRatio) {
    return aspectRatio === '21:9' ? '2688x1152' : '1536x1024';
}

function getPrintAiMimeType(filePath) {
    const ext = path.extname(filePath || '').toLowerCase();
    if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
    if (ext === '.webp') return 'image/webp';
    if (ext === '.gif') return 'image/gif';
    if (ext === '.avif') return 'image/avif';
    return 'image/png';
}

function printAiExtensionForMime(mimeType) {
    const normalized = String(mimeType || '').toLowerCase();
    if (normalized.includes('jpeg') || normalized.includes('jpg')) return '.jpg';
    if (normalized.includes('webp')) return '.webp';
    if (normalized.includes('gif')) return '.gif';
    if (normalized.includes('avif')) return '.avif';
    return '.png';
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRetryablePrintAiStatus(status) {
    return [408, 409, 425, 429, 500, 502, 503, 504, 520, 522, 524].includes(Number(status));
}

function isRetryablePrintAiError(error) {
    if (isRetryablePrintAiStatus(error?.status)) return true;
    const name = String(error?.name || '');
    const message = String(error?.message || '');
    return ['AbortError', 'TimeoutError'].includes(name) || /fetch failed|network|timeout|timed out/i.test(message);
}

function compactPrintAiErrorMessage(error, fallback = '图片生成失败', actionText = '请稍后重试或切换模型') {
    const raw = String(error?.message || fallback || '').replace(/\s+/g, ' ').trim();
    const status = Number(error?.status) || Number((raw.match(/：(\d{3})/) || raw.match(/\b([45]\d{2})\b/) || [])[1]) || 0;
    const suffix = actionText ? `，${actionText}` : '';
    if (status === 524) return `图片生成接口超时 524${suffix}`;
    if (status === 429) return `图片生成接口限流 429${suffix}`;
    if (status >= 500) return `图片生成接口异常 ${status}${suffix}`;
    if (raw && /<!doctype|<html/i.test(raw)) {
        return `图片生成接口返回了 HTML 错误页${status ? ` ${status}` : ''}${suffix}`;
    }
    return raw.slice(0, 180) || fallback;
}

async function requestPrintAiImageEdit(input) {
    const apiUrl = resolvePrintAiApiUrl(input.baseUrl, '/images/edits');
    if (!apiUrl || !input.model) {
        throw new Error('请先填写 API 地址和模型名称');
    }
    const bytes = fs.readFileSync(input.imagePath);
    const mimeType = getPrintAiMimeType(input.imagePath);
    const headers = {};
    if (input.apiKey) {
        headers.Authorization = `Bearer ${input.apiKey}`;
    }
    const attempts = Math.max(1, Math.min(4, Math.round(Number(input.retryAttempts) || 3)));
    let lastError = null;
    for (let attempt = 1; attempt <= attempts; attempt += 1) {
        try {
            const form = new FormData();
            form.set('model', input.model);
            form.set('prompt', input.prompt);
            form.set('response_format', 'b64_json');
            if (input.size) {
                form.set('size', input.size);
            }
            form.set('image', new Blob([bytes], { type: mimeType }), path.basename(input.imagePath));
            const response = await fetch(apiUrl, {
                method: 'POST',
                headers,
                body: form,
                signal: AbortSignal.timeout(input.timeoutMs)
            });
            const rawText = await response.text().catch(() => '');
            if (!response.ok) {
                const shortBody = rawText && /<!doctype|<html/i.test(rawText)
                    ? '接口返回 HTML 错误页'
                    : rawText.slice(0, 260);
                const error = new Error(`图片生成失败：${response.status}${shortBody ? ` ${shortBody}` : ''}`);
                error.status = response.status;
                error.rawBody = rawText;
                throw error;
            }
            let payload = {};
            try {
                payload = rawText ? JSON.parse(rawText) : {};
            } catch {
                throw new Error(`图片生成失败：接口返回的不是合法 JSON${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
            }
            const first = Array.isArray(payload?.data) ? payload.data[0] : null;
            if (first?.b64_json) {
                return { bytes: Buffer.from(first.b64_json, 'base64'), mimeType: 'image/png' };
            }
            if (first?.url) {
                const imageResponse = await fetch(first.url, input.apiKey ? { headers: { Authorization: `Bearer ${input.apiKey}` } } : {});
                if (!imageResponse.ok) {
                    const error = new Error(`下载生成图失败：${imageResponse.status}`);
                    error.status = imageResponse.status;
                    throw error;
                }
                return {
                    bytes: Buffer.from(await imageResponse.arrayBuffer()),
                    mimeType: imageResponse.headers.get('content-type') || 'image/png'
                };
            }
            throw new Error('模型没有返回图片数据');
        } catch (error) {
            lastError = error;
            if (attempt >= attempts || !isRetryablePrintAiError(error)) {
                throw error;
            }
            await sleep(Math.min(12000, 1200 * (2 ** (attempt - 1))) + Math.round(Math.random() * 500));
        }
    }
    throw lastError || new Error('图片生成失败');
}

function savePrintAiGeneratedImage(bytes, mimeType, section, nameHint = '') {
    const outputDir = ensureDir(path.join(PRINT_AI_STORAGE_DIR, section));
    const safeName = sanitizePathSegment(path.basename(nameHint || section, path.extname(nameHint || ''))) || section;
    const filePath = ensureUniqueFilePath(path.join(outputDir, `${safeName}-${Date.now()}${printAiExtensionForMime(mimeType)}`));
    fs.writeFileSync(filePath, bytes);
    return filePath;
}

async function runLimitedPrintAiJobs(jobs, concurrency) {
    const safeConcurrency = Math.max(1, Math.min(Math.round(Number(concurrency) || 1), jobs.length || 1));
    let nextIndex = 0;
    const workers = Array.from({ length: safeConcurrency }, async () => {
        while (nextIndex < jobs.length) {
            const currentIndex = nextIndex;
            nextIndex += 1;
            await jobs[currentIndex]();
        }
    });
    await Promise.all(workers);
}

function importPrintAiImages(filePaths) {
    const data = loadPrintAiData();
    const uploadDir = ensureDir(path.join(PRINT_AI_STORAGE_DIR, 'uploads'));
    const now = new Date().toISOString();
    const imported = [];
    (Array.isArray(filePaths) ? filePaths : []).forEach((filePath) => {
        const sourcePath = String(filePath || '').trim();
        if (!sourcePath || !fs.existsSync(sourcePath)) return;
        const ext = path.extname(sourcePath).toLowerCase();
        if (!PRINT_AI_IMAGE_EXTS.has(ext)) return;
        const copiedPath = ensureUniqueFilePath(path.join(uploadDir, path.basename(sourcePath)));
        fs.copyFileSync(sourcePath, copiedPath);
        const task = normalizePrintAiTask({
            id: `print-ai-task-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`,
            sourceName: path.basename(sourcePath),
            sourcePath: copiedPath,
            status: 'pending',
            variants: [],
            createdAt: now,
            updatedAt: now
        });
        data.tasks.unshift(task);
        imported.push(task);
    });
    savePrintAiData(data);
    return { imported, ...data };
}

async function runPrintAiTask(task, cfg, sender) {
    const size = getPrintAiImageSize(cfg.aspectRatio);
    const extractionPrompt = [
        cfg.extractionPrompt,
        '',
        `Output aspect ratio: ${cfg.aspectRatio}. The extracted standalone print image must use this exact canvas ratio.`
    ].join('\n');
    const extracted = await requestPrintAiImageEdit({
        baseUrl: cfg.baseUrl,
        apiKey: cfg.apiKey,
        model: cfg.extractModel,
        timeoutMs: cfg.timeoutMs,
        imagePath: task.sourcePath,
        prompt: extractionPrompt,
        size
    });
    task.extractedPath = savePrintAiGeneratedImage(extracted.bytes, extracted.mimeType, 'extracted', task.sourceName);
    task.status = 'extracted';
    task.error = '';
    task.updatedAt = new Date().toISOString();
    const dataAfterExtract = loadPrintAiData();
    dataAfterExtract.tasks = dataAfterExtract.tasks.map((item) => item.id === task.id ? task : item);
    savePrintAiData(dataAfterExtract);
    broadcastPrintAiTasks(sender, dataAfterExtract);

    const prompts = cfg.variationPrompts.slice(0, 12);
    const variants = [];
    const variantJobs = [];
    for (const prompt of prompts) {
        for (let index = 0; index < cfg.variationCount; index += 1) {
            const variant = {
                id: `${task.id}-${prompt.id}-${index + 1}`,
                promptId: prompt.id,
                promptName: prompt.name,
                status: 'running',
                imagePath: '',
                error: ''
            };
            variants.push(variant);
            variantJobs.push(async () => {
                variant.status = 'running';
                variant.error = '';
                task.variants = variants.slice();
                task.updatedAt = new Date().toISOString();
                const runningData = loadPrintAiData();
                runningData.tasks = runningData.tasks.map((item) => item.id === task.id ? task : item);
                savePrintAiData(runningData);
                broadcastPrintAiTasks(sender, runningData);

                try {
                    const variationPrompt = [
                        prompt.prompt,
                        '',
                        `Aspect ratio: ${cfg.aspectRatio}. Use the extracted print image as reference and preserve the canvas ratio.`
                    ].join('\n');
                    const generated = await requestPrintAiImageEdit({
                        baseUrl: cfg.baseUrl,
                        apiKey: cfg.apiKey,
                        model: cfg.variationModel,
                        timeoutMs: cfg.timeoutMs,
                        imagePath: task.extractedPath,
                        prompt: variationPrompt,
                        size
                    });
                    variant.status = 'done';
                    variant.imagePath = savePrintAiGeneratedImage(generated.bytes, generated.mimeType, 'variants', `${task.sourceName}-${prompt.name}`);
                } catch (error) {
                    variant.status = 'failed';
                    variant.error = error.message || '裂变失败';
                }
                task.variants = variants.slice();
                task.updatedAt = new Date().toISOString();
                const doneData = loadPrintAiData();
                doneData.tasks = doneData.tasks.map((item) => item.id === task.id ? task : item);
                savePrintAiData(doneData);
                broadcastPrintAiTasks(sender, doneData);
            });
        }
    }
    task.variants = variants.slice();
    task.updatedAt = new Date().toISOString();
    const queuedData = loadPrintAiData();
    queuedData.tasks = queuedData.tasks.map((item) => item.id === task.id ? task : item);
    savePrintAiData(queuedData);
    broadcastPrintAiTasks(sender, queuedData);

    await runLimitedPrintAiJobs(variantJobs, cfg.concurrency);
    task.variants = variants;
    task.status = variants.some((item) => item.status === 'done') ? 'done' : 'failed';
    task.error = task.status === 'failed' ? (variants.find((item) => item.error)?.error || '裂变失败') : '';
    task.updatedAt = new Date().toISOString();
    const finalData = loadPrintAiData();
    finalData.tasks = finalData.tasks.map((item) => item.id === task.id ? task : item);
    savePrintAiData(finalData);
    broadcastPrintAiTasks(sender, finalData);
}

async function startPrintAiRun(payload, sender) {
    const cfg = savePrintAiConfig({
        ...loadPrintAiConfig(),
        ...(payload?.config || {})
    });
    if (!cfg.baseUrl || !cfg.extractModel || !cfg.variationModel) {
        throw new Error('请先填写 API 地址、提取模型和裂变模型');
    }
    const selectedIds = new Set(Array.isArray(payload?.taskIds) ? payload.taskIds.map((id) => String(id || '').trim()).filter(Boolean) : []);
    if (!selectedIds.size) {
        throw new Error('请先选择要运行的印花任务');
    }
    const data = loadPrintAiData();
    const tasks = data.tasks.filter((task) => selectedIds.has(task.id));
    if (!tasks.length) {
        throw new Error('没有找到可运行的印花任务');
    }
    data.tasks = data.tasks.map((task) => selectedIds.has(task.id)
        ? { ...task, status: 'running', error: '', variants: [], updatedAt: new Date().toISOString() }
        : task);
    savePrintAiData(data);
    broadcastPrintAiTasks(sender, data);

    void (async () => {
        await runLimitedPrintAiJobs(tasks.map((original) => async () => {
            const latest = loadPrintAiData().tasks.find((item) => item.id === original.id) || original;
            try {
                await runPrintAiTask({ ...latest, status: 'running', error: '', variants: [] }, cfg, sender);
            } catch (error) {
                const failedData = loadPrintAiData();
                failedData.tasks = failedData.tasks.map((item) => item.id === original.id
                    ? {
                        ...item,
                        status: 'failed',
                        error: error.message || '印花裂变失败',
                        updatedAt: new Date().toISOString()
                    }
                    : item);
                savePrintAiData(failedData);
                broadcastPrintAiTasks(sender, failedData);
            }
        }), cfg.concurrency);
        broadcastPrintAiTasks(sender);
    })();

    return loadPrintAiData();
}

const WORKFLOW_NODE_DEFS = [
    { type: 'input', label: '导入图片/文件夹', order: 0, branchable: false },
    { type: 'extract', label: '提取图案', order: 10, branchable: false },
    { type: 'variation', label: '印花裂变', order: 20, branchable: true },
    { type: 'slice', label: '智能切片', order: 30, branchable: false },
    { type: 'template', label: '智能模板', order: 40, branchable: true },
    { type: 'watermark', label: '智能水印', order: 50, branchable: false },
    { type: 'publish', label: '产品发布', order: 60, branchable: false },
    { type: 'export', label: '导出保存', order: 70, branchable: false }
];

function getWorkflowNodeDef(type) {
    return WORKFLOW_NODE_DEFS.find((item) => item.type === type) || null;
}

function createDefaultWorkflowGraph() {
    const nodes = [
        { id: 'node-input', type: 'input', title: '导入图片/文件夹', x: -200, y: 140, config: {} },
        { id: 'node-extract', type: 'extract', title: '提取图案', x: 80, y: 140, config: {} },
        { id: 'node-template', type: 'template', title: '智能模板', x: 360, y: 140, config: {} },
        { id: 'node-watermark', type: 'watermark', title: '智能水印', x: 640, y: 140, config: {} },
        { id: 'node-publish', type: 'publish', title: '产品发布', x: 920, y: 140, config: {} },
        { id: 'node-export', type: 'export', title: '导出保存', x: 1200, y: 140, config: {} }
    ];
    return {
        nodes,
        links: [
            { from: 'node-input', to: 'node-extract' },
            { from: 'node-extract', to: 'node-template' },
            { from: 'node-template', to: 'node-watermark' },
            { from: 'node-watermark', to: 'node-publish' },
            { from: 'node-publish', to: 'node-export' }
        ]
    };
}

function normalizeWorkflowNode(node, index = 0) {
    const type = String(node?.type || '').trim();
    const def = getWorkflowNodeDef(type) || WORKFLOW_NODE_DEFS[0];
    const id = String(node?.id || `workflow-node-${Date.now()}-${index + 1}`).trim();
    return {
        id,
        type: def.type,
        title: String(node?.title || def.label).trim() || def.label,
        x: Number.isFinite(Number(node?.x)) ? Number(node.x) : 80 + index * 240,
        y: Number.isFinite(Number(node?.y)) ? Number(node.y) : 120,
        config: node?.config && typeof node.config === 'object' ? node.config : {}
    };
}

function normalizeWorkflowGraph(graph) {
    const fallback = createDefaultWorkflowGraph();
    const hasExplicitNodes = Array.isArray(graph?.nodes);
    const nodes = Array.isArray(graph?.nodes)
        ? graph.nodes.map((node, index) => normalizeWorkflowNode(node, index))
        : fallback.nodes.map((node, index) => normalizeWorkflowNode(node, index));
    const nodeIds = new Set(nodes.map((node) => node.id));
    const rawLinks = Array.isArray(graph?.links) ? graph.links : (hasExplicitNodes ? [] : fallback.links);
    const links = rawLinks
        .map((link) => ({
            from: String(link?.from || '').trim(),
            to: String(link?.to || '').trim(),
            fromSlot: Math.max(0, Number(link?.fromSlot || 0)),
            toSlot: Math.max(0, Number(link?.toSlot || 0))
        })).filter((link) => nodeIds.has(link.from) && nodeIds.has(link.to) && link.from !== link.to);
    return { nodes, links };
}

function createDefaultWorkflowConfig() {
    const defaultGraph = createDefaultWorkflowGraph();
    return {
        activePresetId: 'workflow-default',
        presets: [
            {
                id: 'workflow-default',
                name: '提取后套模板发布',
                graph: defaultGraph,
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString()
            }
        ]
    };
}

function normalizeWorkflowConfig(config) {
    const fallback = createDefaultWorkflowConfig();
    const presets = Array.isArray(config?.presets) && config.presets.length
        ? config.presets.map((preset, index) => ({
            id: String(preset?.id || `workflow-preset-${Date.now()}-${index + 1}`).trim(),
            name: String(preset?.name || `工作流 ${index + 1}`).trim(),
            graph: normalizeWorkflowGraph(preset?.graph),
            createdAt: String(preset?.createdAt || new Date().toISOString()),
            updatedAt: String(preset?.updatedAt || preset?.createdAt || new Date().toISOString())
        })).filter((preset) => preset.id && preset.name)
        : fallback.presets;
    const activePresetId = String(config?.activePresetId || '').trim();
    return {
        activePresetId: presets.some((preset) => preset.id === activePresetId) ? activePresetId : presets[0].id,
        presets
    };
}

function loadWorkflowConfig() {
    return normalizeWorkflowConfig(readJsonFile(WORKFLOW_CONFIG_FILE, null));
}

function saveWorkflowConfig(config) {
    const nextConfig = normalizeWorkflowConfig(config || {});
    writeJsonFile(WORKFLOW_CONFIG_FILE, nextConfig);
    return nextConfig;
}

function createDefaultWorkflowData() {
    return { runs: [] };
}

function normalizeWorkflowRun(run, index = 0) {
    const id = String(run?.id || `workflow-run-${Date.now()}-${index + 1}`).trim();
    return {
        id,
        name: String(run?.name || `任务 ${index + 1}`).trim(),
        workflowPresetId: String(run?.workflowPresetId || '').trim(),
        sourceInputNodeId: String(run?.sourceInputNodeId || run?.inputNodeId || '').trim(),
        sourcePath: String(run?.sourcePath || '').trim(),
        sourceType: ['file', 'folder'].includes(run?.sourceType) ? run.sourceType : 'file',
        status: ['pending', 'running', 'paused', 'done', 'failed'].includes(run?.status) ? run.status : 'pending',
        currentNodeId: String(run?.currentNodeId || '').trim(),
        currentNodeType: String(run?.currentNodeType || '').trim(),
        progress: Math.max(0, Math.min(100, Number(run?.progress || 0))),
        artifacts: run?.artifacts && typeof run.artifacts === 'object' ? run.artifacts : {},
        nodeStates: run?.nodeStates && typeof run.nodeStates === 'object' ? run.nodeStates : {},
        error: String(run?.error || '').trim(),
        createdAt: String(run?.createdAt || new Date().toISOString()),
        updatedAt: String(run?.updatedAt || run?.createdAt || new Date().toISOString())
    };
}

function loadWorkflowData() {
    const data = readJsonFile(WORKFLOW_DATA_FILE, createDefaultWorkflowData());
    return {
        runs: Array.isArray(data?.runs)
            ? data.runs.map((run, index) => normalizeWorkflowRun(run, index))
            : []
    };
}

function saveWorkflowData(data) {
    const nextData = {
        runs: Array.isArray(data?.runs)
            ? data.runs.map((run, index) => normalizeWorkflowRun(run, index))
            : []
    };
    writeJsonFile(WORKFLOW_DATA_FILE, nextData);
    return nextData;
}

function broadcastWorkflowData(sender, data = loadWorkflowData()) {
    const target = sender || mainWindow?.webContents;
    if (target) {
        safeSend(target, 'workflow:data', data);
    }
}

function getWorkflowOutgoingLinks(graph, nodeId) {
    const normalizedNodeId = String(nodeId || '').trim();
    return (Array.isArray(graph?.links) ? graph.links : []).filter((link) => (
        String(link?.from || '').trim() === normalizedNodeId
    ));
}

function collectWorkflowReachableNodeIds(graph, startIds) {
    const nodesById = new Map((Array.isArray(graph?.nodes) ? graph.nodes : []).map((node) => [node.id, node]));
    const reachable = new Set();
    const stack = (Array.isArray(startIds) ? startIds : [startIds])
        .map((id) => String(id || '').trim())
        .filter((id) => id && nodesById.has(id));
    while (stack.length) {
        const nodeId = stack.shift();
        if (!nodeId || reachable.has(nodeId)) continue;
        reachable.add(nodeId);
        getWorkflowOutgoingLinks(graph, nodeId).forEach((link) => {
            const targetId = String(link?.to || '').trim();
            if (targetId && nodesById.has(targetId) && !reachable.has(targetId)) {
                stack.push(targetId);
            }
        });
    }
    return reachable;
}

function getWorkflowExecutionGraph(graph, startNodeId = '') {
    const normalized = normalizeWorkflowGraph(graph);
    const normalizedStartNodeId = String(startNodeId || '').trim();
    const nodesById = new Map(normalized.nodes.map((node) => [node.id, node]));
    let startIds = [];
    if (normalizedStartNodeId) {
        if (!nodesById.has(normalizedStartNodeId)) {
            throw new Error('找不到要运行的节点');
        }
        const startNode = nodesById.get(normalizedStartNodeId);
        if (startNode?.type === 'input' && !getWorkflowOutgoingLinks(normalized, normalizedStartNodeId).length) {
            throw new Error('导入图片节点未连接到后续节点，无法单独运行');
        }
        if (startNode?.type !== 'input' && !getWorkflowIncomingNodeIds(normalized, normalizedStartNodeId).length) {
            throw new Error('节点未接入输入链路，无法运行');
        }
        startIds = [normalizedStartNodeId];
    } else {
        const inputNodes = normalized.nodes.filter((node) => node.type === 'input');
        const connectedInputNodes = inputNodes.filter((node) => getWorkflowOutgoingLinks(normalized, node.id).length > 0);
        startIds = connectedInputNodes.map((node) => node.id);
    }
    const reachable = collectWorkflowReachableNodeIds(normalized, startIds);
    const indexedNodes = normalized.nodes.map((node, index) => ({ node, index }));
    const nodes = indexedNodes
        .filter(({ node }) => reachable.has(node.id))
        .sort((a, b) => {
            const orderA = getWorkflowNodeDef(a.node.type)?.order || 0;
            const orderB = getWorkflowNodeDef(b.node.type)?.order || 0;
            return orderA - orderB || a.index - b.index;
        })
        .map(({ node }) => node);
    const links = normalized.links.filter((link) => reachable.has(link.from) && reachable.has(link.to));
    return { ...normalized, nodes, links };
}

function getWorkflowGraphTopo(graph) {
    const executionGraph = getWorkflowExecutionGraph(graph);
    if (!executionGraph.nodes.length) {
        throw new Error('没有可执行节点，请先连接导入节点和后续节点');
    }
    return executionGraph;
}

function getWorkflowNodesFrom(graph, startNodeId = '') {
    const executionGraph = getWorkflowExecutionGraph(graph, startNodeId);
    if (!executionGraph.nodes.length) {
        throw new Error('没有可执行节点');
    }
    const executionNodeIds = new Set(executionGraph.nodes.map((node) => node.id));
    const normalized = normalizeWorkflowGraph(graph);
    return {
        ...executionGraph,
        links: normalized.links.filter((link) => executionNodeIds.has(link.to))
    };
}

function getWorkflowIncomingNodeIds(graph, nodeId) {
    const normalizedNodeId = String(nodeId || '').trim();
    const ids = [];
    (Array.isArray(graph?.links) ? graph.links : []).forEach((link) => {
        if (String(link?.to || '').trim() !== normalizedNodeId) return;
        const fromId = String(link?.from || '').trim();
        if (fromId && !ids.includes(fromId)) ids.push(fromId);
    });
    return ids;
}

function mergeWorkflowArtifacts(...items) {
    const merged = {};
    items.forEach((item) => {
        if (!item || typeof item !== 'object') return;
        Object.entries(item).forEach(([key, value]) => {
            if (key === 'nodeArtifacts') return;
            if (value === undefined || value === null || value === '') return;
            if (Array.isArray(value)) {
                const current = Array.isArray(merged[key]) ? merged[key] : [];
                const next = current.concat(value);
                merged[key] = next.every((entry) => typeof entry === 'string')
                    ? Array.from(new Set(next.map((entry) => String(entry || '').trim()).filter(Boolean)))
                    : next;
                return;
            }
            if (value && typeof value === 'object') {
                const current = merged[key] && typeof merged[key] === 'object' && !Array.isArray(merged[key])
                    ? merged[key]
                    : {};
                merged[key] = { ...current, ...value };
                return;
            }
            merged[key] = value;
        });
    });
    return merged;
}

function getWorkflowNodeInputArtifacts(run, graph, node, nodeArtifacts) {
    if (node.type === 'input') return {};
    const incomingIds = getWorkflowIncomingNodeIds(graph, node.id);
    return mergeWorkflowArtifacts(...incomingIds.map((id) => nodeArtifacts?.[id]).filter(Boolean));
}

function mergeWorkflowDisplayArtifacts(currentArtifacts, nodeOutput, nodeId, nodeArtifacts) {
    const merged = mergeWorkflowArtifacts(currentArtifacts || {}, nodeOutput || {});
    if (Array.isArray(nodeOutput?.templateOutputItems)) {
        const seen = new Set();
        merged.templateOutputItems = nodeOutput.templateOutputItems.filter((item) => {
            const key = String(item?.path || '').trim();
            if (!key || seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }
    return {
        ...merged,
        nodeArtifacts: {
            ...((currentArtifacts || {}).nodeArtifacts || {}),
            ...(nodeArtifacts || {}),
            ...(nodeId ? { [nodeId]: nodeOutput || {} } : {})
        }
    };
}

function validateWorkflowGraph(graph) {
    const normalized = normalizeWorkflowGraph(graph);
    const errors = [];
    const warnings = [];
    const issues = [];
    const addIssue = (level, message, nodeIds = []) => {
        const normalizedMessage = String(message || '').trim();
        if (!normalizedMessage) return;
        const issue = {
            level: level === 'error' ? 'error' : 'warning',
            message: normalizedMessage,
            nodeIds: Array.from(new Set((Array.isArray(nodeIds) ? nodeIds : [nodeIds])
                .map((id) => String(id || '').trim())
                .filter(Boolean)))
        };
        issues.push(issue);
        if (issue.level === 'error') {
            errors.push(normalizedMessage);
        } else {
            warnings.push(normalizedMessage);
        }
    };
    if (!normalized.nodes.length) {
        addIssue('error', '工作流至少需要一个节点');
    }
    const nodesById = new Map(normalized.nodes.map((node) => [node.id, node]));
    const outCount = new Map();
    const outgoingLinksByFrom = new Map();
    normalized.links.forEach((link) => {
        const from = nodesById.get(link.from);
        const to = nodesById.get(link.to);
        if (!from || !to) return;
        const fromDef = getWorkflowNodeDef(from.type);
        const toDef = getWorkflowNodeDef(to.type);
        if ((fromDef?.order || 0) > (toDef?.order || 0)) {
            addIssue('error', `节点顺序错误：${from.title} 不能连接到 ${to.title}`, [from.id, to.id]);
        }
        outCount.set(from.id, (outCount.get(from.id) || 0) + 1);
        if (!outgoingLinksByFrom.has(from.id)) outgoingLinksByFrom.set(from.id, []);
        outgoingLinksByFrom.get(from.id).push(link);
    });
    normalized.nodes.forEach((node) => {
        const def = getWorkflowNodeDef(node.type);
        if (!def) {
            addIssue('error', `未知节点：${node.title}`, node.id);
            return;
        }
        if (node.type === 'input') {
            const hasOutput = normalized.links.some((link) => link.from === node.id);
            if (!hasOutput) addIssue('warning', '导入图片/文件夹节点尚未连接到后续节点', node.id);
        }
        const outgoingLinks = outgoingLinksByFrom.get(node.id) || [];
        const nonTemplateTargetCount = outgoingLinks.filter((link) => nodesById.get(link.to)?.type !== 'template').length;
        if ((outCount.get(node.id) || 0) > 1 && !def.branchable && nonTemplateTargetCount > 1) {
            addIssue('error', `${def.label} 不允许分支输出`, node.id);
        }
        if (node.type === 'extract') {
            const cfg = loadPrintAiConfig();
            const model = String(node.config?.model || cfg.extractModel || '').trim();
            if (!model) addIssue('warning', '提取图案节点未选择模型，将无法运行', node.id);
        }
        if (node.type === 'variation') {
            const cfg = loadPrintAiConfig();
            const model = String(node.config?.model || cfg.variationModel || '').trim();
            const prompt = String(node.config?.prompt || '').trim();
            if (!model) addIssue('warning', '印花裂变节点未选择模型，将使用失败状态提示', node.id);
            if (!prompt) addIssue('warning', '印花裂变节点未填写提示词，将使用印花裂变默认提示词', node.id);
        }
        if (node.type === 'template' && (!Array.isArray(node.config?.selectedTemplates) || !node.config.selectedTemplates.length)) {
            addIssue('warning', '智能模板节点未指定模板，运行到该节点会失败', node.id);
        }
        if (node.type === 'export' && !String(node.config?.outputDir || '').trim()) {
            addIssue('error', '导出保存节点缺少导出目录', node.id);
        }
    });
    return {
        ok: errors.length === 0,
        errors,
        warnings,
        issues,
        graph: normalized
    };
}

function listWorkflowFolderImages(folderPath) {
    const files = [];
    if (!folderPath || !fs.existsSync(folderPath)) return files;
    const stack = [folderPath];
    while (stack.length) {
        const current = stack.pop();
        fs.readdirSync(current, { withFileTypes: true }).forEach((entry) => {
            const fullPath = path.join(current, entry.name);
            if (entry.isDirectory()) {
                stack.push(fullPath);
            } else if (entry.isFile() && WORKFLOW_IMAGE_EXTS.has(path.extname(entry.name).toLowerCase())) {
                files.push(fullPath);
            }
        });
    }
    return files.sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
}

function importWorkflowInputs(filePaths, workflowPresetId = '', sourceInputNodeId = '') {
    const data = loadWorkflowData();
    const uploadRoot = ensureDir(path.join(WORKFLOW_STORAGE_DIR, 'inputs'));
    const imported = [];
    const now = new Date().toISOString();
    (Array.isArray(filePaths) ? filePaths : []).forEach((rawPath) => {
        const sourcePath = String(rawPath || '').trim();
        if (!sourcePath || !fs.existsSync(sourcePath)) return;
        const stat = fs.statSync(sourcePath);
        const sourceType = stat.isDirectory() ? 'folder' : 'file';
        if (sourceType === 'file' && !WORKFLOW_IMAGE_EXTS.has(path.extname(sourcePath).toLowerCase())) return;
        if (sourceType === 'folder' && !listWorkflowFolderImages(sourcePath).length) return;
        const safeBase = sanitizePathSegment(path.basename(sourcePath, path.extname(sourcePath)), 'workflow-input');
        const taskRoot = ensureDir(path.join(uploadRoot, `${safeBase}-${Date.now()}-${crypto.randomBytes(2).toString('hex')}`));
        let copiedPath = '';
        if (sourceType === 'folder') {
            copiedPath = taskRoot;
            listWorkflowFolderImages(sourcePath).forEach((imagePath) => {
                const targetPath = ensureUniqueFilePath(path.join(taskRoot, path.basename(imagePath)));
                fs.copyFileSync(imagePath, targetPath);
            });
        } else {
            copiedPath = path.join(taskRoot, path.basename(sourcePath));
            fs.copyFileSync(sourcePath, copiedPath);
        }
        const run = normalizeWorkflowRun({
            id: `workflow-run-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`,
            name: path.basename(sourcePath),
            workflowPresetId,
            sourceInputNodeId: String(sourceInputNodeId || '').trim(),
            sourcePath: copiedPath,
            sourceType,
            status: 'pending',
            progress: 0,
            artifacts: {
                inputPath: copiedPath,
                inputImages: sourceType === 'folder' ? listWorkflowFolderImages(copiedPath) : [copiedPath]
            },
            nodeStates: {},
            createdAt: now,
            updatedAt: now
        });
        data.runs.unshift(run);
        imported.push(run);
    });
    const saved = saveWorkflowData(data);
    return { imported, ...saved };
}

function updateWorkflowRunState(runId, patch, sender) {
    const data = loadWorkflowData();
    data.runs = data.runs.map((run) => run.id === runId
        ? normalizeWorkflowRun({ ...run, ...patch, updatedAt: new Date().toISOString() })
        : run);
    const saved = saveWorkflowData(data);
    broadcastWorkflowData(sender, saved);
    return saved.runs.find((run) => run.id === runId) || null;
}

async function runWorkflowExtractImage(run, node, sender, imagePath, index = 0) {
    const cfg = loadPrintAiConfig();
    if (!imagePath || !fs.existsSync(imagePath)) {
        throw new Error('提取图案缺少输入图片');
    }
    const model = String(node.config?.model || cfg.extractModel || '').trim();
    if (!cfg.baseUrl || !model) {
        throw new Error('提取图案缺少 API 地址或模型');
    }
    const prompt = String(node.config?.prompt || cfg.extractionPrompt || '').trim() || createDefaultPrintAiConfig().extractionPrompt;
    const aspectRatio = String(node.config?.aspectRatio || cfg.aspectRatio || '3:2');
    const generated = await requestPrintAiImageEdit({
        baseUrl: cfg.baseUrl,
        apiKey: cfg.apiKey,
        model,
        timeoutMs: cfg.timeoutMs,
        imagePath,
        prompt: [
            prompt,
            '',
            `Output aspect ratio: ${aspectRatio}. The extracted standalone print image must use this exact canvas ratio.`
        ].join('\n'),
        size: getPrintAiImageSize(aspectRatio)
    });
    return savePrintAiGeneratedImage(generated.bytes, generated.mimeType, 'extracted', index > 0 ? `${run.name}-${index + 1}` : run.name);
}

async function runWorkflowExtractNode(run, node, sender, artifacts = {}) {
    const imagePaths = (Array.isArray(artifacts.inputImages) ? artifacts.inputImages : [])
        .map((filePath) => String(filePath || '').trim())
        .filter((filePath) => filePath && fs.existsSync(filePath) && fs.statSync(filePath).isFile());
    if (!imagePaths.length) {
        throw new Error('提取图案缺少输入图片');
    }
    const extractedPaths = [];
    for (let index = 0; index < imagePaths.length; index += 1) {
        sendTemplateLog(sender, 'info', `工作流提取图案：${index + 1}/${imagePaths.length}`);
        try {
            extractedPaths.push(await runWorkflowExtractImage(run, node, sender, imagePaths[index], index));
        } catch (error) {
            const message = compactPrintAiErrorMessage(error, '提取图案失败', '提取图案失败，请稍后重试或切换模型');
            const wrapped = new Error(message);
            wrapped.status = error?.status;
            throw wrapped;
        }
    }
    return extractedPaths;
}

async function runWorkflowVariationNode(run, node, artifacts = {}) {
    const cfg = loadPrintAiConfig();
    const imagePath = String(
        (Array.isArray(artifacts.extractedPaths) ? artifacts.extractedPaths[0] : '')
        || artifacts.extractedPath
        || (Array.isArray(artifacts.inputImages) ? artifacts.inputImages[0] : '')
        || ''
    ).trim();
    if (!imagePath || !fs.existsSync(imagePath)) {
        throw new Error('印花裂变缺少提取图案输入');
    }
    const model = String(node.config?.model || cfg.variationModel || '').trim();
    if (!cfg.baseUrl || !model) {
        throw new Error('印花裂变缺少 API 地址或模型');
    }
    const defaultPrompt = cfg.variationPrompts?.[0]?.prompt || createDefaultPrintAiConfig().variationPrompts[0].prompt;
    const prompt = String(node.config?.prompt || defaultPrompt || '').trim();
    const aspectRatio = String(node.config?.aspectRatio || cfg.aspectRatio || '3:2');
    const generated = await requestPrintAiImageEdit({
        baseUrl: cfg.baseUrl,
        apiKey: cfg.apiKey,
        model,
        timeoutMs: cfg.timeoutMs,
        imagePath,
        prompt: [
            prompt,
            '',
            `Aspect ratio: ${aspectRatio}. Use the extracted print image as reference and preserve the canvas ratio.`
        ].join('\n'),
        size: getPrintAiImageSize(aspectRatio)
    });
    return savePrintAiGeneratedImage(generated.bytes, generated.mimeType, 'variants', `${run.name}-${node.title}`);
}

async function runWorkflowVariationNodes(run, node, artifacts = {}) {
    const count = Math.max(1, Math.min(4, Math.round(Number(node.config?.count || 1) || 1)));
    const paths = [];
    const sourcePaths = (Array.isArray(artifacts.extractedPaths) && artifacts.extractedPaths.length
        ? artifacts.extractedPaths
        : (artifacts.extractedPath ? [artifacts.extractedPath] : (Array.isArray(artifacts.inputImages) ? artifacts.inputImages : [])))
        .map((filePath) => String(filePath || '').trim())
        .filter((filePath) => filePath && fs.existsSync(filePath) && fs.statSync(filePath).isFile());
    for (let sourceIndex = 0; sourceIndex < sourcePaths.length; sourceIndex += 1) {
        for (let index = 0; index < count; index += 1) {
            const variantPath = await runWorkflowVariationNode(
                {
                    ...run,
                    name: sourcePaths.length > 1 || count > 1 ? `${run.name}-${sourceIndex + 1}-${index + 1}` : run.name
                },
                node,
                { ...artifacts, extractedPath: sourcePaths[sourceIndex], extractedPaths: [sourcePaths[sourceIndex]] }
            );
            paths.push(variantPath);
        }
    }
    return paths;
}

function clampWorkflowNumber(value, min, max) {
    return Math.max(min, Math.min(max, Number(value)));
}

function buildWorkflowEvenRanges(start, end, grid) {
    const length = Math.max(1, end - start + 1);
    return Array.from({ length: grid }, (_, index) => {
        const rangeStart = start + Math.round((length * index) / grid);
        const rangeEnd = index === grid - 1
            ? end
            : start + Math.round((length * (index + 1)) / grid) - 1;
        return {
            start: clampWorkflowNumber(rangeStart, start, end),
            end: clampWorkflowNumber(rangeEnd, start, end)
        };
    });
}

function smoothWorkflowSignal(values, radius = 2) {
    const source = Array.isArray(values) ? values : [];
    if (!source.length) return [];
    return source.map((_, index) => {
        let total = 0;
        let count = 0;
        for (let cursor = Math.max(0, index - radius); cursor <= Math.min(source.length - 1, index + radius); cursor += 1) {
            total += Number(source[cursor] || 0);
            count += 1;
        }
        return count ? total / count : 0;
    });
}

function buildWorkflowSliceRangesFromBands(start, end, bands, grid) {
    if (!Array.isArray(bands) || bands.length !== grid - 1) return [];
    const ranges = [];
    let cursor = start;
    bands.forEach((band) => {
        ranges.push({ start: cursor, end: Math.max(cursor, band.start - 1) });
        cursor = Math.min(end, band.end + 1);
    });
    ranges.push({ start: cursor, end });
    return ranges.length === grid && ranges.every((range) => range.end - range.start >= 2) ? ranges : [];
}

function insetWorkflowSliceRanges(ranges, inset = 0) {
    return (Array.isArray(ranges) ? ranges : []).map((range) => {
        const maxInset = Math.max(0, Math.floor((range.end - range.start) / 4));
        const appliedInset = Math.min(Math.max(0, Number(inset) || 0), maxInset);
        return { start: range.start + appliedInset, end: range.end - appliedInset };
    });
}

function analyzeWorkflowSliceSignals(raw, width, height) {
    const grayscale = new Float32Array(width * height);
    const columnMean = new Array(width).fill(0);
    const rowMean = new Array(height).fill(0);
    const columnBrightRatio = new Array(width).fill(0);
    const rowBrightRatio = new Array(height).fill(0);
    for (let y = 0; y < height; y += 1) {
        for (let x = 0; x < width; x += 1) {
            const offset = (y * width + x) * 4;
            const alpha = raw[offset + 3] / 255;
            const gray = ((raw[offset] * 0.299) + (raw[offset + 1] * 0.587) + (raw[offset + 2] * 0.114)) * alpha;
            grayscale[y * width + x] = gray;
            columnMean[x] += gray;
            rowMean[y] += gray;
            if (alpha > 0.72 && gray >= 218) {
                columnBrightRatio[x] += 1;
                rowBrightRatio[y] += 1;
            }
        }
    }
    for (let x = 0; x < width; x += 1) {
        columnMean[x] /= Math.max(1, height);
        columnBrightRatio[x] /= Math.max(1, height);
    }
    for (let y = 0; y < height; y += 1) {
        rowMean[y] /= Math.max(1, width);
        rowBrightRatio[y] /= Math.max(1, width);
    }
    const columnVariance = new Array(width).fill(0);
    const rowVariance = new Array(height).fill(0);
    const columnEdge = new Array(width).fill(0);
    const rowEdge = new Array(height).fill(0);
    for (let y = 0; y < height; y += 1) {
        for (let x = 0; x < width; x += 1) {
            const gray = grayscale[y * width + x];
            columnVariance[x] += Math.abs(gray - columnMean[x]);
            rowVariance[y] += Math.abs(gray - rowMean[y]);
            if (x > 0) columnEdge[x] += Math.abs(gray - grayscale[y * width + x - 1]);
            if (y > 0) rowEdge[y] += Math.abs(gray - grayscale[(y - 1) * width + x]);
        }
    }
    for (let x = 0; x < width; x += 1) {
        columnVariance[x] /= Math.max(1, height);
        columnEdge[x] /= Math.max(1, height);
    }
    for (let y = 0; y < height; y += 1) {
        rowVariance[y] /= Math.max(1, width);
        rowEdge[y] /= Math.max(1, width);
    }
    return {
        width,
        height,
        columnSeparator: smoothWorkflowSignal(columnBrightRatio.map((bright, index) => {
            const varianceScore = 1 - clampWorkflowNumber(columnVariance[index] / 72, 0, 1);
            const edgeScore = 1 - clampWorkflowNumber(columnEdge[index] / 44, 0, 1);
            return bright * 0.58 + varianceScore * 0.24 + edgeScore * 0.18;
        }), 2),
        rowSeparator: smoothWorkflowSignal(rowBrightRatio.map((bright, index) => {
            const varianceScore = 1 - clampWorkflowNumber(rowVariance[index] / 72, 0, 1);
            const edgeScore = 1 - clampWorkflowNumber(rowEdge[index] / 44, 0, 1);
            return bright * 0.58 + varianceScore * 0.24 + edgeScore * 0.18;
        }), 2)
    };
}

function findWorkflowSliceBands(signal, grid) {
    const length = signal.length;
    const bands = [];
    const cellSize = length / Math.max(1, grid);
    const searchRadius = Math.max(4, Math.round(cellSize * 0.2));
    for (let index = 1; index < grid; index += 1) {
        const expected = cellSize * index;
        const start = Math.max(2, Math.round(expected - searchRadius));
        const end = Math.min(length - 3, Math.round(expected + searchRadius));
        let best = start;
        let bestScore = -Infinity;
        for (let cursor = start; cursor <= end; cursor += 1) {
            const distancePenalty = Math.abs(cursor - expected) / Math.max(1, searchRadius);
            const score = Number(signal[cursor] || 0) - distancePenalty * 0.28;
            if (score > bestScore) {
                bestScore = score;
                best = cursor;
            }
        }
        if (bestScore < 0.58) return [];
        const threshold = Math.max(0.52, Number(signal[best] || 0) * 0.82);
        let bandStart = best;
        let bandEnd = best;
        while (bandStart > start && Number(signal[bandStart - 1] || 0) >= threshold) bandStart -= 1;
        while (bandEnd < end && Number(signal[bandEnd + 1] || 0) >= threshold) bandEnd += 1;
        bands.push({ start: bandStart, end: bandEnd, score: bestScore });
    }
    return bands;
}

function detectWorkflowSliceLayout(signals, requestedGrid = 'auto') {
    const candidateGrids = requestedGrid === 'auto' ? [3, 2, 4] : [Number(requestedGrid) || 3];
    const candidates = candidateGrids.map((grid) => {
        const xBands = findWorkflowSliceBands(signals.columnSeparator, grid);
        const yBands = findWorkflowSliceBands(signals.rowSeparator, grid);
        const complete = xBands.length === grid - 1 && yBands.length === grid - 1;
        const score = complete
            ? (xBands.concat(yBands).reduce((sum, band) => sum + band.score, 0) / Math.max(1, xBands.length + yBands.length)) + 0.3
            : 0;
        return { grid, xBands, yBands, complete, score };
    }).sort((a, b) => b.score - a.score);
    const best = candidates[0] || { grid: 3, complete: false, xBands: [], yBands: [] };
    const grid = best.grid || 3;
    const xRanges = best.complete
        ? buildWorkflowSliceRangesFromBands(0, signals.width - 1, best.xBands, grid)
        : buildWorkflowEvenRanges(0, signals.width - 1, grid);
    const yRanges = best.complete
        ? buildWorkflowSliceRangesFromBands(0, signals.height - 1, best.yBands, grid)
        : buildWorkflowEvenRanges(0, signals.height - 1, grid);
    return {
        grid,
        xRanges: xRanges.length ? xRanges : buildWorkflowEvenRanges(0, signals.width - 1, grid),
        yRanges: yRanges.length ? yRanges : buildWorkflowEvenRanges(0, signals.height - 1, grid),
        confidence: best.complete ? clampWorkflowNumber(best.score, 0.45, 0.98) : 0.35,
        mode: best.complete ? 'separator' : 'equal'
    };
}

async function runWorkflowAutoSliceImage(filePath, node, run) {
    const cfg = loadSliceConfig();
    const baseOutputDir = cfg.outputDir || getDefaultSliceOutputDir();
    fs.mkdirSync(baseOutputDir, { recursive: true });
    const source = sharp(filePath, { limitInputPixels: false }).rotate();
    const metadata = await source.metadata();
    const sourceWidth = Math.max(1, Number(metadata.width) || 1);
    const sourceHeight = Math.max(1, Number(metadata.height) || 1);
    const sampleMaxSide = Number(node.config?.sampleMaxSide || 0) || Math.max(sourceWidth, sourceHeight);
    const sampleRatio = Math.min(1, sampleMaxSide / Math.max(sourceWidth, sourceHeight));
    const sampleWidth = Math.max(1, Math.round(sourceWidth * sampleRatio));
    const sampleHeight = Math.max(1, Math.round(sourceHeight * sampleRatio));
    const { data, info } = await sharp(filePath, { limitInputPixels: false })
        .rotate()
        .resize(sampleWidth, sampleHeight, { fit: 'fill' })
        .ensureAlpha()
        .raw()
        .toBuffer({ resolveWithObject: true });
    const signals = analyzeWorkflowSliceSignals(data, info.width, info.height);
    const requestedGrid = String(node.config?.grid || node.config?.gridSize || 'auto').trim() || 'auto';
    const layout = detectWorkflowSliceLayout(signals, requestedGrid);
    const inset = Math.max(0, Number(node.config?.shrink ?? node.config?.sliceInsetPx ?? 2) || 0);
    const xRanges = insetWorkflowSliceRanges(layout.xRanges, Math.round(inset * sampleRatio));
    const yRanges = insetWorkflowSliceRanges(layout.yRanges, Math.round(inset * sampleRatio));
    const baseName = sanitizePathSegment(path.basename(filePath, path.extname(filePath)), 'slice-task');
    const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\..+$/, '').replace('T', '_');
    const taskDir = path.join(baseOutputDir, `${baseName}_${stamp}`);
    fs.mkdirSync(taskDir, { recursive: true });
    const files = [];
    for (let row = 0; row < yRanges.length; row += 1) {
        for (let col = 0; col < xRanges.length; col += 1) {
            const xRange = xRanges[col];
            const yRange = yRanges[row];
            const left = clampWorkflowNumber(Math.round(xRange.start / sampleRatio), 0, sourceWidth - 1);
            const top = clampWorkflowNumber(Math.round(yRange.start / sampleRatio), 0, sourceHeight - 1);
            const right = clampWorkflowNumber(Math.round(xRange.end / sampleRatio), left, sourceWidth - 1);
            const bottom = clampWorkflowNumber(Math.round(yRange.end / sampleRatio), top, sourceHeight - 1);
            const width = Math.max(1, right - left + 1);
            const height = Math.max(1, bottom - top + 1);
            const targetPath = path.join(taskDir, `${baseName}_${row + 1}-${col + 1}.png`);
            await sharp(filePath, { limitInputPixels: false })
                .rotate()
                .extract({ left, top, width, height })
                .png()
                .toFile(targetPath);
            files.push(targetPath);
        }
    }
    return {
        sourcePath: filePath,
        taskDir,
        files,
        gridSize: layout.grid,
        mode: layout.mode,
        confidence: layout.confidence,
        runId: run.id
    };
}

async function runWorkflowSliceNode(run, node, artifacts) {
    const inputPaths = (Array.isArray(artifacts.variantPaths) && artifacts.variantPaths.length)
        ? artifacts.variantPaths
        : (Array.isArray(artifacts.extractedPaths) && artifacts.extractedPaths.length
            ? artifacts.extractedPaths
            : (artifacts.extractedPath ? [artifacts.extractedPath] : (artifacts.inputImages || [])));
    const imagePaths = inputPaths
        .map((filePath) => String(filePath || '').trim())
        .filter((filePath) => filePath && fs.existsSync(filePath) && fs.statSync(filePath).isFile());
    if (!imagePaths.length) {
        throw new Error('智能切片缺少可用输入图片');
    }
    const results = [];
    for (const imagePath of imagePaths) {
        results.push(await runWorkflowAutoSliceImage(imagePath, node, run));
    }
    return {
        sliceInputPaths: imagePaths,
        sliceResultPaths: results.flatMap((item) => item.files),
        sliceOutputDirs: results.map((item) => item.taskDir),
        sliceSummary: results.map((item) => ({
            sourcePath: item.sourcePath,
            taskDir: item.taskDir,
            fileCount: item.files.length,
            gridSize: item.gridSize,
            mode: item.mode,
            confidence: item.confidence
        }))
    };
}

function getWorkflowTemplateInputPaths(artifacts) {
    const templateInputPathSet = new Set();
    const pushTemplateInput = (value) => {
        (Array.isArray(value) ? value : [value]).forEach((item) => {
            const filePath = String(item || '').trim();
            if (filePath) templateInputPathSet.add(filePath);
        });
    };
    pushTemplateInput(artifacts.sliceResultPaths);
    pushTemplateInput(artifacts.variantPaths);
    pushTemplateInput(artifacts.extractedPaths);
    pushTemplateInput(artifacts.extractedPath);
    pushTemplateInput(artifacts.inputImages);
    return Array.from(templateInputPathSet);
}

async function runWorkflowTemplateNode(run, node, artifacts, sender) {
    if (templateProcess) {
        throw new Error('智能模板任务正在运行，请等待当前模板任务结束');
    }

    const templateInputPaths = getWorkflowTemplateInputPaths(artifacts);
    const designs = templateInputPaths
        .map((filePath) => String(filePath || '').trim())
        .filter((filePath) => filePath && fs.existsSync(filePath) && fs.statSync(filePath).isFile())
        .map((filePath) => ({
            name: path.basename(filePath),
            path: filePath
        }));

    if (!designs.length) {
        throw new Error('智能模板缺少可用输入图片');
    }

    const templateRendererScript = getTemplateRendererScriptPath();
    if (!fs.existsSync(templateRendererScript)) {
        throw new Error('缺少 template_renderer.py');
    }

    const cfg = loadTemplateConfig();
    const selectedTemplates = Array.isArray(node.config?.selectedTemplates) ? node.config.selectedTemplates : [];
    const activeTemplateGroups = selectedTemplates
        .map((item) => String(item || '').trim())
        .filter(Boolean);

    if (!activeTemplateGroups.length) {
        throw new Error('智能模板节点未选择模板');
    }

    const templateGroups = listTemplateFolders();
    const resolvedTemplateGroups = activeTemplateGroups
        .map((name) => templateGroups.find((item) => item.name === name))
        .filter(Boolean)
        .map((group) => ({
            name: group.name,
            scenes: (group.scenes || []).filter((scene) => scene.valid).map((scene) => ({
                name: scene.name,
                relativePath: scene.relativePath
            }))
        }))
        .filter((group) => group.scenes.length > 0);

    if (!resolvedTemplateGroups.length) {
        throw new Error('智能模板节点选择的模板不存在或没有可用场景');
    }

    const pythonRuntime = getPythonRuntime(templateRendererScript);
    if (!pythonRuntime) {
        throw new Error('未检测到可用 Python 运行环境，请安装 Python 或 py 启动器');
    }

    const watermarkPresetId = String(node.config?.watermarkPresetId || artifacts.watermarkPresetId || cfg.watermarkPresetId || '').trim();
    const savedPresets = loadWatermarkPresets();
    const selectedPreset = savedPresets.find((item) => item.id === watermarkPresetId) || null;
    const parameterPresetId = String(node.config?.parameterPresetId || cfg.parameterPresetId || '').trim();
    const parameterPresets = loadTemplateParameterPresets();
    const selectedParameterPreset = parameterPresetId
        ? (parameterPresets.find((item) => item.id === parameterPresetId) || {}).effects || null
        : null;
    const resolvedOutputDir = String(node.config?.outputDir || cfg.outputDir || getDefaultTemplateOutputDir()).trim();
    ensureDir(resolvedOutputDir);

    const jobPayload = {
        outputDir: resolvedOutputDir,
        templateRootDir: getTemplateRootDir(),
        templateGroups: resolvedTemplateGroups,
        designs,
        watermarkPreset: selectedPreset,
        effectPreset: selectedParameterPreset
    };

    sendTemplateLog(
        sender,
        'info',
        `工作流智能模板开始：${run.name}，${designs.length} 张图，${resolvedTemplateGroups.length} 个模板`
    );

    const outputPaths = [];
    const outputItems = [];
    const outputDirs = new Set();
    const publishTemplateProgress = () => {
        if (!outputPaths.length) return;
        const latestRun = loadWorkflowData().runs.find((item) => item.id === run.id) || run;
        const currentNodeStates = latestRun.nodeStates && typeof latestRun.nodeStates === 'object'
            ? latestRun.nodeStates
            : {};
        const currentNodeState = currentNodeStates[node.id] || {};
        const partialArtifacts = {
            templateInputPaths,
            templateOutputPaths: outputPaths.slice(),
            templateOutputDirs: Array.from(outputDirs).filter(Boolean),
            templateOutputItems: outputItems.slice()
        };
        updateWorkflowRunState(run.id, {
            artifacts: mergeWorkflowDisplayArtifacts(latestRun.artifacts || run.artifacts, partialArtifacts, node.id),
            nodeStates: {
                ...currentNodeStates,
                [node.id]: {
                    ...currentNodeState,
                    status: 'running',
                    message: `已生成 ${outputPaths.length} 张模板图`,
                    updatedAt: new Date().toISOString()
                }
            }
        }, sender);
    };
    const doneSummary = await new Promise((resolve, reject) => {
        const child = spawn(pythonRuntime.command, pythonRuntime.scriptArgs, {
            cwd: path.dirname(templateRendererScript),
            windowsHide: true,
            stdio: ['pipe', 'pipe', 'pipe'],
            env: {
                ...process.env,
                PYTHONUTF8: '1'
            }
        });

        let stdoutBuffer = '';
        let stderrBuffer = '';
        let doneMessage = null;

        const handleMessage = (message) => {
            if (!message || typeof message !== 'object') return;
            if (message.outputDir) {
                outputDirs.add(String(message.outputDir || '').trim());
            }
            if (message.outputPath) {
                const outputPath = String(message.outputPath || '').trim();
                if (outputPath && fs.existsSync(outputPath) && !outputPaths.includes(outputPath)) {
                    outputPaths.push(outputPath);
                    outputItems.push({
                        path: outputPath,
                        outputDir: String(message.outputDir || path.dirname(outputPath) || '').trim(),
                        groupName: String(message.groupName || '').trim(),
                        sceneName: String(message.sceneName || '').trim(),
                        designName: String(message.designName || '').trim(),
                        designPath: String(message.designPath || '').trim()
                    });
                    publishTemplateProgress();
                }
            }
            if (message.type === 'log') {
                sendTemplateLog(sender, message.level || 'info', message.message || '', message);
            } else if (message.type === 'progress') {
                safeSend(sender, 'template:progress', message);
            } else if (message.type === 'done') {
                doneMessage = message;
            }
        };

        child.stdout.on('data', (chunk) => {
            stdoutBuffer += chunk.toString('utf-8');
            const lines = stdoutBuffer.split(/\r?\n/);
            stdoutBuffer = lines.pop() || '';
            lines.forEach((line) => {
                const text = line.trim();
                if (!text) return;
                try {
                    handleMessage(JSON.parse(text));
                } catch {
                    sendTemplateLog(sender, 'info', text);
                }
            });
        });

        child.stderr.on('data', (chunk) => {
            stderrBuffer += chunk.toString('utf-8');
        });

        child.on('error', reject);

        child.on('close', (code, signal) => {
            if (stdoutBuffer.trim()) {
                try {
                    handleMessage(JSON.parse(stdoutBuffer.trim()));
                } catch {
                    sendTemplateLog(sender, 'info', stdoutBuffer.trim());
                }
            }
            if (code !== 0) {
                reject(new Error(stderrBuffer.trim() || `智能模板任务异常结束 (code=${code ?? 'null'}, signal=${signal ?? 'null'})`));
                return;
            }
            resolve(doneMessage || { processed: outputPaths.length, failed: 0, total: outputPaths.length });
        });

        child.stdin.end(JSON.stringify(jobPayload), 'utf-8');
    });

    if (!outputPaths.length) {
        throw new Error('智能模板没有生成任何结果');
    }

    sendTemplateLog(sender, 'success', `工作流智能模板完成：生成 ${outputPaths.length} 张图`);

    return {
        templateInputPaths,
        templateOutputPaths: outputPaths,
        templateOutputItems: outputItems,
        templateOutputDirs: Array.from(outputDirs).filter(Boolean),
        templateSummary: {
            processed: Number(doneSummary.processed || outputPaths.length),
            failed: Number(doneSummary.failed || 0),
            total: Number(doneSummary.total || outputPaths.length)
        }
    };
}

function collectWorkflowExportFiles(artifacts) {
    const files = [];
    const pushFile = (filePath) => {
        const normalized = String(filePath || '').trim();
        if (normalized && fs.existsSync(normalized) && fs.statSync(normalized).isFile() && !files.includes(normalized)) {
            files.push(normalized);
        }
    };
    [
        artifacts.extractedPath,
        ...(Array.isArray(artifacts.extractedPaths) ? artifacts.extractedPaths : []),
        ...(Array.isArray(artifacts.variantPaths) ? artifacts.variantPaths : []),
        ...(Array.isArray(artifacts.sliceResultPaths) ? artifacts.sliceResultPaths : []),
        ...(Array.isArray(artifacts.templateOutputPaths) ? artifacts.templateOutputPaths : []),
        ...(Array.isArray(artifacts.publishOutputPaths) ? artifacts.publishOutputPaths : [])
    ].forEach(pushFile);
    if (!files.length) {
        (Array.isArray(artifacts.inputImages) ? artifacts.inputImages : []).forEach(pushFile);
    }
    return files;
}

function exportWorkflowArtifacts(run, node, artifacts) {
    const outputDir = String(node.config?.outputDir || '').trim();
    if (!outputDir) {
        throw new Error('导出保存节点缺少导出目录');
    }
    const files = collectWorkflowExportFiles(artifacts);
    if (!files.length) {
        throw new Error('导出保存节点没有可导出的文件');
    }
    const runFolder = ensureDir(path.join(outputDir, sanitizePathSegment(path.basename(run.name, path.extname(run.name)), 'workflow-export')));
    const exportedPaths = files.map((filePath) => {
        const targetPath = ensureUniqueFilePath(path.join(runFolder, path.basename(filePath)));
        fs.copyFileSync(filePath, targetPath);
        return targetPath;
    });
    return { exportDir: runFolder, exportedPaths };
}

async function runWorkflowNode(run, node, graph, sender, inputArtifacts = {}) {
    const artifacts = { ...(inputArtifacts || {}) };
    if (node.type === 'input') {
        const inputImages = run.sourceType === 'folder' ? listWorkflowFolderImages(run.sourcePath) : [run.sourcePath].filter(Boolean);
        return {
            status: 'done',
            artifacts: {
                inputPath: run.sourcePath,
                inputImages
            }
        };
    }
    if (node.type === 'extract') {
        const extractedPaths = await runWorkflowExtractNode(run, node, sender, artifacts);
        return {
            status: 'done',
            artifacts: {
                extractedPath: extractedPaths[0] || '',
                extractedPaths
            }
        };
    }
    if (node.type === 'variation') {
        try {
            const variantPaths = await runWorkflowVariationNodes(run, node, artifacts);
            return { status: 'done', artifacts: { variantPaths, variationError: '' } };
        } catch (error) {
            const message = compactPrintAiErrorMessage(error, '印花裂变失败', '已跳过本次裂变');
            return { status: 'done', artifacts: { ...artifacts, variationError: message }, message, level: 'warning' };
        }
    }
    if (node.type === 'slice') {
        const result = await runWorkflowSliceNode(run, node, artifacts);
        return { status: 'done', artifacts: result };
    }
    if (node.type === 'template') {
        const result = await runWorkflowTemplateNode(run, node, artifacts, sender);
        return { status: 'done', artifacts: result };
    }
    if (node.type === 'watermark') {
        return {
            status: 'done',
            artifacts: {
                ...artifacts,
                watermarkPresetId: String(node.config?.watermarkPresetId || '').trim()
            }
        };
    }
    if (node.type === 'publish') {
        const result = await runWorkflowPublishNode(run, node, artifacts, sender);
        return { status: 'done', artifacts: result };
    }
    if (node.type === 'export') {
        const result = exportWorkflowArtifacts(run, node, artifacts);
        return { status: 'done', artifacts: result };
    }
    return { status: 'done', artifacts };
}

function clearWorkflowArtifactsFromNode(artifacts, nodeType) {
    const next = { ...(artifacts || {}) };
    const order = getWorkflowNodeDef(nodeType)?.order || 0;
    if (order <= 10) {
        delete next.extractedPath;
        delete next.extractedPaths;
        delete next.extractionError;
    }
    if (order <= 20) {
        delete next.variantPaths;
        delete next.variationError;
    }
    if (order <= 30) {
        delete next.sliceInputPaths;
        delete next.sliceResultPaths;
        delete next.sliceOutputDirs;
        delete next.sliceSummary;
    }
    if (order <= 40) {
        delete next.templateInputPaths;
        delete next.templateOutputPaths;
        delete next.templateOutputDirs;
        delete next.templateSummary;
    }
    if (order <= 50) {
        delete next.watermarkPresetId;
    }
    if (order <= 60) {
        delete next.publishRecords;
        delete next.publishOutputPaths;
        delete next.publishOutputDir;
    }
    if (order <= 70) {
        delete next.exportDir;
        delete next.exportedPaths;
    }
    return next;
}

async function startWorkflowRun(payload, sender) {
    const graph = normalizeWorkflowGraph(payload?.graph);
    const validation = validateWorkflowGraph(graph);
    if (!validation.ok) {
        throw new Error(validation.errors.join('\n'));
    }
    const ids = new Set(Array.isArray(payload?.runIds) ? payload.runIds.map((id) => String(id || '').trim()).filter(Boolean) : []);
    if (!ids.size) {
        throw new Error('请先导入并选择要运行的工作流任务');
    }
    const topo = getWorkflowGraphTopo(graph);
    let data = loadWorkflowData();
    data.runs = data.runs.map((run) => ids.has(run.id)
        ? { ...run, status: 'running', error: '', progress: 0, currentNodeId: '', currentNodeType: '', artifacts: {}, nodeStates: {}, updatedAt: new Date().toISOString() }
        : run);
    saveWorkflowData(data);
    broadcastWorkflowData(sender, data);

    void (async () => {
        for (const runId of ids) {
            let run = updateWorkflowRunState(runId, { status: 'running', artifacts: {}, nodeStates: {}, progress: 0, error: '' }, sender);
            if (!run) continue;
            const nodeStates = { ...(run.nodeStates || {}) };
            const nodeArtifacts = {};
            for (let index = 0; index < topo.nodes.length; index += 1) {
                const node = topo.nodes[index];
                const startedAt = new Date().toISOString();
                const startedAtMs = Date.now();
                try {
                    nodeStates[node.id] = { status: 'running', message: '', startedAt, updatedAt: startedAt };
                    run = updateWorkflowRunState(runId, {
                        status: 'running',
                        currentNodeId: node.id,
                        currentNodeType: node.type,
                        nodeStates,
                        progress: Math.round((index / Math.max(1, topo.nodes.length)) * 100)
                    }, sender) || run;
                    const inputArtifacts = getWorkflowNodeInputArtifacts(run, topo, node, nodeArtifacts);
                    const result = await runWorkflowNode(run, node, topo, sender, inputArtifacts);
                    nodeArtifacts[node.id] = result.artifacts || {};
                    const finishedAt = new Date().toISOString();
                    nodeStates[node.id] = {
                        status: result.status,
                        message: result.message || '',
                        startedAt,
                        finishedAt,
                        durationMs: Math.max(0, Date.now() - startedAtMs),
                        updatedAt: finishedAt
                    };
                    run = updateWorkflowRunState(runId, {
                        status: result.status === 'paused' ? 'paused' : 'running',
                        artifacts: mergeWorkflowDisplayArtifacts(run.artifacts, result.artifacts || {}, node.id, nodeArtifacts),
                        nodeStates,
                        error: result.level === 'warning' ? '' : (result.message || ''),
                        progress: Math.round(((index + 1) / Math.max(1, topo.nodes.length)) * 100)
                    }, sender) || run;
                    if (result.status === 'paused') break;
                } catch (error) {
                    const finishedAt = new Date().toISOString();
                    const message = compactPrintAiErrorMessage(error, '节点执行失败', '请检查节点配置或稍后重试');
                    nodeStates[node.id] = {
                        status: 'failed',
                        message,
                        startedAt,
                        finishedAt,
                        durationMs: Math.max(0, Date.now() - startedAtMs),
                        updatedAt: finishedAt
                    };
                    updateWorkflowRunState(runId, {
                        status: 'failed',
                        currentNodeId: node.id,
                        currentNodeType: node.type,
                        nodeStates,
                        error: message
                    }, sender);
                    break;
                }
            }
            const latest = loadWorkflowData().runs.find((item) => item.id === runId);
            if (latest && latest.status === 'running') {
                updateWorkflowRunState(runId, { status: 'done', progress: 100, currentNodeId: '', currentNodeType: '', error: '' }, sender);
            }
        }
        broadcastWorkflowData(sender);
    })();

    return loadWorkflowData();
}

async function retryWorkflowNode(payload, sender) {
    const graph = normalizeWorkflowGraph(payload?.graph);
    const validation = validateWorkflowGraph(graph);
    if (!validation.ok) {
        throw new Error(validation.errors.join('\n'));
    }
    const runId = String(payload?.runId || '').trim();
    const nodeId = String(payload?.nodeId || '').trim();
    if (!runId) throw new Error('请先选择要重新运行的任务');
    if (!nodeId) throw new Error('请先选择要重新运行的节点');
    const topo = getWorkflowNodesFrom(graph, nodeId);
    const startNode = topo.nodes[0];
    let data = loadWorkflowData();
    const currentRun = data.runs.find((run) => run.id === runId);
    if (!currentRun) throw new Error('找不到要重新运行的任务');
    const executionNodeIds = new Set(topo.nodes.map((node) => node.id));
    const nodeStates = { ...(currentRun.nodeStates || {}) };
    Object.keys(nodeStates).forEach((id) => {
        if (executionNodeIds.has(id)) {
            delete nodeStates[id];
        }
    });
    const nodeArtifacts = {};
    Object.entries(currentRun.artifacts?.nodeArtifacts || {}).forEach(([id, artifact]) => {
        if (!executionNodeIds.has(id)) {
            nodeArtifacts[id] = artifact;
        }
    });
    const clearedArtifacts = { nodeArtifacts };
    clearedArtifacts.nodeArtifacts = nodeArtifacts;
    data.runs = data.runs.map((run) => run.id === runId
        ? {
            ...run,
            status: 'running',
            error: '',
            currentNodeId: startNode.id,
            currentNodeType: startNode.type,
            artifacts: clearedArtifacts,
            nodeStates,
            updatedAt: new Date().toISOString()
        }
        : run);
    saveWorkflowData(data);
    broadcastWorkflowData(sender, data);

    void (async () => {
        let run = updateWorkflowRunState(runId, { status: 'running', artifacts: clearedArtifacts, nodeStates, error: '' }, sender);
        if (!run) return;
        for (let index = 0; index < topo.nodes.length; index += 1) {
            const node = topo.nodes[index];
            const startedAt = new Date().toISOString();
            const startedAtMs = Date.now();
            try {
                nodeStates[node.id] = { status: 'running', message: '', startedAt, updatedAt: startedAt };
                run = updateWorkflowRunState(runId, {
                    status: 'running',
                    currentNodeId: node.id,
                    currentNodeType: node.type,
                    nodeStates,
                    progress: Math.round((index / Math.max(1, topo.nodes.length)) * 100)
                }, sender) || run;
                const inputArtifacts = getWorkflowNodeInputArtifacts(run, topo, node, nodeArtifacts);
                const result = await runWorkflowNode(run, node, topo, sender, inputArtifacts);
                nodeArtifacts[node.id] = result.artifacts || {};
                const finishedAt = new Date().toISOString();
                nodeStates[node.id] = {
                    status: result.status,
                    message: result.message || '',
                    startedAt,
                    finishedAt,
                    durationMs: Math.max(0, Date.now() - startedAtMs),
                    updatedAt: finishedAt
                };
                run = updateWorkflowRunState(runId, {
                    status: result.status === 'paused' ? 'paused' : 'running',
                    artifacts: mergeWorkflowDisplayArtifacts(run.artifacts, result.artifacts || {}, node.id, nodeArtifacts),
                    nodeStates,
                    error: result.level === 'warning' ? '' : (result.message || ''),
                    progress: Math.round(((index + 1) / Math.max(1, topo.nodes.length)) * 100)
                }, sender) || run;
                if (result.status === 'paused') break;
            } catch (error) {
                const finishedAt = new Date().toISOString();
                const message = compactPrintAiErrorMessage(error, '节点执行失败', '请检查节点配置或稍后重试');
                nodeStates[node.id] = {
                    status: 'failed',
                    message,
                    startedAt,
                    finishedAt,
                    durationMs: Math.max(0, Date.now() - startedAtMs),
                    updatedAt: finishedAt
                };
                updateWorkflowRunState(runId, {
                    status: 'failed',
                    currentNodeId: node.id,
                    currentNodeType: node.type,
                    nodeStates,
                    error: message
                }, sender);
                break;
            }
        }
        const latest = loadWorkflowData().runs.find((item) => item.id === runId);
        if (latest && latest.status === 'running') {
            updateWorkflowRunState(runId, { status: 'done', progress: 100, currentNodeId: '', currentNodeType: '', error: '' }, sender);
        }
        broadcastWorkflowData(sender);
    })();

    return loadWorkflowData();
}

function normalizeProductPublishImage(item, index = 0) {
    const imagePath = String(item?.path || '').trim();
    const sceneName = String(item?.sceneName || '').trim();
    const name = String(item?.name || path.basename(imagePath || `image-${index + 1}`)).trim() || `image-${index + 1}`;
    return {
        id: String(item?.id || `image-${index + 1}`).trim() || `image-${index + 1}`,
        name,
        path: imagePath,
        sceneName
    };
}

function inferProductPublishTypeFromNames(values, mappings = createDefaultProductPublishTypeMappings()) {
    const joined = values
        .flatMap((value) => Array.isArray(value) ? value : [value])
        .map((value) => String(value || '').trim().toLowerCase())
        .filter(Boolean)
        .join(' ');
    if (!joined) return '';
    for (const mapping of Array.isArray(mappings) ? mappings : []) {
        const name = String(mapping?.name || '').trim();
        const keywords = getProductPublishTypeKeywords(mapping);
        if (!name || !keywords.length) continue;
        const matched = keywords.some((keyword) => {
            const normalized = String(keyword || '').trim().toLowerCase();
            return normalized && joined.includes(normalized);
        });
        if (matched) return name;
    }
    return '其他';
}

function normalizeProductPublishRecord(record, index = 0) {
    const id = String(record?.id || `product-${Date.now()}-${index + 1}`).trim() || `product-${Date.now()}-${index + 1}`;
    const legacyTitle = String(record?.title || '').trim();
    const titleEn = String(record?.titleEn || '').trim();
    const titleZh = String(record?.titleZh || '').trim() || ((!titleEn && legacyTitle) ? legacyTitle : '');
    const titleStatus = ['pending', 'generated', 'edited'].includes(record?.titleStatus)
        ? record.titleStatus
        : ((titleEn || titleZh) ? 'generated' : 'pending');
    const urls = Array.isArray(record?.urls)
        ? record.urls.map((item) => String(item || '').trim()).filter(Boolean)
        : [];
    const images = Array.isArray(record?.images)
        ? record.images.map((item, itemIndex) => normalizeProductPublishImage(item, itemIndex)).filter((item) => item.path)
        : [];
    const titleHistory = Array.isArray(record?.titleHistory)
        ? record.titleHistory
            .map((item, itemIndex) => {
                const historyTitleEn = String(item?.titleEn || '').trim();
                const historyTitleZh = String(item?.titleZh || '').trim();
                if (!historyTitleEn && !historyTitleZh) return null;
                return {
                    id: String(item?.id || `title-history-${id}-${itemIndex + 1}`).trim() || `title-history-${id}-${itemIndex + 1}`,
                    titleEn: historyTitleEn,
                    titleZh: historyTitleZh,
                    createdAt: String(item?.createdAt || item?.updatedAt || new Date().toISOString()).trim() || new Date().toISOString()
                };
            })
            .filter(Boolean)
            .slice(0, 20)
        : [];
    const inferredProductType = inferProductPublishTypeFromNames([
        record?.groupName || '',
        record?.designName || '',
        Array.isArray(record?.sceneNames) ? record.sceneNames : [],
        images.map((item) => item.name),
        images.map((item) => item.sceneName)
    ]);
    return {
        id,
        sourceTaskKey: String(record?.sourceTaskKey || '').trim(),
        groupName: String(record?.groupName || '').trim(),
        productType: String(inferredProductType || record?.productType || '').trim(),
        categoryId: String(record?.categoryId || '').trim(),
        mainCode: String(record?.mainCode || record?.groupName || '').trim(),
        shipLeadTime: String(record?.shipLeadTime || '').trim(),
        originPlace: String(record?.originPlace || '').trim(),
        previewImageUrl: String(record?.previewImageUrl || '').trim(),
        customized: String(record?.customized || '否').trim() || '否',
        specName1: String(record?.specName1 || '').trim(),
        specName2: String(record?.specName2 || '').trim(),
        specValue1: String(record?.specValue1 || '').trim(),
        specValue2: String(record?.specValue2 || '').trim(),
        declaredPrice: String(record?.declaredPrice || '').trim(),
        suggestedPrice: String(record?.suggestedPrice || '').trim(),
        lengthCm: String(record?.lengthCm || '').trim(),
        widthCm: String(record?.widthCm || '').trim(),
        heightCm: String(record?.heightCm || '').trim(),
        weightG: String(record?.weightG || '').trim(),
        inventory: String(record?.inventory || '').trim(),
        sensitive: String(record?.sensitive || '否').trim() || '否',
        outputDir: String(record?.outputDir || '').trim(),
        designName: String(record?.designName || '').trim(),
        sceneNames: Array.isArray(record?.sceneNames)
            ? record.sceneNames.map((item) => String(item || '').trim()).filter(Boolean)
            : [],
        images,
        titleEn,
        titleZh,
        titleHistory,
        titleStatus,
        urls,
        urlStatus: String(record?.urlStatus || (urls.length ? 'ready' : 'pending')).trim() || 'pending',
        exportStatus: String(record?.exportStatus || 'idle').trim() || 'idle',
        exportedAt: String(record?.exportedAt || '').trim(),
        exportFilePath: String(record?.exportFilePath || '').trim(),
        exportFileName: String(record?.exportFileName || '').trim(),
        createdAt: String(record?.createdAt || new Date().toISOString()),
        updatedAt: String(record?.updatedAt || new Date().toISOString())
    };
}

function resolveProductPublishTemuTemplatePath() {
    const candidates = [
        path.join(app.getPath('downloads'), PRODUCT_PUBLISH_TEMU_TEMPLATE_NAME),
        path.join(app.getPath('documents'), PRODUCT_PUBLISH_TEMU_TEMPLATE_NAME),
        path.join(process.cwd(), PRODUCT_PUBLISH_TEMU_TEMPLATE_NAME)
    ];
    return candidates.find((filePath) => fs.existsSync(filePath)) || '';
}

function setWorksheetCellValue(worksheet, cellAddress, value) {
    const previous = worksheet[cellAddress] || {};
    const cell = { ...previous };
    if (value === undefined || value === null || value === '') {
        cell.t = 's';
        cell.v = '';
        delete cell.w;
        worksheet[cellAddress] = cell;
        return;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
        cell.t = 'n';
        cell.v = value;
    } else {
        cell.t = 's';
        cell.v = String(value);
    }
    delete cell.w;
    worksheet[cellAddress] = cell;
}

function clearWorksheetDataRows(worksheet, startRowNumber) {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
    const startRowIndex = Math.max(0, startRowNumber - 1);
    for (let row = startRowIndex; row <= range.e.r; row += 1) {
        for (let col = range.s.c; col <= range.e.c; col += 1) {
            const addr = XLSX.utils.encode_cell({ r: row, c: col });
            if (worksheet[addr]) {
                setWorksheetCellValue(worksheet, addr, '');
            }
        }
    }
}

function buildProductPublishTemuWorkbook(records, templatePath) {
    const workbook = XLSX.readFile(templatePath, {
        cellStyles: true,
        cellFormula: true,
        cellNF: true,
        cellDates: true
    });
    const worksheet = workbook.Sheets.Sheet1 || workbook.Sheets[workbook.SheetNames[0]];
    if (!worksheet) {
        throw new Error('模板中没有可写入的工作表');
    }
    clearWorksheetDataRows(worksheet, 3);
    const templateRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:AQ3');
    const startRow = 3;
    const rows = Array.isArray(records) ? records : [];
    rows.forEach((record, index) => {
        const rowNumber = startRow + index;
        const carouselUrls = Array.isArray(record?.urls)
            ? record.urls.map((item) => String(item || '').trim()).filter(Boolean)
            : [];
        const fallbackImageUrl = carouselUrls[0] || '';
        setWorksheetCellValue(worksheet, `B${rowNumber}`, String(record?.categoryId || '').trim());
        setWorksheetCellValue(worksheet, `C${rowNumber}`, String(record?.mainCode || '').trim());
        setWorksheetCellValue(worksheet, `D${rowNumber}`, String(record?.titleZh || '').trim());
        setWorksheetCellValue(worksheet, `E${rowNumber}`, String(record?.titleEn || '').trim());
        setWorksheetCellValue(worksheet, `G${rowNumber}`, String(record?.shipLeadTime || '').trim());
        setWorksheetCellValue(worksheet, `I${rowNumber}`, String(record?.originPlace || '').trim());
        setWorksheetCellValue(worksheet, `K${rowNumber}`, carouselUrls.join('\n'));
        setWorksheetCellValue(worksheet, `L${rowNumber}`, fallbackImageUrl);
        setWorksheetCellValue(worksheet, `M${rowNumber}`, String(record?.customized || '否').trim() || '否');
        setWorksheetCellValue(worksheet, `N${rowNumber}`, String(record?.specName1 || '').trim());
        setWorksheetCellValue(worksheet, `O${rowNumber}`, String(record?.specValue1 || '').trim());
        setWorksheetCellValue(worksheet, `P${rowNumber}`, String(record?.specName2 || '').trim());
        setWorksheetCellValue(worksheet, `Q${rowNumber}`, String(record?.specValue2 || '').trim());
        setWorksheetCellValue(worksheet, `R${rowNumber}`, String(record?.previewImageUrl || fallbackImageUrl).trim());
        setWorksheetCellValue(worksheet, `S${rowNumber}`, String(record?.declaredPrice || '').trim());
        setWorksheetCellValue(worksheet, `T${rowNumber}`, String(record?.suggestedPrice || '').trim());
        setWorksheetCellValue(worksheet, `U${rowNumber}`, String(record?.lengthCm || '').trim());
        setWorksheetCellValue(worksheet, `V${rowNumber}`, String(record?.widthCm || '').trim());
        setWorksheetCellValue(worksheet, `W${rowNumber}`, String(record?.heightCm || '').trim());
        setWorksheetCellValue(worksheet, `X${rowNumber}`, String(record?.weightG || '').trim());
        setWorksheetCellValue(worksheet, `Y${rowNumber}`, String(record?.inventory || '').trim());
        setWorksheetCellValue(worksheet, `AA${rowNumber}`, String(record?.sensitive || '否').trim() || '否');
    });
    const requiredLastRow = Math.max(templateRange.e.r + 1, startRow + rows.length - 1);
    worksheet['!ref'] = XLSX.utils.encode_range({
        s: templateRange.s,
        e: {
            c: Math.max(templateRange.e.c, XLSX.utils.decode_col('AA')),
            r: Math.max(templateRange.e.r, requiredLastRow - 1)
        }
    });
    return workbook;
}

function getProductPublishImageMimeType(filePath) {
    const ext = path.extname(String(filePath || '')).toLowerCase();
    if (ext === '.png') return 'image/png';
    if (ext === '.webp') return 'image/webp';
    if (ext === '.gif') return 'image/gif';
    if (ext === '.bmp') return 'image/bmp';
    return 'image/jpeg';
}

function normalizeProductPublishOssRegion(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';
    const clean = raw.replace(/^https?:\/\//i, '').replace(/\/+$/, '');
    if (clean.startsWith('oss-')) return clean;
    return `oss-${clean}`;
}

function normalizeProductPublishOssPrefix(value) {
    return String(value || '')
        .trim()
        .replace(/^\/+/, '')
        .replace(/\/+$/, '');
}

function encodeProductPublishUrlPath(value) {
    return String(value || '')
        .split('/')
        .map((segment) => encodeURIComponent(segment))
        .join('/');
}

function buildProductPublishOssObjectKey(record, image, index, cfg) {
    const prefix = normalizeProductPublishOssPrefix(cfg?.ossObjectPrefix || '');
    const groupName = String(record?.groupName || 'product').trim() || 'product';
    const rawName = String(
        image?.name
        || path.basename(String(image?.path || ''))
        || `image-${index + 1}.jpg`
    ).trim() || `image-${index + 1}.jpg`;
    const parts = [];
    if (prefix) parts.push(prefix);
    parts.push(groupName, rawName);
    return parts.join('/');
}

function buildProductPublishOssPublicUrl(bucket, region, objectKey) {
    const bucketName = String(bucket || '').trim();
    const regionHost = normalizeProductPublishOssRegion(region);
    const key = String(objectKey || '').trim();
    if (!bucketName || !regionHost || !key) return '';
    return `https://${bucketName}.${regionHost}.aliyuncs.com/${encodeProductPublishUrlPath(key)}`;
}

function getProductPublishOssConfig(cfg) {
    return {
        ossBucket: String(cfg?.ossBucket || '').trim(),
        ossRegion: normalizeProductPublishOssRegion(cfg?.ossRegion || ''),
        ossAccessKeyId: String(cfg?.ossAccessKeyId || '').trim(),
        ossAccessKeySecret: String(cfg?.ossAccessKeySecret || '').trim(),
        ossObjectPrefix: normalizeProductPublishOssPrefix(cfg?.ossObjectPrefix || '')
    };
}

function isProductPublishOssConfigured(cfg) {
    const oss = getProductPublishOssConfig(cfg);
    return Boolean(oss.ossBucket && oss.ossRegion && oss.ossAccessKeyId && oss.ossAccessKeySecret);
}

function uploadBufferToOss(buffer, contentType, objectKey, ossCfg) {
    return new Promise((resolve, reject) => {
        const bucket = String(ossCfg?.ossBucket || '').trim();
        const region = normalizeProductPublishOssRegion(ossCfg?.ossRegion || '');
        const accessKeyId = String(ossCfg?.ossAccessKeyId || '').trim();
        const accessKeySecret = String(ossCfg?.ossAccessKeySecret || '').trim();
        const key = String(objectKey || '').trim();
        if (!bucket || !region || !accessKeyId || !accessKeySecret || !key) {
            reject(new Error('OSS 配置不完整，无法上传图片'));
            return;
        }
        const date = new Date().toUTCString();
        const canonicalResource = `/${bucket}/${key}`;
        const stringToSign = `PUT\n\n${contentType}\n${date}\n${canonicalResource}`;
        const signature = crypto
            .createHmac('sha1', accessKeySecret)
            .update(stringToSign, 'utf8')
            .digest('base64');
        const request = https.request({
            hostname: `${bucket}.${region}.aliyuncs.com`,
            method: 'PUT',
            path: `/${encodeProductPublishUrlPath(key)}`,
            headers: {
                Date: date,
                Host: `${bucket}.${region}.aliyuncs.com`,
                'Content-Type': contentType,
                'Content-Length': buffer.length,
                Authorization: `OSS ${accessKeyId}:${signature}`
            }
        }, (response) => {
            const chunks = [];
            response.on('data', (chunk) => chunks.push(chunk));
            response.on('end', () => {
                const body = Buffer.concat(chunks).toString('utf8');
                if (response.statusCode >= 200 && response.statusCode < 300) {
                    resolve({
                        objectKey: key,
                        url: buildProductPublishOssPublicUrl(bucket, region, key)
                    });
                    return;
                }
                reject(new Error(`OSS 上传失败：${response.statusCode}${body ? ` ${body}` : ''}`));
            });
        });
        request.on('error', (error) => reject(error));
        request.write(buffer);
        request.end();
    });
}

async function uploadProductPublishRecordImagesToOss(record, cfg, onProgress = null) {
    const ossCfg = getProductPublishOssConfig(cfg);
    if (!isProductPublishOssConfigured(ossCfg)) {
        return [];
    }
    const images = Array.isArray(record?.images) ? record.images : [];
    const urls = [];
    for (let index = 0; index < images.length; index += 1) {
        const image = images[index];
        const imagePath = String(image?.path || '').trim();
        if (!imagePath || !fs.existsSync(imagePath)) continue;
        if (typeof onProgress === 'function') {
            onProgress({
                phase: 'uploading',
                recordName: String(record?.groupName || '').trim() || '未命名产品',
                imageName: String(image?.name || path.basename(imagePath)).trim() || path.basename(imagePath),
                imageIndex: index + 1,
                imageTotal: images.length
            });
        }
        const buffer = fs.readFileSync(imagePath);
        const contentType = getProductPublishImageMimeType(imagePath);
        const objectKey = buildProductPublishOssObjectKey(record, image, index, ossCfg);
        const uploaded = await uploadBufferToOss(buffer, contentType, objectKey, ossCfg);
        if (uploaded?.url) {
            urls.push(uploaded.url);
            if (typeof onProgress === 'function') {
                onProgress({
                    phase: 'uploaded',
                    recordName: String(record?.groupName || '').trim() || '未命名产品',
                    imageName: String(image?.name || path.basename(imagePath)).trim() || path.basename(imagePath),
                    imageIndex: index + 1,
                    imageTotal: images.length,
                    url: uploaded.url
                });
            }
        }
    }
    return urls;
}

function buildProductPublishExportFileName(count = 0) {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, '0');
    const stamp = `${now.getFullYear()}年${pad(now.getMonth() + 1)}月${pad(now.getDate())}日_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    const total = Math.max(0, Number(count) || 0);
    return `ImageFlow_${stamp}_产品${total}.xlsx`;
}

function ensureUniqueFilePath(filePath) {
    if (!fs.existsSync(filePath)) {
        return filePath;
    }
    const parsed = path.parse(filePath);
    let index = 2;
    while (true) {
        const candidate = path.join(parsed.dir, `${parsed.name}_${index}${parsed.ext}`);
        if (!fs.existsSync(candidate)) {
            return candidate;
        }
        index += 1;
    }
}

function buildProductPublishVisionInputs(images) {
    return (Array.isArray(images) ? images : [])
        .filter((item) => item && item.path && fs.existsSync(item.path))
        .slice(0, 1)
        .map((item) => {
            const buffer = fs.readFileSync(item.path);
            const mime = getProductPublishImageMimeType(item.path);
            return {
                type: 'image_url',
                image_url: {
                    url: `data:${mime};base64,${buffer.toString('base64')}`
                }
            };
        });
}

function buildProductPublishGeminiVisionParts(images) {
    return (Array.isArray(images) ? images : [])
        .filter((item) => item && item.path && fs.existsSync(item.path))
        .slice(0, 1)
        .map((item) => {
            const buffer = fs.readFileSync(item.path);
            const mime = getProductPublishImageMimeType(item.path);
            return {
                inline_data: {
                    mime_type: mime,
                    data: buffer.toString('base64')
                }
            };
        });
}

function loadProductPublishData() {
    try {
        const parsed = JSON.parse(fs.readFileSync(PRODUCT_PUBLISH_DATA_FILE, 'utf-8'));
        return {
            records: Array.isArray(parsed?.records)
                ? parsed.records.map((item, index) => normalizeProductPublishRecord(item, index))
                : []
        };
    } catch {
        return { records: [] };
    }
}

function saveProductPublishData(data) {
    const nextData = {
        records: Array.isArray(data?.records)
            ? data.records.map((item, index) => normalizeProductPublishRecord(item, index))
            : []
    };
    fs.writeFileSync(PRODUCT_PUBLISH_DATA_FILE, JSON.stringify(nextData, null, 2), 'utf-8');
    return nextData;
}

function importProductPublishRecordFromTemplateTask(payload) {
    const currentData = loadProductPublishData();
    const sourceTaskKey = String(payload?.sourceTaskKey || '').trim();
    const groupName = String(payload?.groupName || '').trim();
    if (!sourceTaskKey || !groupName) {
        throw new Error('缺少模板任务信息，无法导入产品发布');
    }
    const cfg = loadProductPublishConfig();
    const normalizedImages = Array.isArray(payload?.images)
        ? payload.images.map((item, index) => normalizeProductPublishImage(item, index)).filter((item) => item.path)
        : [];
    const inferredProductType = inferProductPublishTypeFromNames([
        payload?.groupName || '',
        payload?.designName || '',
        Array.isArray(payload?.sceneNames) ? payload.sceneNames : [],
        normalizedImages.map((item) => item.name),
        normalizedImages.map((item) => item.sceneName)
    ], cfg.productTypeMappings);
    const existingIndex = currentData.records.findIndex((item) => item.sourceTaskKey === sourceTaskKey);
    const now = new Date().toISOString();
    if (existingIndex >= 0) {
        const existing = currentData.records[existingIndex];
        currentData.records[existingIndex] = normalizeProductPublishRecord({
            ...existing,
            groupName,
            productType: inferredProductType || existing.productType || '',
            mainCode: String(existing.mainCode || groupName || '').trim(),
            outputDir: String(payload?.outputDir || existing.outputDir || '').trim(),
            designName: String(payload?.designName || existing.designName || '').trim(),
            sceneNames: Array.isArray(payload?.sceneNames) && payload.sceneNames.length ? payload.sceneNames : existing.sceneNames,
            images: normalizedImages.length ? normalizedImages : existing.images,
            exportStatus: 'idle',
            exportedAt: '',
            exportFilePath: '',
            exportFileName: '',
            updatedAt: now
        }, existingIndex);
    } else {
        currentData.records.unshift(normalizeProductPublishRecord({
            id: `product-${Date.now()}`,
            sourceTaskKey,
            groupName,
            productType: inferredProductType,
            mainCode: groupName,
            outputDir: String(payload?.outputDir || '').trim(),
            designName: String(payload?.designName || '').trim(),
            sceneNames: Array.isArray(payload?.sceneNames) ? payload.sceneNames : [],
            images: normalizedImages,
            titleEn: '',
            titleZh: '',
            titleStatus: 'pending',
            urls: [],
            urlStatus: 'pending',
            exportStatus: 'idle',
            categoryId: '',
            shipLeadTime: '',
            originPlace: '',
            previewImageUrl: '',
            customized: '否',
            specName1: '',
            specName2: '',
            specValue1: '',
            specValue2: '',
            declaredPrice: '',
            suggestedPrice: '',
            lengthCm: '',
            widthCm: '',
            heightCm: '',
            weightG: '',
            inventory: '',
            sensitive: '否',
            createdAt: now,
            updatedAt: now
        }));
    }
    return saveProductPublishData(currentData);
}

function getProductPublishWorkflowExportDefaults(cfg, nodeConfig = {}) {
    const defaults = {
        ...createDefaultProductPublishConfig().exportTemplateDefaults,
        ...(cfg?.exportTemplateDefaults || {})
    };
    const profileId = String(nodeConfig?.exportProfileId || '').trim();
    const profile = profileId
        ? (Array.isArray(cfg?.exportTemplateProfiles) ? cfg.exportTemplateProfiles : []).find((item) => item.id === profileId)
        : null;
    return {
        ...defaults,
        ...(profile?.fields || {}),
        mainCodePrefix: String((profile?.fields || defaults).mainCodePrefix || 'A').trim() || 'A',
        categoryId: String((profile?.fields || defaults).categoryId || '124300').trim() || '124300',
        outputDir: normalizeDirectoryPath((profile?.fields || defaults).outputDir, getDefaultProductPublishOutputDir()),
        urlPrefix: String((profile?.fields || defaults).urlPrefix || '').trim(),
        ossBucket: String((profile?.fields || defaults).ossBucket || '').trim(),
        ossRegion: String((profile?.fields || defaults).ossRegion || '').trim(),
        ossAccessKeyId: String((profile?.fields || defaults).ossAccessKeyId || '').trim(),
        ossAccessKeySecret: String((profile?.fields || defaults).ossAccessKeySecret || '').trim(),
        ossObjectPrefix: String((profile?.fields || defaults).ossObjectPrefix || 'products').trim() || 'products',
        shipLeadTime: String((profile?.fields || defaults).shipLeadTime || '2').trim() || '2',
        originPlace: String((profile?.fields || defaults).originPlace || '中国-浙江省').trim() || '中国-浙江省',
        customized: String((profile?.fields || defaults).customized || '否').trim() || '否',
        specName1: String((profile?.fields || defaults).specName1 || '尺寸').trim() || '尺寸',
        specName2: String((profile?.fields || defaults).specName2 || '颜色').trim() || '颜色',
        specValue1: String((profile?.fields || defaults).specValue1 || '白色').trim() || '白色',
        specValue2: String((profile?.fields || defaults).specValue2 || '0').trim() || '0',
        declaredPrice: String((profile?.fields || defaults).declaredPrice || '0.01').trim() || '0.01',
        suggestedPrice: String((profile?.fields || defaults).suggestedPrice || '').trim(),
        lengthCm: String((profile?.fields || defaults).lengthCm || '0').trim() || '0',
        widthCm: String((profile?.fields || defaults).widthCm || '0').trim() || '0',
        heightCm: String((profile?.fields || defaults).heightCm || '0').trim() || '0',
        weightG: String((profile?.fields || defaults).weightG || '0').trim() || '0',
        inventory: String((profile?.fields || defaults).inventory || '0').trim() || '0',
        sensitive: String((profile?.fields || defaults).sensitive || '否').trim() || '否'
    };
}

function buildProductPublishWorkflowUrlPrefixUrls(record, bulk) {
    const prefix = String(bulk?.urlPrefix || '').trim().replace(/\/+$/, '');
    if (!prefix) return [];
    const folderName = encodeURIComponent(String(record?.groupName || 'product').trim() || 'product');
    return (Array.isArray(record?.images) ? record.images : []).map((image, index) => {
        const rawName = String(
            image?.name
            || path.basename(String(image?.path || ''))
            || `image-${index + 1}.jpg`
        ).trim() || `image-${index + 1}.jpg`;
        return `${prefix}/${folderName}/${encodeURIComponent(rawName)}`;
    }).filter(Boolean);
}

function applyProductPublishWorkflowExportFields(records, bulk) {
    const prefix = String(bulk?.mainCodePrefix || 'A').trim() || 'A';
    return (Array.isArray(records) ? records : []).map((record, index) => {
        const urls = Array.isArray(record.urls) ? record.urls.map((item) => String(item || '').trim()).filter(Boolean) : [];
        const firstUrl = String(urls[0] || '').trim();
        return normalizeProductPublishRecord({
            ...record,
            urls,
            urlStatus: firstUrl ? 'ready' : 'pending',
            categoryId: String(bulk?.categoryId || '').trim(),
            mainCode: `${prefix}${index + 1}`,
            shipLeadTime: String(bulk?.shipLeadTime || '2').trim() || '2',
            originPlace: String(bulk?.originPlace || '中国-浙江省').trim() || '中国-浙江省',
            previewImageUrl: firstUrl,
            customized: String(bulk?.customized || '否').trim() || '否',
            specName1: String(bulk?.specName1 || '').trim(),
            specName2: String(bulk?.specName2 || '').trim(),
            specValue1: String(bulk?.specValue1 || '').trim(),
            specValue2: String(bulk?.specValue2 || '').trim(),
            declaredPrice: String(bulk?.declaredPrice || '').trim(),
            suggestedPrice: String(bulk?.suggestedPrice || '').trim(),
            lengthCm: String(bulk?.lengthCm || '0').trim() || '0',
            widthCm: String(bulk?.widthCm || '0').trim() || '0',
            heightCm: String(bulk?.heightCm || '0').trim() || '0',
            weightG: String(bulk?.weightG || '0').trim() || '0',
            inventory: String(bulk?.inventory || '0').trim() || '0',
            sensitive: String(bulk?.sensitive || '否').trim() || '否',
            updatedAt: new Date().toISOString()
        }, index);
    });
}

function buildProductPublishWorkflowRecords(run, artifacts, cfg) {
    const outputPaths = (Array.isArray(artifacts?.templateOutputPaths) && artifacts.templateOutputPaths.length
        ? artifacts.templateOutputPaths
        : (Array.isArray(artifacts?.inputImages) ? artifacts.inputImages : []))
        .map((item) => String(item || '').trim())
        .filter((filePath) => filePath && fs.existsSync(filePath) && fs.statSync(filePath).isFile());
    if (!outputPaths.length) {
        throw new Error('产品发布节点缺少模板输出图片');
    }

    const groups = new Map();
    outputPaths.forEach((filePath) => {
        const folder = path.dirname(filePath);
        if (!groups.has(folder)) groups.set(folder, []);
        groups.get(folder).push(filePath);
    });

    return Array.from(groups.entries()).map(([folder, files], index) => {
        const groupName = path.basename(folder) || path.basename(run.name, path.extname(run.name)) || `产品${index + 1}`;
        const images = files
            .sort((a, b) => a.localeCompare(b, 'zh-CN'))
            .map((filePath, imageIndex) => ({
                id: `image-${imageIndex + 1}`,
                name: path.basename(filePath),
                path: filePath,
                sceneName: path.basename(filePath, path.extname(filePath))
            }));
        const sceneNames = images.map((item) => item.sceneName).filter(Boolean);
        const productType = inferProductPublishTypeFromNames([
            groupName,
            run.name,
            sceneNames,
            images.map((item) => item.name)
        ], cfg.productTypeMappings);
        return normalizeProductPublishRecord({
            id: `product-workflow-${Date.now()}-${index + 1}-${crypto.randomBytes(2).toString('hex')}`,
            sourceTaskKey: `workflow::${run.id}::${folder}`,
            groupName,
            productType,
            mainCode: groupName,
            outputDir: folder,
            designName: path.basename(run.name, path.extname(run.name)),
            sceneNames,
            images,
            titleEn: '',
            titleZh: '',
            titleStatus: 'pending',
            urls: [],
            urlStatus: 'pending',
            exportStatus: 'idle',
            customized: '否',
            sensitive: '否'
        }, index);
    });
}

function saveProductPublishWorkflowRecords(records) {
    const data = loadProductPublishData();
    const nextRecords = Array.isArray(data.records) ? data.records.slice() : [];
    (Array.isArray(records) ? records : []).forEach((record) => {
        const normalized = normalizeProductPublishRecord(record);
        const existingIndex = nextRecords.findIndex((item) => item.sourceTaskKey === normalized.sourceTaskKey);
        if (existingIndex >= 0) {
            nextRecords[existingIndex] = normalizeProductPublishRecord({
                ...nextRecords[existingIndex],
                ...normalized,
                id: nextRecords[existingIndex].id || normalized.id,
                createdAt: nextRecords[existingIndex].createdAt || normalized.createdAt
            }, existingIndex);
        } else {
            nextRecords.unshift(normalized);
        }
    });
    return saveProductPublishData({ records: nextRecords });
}

function validateProductPublishWorkflowRecords(records, bulk) {
    const errors = [];
    const titleLimit = 250;
    (Array.isArray(records) ? records : []).forEach((record) => {
        const name = String(record?.groupName || '未命名产品').trim() || '未命名产品';
        if (!Array.isArray(record?.urls) || !record.urls.length) {
            errors.push(`“${name}”缺少图片 URL`);
        }
        if (!String(record?.titleZh || '').trim()) {
            errors.push(`“${name}”缺少中文标题`);
        }
        if (!String(record?.titleEn || '').trim()) {
            errors.push(`“${name}”缺少英文标题`);
        }
        if (Array.from(String(record?.titleZh || '')).length > titleLimit) {
            errors.push(`“${name}”中文标题超过 ${titleLimit} 字符`);
        }
        if (Array.from(String(record?.titleEn || '')).length > titleLimit) {
            errors.push(`“${name}”英文标题超过 ${titleLimit} 字符`);
        }
    });
    const declaredPrice = Number(String(bulk?.declaredPrice || '').trim());
    if (!Number.isFinite(declaredPrice) || declaredPrice < 0.01) {
        errors.push('申报价必须大于等于 0.01');
    }
    return errors;
}

async function runWorkflowPublishNode(run, node, artifacts, sender) {
    const cfg = loadProductPublishConfig();
    const bulk = getProductPublishWorkflowExportDefaults(cfg, node.config || {});
    let records = buildProductPublishWorkflowRecords(run, artifacts, cfg);

    const promptPresetId = String(node.config?.promptPresetId || cfg.titlePromptPresetId || '').trim();
    const promptPreset = (Array.isArray(cfg.titlePromptPresets) ? cfg.titlePromptPresets : []).find((item) => item.id === promptPresetId);
    const titleCfg = {
        ...cfg,
        titlePromptPresetId: promptPresetId || cfg.titlePromptPresetId,
        titlePromptDoc: String(promptPreset?.doc || cfg.titlePromptDoc || DEFAULT_PRODUCT_PUBLISH_PROMPT_DOC).trim()
    };

    sendTemplateLog(sender, 'info', `工作流产品发布开始：${records.length} 个父级产品`);
    for (let index = 0; index < records.length; index += 1) {
        const record = records[index];
        sendTemplateLog(sender, 'info', `正在生成标题：${record.groupName || `产品${index + 1}`}`);
        const title = await generateProductPublishTitle(record, titleCfg);
        records[index] = normalizeProductPublishRecord({
            ...record,
            titleEn: title.titleEn,
            titleZh: title.titleZh,
            titleStatus: 'generated',
            titleHistory: [
                ...(Array.isArray(record.titleHistory) ? record.titleHistory : []),
                {
                    id: `title-history-${Date.now()}-${index + 1}`,
                    titleEn: title.titleEn,
                    titleZh: title.titleZh,
                    createdAt: new Date().toISOString()
                }
            ]
        }, index);
    }

    if (isProductPublishOssConfigured(bulk)) {
        for (let index = 0; index < records.length; index += 1) {
            const record = records[index];
            sendTemplateLog(sender, 'info', `正在上传图片：${record.groupName || `产品${index + 1}`}`);
            const uploadedUrls = await uploadProductPublishRecordImagesToOss(record, bulk);
            records[index] = normalizeProductPublishRecord({
                ...record,
                urls: uploadedUrls,
                urlStatus: uploadedUrls.length ? 'ready' : 'pending',
                previewImageUrl: uploadedUrls[0] || ''
            }, index);
        }
    } else {
        records = records.map((record, index) => normalizeProductPublishRecord({
            ...record,
            urls: buildProductPublishWorkflowUrlPrefixUrls(record, bulk),
            urlStatus: buildProductPublishWorkflowUrlPrefixUrls(record, bulk).length ? 'ready' : 'pending'
        }, index));
    }

    records = applyProductPublishWorkflowExportFields(records, bulk);
    const validationErrors = validateProductPublishWorkflowRecords(records, bulk);
    if (validationErrors.length) {
        throw new Error(validationErrors.join('\n'));
    }

    const templatePath = resolveProductPublishTemuTemplatePath();
    if (!templatePath) {
        throw new Error(`未找到 ${PRODUCT_PUBLISH_TEMU_TEMPLATE_NAME}，请放到下载、文档或项目根目录`);
    }

    const workbook = buildProductPublishTemuWorkbook(records, templatePath);
    const outputDir = ensureDir(normalizeDirectoryPath(bulk.outputDir, getDefaultProductPublishOutputDir()));
    const filePath = ensureUniqueFilePath(path.join(outputDir, buildProductPublishExportFileName(records.length)));
    XLSX.writeFile(workbook, filePath, { compression: true });

    const now = new Date().toISOString();
    records = records.map((record, index) => normalizeProductPublishRecord({
        ...record,
        exportStatus: 'done',
        exportedAt: now,
        exportFilePath: filePath,
        exportFileName: path.basename(filePath)
    }, index));
    saveProductPublishWorkflowRecords(records);
    sendTemplateLog(sender, 'success', `工作流产品发布完成：${filePath}`);

    return {
        publishRecords: records,
        publishOutputPaths: [filePath],
        publishOutputDir: outputDir
    };
}

function resolveAiChatCompletionsUrl(rawUrl) {
    const value = String(rawUrl || '').trim();
    if (!value) return '';
    const normalized = value.replace(/\/+$/, '');
    if (/\/chat\/completions\/?$/i.test(normalized)) return normalized;
    if (/\/v1\/?$/i.test(normalized)) return `${normalized}/chat/completions`;
    if (/\/v1\/chat\/completions\/?$/i.test(normalized)) return normalized;
    return `${normalized}/v1/chat/completions`;
}

function resolveAiModelsUrl(rawUrl) {
    const value = String(rawUrl || '').trim();
    if (!value) return '';
    const normalized = value.replace(/\/+$/, '');
    if (/\/chat\/completions$/i.test(normalized)) {
        return normalized.replace(/\/chat\/completions$/i, '/models');
    }
    if (/\/v1\/chat\/completions$/i.test(normalized)) {
        return normalized.replace(/\/chat\/completions$/i, '/models');
    }
    if (/\/models$/i.test(normalized)) {
        return normalized;
    }
    if (/\/v1\/?$/i.test(normalized)) {
        return `${normalized}/models`;
    }
    return `${normalized}/v1/models`;
}

function resolveAiModelsUrlCandidates(rawUrl) {
    const value = String(rawUrl || '').trim();
    if (!value) return [];
    const normalized = value.replace(/\/+$/, '');
    const candidates = [];
    if (/\/v1\/models$/i.test(normalized) || /\/models$/i.test(normalized)) {
        candidates.push(normalized);
    } else if (/\/v1\/chat\/completions$/i.test(normalized)) {
        candidates.push(normalized.replace(/\/chat\/completions$/i, '/models'));
        candidates.push(normalized.replace(/\/v1\/chat\/completions$/i, '/models'));
    } else if (/\/chat\/completions$/i.test(normalized)) {
        candidates.push(normalized.replace(/\/chat\/completions$/i, '/models'));
        candidates.push(normalized.replace(/\/chat\/completions$/i, '/v1/models'));
    } else if (/\/v1$/i.test(normalized)) {
        candidates.push(`${normalized}/models`);
        candidates.push(normalized.replace(/\/v1$/i, '/models'));
    } else {
        candidates.push(`${normalized}/v1/models`);
        candidates.push(`${normalized}/models`);
    }
    return [...new Set(candidates.filter(Boolean))];
}

function isLikelyMultiProviderGateway(rawUrl) {
    const value = String(rawUrl || '').trim().toLowerCase();
    if (!value) return false;
    return !/api\.openai\.com|api\.anthropic\.com|generativelanguage\.googleapis\.com/.test(value);
}

function inferProductPublishAiProviderFromModel(modelName) {
    const model = String(modelName || '').trim().toLowerCase();
    if (!model) return 'openai';
    if (model.includes('gemini')) return 'gemini';
    if (model.includes('claude')) return 'claude';
    return 'openai';
}

function resolveProductPublishAiProvider(cfg, modelName) {
    const explicit = normalizeProductPublishAiProvider(cfg?.aiProvider);
    if (explicit !== 'auto') return explicit;
    return inferProductPublishAiProviderFromModel(modelName || cfg?.aiModel);
}

function buildProductPublishUserPrompt(record, promptDoc = '') {
    const sceneNames = Array.isArray(record?.sceneNames) && record.sceneNames.length
        ? record.sceneNames.join('、')
        : '未命名场景';
    const imageCount = Array.isArray(record?.images) ? record.images.length : 0;
    const productType = String(record?.productType || '').trim();
    const sections = [
        String(promptDoc || '').trim(),
        `产品模板组：${record?.groupName || '未命名产品'}`,
        `产品类型：${productType || '其他（已按文件夹名或图片名锁定，请严格服从，不要重新识别产品类型）'}`,
        `场景列表：${sceneNames}`,
        `图片数量：${imageCount}`,
        '请严格依据以上用户提示词与当前上传图片完成任务，不要额外发挥。'
    ].filter(Boolean);
    return sections.join('\n\n');
}

function shouldUseProductPublish2DPrefix(record) {
    const joined = [
        record?.productType || '',
        record?.groupName || '',
        ...(Array.isArray(record?.sceneNames) ? record.sceneNames : [])
    ].join(' ').toLowerCase();
    return /2d|flat print|地垫|门垫|浴室垫|floor mat|doormat|accent rug|bathroom rug|entryway mat|rug/.test(joined);
}

function stripProductPublishPrefixes(titleEn, titleZh) {
    const cleanEn = String(titleEn || '')
        .replace(/^\s*\[2d flat print\]\s*1pc\s*/i, '')
        .replace(/^\s*1pc\s*/i, '')
        .trim();
    const cleanZh = String(titleZh || '')
        .replace(/^\s*【2d平面打印】\s*一件\s*/i, '')
        .replace(/^\s*一件\s*/i, '')
        .trim();
    return { cleanEn, cleanZh };
}

function enforceProductPublishTitleRules(record, result) {
    const use2D = shouldUseProductPublish2DPrefix(record);
    const { cleanEn, cleanZh } = stripProductPublishPrefixes(result?.titleEn, result?.titleZh);
    return {
        titleEn: `${use2D ? '[2D Flat Print]1pc ' : '1pc '}${cleanEn}`.trim(),
        titleZh: `${use2D ? '【2D平面打印】一件' : '一件'}${cleanZh}`.trim()
    };
}

function parseProductPublishTitleResult(rawContent) {
    const raw = String(rawContent || '').trim();
    const lines = raw
        .trim()
        .replace(/\r/g, '\n')
        .split('\n')
        .map((line) => String(line || '').trim())
        .filter(Boolean)
        .map((line) => line.replace(/^(EN|CN|英文标题|中文标题)\s*[:：]\s*/i, '').trim())
        .filter(Boolean);
    let titleEn = '';
    let titleZh = '';
    for (const line of lines) {
        if (!titleEn && /[A-Za-z]/.test(line) && !/[\u4E00-\u9FFF]/.test(line)) {
            titleEn = line.replace(/^["“”']+|["“”']+$/g, '');
            continue;
        }
        if (!titleZh && /[\u4E00-\u9FFF]/.test(line)) {
            titleZh = line.replace(/^["“”']+|["“”']+$/g, '');
            continue;
        }
    }
    if ((!titleEn || !titleZh) && lines.length >= 2) {
        titleEn = titleEn || lines[0].replace(/^["“”']+|["“”']+$/g, '');
        titleZh = titleZh || lines[1].replace(/^["“”']+|["“”']+$/g, '');
    }
    if (!titleEn || !titleZh) {
        const rawPreview = raw ? raw.slice(0, 1200) : '(空)';
        throw new Error(`AI 未返回可用的中英双标题。\n\nAI 原始返回：\n${rawPreview}`);
    }
    return { titleEn, titleZh };
}

function isProductPublishVisionUnsupportedError(text) {
    const content = String(text || '');
    return /unknown variant\s+[`'"]?image_url|image_url|vision|multimodal|does not support images|not support image|image input/i.test(content);
}

async function requestProductPublishChatCompletion(apiUrl, headers, body, model) {
    const response = await fetch(apiUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(body)
    });
    const rawText = await response.text().catch(() => '');
    if (!response.ok) {
        if (/model not exist|model_not_found|does not exist|invalid model/i.test(rawText)) {
            throw new Error(`模型不存在：${model}`);
        }
        throw new Error(`AI 标题生成失败：${response.status}${rawText ? ` ${rawText}` : ''}`);
    }
    let payload = {};
    try {
        payload = rawText ? JSON.parse(rawText) : {};
    } catch {
        throw new Error(`AI 标题生成失败：接口返回的不是合法 JSON${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
    }
    return parseProductPublishTitleResult(payload?.choices?.[0]?.message?.content || '');
}

function resolveAiGeminiGenerateContentUrl(rawUrl, model) {
    const value = String(rawUrl || '').trim();
    const modelName = String(model || '').trim();
    if (!value || !modelName) return '';
    const normalized = value.replace(/\/+$/, '');
    if (/\/v1beta\/models\/[^/]+:generateContent$/i.test(normalized)) {
        return normalized.replace(/\/v1beta\/models\/[^/]+:generateContent$/i, `/v1beta/models/${modelName}:generateContent`);
    }
    if (/\/v1beta$/i.test(normalized)) {
        return `${normalized}/models/${modelName}:generateContent`;
    }
    if (/\/v1$/i.test(normalized)) {
        return `${normalized.replace(/\/v1$/i, '/v1beta')}/models/${modelName}:generateContent`;
    }
    return `${normalized}/v1beta/models/${modelName}:generateContent`;
}

function appendApiKeyQueryParam(url, apiKey) {
    const trimmedUrl = String(url || '').trim();
    const trimmedKey = String(apiKey || '').trim();
    if (!trimmedUrl || !trimmedKey) return trimmedUrl;
    return `${trimmedUrl}${trimmedUrl.includes('?') ? '&' : '?'}key=${encodeURIComponent(trimmedKey)}`;
}

function parseProductPublishGeminiTitleResult(payload) {
    const parts = Array.isArray(payload?.candidates?.[0]?.content?.parts)
        ? payload.candidates[0].content.parts
        : [];
    const text = parts
        .map((item) => String(item?.text || '').trim())
        .filter(Boolean)
        .join('\n');
    return parseProductPublishTitleResult(text);
}

async function requestProductPublishGeminiGenerateContent(apiUrl, apiKey, userPrompt, images, model) {
    const finalUrl = appendApiKeyQueryParam(apiUrl, apiKey);
    const response = await fetch(finalUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            generationConfig: {
                temperature: 0.2,
                topP: 0.2
            },
            contents: [
                {
                    role: 'user',
                    parts: [
                        { text: userPrompt },
                        ...buildProductPublishGeminiVisionParts(images)
                    ]
                }
            ]
        })
    });
    const rawText = await response.text().catch(() => '');
    if (!response.ok) {
        if (/model not exist|model_not_found|does not exist|invalid model/i.test(rawText)) {
            throw new Error(`模型不存在：${model}`);
        }
        throw new Error(`AI 标题生成失败：${response.status}${rawText ? ` ${rawText}` : ''}`);
    }
    let payload = {};
    try {
        payload = rawText ? JSON.parse(rawText) : {};
    } catch {
        throw new Error(`AI 标题生成失败：接口返回的不是合法 JSON${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
    }
    return parseProductPublishGeminiTitleResult(payload);
}

async function detectProductPublishModels(cfg) {
    const config = {
        ...createDefaultProductPublishConfig(),
        ...(cfg || {})
    };
    const provider = resolveProductPublishAiProvider(config, config.aiModel);
    if (provider === 'text') {
        throw new Error('纯文本接口模式不支持自动识别模型，请手动填写模型名');
    }
    const urls = resolveAiModelsUrlCandidates(config.aiApiUrl);
    if (!urls.length) {
        throw new Error('请先填写 AI 接口地址');
    }
    const headers = {
        'Content-Type': 'application/json'
    };
    const apiKey = String(config.aiApiKey || '').trim();
    if (apiKey) {
        headers.Authorization = `Bearer ${apiKey}`;
    }
    let lastError = null;
    let detectedModels = [];
    for (const apiUrl of urls) {
        try {
            const response = await fetch(apiUrl, {
                method: 'GET',
                headers
            });
            const rawText = await response.text();
            if (!response.ok) {
                lastError = new Error(`模型识别失败：${response.status}${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
                continue;
            }
            let payload = {};
            try {
                payload = rawText ? JSON.parse(rawText) : {};
            } catch {
                lastError = new Error('模型识别失败：接口返回的不是合法 JSON');
                continue;
            }
            const models = Array.isArray(payload?.data)
                ? payload.data.map((item) => String(item?.id || '').trim()).filter(Boolean)
                : [];
            if (models.length) {
                detectedModels = models;
                break;
            }
            lastError = new Error('模型识别失败：接口没有返回可用模型');
        } catch (error) {
            lastError = error;
        }
    }
    const mergedModels = [...new Set([
        ...detectedModels,
        ...(isLikelyMultiProviderGateway(config.aiApiUrl) ? PRODUCT_PUBLISH_GATEWAY_EXTRA_MODELS : []),
        ...(isLikelyMultiProviderGateway(config.aiApiUrl) ? PRODUCT_PUBLISH_CLAUDE_MODELS : []),
        ...(isLikelyMultiProviderGateway(config.aiApiUrl) ? PRODUCT_PUBLISH_GEMINI_MODELS : [])
    ])].filter(Boolean);
    if (mergedModels.length) {
        return {
            models: mergedModels,
            preferredModel: detectedModels[0] || mergedModels[0],
            provider: inferProductPublishAiProviderFromModel(detectedModels[0] || mergedModels[0])
        };
    }
    throw lastError || new Error('模型识别失败');
}

async function testProductPublishModel(cfg) {
    const config = {
        ...createDefaultProductPublishConfig(),
        ...(cfg || {})
    };
    const apiKey = String(config.aiApiKey || '').trim();
    const model = String(config.aiModel || '').trim();
    const provider = resolveProductPublishAiProvider(config, model);
    const apiUrl = provider === 'gemini'
        ? resolveAiGeminiGenerateContentUrl(config.aiApiUrl, model)
        : resolveAiChatCompletionsUrl(config.aiApiUrl);
    if (!apiUrl || !model) {
        throw new Error('请先填写 AI 接口地址和模型名称');
    }
    if (provider === 'gemini') {
        return requestProductPublishGeminiGenerateContent(
            apiUrl,
            apiKey,
            'You are a model connectivity checker.',
            'Reply with OK only.',
            [],
            model
        );
    }
    const headers = {
        'Content-Type': 'application/json'
    };
    if (apiKey) {
        headers.Authorization = `Bearer ${apiKey}`;
    }
    const response = await fetch(apiUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify({
            model,
            temperature: 0,
            max_tokens: 12,
            messages: [
                { role: 'system', content: 'You are a model connectivity checker.' },
                { role: 'user', content: 'Reply with OK only.' }
            ]
        })
    });
    const rawText = await response.text().catch(() => '');
    if (!response.ok) {
        if (/model not exist|model_not_found|does not exist|invalid model/i.test(rawText)) {
            throw new Error(`模型不存在：${model}`);
        }
        throw new Error(`测试模型失败：${response.status}${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
    }
    let payload = {};
    try {
        payload = rawText ? JSON.parse(rawText) : {};
    } catch {
        throw new Error(`测试模型失败：接口返回的不是合法 JSON${rawText ? ` ${rawText.slice(0, 300)}` : ''}`);
    }
    return {
        ok: true,
        content: String(payload?.choices?.[0]?.message?.content || '').trim()
    };
}

async function generateProductPublishTitle(record, cfg) {
    const config = {
        ...createDefaultProductPublishConfig(),
        ...(cfg || {})
    };
    const apiKey = String(config.aiApiKey || '').trim();
    const model = String(config.aiModel || '').trim();
    const provider = resolveProductPublishAiProvider(config, model);
    const apiUrl = provider === 'gemini'
        ? resolveAiGeminiGenerateContentUrl(config.aiApiUrl, model)
        : resolveAiChatCompletionsUrl(config.aiApiUrl);
    if (!apiUrl || !model) {
        throw new Error('请先填写 AI 接口地址和模型名称');
    }
    const promptDoc = String(config.titlePromptDoc || '').trim() || createDefaultProductPublishConfig().titlePromptDoc;
    const userPrompt = buildProductPublishUserPrompt(record, promptDoc);
    const visionInputs = buildProductPublishVisionInputs(record?.images);
    if (!visionInputs.length) {
        throw new Error('当前记录没有可供识别的图片');
    }
    const textOnlyBody = {
        model,
        temperature: 0.2,
        top_p: 0.2,
        messages: [
            {
                role: 'user',
                content: userPrompt
            }
        ]
    };
    if (provider === 'gemini') {
        const result = await requestProductPublishGeminiGenerateContent(
            apiUrl,
            apiKey,
            userPrompt,
            record?.images,
            model
        );
        return enforceProductPublishTitleRules(record, result);
    }
    const headers = {
        'Content-Type': 'application/json'
    };
    if (apiKey) {
        headers.Authorization = `Bearer ${apiKey}`;
    }
    if (provider === 'text') {
        const result = await requestProductPublishChatCompletion(apiUrl, headers, textOnlyBody, model);
        return enforceProductPublishTitleRules(record, result);
    }
    const visionBody = {
            model,
            temperature: 0.2,
            top_p: 0.2,
            messages: [
                {
                    role: 'user',
                    content: [
                        {
                            type: 'text',
                            text: userPrompt
                        },
                        ...visionInputs
                    ]
                }
            ]
        };
    try {
        const result = await requestProductPublishChatCompletion(apiUrl, headers, visionBody, model);
        return enforceProductPublishTitleRules(record, result);
    } catch (error) {
        if (provider === 'openai' && isProductPublishVisionUnsupportedError(error.message || '')) {
            const result = await requestProductPublishChatCompletion(apiUrl, headers, textOnlyBody, model);
            return enforceProductPublishTitleRules(record, result);
        }
        if (provider === 'auto' && isProductPublishVisionUnsupportedError(error.message || '')) {
            const result = await requestProductPublishChatCompletion(apiUrl, headers, textOnlyBody, model);
            return enforceProductPublishTitleRules(record, result);
        }
        throw error;
    }
}

function escapeCsvCell(value) {
    const text = String(value ?? '');
    if (/[",\r\n]/.test(text)) {
        return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
}

function buildProductPublishCsv(records) {
    const header = ['产品名称', '图片URL', '中文标题', '英文标题', '来源模板', '场景列表', 'URL状态'];
    const lines = [header.map(escapeCsvCell).join(',')];
    (Array.isArray(records) ? records : []).forEach((record) => {
        const urls = Array.isArray(record?.urls) && record.urls.length ? record.urls : [''];
        urls.forEach((url) => {
            lines.push([
                record?.groupName || '',
                url,
                record?.titleZh || '',
                record?.titleEn || '',
                record?.sourceTaskKey || '',
                Array.isArray(record?.sceneNames) ? record.sceneNames.join(' / ') : '',
                record?.urlStatus || 'pending'
            ].map(escapeCsvCell).join(','));
        });
    });
    return `\uFEFF${lines.join('\r\n')}`;
}

function walkDirectoryFiles(dirPath) {
    const files = [];
    const stack = [dirPath];
    while (stack.length) {
        const current = stack.pop();
        if (!current || !fs.existsSync(current)) continue;
        for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
            const entryPath = path.join(current, entry.name);
            if (entry.isDirectory()) {
                stack.push(entryPath);
            } else if (entry.isFile()) {
                files.push(entryPath);
            }
        }
    }
    return files;
}

function buildProductPublishRecordFromFolder(folderPath) {
    const resolvedFolder = path.resolve(folderPath);
    const groupName = path.basename(resolvedFolder);
    const files = walkDirectoryFiles(resolvedFolder)
        .filter((filePath) => PRODUCT_PUBLISH_IMAGE_EXTS.has(path.extname(filePath).toLowerCase()))
        .sort((a, b) => a.localeCompare(b, 'zh-CN'));
    const images = files.map((filePath, index) => ({
        id: `image-${index + 1}`,
        name: path.basename(filePath),
        path: filePath,
        sceneName: path.basename(filePath, path.extname(filePath))
    }));
    const cfg = loadProductPublishConfig();
    const productType = inferProductPublishTypeFromNames([
        groupName,
        images.map((item) => item.name),
        images.map((item) => item.sceneName)
    ], cfg.productTypeMappings);
    return normalizeProductPublishRecord({
        id: `product-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        sourceTaskKey: `folder::${resolvedFolder}`,
        groupName,
        productType,
        mainCode: groupName,
        outputDir: resolvedFolder,
        designName: '',
        sceneNames: images.map((item) => item.sceneName).filter(Boolean),
        images,
        titleEn: '',
        titleZh: '',
        titleStatus: 'pending',
        urls: [],
        urlStatus: 'pending',
        exportStatus: 'idle',
        categoryId: '',
        shipLeadTime: '',
        originPlace: '',
        previewImageUrl: '',
        customized: '否',
        specName1: '',
        specName2: '',
        specValue1: '',
        specValue2: '',
        declaredPrice: '',
        suggestedPrice: '',
        lengthCm: '',
        widthCm: '',
        heightCm: '',
        weightG: '',
        inventory: '',
        sensitive: '否',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
    });
}

function normalizeUpdateSource(value) {
    return String(value || '').trim().replace(/[\\/]+$/, '');
}

function parseGitHubRepoSource(value) {
    const raw = normalizeUpdateSource(value)
        .replace(/^git\+/, '')
        .replace(/\.git$/i, '');
    if (!raw) return null;

    let owner = '';
    let repo = '';

    const githubUrlMatch = raw.match(/^https?:\/\/github\.com\/([^\/]+)\/([^\/?#]+)(?:[\/?#].*)?$/i);
    if (githubUrlMatch) {
        owner = githubUrlMatch[1];
        repo = githubUrlMatch[2];
    } else {
        const shortMatch = raw.match(/^([A-Za-z0-9_.-]+)\/([A-Za-z0-9_.-]+)$/);
        if (shortMatch) {
            owner = shortMatch[1];
            repo = shortMatch[2];
        }
    }

    if (!owner || !repo) return null;
    return {
        provider: 'github',
        owner,
        repo,
        label: `${owner}/${repo}`
    };
}

function loadPackageRepositorySource() {
    try {
        const pkg = JSON.parse(fs.readFileSync(PACKAGE_JSON_FILE, 'utf-8'));
        const repository = pkg && pkg.repository;
        if (typeof repository === 'string') {
            return normalizeUpdateSource(repository);
        }
        if (repository && typeof repository.url === 'string') {
            return normalizeUpdateSource(repository.url);
        }
    } catch {
        // ignore
    }
    return '';
}

function resolveUpdateSource(cfg = {}) {
    const rawSource = normalizeUpdateSource(cfg.source !== undefined ? cfg.source : cfg.feedUrl);
    const githubFromSource = parseGitHubRepoSource(rawSource);
    if (githubFromSource) {
        return githubFromSource;
    }

    if (/^https?:\/\//i.test(rawSource)) {
        return {
            provider: 'generic',
            url: rawSource,
            label: rawSource
        };
    }

    const packageRepoSource = loadPackageRepositorySource();
    const githubFromPackage = parseGitHubRepoSource(packageRepoSource);
    if (githubFromPackage) {
        return githubFromPackage;
    }

    return null;
}

function loadUpdateConfig() {
    const defaults = {
        source: '',
        feedUrl: '',
        autoCheckOnStartup: false
    };
    try {
        const cfg = JSON.parse(fs.readFileSync(UPDATE_CONFIG_FILE, 'utf-8'));
        const source = normalizeUpdateSource((cfg && (cfg.source !== undefined ? cfg.source : cfg.feedUrl)) || '');
        return {
            ...defaults,
            ...(cfg || {}),
            source,
            feedUrl: source,
            autoCheckOnStartup: Boolean(cfg && cfg.autoCheckOnStartup)
        };
    } catch {
        return defaults;
    }
}

function saveUpdateConfig(cfg) {
    const current = loadUpdateConfig();
    const nextSource = normalizeUpdateSource(
        cfg && cfg.source !== undefined
            ? cfg.source
            : (cfg && cfg.feedUrl !== undefined ? cfg.feedUrl : current.source)
    );
    const nextCfg = {
        ...current,
        ...(cfg || {}),
        source: nextSource,
        feedUrl: nextSource,
        autoCheckOnStartup: Boolean(cfg && cfg.autoCheckOnStartup !== undefined ? cfg.autoCheckOnStartup : current.autoCheckOnStartup)
    };
    fs.writeFileSync(UPDATE_CONFIG_FILE, JSON.stringify(nextCfg, null, 2), 'utf-8');
    return nextCfg;
}

const TEMPLATE_REQUIRED_FILES = ['base.png', 'mask.png', 'config.json'];
const TEMPLATE_LAYER_REQUIREMENT_TEXT = 'texture.png + highlight.png';

function toTemplateRelativePath(targetPath) {
    return path.relative(getTemplateRootDir(), targetPath).split(path.sep).join('/');
}

function getTemplatePreviewPath(templateDir) {
    const previewCandidates = ['preview.png', 'preview.jpg', 'preview.jpeg', 'base.png'];
    const previewFile = previewCandidates.find((fileName) => fs.existsSync(path.join(templateDir, fileName)));
    return previewFile ? path.join(templateDir, previewFile) : '';
}

function getTemplatePlacement(config = {}) {
    const placement = config && typeof config === 'object' && config.placement && typeof config.placement === 'object'
        ? config.placement
        : {};
    return {
        scale: Number.isFinite(Number(placement.scale)) ? Number(placement.scale) : 1,
        offsetX: Number.isFinite(Number(placement.offsetX)) ? Number(placement.offsetX) : 0,
        offsetY: Number.isFinite(Number(placement.offsetY)) ? Number(placement.offsetY) : 0,
        rotation: Number.isFinite(Number(placement.rotation)) ? Number(placement.rotation) : 0
    };
}

function getTemplateEffects(config = {}) {
    const effects = config && typeof config === 'object' && config.effects && typeof config.effects === 'object'
        ? config.effects
        : {};
    const designBlendMode = String(effects.designBlendMode || '').trim().toLowerCase() === 'normal'
        ? 'normal'
        : 'multiply';
    return {
        designBlendMode,
        designOpacity: Number.isFinite(Number(effects.designOpacity)) ? Number(effects.designOpacity) : 1,
        designBrightness: Number.isFinite(Number(effects.designBrightness)) ? Number(effects.designBrightness) : 1,
        textureOpacity: Number.isFinite(Number(effects.textureOpacity)) ? Number(effects.textureOpacity) : 1,
        highlightOpacity: Number.isFinite(Number(effects.highlightOpacity)) ? Number(effects.highlightOpacity) : 1
    };
}

function parseTemplatePoint(item) {
    if (Array.isArray(item) && item.length >= 2) {
        return {
            x: Number(item[0]) || 0,
            y: Number(item[1]) || 0
        };
    }
    if (item && typeof item === 'object') {
        return {
            x: Number(item.x) || 0,
            y: Number(item.y) || 0
        };
    }
    return { x: 0, y: 0 };
}

function getTemplatePoints(config = {}) {
    let candidate = null;
    for (const key of ['points', 'vertices', 'corners', 'quad']) {
        if (Array.isArray(config[key])) {
            candidate = config[key];
            break;
        }
    }

    if (!candidate && ['topLeft', 'topRight', 'bottomRight', 'bottomLeft'].every((key) => config[key])) {
        candidate = ['topLeft', 'topRight', 'bottomRight', 'bottomLeft'].map((key) => config[key]);
    }

    if (!Array.isArray(candidate) || candidate.length !== 4) {
        return [];
    }

    return candidate.map(parseTemplatePoint);
}

function describeTemplateScene(templateDir, groupName, sceneName) {
    const configPath = path.join(templateDir, 'config.json');
    const missing = TEMPLATE_REQUIRED_FILES.filter((fileName) => !fs.existsSync(path.join(templateDir, fileName)));
    const hasTextureStack = fs.existsSync(path.join(templateDir, 'texture.png'))
        && fs.existsSync(path.join(templateDir, 'highlight.png'));
    if (!hasTextureStack) {
        missing.push(TEMPLATE_LAYER_REQUIREMENT_TEXT);
    }

    let config = {};
    if (fs.existsSync(configPath)) {
        try {
            config = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
        } catch {
            missing.push('config.json 解析失败');
        }
    }

    return {
        name: sceneName,
        groupName,
        relativePath: toTemplateRelativePath(templateDir),
        previewPath: getTemplatePreviewPath(templateDir),
        basePath: fs.existsSync(path.join(templateDir, 'base.png')) ? path.join(templateDir, 'base.png') : '',
        maskPath: fs.existsSync(path.join(templateDir, 'mask.png')) ? path.join(templateDir, 'mask.png') : '',
        texturePath: fs.existsSync(path.join(templateDir, 'texture.png')) ? path.join(templateDir, 'texture.png') : '',
        highlightPath: fs.existsSync(path.join(templateDir, 'highlight.png')) ? path.join(templateDir, 'highlight.png') : '',
        placement: getTemplatePlacement(config),
        effects: getTemplateEffects(config),
        points: getTemplatePoints(config),
        missing: Array.from(new Set(missing)),
        valid: missing.length === 0
    };
}

function listTemplateFolders() {
    const cfg = loadTemplateConfig();
    const templateRootDir = getTemplateRootDir(cfg);
    ensureDir(templateRootDir);
    const groups = fs.readdirSync(templateRootDir, { withFileTypes: true })
        .filter((entry) => entry.isDirectory())
        .map((entry) => {
            const groupDir = path.join(templateRootDir, entry.name);
            const childDirs = fs.readdirSync(groupDir, { withFileTypes: true })
                .filter((item) => item.isDirectory())
                .map((item) => ({
                    name: item.name,
                    dir: path.join(groupDir, item.name)
                }));

            const directLooksLikeScene = TEMPLATE_REQUIRED_FILES.some((fileName) => fs.existsSync(path.join(groupDir, fileName)))
                || (fs.existsSync(path.join(groupDir, 'texture.png')) && fs.existsSync(path.join(groupDir, 'highlight.png')));

            let scenes = [];
            if (childDirs.length > 0 && !directLooksLikeScene) {
                scenes = childDirs.map((item) => describeTemplateScene(item.dir, entry.name, item.name));
            } else if (directLooksLikeScene) {
                scenes = [describeTemplateScene(groupDir, entry.name, entry.name)];
            } else {
                scenes = [];
            }

            const validScenes = scenes.filter((item) => item.valid);
            const missing = validScenes.length > 0
                ? []
                : Array.from(new Set(scenes.flatMap((item) => item.missing)));

            return {
                name: entry.name,
                previewPath: (scenes.find((item) => item.previewPath) || {}).previewPath || '',
                sceneCount: scenes.length,
                scenes,
                missing,
                valid: validScenes.length > 0
            };
        });
    const orderMap = new Map((Array.isArray(cfg.templateOrder) ? cfg.templateOrder : []).map((name, index) => [name, index]));
    return groups.sort((a, b) => {
        const aOrder = orderMap.has(a.name) ? orderMap.get(a.name) : Number.MAX_SAFE_INTEGER;
        const bOrder = orderMap.has(b.name) ? orderMap.get(b.name) : Number.MAX_SAFE_INTEGER;
        if (aOrder !== bOrder) {
            return aOrder - bOrder;
        }
        return String(a.name || '').localeCompare(String(b.name || ''), 'zh-CN');
    });
}

function moveTemplateGroupOrder(groupName, direction) {
    const name = String(groupName || '').trim();
    const move = String(direction || '').trim().toLowerCase();
    if (!name) {
        throw new Error('缺少模板名称');
    }
    if (!['up', 'down'].includes(move)) {
        throw new Error('无效的移动方向');
    }
    const groups = listTemplateFolders();
    const groupNames = groups.map((item) => item.name);
    if (!groupNames.includes(name)) {
        throw new Error('当前模板不存在');
    }
    const nextOrder = groupNames.slice();
    const index = nextOrder.indexOf(name);
    const targetIndex = move === 'up' ? index - 1 : index + 1;
    if (targetIndex < 0 || targetIndex >= nextOrder.length) {
        return {
            config: saveTemplateConfig({
                ...loadTemplateConfig(),
                templateOrder: nextOrder
            }),
            templates: groups
        };
    }
    const [moved] = nextOrder.splice(index, 1);
    nextOrder.splice(targetIndex, 0, moved);
    const nextCfg = saveTemplateConfig({
        ...loadTemplateConfig(),
        templateOrder: nextOrder
    });
    return {
        config: nextCfg,
        templates: listTemplateFolders()
    };
}

function resolveTemplateSceneDir(relativePath) {
    const normalized = String(relativePath || '').replace(/[\\/]+/g, path.sep);
    const templateRootDir = getTemplateRootDir();
    const rootDir = path.resolve(templateRootDir);
    const resolvedPath = path.resolve(templateRootDir, normalized);
    if (resolvedPath !== rootDir && !resolvedPath.startsWith(`${rootDir}${path.sep}`)) {
        throw new Error('模板场景路径无效');
    }
    return resolvedPath;
}

function sanitizeTemplateSegment(value, fallback = '未命名模板') {
    const trimmed = String(value || '').trim();
    const sanitized = trimmed
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, ' ')
        .replace(/\s+/g, ' ')
        .replace(/\.+$/g, '')
        .trim();
    return sanitized || fallback;
}

async function getDefaultTemplateScenePoints(basePath, maskPath = '') {
    try {
        const resolvedMaskPath = String(maskPath || '').trim();
        if (resolvedMaskPath && fs.existsSync(resolvedMaskPath)) {
            const { data, info } = await sharp(resolvedMaskPath)
                .ensureAlpha()
                .raw()
                .toBuffer({ resolveWithObject: true });
            const width = Number(info.width) || 0;
            const height = Number(info.height) || 0;
            if (width > 0 && height > 0) {
                let minX = width;
                let minY = height;
                let maxX = -1;
                let maxY = -1;
                for (let y = 0; y < height; y += 1) {
                    for (let x = 0; x < width; x += 1) {
                        const offset = (y * width + x) * 4;
                        const alpha = data[offset + 3];
                        const red = data[offset];
                        const green = data[offset + 1];
                        const blue = data[offset + 2];
                        const luminance = (red + green + blue) / 3;
                        if (alpha <= 8 || luminance <= 12) continue;
                        if (x < minX) minX = x;
                        if (y < minY) minY = y;
                        if (x > maxX) maxX = x;
                        if (y > maxY) maxY = y;
                    }
                }
                if (maxX >= minX && maxY >= minY) {
                    const points = [
                        { x: minX, y: minY },
                        { x: maxX, y: minY },
                        { x: maxX, y: maxY },
                        { x: minX, y: maxY }
                    ];
                    const centerX = points.reduce((sum, point) => sum + point.x, 0) / points.length;
                    const centerY = points.reduce((sum, point) => sum + point.y, 0) / points.length;
                    return points.map((point) => ({
                        x: Math.round(centerX + (point.x - centerX) * 1.05),
                        y: Math.round(centerY + (point.y - centerY) * 1.05)
                    }));
                }
            }
        }

        const meta = await sharp(basePath).metadata();
        const width = Number(meta.width);
        const height = Number(meta.height);
        if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
            const insetX = Math.round(width * 0.18);
            const insetY = Math.round(height * 0.18);
            return [
                { x: insetX, y: insetY },
                { x: width - insetX, y: insetY },
                { x: width - insetX, y: height - insetY },
                { x: insetX, y: height - insetY }
            ];
        }
    } catch {}
    return [
        { x: 240, y: 240 },
        { x: 1808, y: 240 },
        { x: 1808, y: 1808 },
        { x: 240, y: 1808 }
    ];
}

function runTemplatePreviewJob(jobPayload) {
    return new Promise((resolve, reject) => {
        const scriptPath = getTemplateRendererScriptPath();
        const pythonRuntime = getPythonRuntime(scriptPath);
        if (!pythonRuntime) {
            reject(new Error('未检测到可用 Python 运行环境，请安装 Python 或 py 启动器'));
            return;
        }

        const child = spawn(pythonRuntime.command, pythonRuntime.scriptArgs, {
            cwd: path.dirname(scriptPath),
            windowsHide: true,
            stdio: ['pipe', 'pipe', 'pipe'],
            env: {
                ...process.env,
                PYTHONUTF8: '1'
            }
        });

        let stdoutBuffer = '';
        let stderrBuffer = '';
        let resolved = false;

        child.stdout.on('data', (chunk) => {
            stdoutBuffer += chunk.toString('utf-8');
            const lines = stdoutBuffer.split(/\r?\n/);
            stdoutBuffer = lines.pop() || '';
            lines.forEach((line) => {
                const text = line.trim();
                if (!text) return;
                try {
                    const message = JSON.parse(text);
                    if (message.type === 'done' && message.outputPath) {
                        resolved = true;
                        resolve(message);
                    }
                } catch {}
            });
        });

        child.stderr.on('data', (chunk) => {
            stderrBuffer += chunk.toString('utf-8');
        });

        child.on('error', (error) => {
            reject(error);
        });

        child.on('close', (code) => {
            if (resolved) return;
            if (stdoutBuffer.trim()) {
                try {
                    const message = JSON.parse(stdoutBuffer.trim());
                    if (message.type === 'done' && message.outputPath) {
                        resolve(message);
                        return;
                    }
                    if (message.message) {
                        reject(new Error(message.message));
                        return;
                    }
                } catch {}
            }
            reject(new Error(stderrBuffer.trim() || `预览生成失败 (code=${code ?? 'null'})`));
        });

        child.stdin.end(JSON.stringify(jobPayload), 'utf-8');
    });
}

function getPythonRuntimeForScript(scriptPath) {
    const candidates = [
        ...getPythonRuntimeCandidates().map((pythonPath) => ({
            command: pythonPath,
            versionArgs: ['--version'],
            scriptArgs: ['-u', scriptPath]
        })),
        { command: 'python', versionArgs: ['--version'], scriptArgs: ['-u', scriptPath] },
        { command: 'py', versionArgs: ['-3', '--version'], scriptArgs: ['-3', '-u', scriptPath] },
        { command: 'python3', versionArgs: ['--version'], scriptArgs: ['-u', scriptPath] }
    ];

    for (const candidate of candidates) {
        try {
            if (candidate.command.endsWith('.exe') && !fs.existsSync(candidate.command)) {
                continue;
            }
            const result = spawnSync(candidate.command, candidate.versionArgs, {
                encoding: 'utf-8',
                windowsHide: true
            });
            if (!result.error && result.status === 0) {
                return candidate;
            }
        } catch {}
    }

    return null;
}

function getPythonRuntime(scriptPath = getTemplateRendererScriptPath()) {
    return getPythonRuntimeForScript(scriptPath);
}

function sanitizePathSegment(value, fallback = 'item') {
    const cleaned = String(value || '')
        .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
        .replace(/\s+/g, ' ')
        .trim()
        .replace(/[. ]+$/g, '');
    return cleaned || fallback;
}

// --- DirectoryWatcher ---
class DirectoryWatcher {
    constructor(dir, onFile) {
        this.dir = dir;
        this.onFile = onFile;
        this.processing = new Set();
        this.pending = new Map();
        this.closed = false;
        this.watcher = null;
    }

    start() {
        this.closed = false;
        this.watcher = chokidar.watch(this.dir, {
            depth: 0,
            ignoreInitial: false,
            awaitWriteFinish: {
                stabilityThreshold: 1500,
                pollInterval: 250
            }
        });
        this.watcher.on('add', (fp) => {
            console.log('[Watcher] File added:', fp);
            this._queue(fp);
        });
        this.watcher.on('change', (fp) => {
            console.log('[Watcher] File changed:', fp);
            this._queue(fp);
        });
        this.watcher.on('ready', () => {
            console.log('[Watcher] Ready, watching:', this.dir);
        });
        this.watcher.on('error', (err) => {
            console.error('[Watcher] Error:', err);
        });
    }

    _shouldSkip(fp) {
        const baseName = path.basename(fp).toLowerCase();
        return (
            baseName.startsWith('~$') ||
            baseName.endsWith('.tmp') ||
            baseName.endsWith('.temp') ||
            baseName.endsWith('.part') ||
            baseName.endsWith('.crdownload') ||
            baseName.endsWith('.download')
        );
    }

    _queue(fp) {
        if (this.closed || this._shouldSkip(fp)) {
            return;
        }
        const existing = this.pending.get(fp);
        if (existing) {
            clearTimeout(existing);
        }
        const timer = setTimeout(() => {
            this.pending.delete(fp);
            this._handle(fp);
        }, 900);
        this.pending.set(fp, timer);
    }

    async _waitForFileStable(fp) {
        let lastSignature = '';
        let stableCount = 0;
        const deadline = Date.now() + 45000;

        while (!this.closed && Date.now() < deadline) {
            try {
                const stat = await fs.promises.stat(fp);
                if (!stat.isFile()) {
                    return false;
                }
                const signature = `${stat.size}:${Math.floor(stat.mtimeMs)}`;
                if (signature === lastSignature) {
                    stableCount += 1;
                    if (stableCount >= 2) {
                        return true;
                    }
                } else {
                    lastSignature = signature;
                    stableCount = 0;
                }
            } catch (error) {
                if (error && error.code !== 'ENOENT') {
                    console.warn('[Watcher] Waiting for stable file failed:', fp, error.message);
                }
                lastSignature = '';
                stableCount = 0;
            }

            await new Promise(resolve => setTimeout(resolve, 600));
        }

        console.warn('[Watcher] File did not stabilize in time, skip:', fp);
        return false;
    }

    async _handle(fp) {
        if (this.closed) return;
        if (this.processing.has(fp)) {
            this._queue(fp);
            return;
        }
        this.processing.add(fp);
        try {
            const stable = await this._waitForFileStable(fp);
            if (!stable || this.closed) {
                return;
            }
            await this.onFile(fp);
        } finally {
            this.processing.delete(fp);
        }
    }

    stop() {
        this.closed = true;
        for (const timer of this.pending.values()) {
            clearTimeout(timer);
        }
        this.pending.clear();
        if (this.watcher) {
            this.watcher.close();
            this.watcher = null;
        }
    }
}

// --- Compression engine ---
const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.tiff', '.tif']);

async function binarySearchQuality(inputBuffer, targetBytes, format) {
    let lo = 1, hi = 100, bestBuf = null;
    while (lo <= hi) {
        const mid = Math.floor((lo + hi) / 2);
        let buf;
        if (format === 'jpeg') {
            buf = await sharp(inputBuffer).jpeg({ quality: mid }).toBuffer();
        } else {
            buf = await sharp(inputBuffer).webp({ quality: mid }).toBuffer();
        }
        if (buf.length <= targetBytes) {
            bestBuf = buf;
            lo = mid + 1;
        } else {
            hi = mid - 1;
        }
    }
    return bestBuf;
}

async function compressImage(filePath, thresholdBytes, sendLog) {
    console.log('[Compress] Processing:', filePath, 'threshold:', thresholdBytes);
    const ext = path.extname(filePath).toLowerCase();
    console.log('[Compress] Extension:', ext);
    if (!IMAGE_EXTS.has(ext)) {
        console.log('[Compress] Not an image file, skipping');
        return;
    }

    let stat;
    try { stat = fs.statSync(filePath); } catch (err) {
        console.log('[Compress] Cannot stat file:', err.message);
        return;
    }
    console.log('[Compress] File size:', stat.size, 'bytes');
    if (stat.size <= thresholdBytes) {
        console.log('[Compress] File size below threshold, skipping');
        return;
    }

    const fileName = path.basename(filePath);
    const originalMB = (stat.size / 1024 / 1024).toFixed(2);
    sendLog('info', `检测到 ${fileName} (${originalMB} MB)，开始压缩...`);

    try {
        const inputBuffer = fs.readFileSync(filePath);
        const meta = await sharp(inputBuffer).metadata();

        // Determine output format
        let outFormat = 'jpeg';
        if (ext === '.webp') outFormat = 'webp';
        // PNG -> convert to JPEG for effective compression
        if (ext === '.png') {
            sendLog('info', `PNG 格式将转为 JPEG 进行压缩`);
        }

        // Try binary search on quality
        let result = await binarySearchQuality(inputBuffer, thresholdBytes, outFormat);

        // Fallback: if quality=1 still too large, scale down
        if (!result) {
            sendLog('info', `质量调整不足，尝试缩小尺寸...`);
            let scale = 0.9;
            while (scale > 0.1) {
                const w = Math.round(meta.width * scale);
                const h = Math.round(meta.height * scale);
                let buf;
                if (outFormat === 'jpeg') {
                    buf = await sharp(inputBuffer).resize(w, h).jpeg({ quality: 50 }).toBuffer();
                } else {
                    buf = await sharp(inputBuffer).resize(w, h).webp({ quality: 50 }).toBuffer();
                }
                if (buf.length <= thresholdBytes) {
                    result = buf;
                    break;
                }
                scale -= 0.1;
            }
        }

        if (!result) {
            sendLog('error', `${fileName} 无法压缩到目标大小以内`);
            return;
        }

        // Write back — change extension to .jpg if was .png
        let outPath = filePath;
        if (ext === '.png') {
            outPath = filePath.replace(/\.png$/i, '.jpg');
        }
        fs.writeFileSync(outPath, result);
        // Remove original .png if converted
        if (ext === '.png' && outPath !== filePath) {
            fs.unlinkSync(filePath);
        }

        const newMB = (result.length / 1024 / 1024).toFixed(2);
        sendLog('success', `${fileName} 压缩完成: ${originalMB} MB → ${newMB} MB`);
    } catch (err) {
        sendLog('error', `${fileName} 压缩失败: ${err.message}`);
    }
}

// --- Safe IPC send (prevents EPIPE crash) ---
function safeSend(sender, channel, data) {
    try {
        if (sender && !sender.isDestroyed()) {
            sender.send(channel, data);
        }
    } catch (e) {
        // ignore broken pipe
    }
}

function getUpdateSnapshot() {
    const cfg = loadUpdateConfig();
    const resolvedSource = resolveUpdateSource(cfg);
    return {
        currentVersion: app.getVersion(),
        packaged: app.isPackaged,
        updateSource: cfg.source,
        updateSourceDisplay: resolvedSource ? resolvedSource.label : '',
        updateProvider: resolvedSource ? resolvedSource.provider : '',
        autoCheckOnStartup: Boolean(cfg.autoCheckOnStartup),
        ...updateState
    };
}

function broadcastUpdateState() {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    safeSend(mainWindow.webContents, 'app-update:state', getUpdateSnapshot());
}

function markUpdateState(patch = {}) {
    updateState = {
        ...updateState,
        ...(patch || {})
    };
    broadcastUpdateState();
}

function bindAutoUpdaterEvents() {
    if (autoUpdaterEventsBound) return;
    autoUpdaterEventsBound = true;
    autoUpdater.autoDownload = false;
    autoUpdater.autoInstallOnAppQuit = false;

    autoUpdater.on('checking-for-update', () => {
        markUpdateState({
            checking: true,
            available: false,
            downloading: false,
            downloaded: false,
            progress: 0,
            error: '',
            status: '正在检查更新...',
            lastCheckedAt: new Date().toISOString()
        });
    });

    autoUpdater.on('update-available', (info) => {
        markUpdateState({
            checking: false,
            available: true,
            downloading: false,
            downloaded: false,
            progress: 0,
            error: '',
            latestVersion: String(info?.version || ''),
            releaseDate: String(info?.releaseDate || ''),
            releaseName: String(info?.releaseName || ''),
            status: `发现新版本 ${String(info?.version || '')}`.trim()
        });
    });

    autoUpdater.on('update-not-available', () => {
        markUpdateState({
            checking: false,
            available: false,
            downloading: false,
            downloaded: false,
            progress: 0,
            error: '',
            latestVersion: '',
            releaseDate: '',
            releaseName: '',
            status: '当前已经是最新版本'
        });
    });

    autoUpdater.on('error', (error) => {
        markUpdateState({
            checking: false,
            downloading: false,
            error: error?.message || '检查更新失败',
            status: error?.message || '检查更新失败'
        });
    });

    autoUpdater.on('download-progress', (progress) => {
        markUpdateState({
            checking: false,
            available: true,
            downloading: true,
            downloaded: false,
            progress: Number(progress?.percent || 0),
            error: '',
            status: `正在下载更新 ${Math.round(Number(progress?.percent || 0))}%`
        });
    });

    autoUpdater.on('update-downloaded', (info) => {
        markUpdateState({
            checking: false,
            available: true,
            downloading: false,
            downloaded: true,
            progress: 100,
            error: '',
            latestVersion: String(info?.version || updateState.latestVersion || ''),
            releaseDate: String(info?.releaseDate || updateState.releaseDate || ''),
            releaseName: String(info?.releaseName || updateState.releaseName || ''),
            status: '更新已下载完成，点击安装更新'
        });
    });
}

function ensureUpdateFeedConfigured() {
    const cfg = loadUpdateConfig();
    const resolvedSource = resolveUpdateSource(cfg);
    if (!resolvedSource) {
        throw new Error('请先在关于中填写 GitHub 仓库或更新地址');
    }
    bindAutoUpdaterEvents();
    const feedKey = resolvedSource.provider === 'github'
        ? `github:${resolvedSource.owner}/${resolvedSource.repo}`
        : `generic:${resolvedSource.url}`;
    if (autoUpdaterFeedUrl !== feedKey) {
        if (resolvedSource.provider === 'github') {
            autoUpdater.setFeedURL({
                provider: 'github',
                owner: resolvedSource.owner,
                repo: resolvedSource.repo,
                private: false
            });
        } else {
            autoUpdater.setFeedURL({
                provider: 'generic',
                url: resolvedSource.url
            });
        }
        autoUpdaterFeedUrl = feedKey;
    }
    return resolvedSource;
}

async function checkForAppUpdates() {
    if (!app.isPackaged) {
        throw new Error('开发环境不支持在线更新，请使用安装版测试');
    }
    ensureUpdateFeedConfigured();
    return autoUpdater.checkForUpdates();
}

async function downloadAppUpdate() {
    if (!app.isPackaged) {
        throw new Error('开发环境不支持在线更新，请使用安装版测试');
    }
    ensureUpdateFeedConfigured();
    if (!updateState.available && !updateState.downloaded) {
        throw new Error('当前没有可下载的更新');
    }
    if (updateState.downloaded) {
        return getUpdateSnapshot();
    }
    await autoUpdater.downloadUpdate();
    return getUpdateSnapshot();
}

function installAppUpdate() {
    if (!app.isPackaged) {
        throw new Error('开发环境不支持在线更新，请使用安装版测试');
    }
    if (!updateState.downloaded) {
        throw new Error('当前没有已下载完成的更新');
    }
    app.isQuitting = true;
    autoUpdater.quitAndInstall(false, true);
}

function sendTemplateLog(sender, level, message, extra = {}) {
    safeSend(sender, 'template:log', {
        level,
        message,
        time: new Date().toLocaleTimeString(),
        ...extra
    });
}

function notifyTemplateStatus(sender, running, extra = {}) {
    safeSend(sender, 'template:status', {
        running,
        ...extra
    });
}

function cleanupTemplateProcess() {
    templateProcess = null;
    templateProcessSender = null;
    templateCancelRequested = false;
}

function stopTemplateProcess() {
    if (!templateProcess) {
        return false;
    }

    templateCancelRequested = true;
    try {
        templateProcess.kill();
        return true;
    } catch {
        return false;
    }
}

// --- Watcher state ---
let activeWatcher = null;
let classifyWatcher = null;
let templateProcess = null;
let templateProcessSender = null;
let templateCancelRequested = false;

// --- Product mapping ---
let PRODUCT_MAP = {
    'ZBDZ': '桌布定制',
    'ZB': '桌布',
    'SBDDZ': '鼠标垫定制',
    'SBD': '鼠标垫',
    'DDDZ': '地垫定制',
    'DD': '地垫',
    'SJTDZ': '浴室三件套定制',
    'SJT': '浴室三件套',
    'KFJDDZ': '咖啡垫定制',
    'KFJD': '咖啡垫',
    'GTDZ': '挂毯定制',
    'GT': '挂毯',
    'TXDZ': 'T恤定制',
    'TX': 'T恤'
};

// Sorted by length descending for priority matching
let PRODUCT_PREFIXES = Object.keys(PRODUCT_MAP).sort((a, b) => b.length - a.length);

function updateProductRules(newRules) {
    PRODUCT_MAP = newRules;
    PRODUCT_PREFIXES = Object.keys(PRODUCT_MAP).sort((a, b) => b.length - a.length);
}

function matchProductPrefix(filename) {
    const nameLower = filename.toLowerCase();
    for (const prefix of PRODUCT_PREFIXES) {
        if (nameLower.startsWith(prefix.toLowerCase())) {
            return { prefix: prefix, productName: PRODUCT_MAP[prefix] };
        }
    }
    return null;
}

function getDateFolders() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return {
        l1: `${year}${month}`,
        l2: `${year}${month}${day}`
    };
}

function findNextIndex(targetDir, productName, dateStr, userName) {
    for (let i = 1; i < 1000; i++) {
        const baseName = `${productName}_${dateStr}_${userName}_${i}`;
        const exists = ['.png', '.jpg', '.jpeg'].some(ext =>
            fs.existsSync(path.join(targetDir, baseName + ext))
        );
        if (!exists) return i;
    }
    return 1;
}

async function classifyFile(filePath, targetBaseDir, userName, sendLog, force = false) {
    const fileName = path.basename(filePath);
    const nameLower = fileName.toLowerCase();

    if (!force) {
        // Check file size - skip if >= 20MB
        try {
            const stat = fs.statSync(filePath);
            const sizeMB = stat.size / 1024 / 1024;
            if (sizeMB >= 20) {
                sendLog('oversize', `${fileName} (${sizeMB.toFixed(1)}MB) 超过20MB，跳过归类`);
                return { skipped: true, reason: 'oversize', fileName };
            }
        } catch (err) {
            return { skipped: true, reason: 'stat_error', fileName };
        }

        // Skip files with 'gai'
        if (nameLower.includes('gai')) {
            console.log('[Classify] Skipping file with gai:', fileName);
            sendLog('warn', `${fileName} 包含 gai，跳过归类`);
            return { skipped: true, reason: 'gai', fileName };
        }
    }

    // Match product prefix
    const match = matchProductPrefix(fileName);
    if (!match) {
        sendLog('error', `${fileName} 无法识别产品前缀`);
        return { skipped: true, reason: 'no_prefix', fileName };
    }

    const { productName } = match;
    const { l1, l2 } = getDateFolders();
    const ext = path.extname(filePath);

    // Create target directory structure
    const targetDir = path.join(targetBaseDir, l1, l2, productName);
    try {
        fs.mkdirSync(targetDir, { recursive: true });
    } catch (err) {
        sendLog('error', `创建目录失败: ${err.message}`);
        return { skipped: true, reason: 'mkdir_error', fileName };
    }

    // Find next available index
    const index = findNextIndex(targetDir, productName, l2, userName);
    const newName = `${productName}_${l2}_${userName}_${index}${ext}`;
    const targetPath = path.join(targetDir, newName);

    try {
        fs.renameSync(filePath, targetPath);
        sendLog('success', `${fileName} → ${newName}`);
        return { success: true, fileName, newName };
    } catch (err) {
        sendLog('error', `${fileName} 归类失败: ${err.message}`);
        return { skipped: true, reason: 'move_error', fileName };
    }
}

function createWindow() {
    // 隐藏默认菜单栏
    Menu.setApplicationMenu(null);

    mainWindow = new BrowserWindow({
        width: 1500,
        height: 900,
        minWidth: 1200,
        minHeight: 700,
        icon: path.join(__dirname, 'logo.png'),
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        },
        frame: false,
        transparent: true,
        backgroundColor: '#00000000',
        show: false
    });

    remoteMain.enable(mainWindow.webContents);

    mainWindow.loadFile('图片工作台.html');

    // 窗口准备好后显示
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    // 关闭窗口时隐藏到托盘
    mainWindow.on('close', (event) => {
        if (!app.isQuitting) {
            event.preventDefault();
            mainWindow.hide();
        }
    });

    // 注册快捷键
    mainWindow.webContents.on('before-input-event', (event, input) => {
        if (input.control && input.key === 'r') {
            mainWindow.reload();
        }
        if (input.control && input.shift && input.key === 'I') {
            mainWindow.webContents.toggleDevTools();
        }
    });
}

function createTray() {
    const iconPath = path.join(__dirname, 'logo.png');
    console.log('[Tray] Creating tray with icon:', iconPath);

    try {
        tray = new Tray(iconPath);
        console.log('[Tray] Tray created successfully');
    } catch (err) {
        console.error('[Tray] Failed to create tray:', err);
        return;
    }

    const contextMenu = Menu.buildFromTemplate([
        {
            label: '显示窗口',
            click: () => {
                console.log('[Tray] Show window clicked');
                mainWindow.show();
            }
        },
        {
            label: '退出',
            click: () => {
                console.log('[Tray] Quit clicked');
                app.isQuitting = true;
                app.quit();
            }
        }
    ]);

    tray.setToolTip('图片工作台');
    tray.setContextMenu(contextMenu);

    // 双击托盘图标显示窗口
    tray.on('double-click', () => {
        console.log('[Tray] Double clicked');
        mainWindow.show();
    });
}

app.whenReady().then(() => {
    createWindow();
    createTray();
    bindAutoUpdaterEvents();
    broadcastUpdateState();

    const updateCfg = loadUpdateConfig();
    if (app.isPackaged && updateCfg.autoCheckOnStartup && resolveUpdateSource(updateCfg)) {
        setTimeout(() => {
            checkForAppUpdates().catch((error) => {
                markUpdateState({
                    checking: false,
                    error: error?.message || '检查更新失败',
                    status: error?.message || '检查更新失败'
                });
            });
        }, 2200);
    }

    // --- IPC handlers ---
    ipcMain.handle('app-update:load-config', () => {
        return getUpdateSnapshot();
    });

    ipcMain.handle('app-update:save-config', (event, cfg) => {
        const saved = saveUpdateConfig(cfg || {});
        const resolvedSavedSource = resolveUpdateSource(saved);
        const savedFeedKey = resolvedSavedSource
            ? (resolvedSavedSource.provider === 'github'
                ? `github:${resolvedSavedSource.owner}/${resolvedSavedSource.repo}`
                : `generic:${resolvedSavedSource.url}`)
            : '';
        if (autoUpdaterFeedUrl && savedFeedKey !== autoUpdaterFeedUrl) {
            autoUpdaterFeedUrl = '';
        }
        broadcastUpdateState();
        return getUpdateSnapshot();
    });

    ipcMain.handle('app-update:check', async () => {
        await checkForAppUpdates();
        return getUpdateSnapshot();
    });

    ipcMain.handle('app-update:download', async () => {
        return downloadAppUpdate();
    });

    ipcMain.handle('app-update:install', () => {
        installAppUpdate();
        return { accepted: true };
    });

    ipcMain.handle('compress:load-config', () => {
        return loadConfig();
    });

    ipcMain.handle('compress:save-config', (event, cfg) => {
        const current = loadConfig();
        const nextCfg = {
            ...current,
            ...(cfg || {})
        };
        saveConfig(nextCfg);
        return nextCfg;
    });

    ipcMain.on('compress:start', (event, { directory, thresholdMB }) => {
        if (activeWatcher) {
            activeWatcher.stop();
            activeWatcher = null;
        }

        const thresholdBytes = thresholdMB * 1024 * 1024;
        const sender = event.sender;

        const sendLog = (type, msg) => {
            safeSend(sender, 'compress:log', { type, msg, time: new Date().toLocaleTimeString() });
        };

        saveConfig({
            ...loadConfig(),
            directory,
            thresholdMB
        });

        activeWatcher = new DirectoryWatcher(directory, async (fp) => {
            await compressImage(fp, thresholdBytes, sendLog);
        });
        activeWatcher.start();

        sendLog('info', `开始监控: ${directory} (阈值: ${thresholdMB} MB)`);
        safeSend(sender, 'compress:status', true);
    });

    ipcMain.on('compress:stop', (event) => {
        if (activeWatcher) {
            activeWatcher.stop();
            activeWatcher = null;
        }
        const sender = event.sender;
        safeSend(sender, 'compress:log', { type: 'info', msg: '已停止监控', time: new Date().toLocaleTimeString() });
        safeSend(sender, 'compress:status', false);
    });

    ipcMain.on('compress:manual', async (event, { directory, thresholdMB }) => {
        const sender = event.sender;
        const thresholdBytes = thresholdMB * 1024 * 1024;

        const sendLog = (type, msg) => {
            safeSend(sender, 'compress:log', { type, msg, time: new Date().toLocaleTimeString() });
        };

        sendLog('info', '开始手动压缩...');

        try {
            const files = fs.readdirSync(directory);
            let count = 0;

            for (const file of files) {
                const filePath = path.join(directory, file);
                const stat = fs.statSync(filePath);
                if (!stat.isFile()) continue;

                await compressImage(filePath, thresholdBytes, sendLog);
                count++;
            }

            sendLog('success', `手动压缩完成，共处理 ${count} 个文件`);
        } catch (err) {
            sendLog('error', `手动压缩失败: ${err.message}`);
        }
    });

    // --- Classify IPC handlers ---
    ipcMain.handle('classify:load-config', () => {
        return loadClassifyConfig();
    });

    ipcMain.handle('classify:save-config', (event, cfg) => {
        const current = loadClassifyConfig();
        const nextCfg = {
            ...current,
            ...(cfg || {})
        };
        saveClassifyConfig(nextCfg);
        return nextCfg;
    });

    ipcMain.handle('slice:load-config', () => {
        return loadSliceConfig();
    });

    ipcMain.handle('slice:save-config', (event, cfg) => {
        saveSliceConfig(cfg);
        return loadSliceConfig();
    });

    ipcMain.handle('slice:select-output-dir', async () => {
        const currentCfg = loadSliceConfig();
        const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory', 'createDirectory'],
            defaultPath: currentCfg.outputDir || getDefaultSliceOutputDir(),
            title: '选择切片输出目录'
        });

        if (result.canceled || result.filePaths.length === 0) {
            return currentCfg;
        }

        const nextCfg = {
            ...currentCfg,
            outputDir: result.filePaths[0]
        };
        saveSliceConfig(nextCfg);
        return nextCfg;
    });

    ipcMain.handle('slice:save-results', async (event, payload) => {
        const { fileName, outputDir, files, taskDir: existingTaskDir, overwrite } = payload || {};
        if (!Array.isArray(files) || files.length === 0) {
            throw new Error('没有可保存的切片结果');
        }

        const cfg = loadSliceConfig();
        const baseOutputDir = outputDir || cfg.outputDir || getDefaultSliceOutputDir();
        fs.mkdirSync(baseOutputDir, { recursive: true });

        const baseName = sanitizePathSegment(path.basename(fileName || 'slice-task', path.extname(fileName || '')));
        let taskDir = '';

        if (overwrite && existingTaskDir) {
            taskDir = existingTaskDir;
            fs.mkdirSync(taskDir, { recursive: true });
            fs.readdirSync(taskDir).forEach((name) => {
                const target = path.join(taskDir, name);
                const stat = fs.statSync(target);
                if (stat.isFile()) {
                    fs.unlinkSync(target);
                }
            });
        } else {
            const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\..+$/, '').replace('T', '_');
            taskDir = path.join(baseOutputDir, `${baseName}_${stamp}`);
            fs.mkdirSync(taskDir, { recursive: true });
        }

        const savedFiles = [];

        files.forEach((item, index) => {
            const ext = path.extname(item.name || '').toLowerCase() || '.png';
            const safeName = sanitizePathSegment(path.basename(item.name || `${baseName}_${index + 1}${ext}`, ext), `${baseName}_${index + 1}`);
            const targetPath = path.join(taskDir, `${safeName}${ext}`);
            const base64Data = String(item.dataUrl || '').replace(/^data:.+;base64,/, '');

            if (!base64Data) {
                throw new Error(`切片 ${index + 1} 缺少图像数据`);
            }

            fs.writeFileSync(targetPath, Buffer.from(base64Data, 'base64'));
            savedFiles.push(targetPath);
        });

        saveSliceConfig({
            ...cfg,
            outputDir: baseOutputDir
        });

        return {
            outputDir: baseOutputDir,
            taskDir,
            files: savedFiles
        };
    });

    ipcMain.handle('slice:overwrite-result-file', async (event, payload) => {
        const filePath = String(payload?.filePath || '').trim();
        const dataUrl = String(payload?.dataUrl || '').trim();
        if (!filePath) {
            throw new Error('缺少切片文件路径');
        }
        if (!dataUrl) {
            throw new Error('缺少切片图像数据');
        }
        const base64Data = dataUrl.replace(/^data:.+;base64,/, '');
        if (!base64Data) {
            throw new Error('切片图像数据无效');
        }
        fs.mkdirSync(path.dirname(filePath), { recursive: true });
        fs.writeFileSync(filePath, Buffer.from(base64Data, 'base64'));
        return { filePath };
    });

    ipcMain.handle('print-ai:load-config', () => {
        return loadPrintAiConfig();
    });

    ipcMain.handle('print-ai:save-config', (event, cfg) => {
        return savePrintAiConfig(cfg || {});
    });

    ipcMain.handle('print-ai:detect-models', async (event, cfg) => {
        const config = normalizePrintAiConfig({
            ...loadPrintAiConfig(),
            ...(cfg || {})
        });
        return detectPrintAiModels(config);
    });

    ipcMain.handle('print-ai:test-model', async (event, cfg) => {
        const config = normalizePrintAiConfig({
            ...loadPrintAiConfig(),
            ...(cfg || {})
        });
        const model = String(cfg?.model || config.extractModel || config.variationModel || '').trim();
        if (!model) {
            throw new Error('请先填写或选择要测试的模型');
        }
        return testPrintAiModelAvailability(config, model);
    });

    ipcMain.handle('print-ai:load-data', () => {
        return loadPrintAiData();
    });

    ipcMain.handle('print-ai:import-images', (event, payload) => {
        const result = importPrintAiImages(payload?.filePaths || []);
        broadcastPrintAiTasks(event.sender, result);
        return result;
    });

    ipcMain.handle('print-ai:start-run', (event, payload) => {
        return startPrintAiRun(payload || {}, event.sender);
    });

    ipcMain.handle('print-ai:delete-tasks', (event, payload) => {
        const ids = new Set(Array.isArray(payload?.taskIds) ? payload.taskIds.map((id) => String(id || '').trim()).filter(Boolean) : []);
        const data = loadPrintAiData();
        data.tasks = ids.size ? data.tasks.filter((task) => !ids.has(task.id)) : data.tasks;
        const saved = savePrintAiData(data);
        broadcastPrintAiTasks(event.sender, saved);
        return saved;
    });

    ipcMain.handle('print-ai:open-output-dir', async () => {
        const cfg = loadPrintAiConfig();
        ensureDir(cfg.outputDir);
        await shell.openPath(cfg.outputDir);
        return { outputDir: cfg.outputDir };
    });

    ipcMain.handle('print-ai:select-output-dir', async () => {
        const cfg = loadPrintAiConfig();
        const result = await dialog.showOpenDialog(mainWindow, {
            title: '选择印花裂变输出目录',
            defaultPath: cfg.outputDir,
            properties: ['openDirectory', 'createDirectory']
        });
        if (result.canceled || !result.filePaths.length) {
            return cfg;
        }
        return savePrintAiConfig({
            ...cfg,
            outputDir: result.filePaths[0]
        });
    });

    ipcMain.handle('print-ai:copy-results-to-output', (event, payload) => {
        const ids = new Set(Array.isArray(payload?.taskIds) ? payload.taskIds.map((id) => String(id || '').trim()).filter(Boolean) : []);
        const cfg = loadPrintAiConfig();
        ensureDir(cfg.outputDir);
        const data = loadPrintAiData();
        const copied = [];
        data.tasks
            .filter((task) => ids.has(task.id))
            .forEach((task) => {
                const files = [
                    task.extractedPath,
                    ...((Array.isArray(task.variants) ? task.variants : []).map((variant) => variant.imagePath))
                ].filter((itemPath) => itemPath && fs.existsSync(itemPath));
                files.forEach((itemPath) => {
                    const targetPath = ensureUniqueFilePath(path.join(cfg.outputDir, path.basename(itemPath)));
                    fs.copyFileSync(itemPath, targetPath);
                    copied.push(targetPath);
                });
            });
        return { copied, outputDir: cfg.outputDir };
    });

    ipcMain.handle('workflow:load-config', () => {
        return loadWorkflowConfig();
    });

    ipcMain.handle('workflow:save-config', (event, cfg) => {
        return saveWorkflowConfig(cfg || {});
    });

    ipcMain.handle('workflow:load-data', () => {
        return loadWorkflowData();
    });

    ipcMain.handle('workflow:validate', (event, payload) => {
        return validateWorkflowGraph(payload?.graph || payload);
    });

    async function selectWorkflowInputs(event, mode = 'files', payload = {}) {
        const workflowPresetId = String(payload?.workflowPresetId || payload?.presetId || '').trim();
        const sourceInputNodeId = String(payload?.sourceInputNodeId || payload?.inputNodeId || '').trim();
        const folderMode = mode === 'folders';
        const result = await dialog.showOpenDialog(mainWindow, {
            title: folderMode ? '选择工作流输入文件夹' : '选择工作流输入图片',
            defaultPath: app.getPath('pictures'),
            properties: folderMode ? ['openDirectory', 'multiSelections'] : ['openFile', 'multiSelections'],
            filters: folderMode ? undefined : [
                { name: '图片', extensions: ['jpg', 'jpeg', 'png', 'webp', 'bmp', 'avif'] },
                { name: '所有文件', extensions: ['*'] }
            ]
        });
        if (result.canceled || !result.filePaths.length) {
            return { canceled: true, ...loadWorkflowData() };
        }
        const imported = importWorkflowInputs(result.filePaths, workflowPresetId, sourceInputNodeId);
        broadcastWorkflowData(event.sender, imported);
        return { canceled: false, ...imported };
    }

    ipcMain.handle('workflow:select-inputs', async (event, payload) => {
        return selectWorkflowInputs(event, 'files', payload || {});
    });

    ipcMain.handle('workflow:select-input-files', async (event, payload) => {
        return selectWorkflowInputs(event, 'files', payload || {});
    });

    ipcMain.handle('workflow:select-input-folders', async (event, payload) => {
        return selectWorkflowInputs(event, 'folders', payload || {});
    });

    ipcMain.handle('workflow:import-inputs', (event, payload) => {
        const result = importWorkflowInputs(
            payload?.filePaths || [],
            String(payload?.workflowPresetId || payload?.presetId || '').trim(),
            String(payload?.sourceInputNodeId || payload?.inputNodeId || '').trim()
        );
        broadcastWorkflowData(event.sender, result);
        return result;
    });

    ipcMain.handle('workflow:select-export-dir', async () => {
        const result = await dialog.showOpenDialog(mainWindow, {
            title: '选择工作流导出目录',
            properties: ['openDirectory', 'createDirectory']
        });
        if (result.canceled || !result.filePaths.length) return { canceled: true };
        return { canceled: false, dir: result.filePaths[0] };
    });

    ipcMain.handle('workflow:start', (event, payload) => {
        return startWorkflowRun(payload || {}, event.sender);
    });

    ipcMain.handle('workflow:retry-node', (event, payload) => {
        return retryWorkflowNode(payload || {}, event.sender);
    });

    ipcMain.handle('workflow:delete-runs', (event, payload) => {
        const ids = new Set(Array.isArray(payload?.runIds) ? payload.runIds.map((id) => String(id || '').trim()).filter(Boolean) : []);
        const data = loadWorkflowData();
        data.runs = ids.size ? data.runs.filter((run) => !ids.has(run.id)) : data.runs;
        const saved = saveWorkflowData(data);
        broadcastWorkflowData(event.sender, saved);
        return saved;
    });

    ipcMain.handle('workflow:pause', () => {
        return { paused: true };
    });

    ipcMain.handle('template:load-config', () => {
        return loadTemplateConfig();
    });

    ipcMain.handle('template:save-config', (event, cfg) => {
        return saveTemplateConfig({
            ...loadTemplateConfig(),
            ...(cfg || {})
        });
    });

    ipcMain.handle('template:list-templates', () => {
        const cfg = loadTemplateConfig();
        return {
            templateRootDir: getTemplateRootDir(cfg),
            watermarkDir: getWatermarkDir(cfg),
            templates: listTemplateFolders()
        };
    });

    ipcMain.handle('template:save-scene-config', (event, payload) => {
        const relativePath = String(payload && payload.relativePath ? payload.relativePath : '').trim();
        if (!relativePath) {
            throw new Error('缺少模板场景路径');
        }

        const sceneDir = resolveTemplateSceneDir(relativePath);
        const configPath = path.join(sceneDir, 'config.json');
        if (!fs.existsSync(configPath)) {
            throw new Error('当前模板场景缺少 config.json');
        }

        const currentConfig = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
        const nextConfig = {
            ...currentConfig
        };
        if (payload && Object.prototype.hasOwnProperty.call(payload, 'placement')) {
            nextConfig.placement = getTemplatePlacement({ placement: payload.placement || {} });
        }
        if (payload && Object.prototype.hasOwnProperty.call(payload, 'effects')) {
            nextConfig.effects = getTemplateEffects({ effects: payload.effects || {} });
        }
        if (payload && Object.prototype.hasOwnProperty.call(payload, 'points')) {
            const nextPoints = getTemplatePoints({ points: payload.points || [] });
            if (nextPoints.length !== 4) {
                throw new Error('模板坐标必须为 4 个点');
            }
            nextConfig.points = nextPoints;
        }
        fs.writeFileSync(configPath, JSON.stringify(nextConfig, null, 2), 'utf-8');

        const allGroups = listTemplateFolders();
        for (const group of allGroups) {
            const scene = (group.scenes || []).find((item) => item.relativePath === toTemplateRelativePath(sceneDir));
            if (scene) {
                return {
                    group,
                    scene
                };
            }
        }

        return {
            scene: describeTemplateScene(sceneDir, path.basename(path.dirname(sceneDir)), path.basename(sceneDir))
        };
    });

    ipcMain.handle('template:create-group', async (event, payload) => {
        const groupName = sanitizeTemplateSegment(payload && payload.groupName, '新产品');
        const templateRootDir = getTemplateRootDir();
        ensureDir(templateRootDir);
        ensureDir(path.join(templateRootDir, groupName));
        const currentCfg = loadTemplateConfig();
        const nextOrder = (Array.isArray(currentCfg.templateOrder) ? currentCfg.templateOrder : [])
            .filter((name) => name !== groupName)
            .concat(groupName);
        saveTemplateConfig({
            ...currentCfg,
            templateOrder: nextOrder
        });
        const allGroups = listTemplateFolders();
        return {
            templateRootDir,
            templates: allGroups,
            group: allGroups.find((item) => item.name === groupName) || null
        };
    });

    ipcMain.handle('template:delete-group', async (event, payload) => {
        const groupName = sanitizeTemplateSegment(payload && payload.groupName, '');
        if (!groupName) {
            throw new Error('缺少模板名称');
        }
        const templateRootDir = getTemplateRootDir();
        const groupDir = path.join(templateRootDir, groupName);
        if (!fs.existsSync(groupDir)) {
            throw new Error('当前模板不存在');
        }
        fs.rmSync(groupDir, { recursive: true, force: true });
        const currentCfg = loadTemplateConfig();
        saveTemplateConfig({
            ...currentCfg,
            templateOrder: (Array.isArray(currentCfg.templateOrder) ? currentCfg.templateOrder : [])
                .filter((name) => name !== groupName)
        });
        const allGroups = listTemplateFolders();
        return {
            templateRootDir,
            templates: allGroups,
            deletedGroupName: groupName
        };
    });

    ipcMain.handle('template:reorder-groups', async (event, payload) => {
        const groupName = sanitizeTemplateSegment(payload && payload.groupName, '');
        const direction = String(payload && payload.direction ? payload.direction : '').trim().toLowerCase();
        const { templates, config } = moveTemplateGroupOrder(groupName, direction);
        return {
            templateRootDir: getTemplateRootDir(config),
            watermarkDir: getWatermarkDir(config),
            templates
        };
    });

    ipcMain.handle('template:create-scene', async (event, payload) => {
        const groupName = sanitizeTemplateSegment(payload && payload.groupName, '新模板组');
        const sceneName = sanitizeTemplateSegment(payload && payload.sceneName, '新场景');
        const files = payload && payload.files && typeof payload.files === 'object' ? payload.files : {};
        const basePath = String(files.basePath || '').trim();
        const maskPath = String(files.maskPath || '').trim();
        const texturePath = String(files.texturePath || '').trim();
        const highlightPath = String(files.highlightPath || '').trim();
        if (!basePath || !fs.existsSync(basePath)) {
            throw new Error('请先上传 base.png 底图');
        }
        if (!maskPath || !fs.existsSync(maskPath)) {
            throw new Error('请先上传 mask.png 蒙版');
        }
        const hasTextureStack = texturePath && highlightPath && fs.existsSync(texturePath) && fs.existsSync(highlightPath);
        if (!hasTextureStack) {
            throw new Error('请同时上传 texture.png 和 highlight.png');
        }

        const templateRootDir = getTemplateRootDir();
        ensureDir(templateRootDir);
        const groupDir = ensureDir(path.join(templateRootDir, groupName));
        const sceneDir = path.join(groupDir, sceneName);
        ensureDir(sceneDir);

        const copyAsset = (sourcePath, targetName) => {
            if (!sourcePath) return;
            fs.copyFileSync(sourcePath, path.join(sceneDir, targetName));
        };

        copyAsset(basePath, 'base.png');
        copyAsset(maskPath, 'mask.png');
        if (hasTextureStack) {
            copyAsset(texturePath, 'texture.png');
            copyAsset(highlightPath, 'highlight.png');
            if (fs.existsSync(path.join(sceneDir, 'shadow.png'))) {
                fs.unlinkSync(path.join(sceneDir, 'shadow.png'));
            }
        }

        const configPath = path.join(sceneDir, 'config.json');
        const defaultPoints = await getDefaultTemplateScenePoints(path.join(sceneDir, 'base.png'), path.join(sceneDir, 'mask.png'));
        const nextConfig = {
            points: defaultPoints,
            placement: getTemplatePlacement({}),
            effects: getTemplateEffects({})
        };
        fs.writeFileSync(configPath, JSON.stringify(nextConfig, null, 2), 'utf-8');

        const allGroups = listTemplateFolders();
        const createdGroup = allGroups.find((item) => item.name === groupName) || null;
        const createdScene = createdGroup
            ? (createdGroup.scenes || []).find((item) => item.relativePath === toTemplateRelativePath(sceneDir)) || null
            : null;

        return {
            group: createdGroup,
            scene: createdScene,
            relativePath: toTemplateRelativePath(sceneDir),
            templateRootDir,
            templates: allGroups
        };
    });

    ipcMain.handle('template:update-scene-assets', async (event, payload) => {
        const relativePath = String(payload && payload.relativePath ? payload.relativePath : '').trim();
        if (!relativePath) {
            throw new Error('缺少模板场景路径');
        }
        const files = payload && payload.files && typeof payload.files === 'object' ? payload.files : {};
        const sceneDir = resolveTemplateSceneDir(relativePath);
        const copyAsset = (sourcePath, targetName) => {
            const resolved = String(sourcePath || '').trim();
            if (!resolved) return false;
            if (!fs.existsSync(resolved)) {
                throw new Error(`素材文件不存在：${targetName}`);
            }
            fs.copyFileSync(resolved, path.join(sceneDir, targetName));
            return true;
        };

        const changedBase = copyAsset(files.basePath, 'base.png');
        const changedMask = copyAsset(files.maskPath, 'mask.png');
        const changedTexture = copyAsset(files.texturePath, 'texture.png');
        const changedHighlight = copyAsset(files.highlightPath, 'highlight.png');

        if ((changedTexture || changedHighlight) && fs.existsSync(path.join(sceneDir, 'shadow.png'))) {
            fs.unlinkSync(path.join(sceneDir, 'shadow.png'));
        }

        const configPath = path.join(sceneDir, 'config.json');
        if ((changedBase || changedMask) && fs.existsSync(configPath)) {
            const currentConfig = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
            if (!Array.isArray(getTemplatePoints(currentConfig)) || getTemplatePoints(currentConfig).length !== 4) {
                currentConfig.points = await getDefaultTemplateScenePoints(path.join(sceneDir, 'base.png'), path.join(sceneDir, 'mask.png'));
                fs.writeFileSync(configPath, JSON.stringify(currentConfig, null, 2), 'utf-8');
            }
        }

        const allGroups = listTemplateFolders();
        for (const group of allGroups) {
            const scene = (group.scenes || []).find((item) => item.relativePath === toTemplateRelativePath(sceneDir));
            if (scene) {
                return { group, scene, templates: allGroups };
            }
        }
        throw new Error('模板场景刷新失败');
    });

    ipcMain.handle('template:delete-scene', async (event, payload) => {
        const relativePath = String(payload && payload.relativePath ? payload.relativePath : '').trim();
        if (!relativePath) {
            throw new Error('缺少模板场景路径');
        }
        const sceneDir = resolveTemplateSceneDir(relativePath);
        if (!fs.existsSync(sceneDir)) {
            throw new Error('当前模板场景不存在');
        }
        fs.rmSync(sceneDir, { recursive: true, force: true });
        const allGroups = listTemplateFolders();
        return {
            templateRootDir: getTemplateRootDir(),
            templates: allGroups,
            deletedRelativePath: relativePath
        };
    });

    ipcMain.handle('template:open-templates-dir', async () => {
        const templateRootDir = getTemplateRootDir();
        ensureDir(templateRootDir);
        await shell.openPath(templateRootDir);
        return { templateRootDir };
    });

    ipcMain.handle('template:select-root-dir', async () => {
        const currentCfg = loadTemplateConfig();
        const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory', 'createDirectory'],
            defaultPath: currentCfg.templateRootDir || getDefaultTemplateRootDir(),
            title: '选择模板存放目录'
        });
        if (result.canceled || result.filePaths.length === 0) {
            return {
                ...currentCfg,
                templateRootDir: getTemplateRootDir(currentCfg),
                templates: listTemplateFolders()
            };
        }
        const nextCfg = saveTemplateConfig({
            ...currentCfg,
            templateRootDir: result.filePaths[0]
        });
        return {
            ...nextCfg,
            templateRootDir: getTemplateRootDir(nextCfg),
            templates: listTemplateFolders()
        };
    });

    ipcMain.handle('template:select-watermark-dir', async () => {
        const currentCfg = loadTemplateConfig();
        const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory', 'createDirectory'],
            defaultPath: currentCfg.watermarkDir || getDefaultWatermarkDir(),
            title: '选择水印存放目录'
        });
        if (result.canceled || result.filePaths.length === 0) {
            const nextCfg = {
                ...currentCfg,
                watermarkDir: getWatermarkDir(currentCfg)
            };
            return {
                ...nextCfg,
                watermarks: loadWatermarkPresets()
            };
        }
        const nextCfg = saveTemplateConfig({
            ...currentCfg,
            watermarkDir: result.filePaths[0]
        });
        return {
            ...nextCfg,
            watermarks: loadWatermarkPresets()
        };
    });

    ipcMain.handle('template:list-watermarks', () => {
        return loadWatermarkPresets();
    });

    ipcMain.handle('template:save-watermarks', (event, presets) => {
        saveWatermarkPresets(presets);
        return loadWatermarkPresets();
    });

    ipcMain.handle('template:list-parameter-presets', () => {
        return loadTemplateParameterPresets();
    });

    ipcMain.handle('template:save-parameter-presets', (event, presets) => {
        saveTemplateParameterPresets(presets);
        return loadTemplateParameterPresets();
    });

    ipcMain.handle('template:select-output-dir', async () => {
        const currentCfg = loadTemplateConfig();
        const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory', 'createDirectory'],
            defaultPath: currentCfg.outputDir || getDefaultTemplateOutputDir(),
            title: '选择智能模板导出目录'
        });

        if (result.canceled || result.filePaths.length === 0) {
            return currentCfg;
        }

        const nextCfg = {
            ...currentCfg,
            outputDir: result.filePaths[0]
        };
        saveTemplateConfig(nextCfg);
        return nextCfg;
    });

    ipcMain.handle('template:open-output-dir', async (event, outputDir) => {
        const targetDir = outputDir || loadTemplateConfig().outputDir || getDefaultTemplateOutputDir();
        ensureDir(targetDir);
        await shell.openPath(targetDir);
        return { outputDir: targetDir };
    });

    ipcMain.handle('template:reveal-output-file', async (event, filePath) => {
        if (!filePath || !fs.existsSync(filePath)) {
            throw new Error('结果文件不存在');
        }
        shell.showItemInFolder(filePath);
        return { filePath };
    });

    ipcMain.handle('template:open-result-file', async (event, filePath) => {
        if (!filePath || !fs.existsSync(filePath)) {
            throw new Error('结果文件不存在');
        }
        await shell.openPath(filePath);
        return { filePath };
    });

    ipcMain.handle('template:render-preview', async (event, payload) => {
        const {
            designPath = '',
            designName = '',
            sceneRelativePath = '',
            sceneName = '',
            placement = {},
            effects = null,
            previewKey = '',
            watermarkPresetId = '',
            watermarkPreset = null
        } = payload || {};

        if (!designPath || !fs.existsSync(designPath)) {
            throw new Error('预览设计图不存在');
        }

        const sceneDir = resolveTemplateSceneDir(sceneRelativePath);
        if (!fs.existsSync(sceneDir)) {
            throw new Error('预览模板场景不存在');
        }

        const savedPresets = loadWatermarkPresets();
        const selectedPreset = watermarkPreset
            || savedPresets.find((item) => item.id === watermarkPresetId)
            || null;

        const previewRoot = path.join(app.getPath('temp'), 'ImageFlow-template-preview');
        ensureDir(previewRoot);

        const message = await runTemplatePreviewJob({
            mode: 'preview',
            outputDir: previewRoot,
            templateRootDir: getTemplateRootDir(),
            designs: [
                {
                    name: designName || path.basename(designPath),
                    path: designPath
                }
            ],
            preview: {
                relativePath: toTemplateRelativePath(sceneDir),
                name: sceneName || path.basename(sceneDir),
                placement: getTemplatePlacement({ placement }),
                effects: getTemplateEffects({ effects: effects || {} }),
                previewKey: String(previewKey || '').trim()
            },
            watermarkPreset: selectedPreset
        });

        return {
            previewPath: message.outputPath || '',
            sceneRelativePath: toTemplateRelativePath(sceneDir)
        };
    });

    ipcMain.handle('template:start-generation', async (event, payload) => {
        if (templateProcess) {
            throw new Error('智能模板任务正在运行');
        }

        const sender = event.sender;
        const {
            designs = [],
            selectedTemplates = [],
            outputDir,
            watermarkPresetId = '',
            watermarkPreset = null,
            parameterPresetId = '',
            effectPreset = null
        } = payload || {};

        if (!Array.isArray(designs) || designs.length === 0) {
            throw new Error('请先导入设计图');
        }

        const templateRendererScript = getTemplateRendererScriptPath();
        if (!fs.existsSync(templateRendererScript)) {
            throw new Error('缺少 template_renderer.py');
        }

        const templateGroups = listTemplateFolders();
        const activeTemplateGroups = Array.isArray(selectedTemplates)
            ? selectedTemplates
                .map((item) => String(item || '').trim())
                .filter(Boolean)
            : [];

        if (activeTemplateGroups.length === 0) {
            throw new Error('请先勾选至少一个模板套组');
        }

        const resolvedTemplateGroups = activeTemplateGroups
            .map((name) => templateGroups.find((item) => item.name === name))
            .filter(Boolean)
            .map((group) => ({
                name: group.name,
                scenes: (group.scenes || []).filter((scene) => scene.valid).map((scene) => ({
                    name: scene.name,
                    relativePath: scene.relativePath
                }))
            }))
            .filter((group) => group.scenes.length > 0);

        if (resolvedTemplateGroups.length === 0) {
            throw new Error('当前所选模板组不存在');
        }

        const cfg = loadTemplateConfig();
        const resolvedOutputDir = outputDir || cfg.outputDir || getDefaultTemplateOutputDir();
        ensureDir(resolvedOutputDir);

        saveTemplateConfig({
            ...cfg,
            outputDir: resolvedOutputDir,
            selectedTemplates: resolvedTemplateGroups.map((item) => item.name),
            watermarkPresetId,
            parameterPresetId
        });

        const pythonRuntime = getPythonRuntime(templateRendererScript);
        if (!pythonRuntime) {
            throw new Error('未检测到可用 Python 运行环境，请安装 Python 或 py 启动器');
        }

        const savedPresets = loadWatermarkPresets();
        const selectedPreset = watermarkPreset
            || savedPresets.find((item) => item.id === watermarkPresetId)
            || null;
        const parameterPresets = loadTemplateParameterPresets();
        const selectedParameterPreset = effectPreset
            || (parameterPresetId ? (parameterPresets.find((item) => item.id === parameterPresetId) || {}).effects : null)
            || null;

        const jobPayload = {
            outputDir: resolvedOutputDir,
            templateRootDir: getTemplateRootDir(),
            templateGroups: resolvedTemplateGroups,
            designs: designs.map((item) => ({
                name: item.name,
                path: item.path
            })),
            watermarkPreset: selectedPreset,
            effectPreset: selectedParameterPreset
        };

        templateCancelRequested = false;
        templateProcessSender = sender;
        notifyTemplateStatus(sender, true, { outputDir: resolvedOutputDir });
        sendTemplateLog(
            sender,
            'info',
            `开始生成，共 ${designs.length} 张设计图，${resolvedTemplateGroups.length} 个模板组，${resolvedTemplateGroups.reduce((sum, group) => sum + group.scenes.length, 0)} 个场景`
        );

        const child = spawn(pythonRuntime.command, pythonRuntime.scriptArgs, {
            cwd: path.dirname(templateRendererScript),
            windowsHide: true,
            stdio: ['pipe', 'pipe', 'pipe'],
            env: {
                ...process.env,
                PYTHONUTF8: '1'
            }
        });

        templateProcess = child;

        let stdoutBuffer = '';
        child.stdout.on('data', (chunk) => {
            stdoutBuffer += chunk.toString('utf-8');
            const lines = stdoutBuffer.split(/\r?\n/);
            stdoutBuffer = lines.pop() || '';

            lines.forEach((line) => {
                const text = line.trim();
                if (!text) return;
                try {
                    const message = JSON.parse(text);
                    if (message.type === 'log') {
                        sendTemplateLog(sender, message.level || 'info', message.message || '', message);
                    } else if (message.type === 'done') {
                        safeSend(sender, 'template:done', message);
                    } else if (message.type === 'progress') {
                        safeSend(sender, 'template:progress', message);
                    }
                } catch {
                    sendTemplateLog(sender, 'info', text);
                }
            });
        });

        child.stderr.on('data', (chunk) => {
            const lines = chunk.toString('utf-8').split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
            lines.forEach((line) => sendTemplateLog(sender, 'error', line));
        });

        child.on('error', (error) => {
            sendTemplateLog(sender, 'error', `智能模板启动失败: ${error.message}`);
            notifyTemplateStatus(sender, false, { outputDir: resolvedOutputDir });
            cleanupTemplateProcess();
        });

        child.on('close', (code, signal) => {
            if (stdoutBuffer.trim()) {
                try {
                    const message = JSON.parse(stdoutBuffer.trim());
                    if (message.type === 'log') {
                        sendTemplateLog(sender, message.level || 'info', message.message || '', message);
                    }
                } catch {
                    sendTemplateLog(sender, 'info', stdoutBuffer.trim());
                }
            }

            if (templateCancelRequested) {
                sendTemplateLog(sender, 'warn', '已停止智能模板任务');
            } else if (code !== 0) {
                sendTemplateLog(sender, 'error', `智能模板任务异常结束 (code=${code ?? 'null'}, signal=${signal ?? 'null'})`);
            } else {
                sendTemplateLog(sender, 'success', '智能模板任务已完成');
            }

            notifyTemplateStatus(sender, false, {
                outputDir: resolvedOutputDir,
                canceled: templateCancelRequested,
                exitCode: code,
                signal
            });
            cleanupTemplateProcess();
        });

        child.stdin.end(JSON.stringify(jobPayload), 'utf-8');

        return {
            started: true,
            outputDir: resolvedOutputDir
        };
    });

    ipcMain.handle('template:cancel-generation', () => {
        const stopped = stopTemplateProcess();
        return { stopped };
    });

    ipcMain.handle('product-publish:load-config', () => {
        return loadProductPublishConfig();
    });

    ipcMain.handle('product-publish:save-config', (event, cfg) => {
        return saveProductPublishConfig(cfg || {});
    });

    ipcMain.handle('product-publish:load-data', () => {
        return loadProductPublishData();
    });

    ipcMain.handle('product-publish:save-data', (event, data) => {
        return saveProductPublishData(data || {});
    });

    ipcMain.handle('product-publish:import-template-task', (event, payload) => {
        return importProductPublishRecordFromTemplateTask(payload || {});
    });

    ipcMain.handle('product-publish:select-import-folders', async () => {
        const result = await dialog.showOpenDialog(mainWindow, {
            title: '选择产品图片文件夹',
            defaultPath: app.getPath('pictures'),
            properties: ['openDirectory', 'multiSelections', 'createDirectory']
        });
        if (result.canceled || !Array.isArray(result.filePaths) || !result.filePaths.length) {
            const data = loadProductPublishData();
            return { canceled: true, ...data };
        }
        const currentData = loadProductPublishData();
        const now = new Date().toISOString();
        result.filePaths.forEach((folderPath) => {
            const nextRecord = buildProductPublishRecordFromFolder(folderPath);
            const existingIndex = currentData.records.findIndex((item) => item.sourceTaskKey === nextRecord.sourceTaskKey);
            if (existingIndex >= 0) {
                const existing = currentData.records[existingIndex];
                currentData.records[existingIndex] = normalizeProductPublishRecord({
                    ...existing,
                    groupName: nextRecord.groupName,
                    outputDir: nextRecord.outputDir,
                    sceneNames: nextRecord.sceneNames,
                    images: nextRecord.images,
                    exportStatus: 'idle',
                    exportedAt: '',
                    exportFilePath: '',
                    exportFileName: '',
                    updatedAt: now
                }, existingIndex);
            } else {
                currentData.records.unshift(nextRecord);
            }
        });
        const saved = saveProductPublishData(currentData);
        return { canceled: false, ...saved };
    });

    ipcMain.handle('product-publish:generate-title', async (event, payload) => {
        const record = normalizeProductPublishRecord(payload?.record || {});
        const cfg = loadProductPublishConfig();
        return generateProductPublishTitle(record, cfg);
    });

    ipcMain.handle('product-publish:detect-models', async () => {
        const cfg = loadProductPublishConfig();
        return detectProductPublishModels(cfg);
    });

    ipcMain.handle('product-publish:test-model', async () => {
        const cfg = loadProductPublishConfig();
        return testProductPublishModel(cfg);
    });

    ipcMain.handle('product-publish:select-output-dir', async () => {
        const currentCfg = loadProductPublishConfig();
        const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory', 'createDirectory'],
            defaultPath: currentCfg.exportTemplateDefaults?.outputDir || getDefaultProductPublishOutputDir(),
            title: '选择产品发布默认导出目录'
        });
        if (result.canceled || !result.filePaths.length) {
            return currentCfg;
        }
        const nextCfg = saveProductPublishConfig({
            ...currentCfg,
            exportTemplateDefaults: {
                ...(currentCfg.exportTemplateDefaults || {}),
                outputDir: result.filePaths[0]
            }
        });
        return nextCfg;
    });

    ipcMain.handle('product-publish:open-output-dir', async (event, dirPath) => {
        const targetDir = normalizeDirectoryPath(dirPath, loadProductPublishConfig().exportTemplateDefaults?.outputDir || getDefaultProductPublishOutputDir());
        ensureDir(targetDir);
        await shell.openPath(targetDir);
        return { outputDir: targetDir };
    });

    ipcMain.handle('product-publish:prepare-temu-sheet', async (event, payload) => {
        const records = Array.isArray(payload?.records)
            ? payload.records.map((item, index) => normalizeProductPublishRecord(item, index))
            : [];
        const exportConfig = {
            ...createDefaultProductPublishConfig().exportTemplateDefaults,
            ...(payload?.bulk || {})
        };
        if (!records.length) {
            throw new Error('当前没有可导出的产品记录');
        }
        let templatePath = resolveProductPublishTemuTemplatePath();
        if (!templatePath) {
            const templateResult = await dialog.showOpenDialog(mainWindow, {
                title: '选择妙手 Temu 导入模板',
                defaultPath: app.getPath('downloads'),
                properties: ['openFile'],
                filters: [
                    { name: 'Excel 模板', extensions: ['xlsx'] }
                ]
            });
            if (templateResult.canceled || !Array.isArray(templateResult.filePaths) || !templateResult.filePaths[0]) {
                return { canceled: true };
            }
            templatePath = templateResult.filePaths[0];
        }
        const sender = event.sender;
        const totalImages = records.reduce((sum, record) => sum + ((Array.isArray(record.images) ? record.images.length : 0)), 0);
        let uploadedCount = 0;
        safeSend(sender, 'product-publish:export-progress', {
            stage: 'running',
            total: totalImages || records.length,
            current: 0,
            message: totalImages ? '开始上传图片并生成 URL...' : '开始整理导出数据...'
        });
        if (isProductPublishOssConfigured(exportConfig)) {
            for (const record of records) {
                const uploadedUrls = await uploadProductPublishRecordImagesToOss(record, exportConfig, (progress) => {
                    if (progress?.phase === 'uploading') {
                        safeSend(sender, 'product-publish:export-progress', {
                            stage: 'running',
                            total: totalImages || records.length,
                            current: Math.min(uploadedCount + 1, totalImages || records.length),
                            message: `正在转换 URL：${progress.recordName} / ${progress.imageName}`
                        });
                    }
                    if (progress?.phase === 'uploaded') {
                        uploadedCount += 1;
                        safeSend(sender, 'product-publish:export-progress', {
                            stage: 'running',
                            total: totalImages || records.length,
                            current: uploadedCount,
                            message: `已生成 URL：${progress.recordName} / ${progress.imageName}`
                        });
                    }
                });
                if (uploadedUrls.length) {
                    record.urls = uploadedUrls;
                    record.urlStatus = 'ready';
                    record.previewImageUrl = uploadedUrls[0];
                }
            }
        }
        safeSend(sender, 'product-publish:export-progress', {
            stage: 'running',
            total: totalImages || records.length || 1,
            current: totalImages || records.length || 1,
            message: '正在写入导出模板...'
        });
        const workbook = buildProductPublishTemuWorkbook(records, templatePath);
        const tempFilePath = path.join(app.getPath('temp'), `imageflow-product-publish-${Date.now()}-${process.pid}.xlsx`);
        XLSX.writeFile(workbook, tempFilePath, { compression: true });
        const defaultFileName = buildProductPublishExportFileName(records.length);
        safeSend(sender, 'product-publish:export-progress', {
            stage: 'done',
            total: totalImages || records.length || 1,
            current: totalImages || records.length || 1,
            message: '导出文件已准备完成，可以保存。'
        });
        return { canceled: false, tempFilePath, defaultFileName, templatePath };
    });

    ipcMain.handle('product-publish:finalize-temu-sheet', async (event, payload) => {
        const tempFilePath = String(payload?.tempFilePath || '').trim();
        if (!tempFilePath || !fs.existsSync(tempFilePath)) {
            throw new Error('临时导出文件不存在，请重新导出。');
        }
        const currentCfg = loadProductPublishConfig();
        const configuredDir = normalizeDirectoryPath(
            payload?.outputDir || currentCfg.exportTemplateDefaults?.outputDir,
            getDefaultProductPublishOutputDir()
        );
        ensureDir(configuredDir);
        const defaultFileName = String(payload?.defaultFileName || '').trim() || buildProductPublishExportFileName(Number(payload?.recordCount) || 0);
        let targetPath = '';
        if (payload?.mode === 'saveAs') {
            const saveResult = await dialog.showSaveDialog(mainWindow, {
                title: '保存妙手 Temu 表格',
                defaultPath: path.join(configuredDir, defaultFileName),
                filters: [
                    { name: 'Excel 表格', extensions: ['xlsx'] }
                ]
            });
            if (saveResult.canceled || !saveResult.filePath) {
                return { canceled: true };
            }
            targetPath = saveResult.filePath;
        } else {
            targetPath = ensureUniqueFilePath(path.join(configuredDir, defaultFileName));
        }
        fs.copyFileSync(tempFilePath, targetPath);
        try {
            fs.unlinkSync(tempFilePath);
        } catch {}
        return { canceled: false, filePath: targetPath };
    });

    ipcMain.handle('product-publish:export-temu-sheet', async (event, payload) => {
        const records = Array.isArray(payload?.records)
            ? payload.records.map((item, index) => normalizeProductPublishRecord(item, index))
            : [];
        const exportConfig = {
            ...createDefaultProductPublishConfig().exportTemplateDefaults,
            ...(payload?.bulk || {})
        };
        if (!records.length) {
            throw new Error('当前没有可导出的产品记录');
        }
        let templatePath = resolveProductPublishTemuTemplatePath();
        if (!templatePath) {
            const templateResult = await dialog.showOpenDialog(mainWindow, {
                title: '选择妙手 Temu 导入模板',
                defaultPath: app.getPath('downloads'),
                properties: ['openFile'],
                filters: [
                    { name: 'Excel 模板', extensions: ['xlsx'] }
                ]
            });
            if (templateResult.canceled || !Array.isArray(templateResult.filePaths) || !templateResult.filePaths[0]) {
                return { canceled: true };
            }
            templatePath = templateResult.filePaths[0];
        }
        if (isProductPublishOssConfigured(exportConfig)) {
            for (const record of records) {
                const uploadedUrls = await uploadProductPublishRecordImagesToOss(record, exportConfig);
                if (uploadedUrls.length) {
                    record.urls = uploadedUrls;
                    record.urlStatus = 'ready';
                    record.previewImageUrl = uploadedUrls[0];
                }
            }
        }
        const workbook = buildProductPublishTemuWorkbook(records, templatePath);
        const saveResult = await dialog.showSaveDialog(mainWindow, {
            title: '保存妙手 Temu 表格',
            defaultPath: path.join(
                normalizeDirectoryPath(exportConfig.outputDir, getDefaultProductPublishOutputDir()),
                buildProductPublishExportFileName(records.length)
            ),
            filters: [
                { name: 'Excel 表格', extensions: ['xlsx'] }
            ]
        });
        if (saveResult.canceled || !saveResult.filePath) {
            return { canceled: true };
        }
        XLSX.writeFile(workbook, saveResult.filePath, { compression: true });
        return { canceled: false, filePath: saveResult.filePath, templatePath };
    });

    ipcMain.on('classify:start-auto', (event, { sourceDir, targetDir, userName }) => {
        if (classifyWatcher) {
            classifyWatcher.stop();
            classifyWatcher = null;
        }

        const sender = event.sender;
        const sendLog = (type, msg) => {
            safeSend(sender, 'classify:log', { type, msg, time: new Date().toLocaleTimeString() });
        };

        saveClassifyConfig({
            ...loadClassifyConfig(),
            sourceDir,
            targetDir,
            userName
        });

        classifyWatcher = new DirectoryWatcher(sourceDir, async (fp) => {
            await classifyFile(fp, targetDir, userName, sendLog);
        });
        classifyWatcher.start();

        sendLog('info', `开始自动归类: ${sourceDir}`);
        safeSend(sender, 'classify:status', true);
    });

    ipcMain.on('classify:stop-auto', (event) => {
        if (classifyWatcher) {
            classifyWatcher.stop();
            classifyWatcher = null;
        }
        const sender = event.sender;
        safeSend(sender, 'classify:log', { type: 'info', msg: '已停止自动归类', time: new Date().toLocaleTimeString() });
        safeSend(sender, 'classify:status', false);
    });

    ipcMain.on('classify:manual', async (event, { sourceDir, targetDir, userName }) => {
        console.log('[Classify Manual] Received request:', { sourceDir, targetDir, userName });

        const sender = event.sender;
        const sendLog = (type, msg) => {
            console.log('[Classify Manual] Sending log:', type, msg);
            safeSend(sender, 'classify:log', { type, msg, time: new Date().toLocaleTimeString() });
        };

        sendLog('info', '开始手动归类...');

        try {
            const files = fs.readdirSync(sourceDir);
            console.log('[Classify Manual] Found files:', files.length);
            const needFix = [];
            let successCount = 0;
            let skipCount = 0;

            for (const file of files) {
                const filePath = path.join(sourceDir, file);
                const stat = fs.statSync(filePath);
                if (!stat.isFile()) continue;

                console.log('[Classify Manual] Processing:', file);
                const result = await classifyFile(filePath, targetDir, userName, sendLog);
                console.log('[Classify Manual] Result:', result);

                if (result.success) {
                    successCount++;
                } else if (result.skipped) {
                    skipCount++;
                    if (result.reason === 'gai') {
                        needFix.push(result.fileName);
                    }
                }
            }

            sendLog('success', `手动归类完成: ${successCount} 个成功, ${skipCount} 个跳过`);
            if (needFix.length > 0) {
                safeSend(sender, 'classify:need-fix-list', needFix);
            }
        } catch (err) {
            console.error('[Classify Manual] Error:', err);
            sendLog('error', `手动归类失败: ${err.message}`);
        }
    });

    ipcMain.on('scan-fix-items', (event, sourceDir) => {
        const sender = event.sender;
        try {
            const files = fs.readdirSync(sourceDir);
            const gaiFiles = files.filter(f => {
                const stat = fs.statSync(path.join(sourceDir, f));
                return stat.isFile() && f.toLowerCase().includes('gai');
            });
            console.log('[Scan] Found gai files:', gaiFiles);
            safeSend(sender, 'classify:log', { type: 'info', msg: `扫描完成，找到 ${gaiFiles.length} 个包含 gai 的文件`, time: new Date().toLocaleTimeString() });
            safeSend(sender, 'scan-fix-items-result', gaiFiles);
        } catch (err) {
            console.error('[Scan] Error:', err);
            safeSend(sender, 'classify:log', { type: 'error', msg: `扫描失败: ${err.message}`, time: new Date().toLocaleTimeString() });
        }
    });

    ipcMain.on('classify:fix-items', async (event, { sourceDir, targetDir, userName, fileNames }) => {
        const sender = event.sender;
        const sendLog = (type, msg) => {
            safeSend(sender, 'classify:log', { type, msg, time: new Date().toLocaleTimeString() });
        };

        sendLog('info', `开始归类修改项 (${fileNames.length} 个文件)...`);
        let successCount = 0;
        let skipCount = 0;

        for (const fileName of fileNames) {
            const filePath = path.join(sourceDir, fileName);
            if (!fs.existsSync(filePath)) {
                sendLog('warn', `${fileName} 不存在，可能已被重命名或移动`);
                skipCount++;
                continue;
            }
            const result = await classifyFile(filePath, targetDir, userName, sendLog, true);
            if (result.success) {
                successCount++;
            } else {
                skipCount++;
            }
        }

        sendLog('success', `归类修改项完成: ${successCount} 个成功, ${skipCount} 个跳过`);
        // 通知前端清理已成功归类的项
        safeSend(sender, 'classify:fix-items-done', { successCount });
    });

    ipcMain.on('update-product-rules', (event, rules) => {
        updateProductRules(rules);
    });

    ipcMain.on('open-folder', (event, dir) => {
        shell.openPath(dir);
    });

    ipcMain.on('open-folder-select-first-gai', (event, dir) => {
        try {
            const files = fs.readdirSync(dir);
            const gaiFile = files.find(f => {
                const stat = fs.statSync(path.join(dir, f));
                return stat.isFile() && f.toLowerCase().includes('gai');
            });
            if (gaiFile) {
                shell.showItemInFolder(path.join(dir, gaiFile));
            } else {
                shell.openPath(dir);
            }
        } catch (err) {
            shell.openPath(dir);
        }
    });

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    // Windows 下不退出，保持托盘运行
    // macOS 下也不退出
});

app.on('will-quit', () => {
    if (activeWatcher) {
        activeWatcher.stop();
        activeWatcher = null;
    }
    if (classifyWatcher) {
        classifyWatcher.stop();
        classifyWatcher = null;
    }
    stopTemplateProcess();
    cleanupTemplateProcess();
});
