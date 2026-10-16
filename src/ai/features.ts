/**
 * AI features for spreadsheets, as data. Collected from a survey of how spreadsheet tools expose
 * AI (desktop and web spreadsheets, notebook-style data tools, spreadsheet add-ins and data
 * cleaning services) and named generically: no product's feature branding is used.
 *
 * Each feature says what it reads (input), what it produces (output) and how to prompt. The runner
 * in src/commands/ai.ts turns every feature into a command.
 */

export type AiInput = 'cell' | 'selection' | 'region' | 'sheet' | 'none';

/**
 * formula    a formula, validated with the calc engine and previewed before it is inserted
 * pane       an answer streamed into the AI pane
 * rows       one answer per cell or row, in batches, previewed and written as one undo step
 * table      a table (JSON) previewed, then written at the active cell or on a new sheet
 * chart      a chart definition added to the sheet
 * cf         conditional formatting rules
 * validation a drop-down list (data validation)
 * macro      a JavaScript macro module
 * analysis   insights plus a suggested chart and PivotTable
 * anomalies  rows flagged as unusual, highlighted
 * headers    a header row for unlabeled data
 */
export type AiOutput = 'formula' | 'pane' | 'rows' | 'table' | 'chart' | 'cf' | 'validation' | 'macro' | 'analysis' | 'anomalies' | 'headers';

export interface RowSpec {
  /** Each item is one cell of the selection, or one row (all columns of the data region). */
  item: 'cell' | 'row';
  /** replace: overwrite the cells; column: one new column; columns: several new columns (keys). */
  mode: 'replace' | 'column' | 'columns';
  /** Only text cells are sent (replace mode). */
  textOnly?: boolean;
  header?: (arg?: string, extra?: string) => string;
  keys?: (arg?: string, extra?: string) => string[];
  /** The per-item instruction. */
  instruction: (arg?: string, extra?: string) => string;
  /** Answers are numbers (written as numbers). */
  numeric?: boolean;
}

export interface FeatureCtx {
  /** Data or cell context assembled by the runner (tables as tab-separated text). */
  text: string;
  arg?: string;
  extra?: string;
  lang: string;
  sheet?: string;
  address?: string;
}

export interface AiFeature {
  id: string;
  label: string;
  icon: string;
  group: string;
  tip: string;
  input: AiInput;
  output: AiOutput;
  tier?: 'fast' | 'quality';
  temperature?: number;
  ask?: { title: string; label: string; placeholder?: string; multiline?: boolean; optional?: boolean };
  options?: { value: string; label: string }[];
  keys?: string[];
  popular?: boolean;
  keywords?: string;
  origin?: string;
  system?: string;
  maxTokens?: number;
  rows?: RowSpec;
  prompt: (ctx: FeatureCtx) => string;
}

export const BASE_SYSTEM = [
  'You are the assistant built into QuillLattice, a spreadsheet application compatible with Excel formulas.',
  'Follow the requested output format exactly. Return only what is asked for: no preamble, no explanation unless asked, no closing remarks.',
  'Formulas use Excel syntax with commas between arguments, English function names, and A1 references. Dynamic arrays (FILTER, SORT, UNIQUE, XLOOKUP, LET, LAMBDA, SEQUENCE) are supported.',
  'Never invent data that is not in the sheet unless asked to generate sample data. Keep names, numbers and dates exactly as given unless asked to change them.',
  'Unless told which variety of English to use, keep the variety (American, British, Canadian, Australian…) of the data and its spelling.',
].join(' ');

export const TONES = ['Professional', 'Formal', 'Casual', 'Friendly', 'Concise', 'Persuasive', 'Neutral', 'Enthusiastic', 'Empathetic', 'Plain language'];
export const LENGTHS = ['Shorter', 'Much shorter (one line)', 'Longer', 'Simpler words', 'More specific'];
export const LANGUAGES = ['English', 'English (American)', 'English (British)', 'English (Canadian)', 'English (Australian)', 'Spanish', 'French', 'German', 'Italian', 'Portuguese', 'Dutch', 'Swedish', 'Norwegian', 'Danish', 'Finnish', 'Polish', 'Czech', 'Hungarian', 'Romanian', 'Greek', 'Turkish', 'Russian', 'Ukrainian', 'Hebrew', 'Arabic', 'Persian', 'Hindi', 'Bengali', 'Urdu', 'Chinese (Simplified)', 'Chinese (Traditional)', 'Japanese', 'Korean', 'Vietnamese', 'Thai', 'Indonesian', 'Malay', 'Filipino', 'Swahili', 'Latin', 'Irish', 'Welsh'];
export const ENTITIES: Record<string, { label: string; keys: string[] }> = {
  names: { label: 'Names (first, last)', keys: ['First name', 'Last name'] },
  people: { label: 'People mentioned', keys: ['People'] },
  dates: { label: 'Dates', keys: ['Date'] },
  amounts: { label: 'Amounts and currency', keys: ['Amount', 'Currency'] },
  emails: { label: 'Email addresses', keys: ['Email'] },
  phones: { label: 'Phone numbers', keys: ['Phone'] },
  addresses: { label: 'Address parts', keys: ['Street', 'City', 'State/Region', 'Postal code', 'Country'] },
  companies: { label: 'Companies and organizations', keys: ['Organization'] },
  places: { label: 'Places', keys: ['Place'] },
  products: { label: 'Products and quantities', keys: ['Product', 'Quantity'] },
};
export const CLEANUPS: Record<string, string> = {
  all: 'Clean each value: fix casing, spacing, obvious typos and inconsistent formats, without changing its meaning',
  names: 'Normalize person names: proper capitalization (McDonald, O\'Brien, van der Berg), "First Last" order unless the column clearly uses "Last, First", remove stray titles only if inconsistent',
  companies: 'Normalize company names to one consistent spelling (e.g. "IBM Corp.", "I.B.M." → "IBM"), keeping legal suffixes consistent',
  addresses: 'Normalize postal addresses: consistent abbreviations (Street → St, Avenue → Ave), capitalization and spacing, postal codes in the right format',
  dates: 'Rewrite each date in ISO format YYYY-MM-DD; keep text that is not a date unchanged; for partial dates use YYYY-MM or YYYY',
  phones: 'Normalize phone numbers to international format (+1 555 123 4567 style) when the country is clear, otherwise a consistent national format',
  casing: 'Fix casing only: Title Case for names and titles, sentence case for sentences, keep acronyms in capitals',
  categories: 'Map each value to one consistent label: merge spelling variants, plurals, abbreviations and typos of the same category into the most common spelling in the list',
  units: 'Normalize units and measurements to one consistent unit and format (e.g. "5 lbs", "5lb", "five pounds" → "5 lb")',
};
export const SUMMARY_KINDS = ['Brief overview', 'Key findings (bullets)', 'Executive summary', 'Trends over time', 'Notable outliers'];
export const TEMPLATES = ['Monthly budget', 'Expense report', 'Invoice', 'Project plan with dates', 'Task tracker', 'Inventory list', 'Sales pipeline', 'Loan amortization schedule', 'Grade book', 'Event guest list', 'Workout log', 'Family tree / genealogy research log', 'Habit tracker', 'Time sheet', 'Savings goal tracker'];
export const FORMULA_DIALECTS = ['Convert from Google Sheets', 'Convert from LibreOffice / semicolons', 'Use dynamic arrays (modern functions)', 'Replace VLOOKUP/HLOOKUP with XLOOKUP', 'Make compatible with older spreadsheets', 'Use LET to name repeated parts'];

const opt = (xs: string[]) => xs.map((x) => ({ value: x, label: x }));
const langNote = (lang: string) => (lang && lang !== 'auto' ? ` Write any text in ${lang}.` : '');

export const FORMULA_RULES = 'Respond with the formula on the first line, starting with "=", and nothing else on that line. On the next lines, give a one-sentence explanation.';
export const JSON_ONLY = 'Respond with JSON only, no code fences.';

export const FEATURES: AiFeature[] = [
  // ---------------------------------------------------------------- Formulas
  {
    id: 'ai.writeFormula', label: 'Write Formula…', icon: 'function', group: 'Formulas', input: 'region', output: 'formula', tier: 'quality', temperature: 0.1, popular: true, keys: ['Mod-Alt-f'],
    tip: 'Describe a calculation in plain words; the AI writes the formula, which is checked before you insert it', keywords: 'formula generator natural language',
    ask: { title: 'Write Formula', label: 'What should the formula calculate?', placeholder: 'e.g. total sales for the East region in 2024, or the number of days between the dates in B and C', multiline: true },
    prompt: ({ text, extra, address }) => `Write one spreadsheet formula for cell ${address} that does this: ${extra}\n\n${FORMULA_RULES}\n\nThe sheet:\n${text}`,
  },
  {
    id: 'ai.explainFormula', label: 'Explain Formula', icon: 'help_center', group: 'Formulas', input: 'cell', output: 'pane', tier: 'fast', temperature: 0.2, popular: true,
    tip: 'Explain the formula in the active cell step by step', keywords: 'explain formula understand',
    prompt: ({ text, address }) => `Explain the formula in cell ${address} in plain language for a spreadsheet user: first one sentence saying what it returns, then each part step by step (Markdown bullets). Mention any possible problems.\n\n${text}`,
  },
  {
    id: 'ai.fixFormula', label: 'Fix Formula', icon: 'build', group: 'Formulas', input: 'cell', output: 'formula', tier: 'quality', temperature: 0.1,
    tip: 'Find out why the active cell shows an error (or the wrong result) and propose a corrected formula', keywords: 'error #N/A #VALUE #REF repair debug',
    ask: { title: 'Fix Formula', label: 'What is wrong or what should it do? (optional)', optional: true },
    prompt: ({ text, address, extra }) => `The formula in cell ${address} returns an error or a wrong result.${extra ? ` The user says: ${extra}.` : ''} Write a corrected formula.\n\n${FORMULA_RULES} The explanation should say what was wrong.\n\n${text}`,
  },
  {
    id: 'ai.simplifyFormula', label: 'Simplify Formula', icon: 'compress', group: 'Formulas', input: 'cell', output: 'formula', tier: 'quality', temperature: 0.1,
    tip: 'Rewrite the formula in the active cell so it is shorter and easier to read, with the same results', keywords: 'refactor nested if simplify',
    prompt: ({ text, address }) => `Rewrite the formula in cell ${address} so it is simpler and easier to read (fewer nested IFs, IFS/SWITCH/XLOOKUP/LET where useful) and returns exactly the same results.\n\n${FORMULA_RULES}\n\n${text}`,
  },
  {
    id: 'ai.convertFormula', label: 'Convert Formula', icon: 'transform', group: 'Formulas', input: 'cell', output: 'formula', tier: 'quality', temperature: 0.1, options: opt(FORMULA_DIALECTS),
    tip: 'Convert a formula from another spreadsheet, switch to modern functions, or make it work in older spreadsheets', keywords: 'google sheets libreoffice convert xlookup compatibility',
    prompt: ({ text, address, arg }) => `Rewrite the formula in cell ${address}: ${arg}. Keep the same results.\n\n${FORMULA_RULES}\n\n${text}`,
  },
  {
    id: 'ai.reviewFormulas', label: 'Review Formulas', icon: 'fact_check', group: 'Formulas', input: 'sheet', output: 'pane', tier: 'quality', temperature: 0.2,
    tip: 'Look for inconsistent formulas, hard-coded numbers, ranges that miss rows and other likely mistakes on this sheet', keywords: 'audit check errors inconsistent',
    prompt: ({ text }) => `Review the formulas on this sheet for likely mistakes: a formula that differs from its neighbors, ranges that stop short of the data, hard-coded numbers that should be references, absolute/relative reference errors, division by zero risks. List each finding with the cell address and a suggested fix (Markdown bullets). If everything looks fine, say so.\n\n${text}`,
  },

  // ---------------------------------------------------------------- Data
  {
    id: 'ai.analyzeData', label: 'Analyze Data', icon: 'insights', group: 'Data', input: 'region', output: 'analysis', tier: 'quality', temperature: 0.3, popular: true,
    tip: 'Insights about the selected data, with a suggested chart and PivotTable you can add with one click', keywords: 'insights explore analyze trends',
    prompt: ({ text, lang }) => `Analyze this table.${langNote(lang)} ${JSON_ONLY} Use this shape:\n{"insights": ["5 to 8 short, specific findings with numbers"], "questions": ["3 follow-up questions worth asking"], "chart": {"type": "column|bar|line|pie|scatter|area", "title": "…", "x": "header of the category column", "y": ["header(s) of value columns"]}, "pivot": {"rows": ["header"], "columns": [], "values": [{"field": "header", "fn": "sum|count|average|max|min"}]}}\nUse only headers that exist in the table.\n\n${text}`,
  },
  {
    id: 'ai.askData', label: 'Ask About Data…', icon: 'quiz', group: 'Data', input: 'region', output: 'pane', tier: 'quality', temperature: 0.2,
    tip: 'Ask a question about the selected data; the answer appears in the AI pane', keywords: 'question answer data chat',
    ask: { title: 'Ask About Data', label: 'Your question', placeholder: 'e.g. Which product grew fastest? What is the average order by region?', multiline: true },
    prompt: ({ text, extra, lang }) => `Answer the question about this table.${langNote(lang)} Show the numbers you used. If a formula would answer it in the sheet, give it too.\n\nQuestion: ${extra}\n\n${text}`,
  },
  {
    id: 'ai.fillColumn', label: 'Fill Column with AI…', icon: 'auto_fix_high', group: 'Data', input: 'region', output: 'rows', tier: 'fast', temperature: 0.2, popular: true,
    tip: 'Fill a new column by an instruction applied to every row (e.g. "country of the city in column B")', keywords: 'generate column per row enrich',
    ask: { title: 'Fill Column with AI', label: 'For each row, what should the new column contain?', placeholder: 'e.g. The country of the city in the City column; a one-line product description', multiline: true },
    rows: { item: 'row', mode: 'column', header: (_a, extra) => (extra || 'AI').split(/[.,;:]/)[0].slice(0, 40), instruction: (_a, extra) => extra || '' },
    prompt: () => '',
  },
  {
    id: 'ai.fillByExample', label: 'Fill by Example', icon: 'model_training', group: 'Data', input: 'region', output: 'rows', tier: 'fast', temperature: 0,
    tip: 'Type a few results in an empty column next to your data; the AI fills the rest of the column the same way', keywords: 'flash fill pattern examples',
    origin: 'Pattern fill in spreadsheets, extended with AI',
    rows: { item: 'row', mode: 'column', instruction: () => 'Produce the value for the target column, following the pattern of the examples exactly (same format, casing and punctuation).' },
    prompt: () => '',
  },
  {
    id: 'ai.classify', label: 'Categorize…', icon: 'label', group: 'Data', input: 'region', output: 'rows', tier: 'fast', temperature: 0,
    tip: 'Put each row (or cell) into one of your categories, or let the AI choose categories', keywords: 'classify categorize tag label',
    ask: { title: 'Categorize', label: 'Categories (comma-separated), or leave blank to let the AI choose', placeholder: 'e.g. Food, Travel, Office, Utilities, Other', optional: true },
    rows: { item: 'row', mode: 'column', header: () => 'Category', instruction: (_a, extra) => (extra ? `Classify the row into exactly one of these categories: ${extra}. Answer with the category name only.` : 'Classify the row into a short category label (one to three words). Use a small, consistent set of categories across all rows.') },
    prompt: () => '',
  },
  {
    id: 'ai.extract', label: 'Extract', icon: 'data_array', group: 'Data', input: 'selection', output: 'rows', tier: 'fast', temperature: 0, options: Object.entries(ENTITIES).map(([value, e]) => ({ value, label: e.label })),
    tip: 'Pull names, dates, amounts, emails, phone numbers or address parts out of text into new columns', keywords: 'extract entities parse names dates amounts address',
    rows: { item: 'cell', mode: 'columns', keys: (arg) => ENTITIES[arg || 'names']?.keys || ['Value'], instruction: (arg) => `Extract ${ENTITIES[arg || 'names']?.label.toLowerCase() || 'the values'} from the text. Use an empty string when a field is not present. Dates as YYYY-MM-DD; amounts as plain numbers.` },
    prompt: () => '',
  },
  {
    id: 'ai.cleanData', label: 'Clean Data', icon: 'mop', group: 'Data', input: 'selection', output: 'rows', tier: 'fast', temperature: 0, options: Object.keys(CLEANUPS).map((k) => ({ value: k, label: { all: 'Everything', names: 'Person names', companies: 'Company names', addresses: 'Addresses', dates: 'Dates', phones: 'Phone numbers', casing: 'Casing only', categories: 'Merge spelling variants', units: 'Units and measurements' }[k]! })),
    tip: 'Normalize names, addresses, dates, phone numbers, casing or spelling variants; you review every change before it is applied', keywords: 'normalize standardize clean dedupe variants',
    rows: { item: 'cell', mode: 'replace', instruction: (arg) => `${CLEANUPS[arg || 'all'] || CLEANUPS.all}. If a value needs no change, return it exactly as it is.` },
    prompt: () => '',
  },
  {
    id: 'ai.detectAnomalies', label: 'Find Anomalies', icon: 'troubleshoot', group: 'Data', input: 'region', output: 'anomalies', tier: 'quality', temperature: 0.1,
    tip: 'Flag rows that look unusual (outliers, typos, impossible values, inconsistent entries) and explain why', keywords: 'outliers anomalies errors unusual detect',
    prompt: ({ text }) => `Find rows in this table that look wrong or unusual: statistical outliers, impossible values (negative ages, future birth dates), likely typos, values inconsistent with the rest of the column, duplicates with small differences. ${JSON_ONLY} Return an array of {"row": <row number as shown>, "column": "column letter or header", "reason": "short explanation"}. Return [] if nothing stands out.\n\n${text}`,
  },
  {
    id: 'ai.dataDictionary', label: 'Describe Columns', icon: 'menu_book', group: 'Data', input: 'region', output: 'table', tier: 'quality', temperature: 0.2,
    tip: 'A data dictionary: what each column means, its type, units, example values and quality notes, on a new sheet', keywords: 'data dictionary documentation schema',
    prompt: ({ text }) => `Write a data dictionary for this table. ${JSON_ONLY} Return an array of objects with the keys "Column", "Description", "Type", "Units or format", "Example", "Quality notes" (one object per column, in order).\n\n${text}`,
  },
  {
    id: 'ai.suggestHeaders', label: 'Suggest Headers', icon: 'title', group: 'Data', input: 'region', output: 'headers', tier: 'fast', temperature: 0.2,
    tip: 'Propose column headers for data that has none, or clearer names for cryptic ones', keywords: 'column names headers label',
    prompt: ({ text }) => `Suggest a short, clear header (1 to 4 words, Title Case) for each column of this data, in order. ${JSON_ONLY} Return {"hasHeaders": true|false (whether the first row already is a header row), "headers": ["…"]}.\n\n${text}`,
  },

  // ---------------------------------------------------------------- Generate
  {
    id: 'ai.generateTable', label: 'Generate Table…', icon: 'table_chart', group: 'Generate', input: 'none', output: 'table', tier: 'quality', temperature: 0.5, popular: true,
    tip: 'Describe a table; the AI creates it with headers (and formulas where useful)', keywords: 'create table generate list',
    ask: { title: 'Generate Table', label: 'Describe the table', placeholder: 'e.g. The 50 US states with capital, population and admission year; a 12-month savings plan', multiline: true },
    prompt: ({ extra, lang }) => `Create this table: ${extra}.${langNote(lang)} ${JSON_ONLY} Return an array of objects, one per row, with the same keys in the same order (the keys become the headers). Values are numbers, text, or Excel formulas starting with "=" that refer to cells assuming the header is in row 1 and column A is the first key. Mark uncertain facts with "(verify)".`,
  },
  {
    id: 'ai.sampleData', label: 'Sample Data…', icon: 'dataset', group: 'Generate', input: 'none', output: 'table', tier: 'fast', temperature: 0.9,
    tip: 'Realistic made-up data for testing and teaching (customers, orders, grades…)', keywords: 'fake test data mock dummy',
    ask: { title: 'Sample Data', label: 'What data, and how many rows?', placeholder: 'e.g. 30 customer orders with date, customer, product, quantity, unit price', multiline: true },
    prompt: ({ extra }) => `Generate realistic but fictional sample data: ${extra}. Use made-up names (not real private people) and plausible values with some variety. ${JSON_ONLY} Return an array of objects with the same keys in the same order. Dates as YYYY-MM-DD.`,
  },
  {
    id: 'ai.suggestChart', label: 'Suggest Chart', icon: 'add_chart', group: 'Generate', input: 'region', output: 'chart', tier: 'fast', temperature: 0.2, popular: true,
    tip: 'Pick the best chart for the selected data and add it to the sheet', keywords: 'recommended chart visualize',
    prompt: ({ text }) => `Choose the single best chart for this data. ${JSON_ONLY} Return {"type": "column|stackedColumn|bar|line|lineMarkers|area|pie|doughnut|scatter|radar", "title": "…", "x": "header of the category (or x) column", "y": ["header(s) of the value columns"], "reason": "one sentence"}. Use only headers that exist.\n\n${text}`,
  },
  {
    id: 'ai.summarize', label: 'Summarize', icon: 'summarize', group: 'Generate', input: 'region', output: 'pane', tier: 'quality', temperature: 0.3, options: opt(SUMMARY_KINDS),
    tip: 'Summarize the selected data or the sheet in words', keywords: 'summary overview narrative report',
    prompt: ({ text, arg, lang }) => `Write a ${arg || 'Brief overview'} of this data for a reader who has not seen it.${langNote(lang)} Use specific numbers. Markdown is fine.\n\n${text}`,
  },
  {
    id: 'ai.buildTemplate', label: 'Build Template', icon: 'dashboard_customize', group: 'Generate', input: 'none', output: 'table', tier: 'quality', temperature: 0.4, options: opt(TEMPLATES),
    tip: 'A ready-to-use sheet (budget, invoice, tracker…) with headers, sample rows and working formulas, on a new sheet', keywords: 'template budget invoice tracker planner',
    ask: { title: 'Build Template', label: 'Anything specific? (optional)', placeholder: 'e.g. in euros, with a column for receipts', optional: true },
    prompt: ({ arg, extra, lang }) => `Design a spreadsheet template: ${arg}${extra ? ` (${extra})` : ''}.${langNote(lang)} ${JSON_ONLY} Return an array of objects, one per row (keys become headers in row 1, column A is the first key), with 5 to 12 example rows and a final totals row where it makes sense. Use Excel formulas starting with "=" for calculated columns and totals, referring to cells by A1 address.`,
  },
  {
    id: 'ai.cfRules', label: 'Conditional Formatting…', icon: 'format_color_fill', group: 'Generate', input: 'selection', output: 'cf', tier: 'quality', temperature: 0.1,
    tip: 'Describe how cells should be highlighted; the AI writes the conditional formatting rules for the selection', keywords: 'highlight rules color format condition',
    ask: { title: 'Conditional Formatting with AI', label: 'How should the selection be highlighted?', placeholder: 'e.g. red when overdue, green when Status is Done, color scale on the totals', multiline: true },
    prompt: ({ text, extra, address }) => `Write conditional formatting rules for the range ${address}: ${extra}\n${JSON_ONLY} Return an array of rules. Each rule: {"type": "cellIs|expression|colorScale|dataBar|iconSet|top10|aboveAverage|duplicateValues|uniqueValues|containsText|containsBlanks|timePeriod", "range": "A1-style range (optional; default is the selection)", "operator": "between|notBetween|equal|notEqual|greaterThan|lessThan|greaterThanOrEqual|lessThanOrEqual (cellIs only)", "formulas": ["formula or value without =; for expression rules a formula relative to the top-left cell of the range"], "text": "for containsText", "rank": 10, "percent": false, "bottom": false, "above": true, "timePeriod": "today|yesterday|last7Days|thisWeek|lastWeek|thisMonth|lastMonth", "colors": ["#hex", "#hex", "#hex"] (colorScale), "barColor": "#hex", "iconSet": "3Arrows|3TrafficLights1|3Flags|3Signs", "style": {"fill": "#hex", "color": "#hex", "bold": true}}.\n\nThe data:\n${text}`,
  },
  {
    id: 'ai.validationList', label: 'Drop-Down List…', icon: 'arrow_drop_down_circle', group: 'Generate', input: 'selection', output: 'validation', tier: 'fast', temperature: 0.3,
    tip: 'Describe the allowed choices; the AI writes the list and adds a drop-down (data validation) to the selection', keywords: 'data validation dropdown list choices',
    ask: { title: 'Drop-Down List with AI', label: 'What choices should the drop-down offer?', placeholder: 'e.g. US states; task statuses; T-shirt sizes', multiline: false },
    prompt: ({ extra, text }) => `List the choices for a drop-down: ${extra}. ${JSON_ONLY} Return {"items": ["…"], "title": "short input title", "message": "short input message"}. At most 100 items, no commas inside items.${text ? `\n\nExisting values in the column (keep their spelling):\n${text}` : ''}`,
  },
  {
    id: 'ai.writeMacro', label: 'Write Macro…', icon: 'code', group: 'Generate', input: 'none', output: 'macro', tier: 'quality', temperature: 0.2,
    tip: 'Describe a task; the AI writes a JavaScript macro for it and opens it in the Macro Editor for you to check and run', keywords: 'macro script automate javascript vba',
    ask: { title: 'Write Macro', label: 'What should the macro do?', placeholder: 'e.g. Bold every row whose Status column says Overdue; copy rows with amounts over 1000 to a new sheet', multiline: true },
    system: 'Macros are JavaScript run as the body of an async function with one variable, api. API: api.cell("B2") returns an object with .value (get/set; strings starting with = set formulas), .formula, .format (number format), .text, .setStyle({bold, italic, color, fill, numFmt}), .offset(rows, cols); api.range("A1:C10") has .values (2-D array get/set), .formulas, .rows, .cols, .clear(), .cell(i, j), .forEach((cell, i, j) => …); api.sheet(name) has .cell(), .range(), .usedRange (address or null), .activate(); api.sheets (names); api.activeSheet; api.addSheet(name); api.select(address); api.selection (address text); api.activeCell (address); api.exec(commandId, arg); api.msgBox(text); await api.inputBox(prompt, default); api.log(...). No DOM, no network, no imports.',
    prompt: ({ extra, text }) => `Write a macro module for this task: ${extra}\nDeclare one or more functions with "function Name() { … }" (use async functions when awaiting inputBox); the first function is the main macro. Add short comments. Return only JavaScript, no code fences.${text ? `\n\nThe active sheet for reference:\n${text}` : ''}`,
  },
  {
    id: 'ai.suggestPivot', label: 'Suggest PivotTable', icon: 'pivot_table_chart', group: 'Generate', input: 'region', output: 'pane', tier: 'fast', temperature: 0.3,
    tip: 'Ideas for PivotTables that answer useful questions about the data, with the fields to use', keywords: 'pivot table recommend summarize',
    prompt: ({ text }) => `Suggest three PivotTables that would answer useful questions about this data. For each: the question it answers, Rows field, Columns field (if any), Values field and function, and any filter. Use the exact headers. Markdown bullets.\n\n${text}`,
  },

  // ---------------------------------------------------------------- Text
  {
    id: 'ai.translateCells', label: 'Translate Cells', icon: 'translate', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0.1, options: opt(LANGUAGES),
    tip: 'Translate the text in the selected cells (review before it replaces them, or add a column)', keywords: 'translate language localize',
    rows: { item: 'cell', mode: 'replace', textOnly: true, header: (arg) => arg || 'Translation', instruction: (arg) => `Translate the text into ${arg || 'English'}. Keep numbers, codes, names and placeholders unchanged.` },
    prompt: () => '',
  },
  {
    id: 'ai.rewriteCells', label: 'Rewrite Cells', icon: 'edit_note', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0.5, options: opt([...TONES.map((t) => `${t} tone`), ...LENGTHS]),
    tip: 'Change the tone or length of the text in the selected cells', keywords: 'rewrite tone shorten paraphrase',
    rows: { item: 'cell', mode: 'replace', textOnly: true, instruction: (arg) => `Rewrite the text: ${arg || 'Professional tone'}. Keep the meaning, facts, names and numbers.` },
    prompt: () => '',
  },
  {
    id: 'ai.proofreadCells', label: 'Proofread Cells', icon: 'spellcheck', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0,
    tip: 'Fix spelling, grammar and punctuation in the selected text cells; review each change first', keywords: 'spelling grammar proofread typos',
    rows: { item: 'cell', mode: 'replace', textOnly: true, instruction: () => 'Correct spelling, grammar and punctuation only. Keep the wording, variety of English, names and codes. If nothing needs fixing, return the text unchanged.' },
    prompt: () => '',
  },
  {
    id: 'ai.summarizeCells', label: 'Summarize Each Cell', icon: 'short_text', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0.2,
    tip: 'A one-line summary of each long text cell (reviews, notes, descriptions) in the next column', keywords: 'summary tldr per row',
    rows: { item: 'cell', mode: 'column', textOnly: true, header: () => 'Summary', instruction: () => 'Summarize the text in one short sentence (at most 15 words).' },
    prompt: () => '',
  },
  {
    id: 'ai.keywords', label: 'Extract Keywords', icon: 'sell', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0.1,
    tip: 'Three to five keywords or tags for each text cell, in the next column', keywords: 'tags keywords topics',
    rows: { item: 'cell', mode: 'column', textOnly: true, header: () => 'Keywords', instruction: () => 'List 3 to 5 keywords or short topics for the text, separated by commas, lowercase.' },
    prompt: () => '',
  },
  {
    id: 'ai.sentiment', label: 'Sentiment', icon: 'sentiment_satisfied', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0,
    tip: 'Score each text cell from -1 (negative) to 1 (positive) with a label, in new columns', keywords: 'sentiment opinion positive negative reviews feedback',
    rows: { item: 'cell', mode: 'columns', textOnly: true, keys: () => ['Sentiment', 'Score'], instruction: () => 'Judge the sentiment of the text. "Sentiment" is Positive, Negative, Neutral or Mixed; "Score" is a number from -1 (very negative) to 1 (very positive) with two decimals.' },
    prompt: () => '',
  },
  {
    id: 'ai.detectLanguage', label: 'Detect Language', icon: 'language', group: 'Text', input: 'selection', output: 'rows', tier: 'fast', temperature: 0,
    tip: 'The language of each text cell, in the next column', keywords: 'language detect identify',
    rows: { item: 'cell', mode: 'column', textOnly: true, header: () => 'Language', instruction: () => 'Name the language of the text in English (e.g. "Spanish"). Use "Unknown" for codes or numbers.' },
    prompt: () => '',
  },
  {
    id: 'ai.explainSheet', label: 'Explain This Sheet', icon: 'school', group: 'Text', input: 'sheet', output: 'pane', tier: 'quality', temperature: 0.3,
    tip: 'A plain-language description of what this sheet does: its sections, inputs, calculations and results', keywords: 'explain document understand workbook onboarding',
    prompt: ({ text, lang }) => `Explain this worksheet to someone opening it for the first time.${langNote(lang)} Describe its purpose, the input areas, the main calculations (refer to cells) and where the results are. Markdown headings and bullets.\n\n${text}`,
  },
];

export function featureById(id: string): AiFeature | undefined {
  return FEATURES.find((f) => f.id === id);
}
