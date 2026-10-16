# AI in QuillLattice

The **AI** tab is optional. Nothing is sent anywhere until you run an AI command, and every AI command is an ordinary command: it can go on the Quick Access Bar, on a custom tab, on a shortcut, or in a macro.

## Setting up

**AI › Set Up AI** opens a wizard that:
1. lists providers grouped as *Free · on device*, *Free · your computer*, *Free tier*, *Free trial* and *Paid*;
2. links to the provider's key page and checks the key with a test request;
3. fetches the provider's model list so you can pick an interactive model and a cheaper "fast" model.

You can add several providers and switch with **Active AI**. **Providers & Models** edits them; **AI & Privacy** shows what each command sends.

### Providers

| Group | Providers |
|---|---|
| On device | QuillLattice Local AI (see below); Chrome's built-in model on desktop Chrome |
| Your computer | Ollama (`http://localhost:11434/v1`), LM Studio (`http://localhost:1234/v1`) |
| Free tiers | OpenRouter (free models), Google Gemini, Groq, Mistral, Cerebras, SambaNova, Z.ai, Cohere, NVIDIA NIM, Cloudflare Workers AI |
| Paid keys | OpenAI, Anthropic, DeepSeek, xAI, Perplexity, Together, Fireworks, Hugging Face Inference Providers, Moonshot, Alibaba Qwen, Azure OpenAI / AI Foundry, Amazon Bedrock, any OpenAI-compatible service |

Free-tier terms are set by each provider and change; the wizard links to the provider's current page rather than promising limits.

### On-device model

QuillLattice Local AI runs a small open instruction model with Transformers.js in a Web Worker, on WebGPU when available, otherwise WebAssembly on the CPU. The model is downloaded once (from the Hugging Face Hub) and then works offline. Choices:

| Model | Download | Notes |
|---|---|---|
| Qwen2.5 0.5B Instruct (recommended) | ≈ 480–510 MB | Best quality for its size; many languages |
| SmolLM2 360M Instruct | ≈ 270–365 MB | Fastest on older phones; English-centric |
| Gemma 3 270M IT | ≈ 270–545 MB | Simple rewrites and summaries |
| Llama 3.2 1B Instruct | ≈ 1.24 GB | Needs WebGPU |

Small models are good at rewriting, classification and extraction; for writing non-trivial formulas a hosted or local-server model is noticeably better. **Delete Downloaded Models** frees the space.

## Commands

**Formulas:** Write Formula (the result is parsed and evaluated by the formula engine before it is offered; formulas that fail to parse are sent back to the model once with the error), Explain Formula, Fix Formula, Simplify Formula, Convert Formula (to and from other spreadsheet dialects), Review Formulas.

**Data:** Analyze Data, Ask About Data, Fill Column with AI, Fill by Example, Categorize, Extract, Clean Data, Find Anomalies, Describe Columns, Suggest Headers, Generate Table, Sample Data, Suggest Chart, Suggest PivotTable, Summarize, Explain This Sheet.

**Building:** Build Template, Conditional Formatting (from a description), Drop-Down List, Write Macro (JavaScript for the macro editor; shown for review, never run automatically).

**Cell text:** Translate, Rewrite, Proofread, Summarize Each Cell, Extract Keywords, Sentiment, Detect Language.

**Other:** AI Assistant (chat pane with the selection and sheet structure as context), Prompt Library and My Prompts (saved prompts with input = selection/sheet/none and output = replace/insert/pane/comment/new workbook), Stop AI.

**Quick Analysis** (formatting, charts, totals, tables and sparklines for the selection) is on the AI tab but does not use AI.

Results that change the workbook are applied as one undoable step.

## Privacy and keys

- Only what a command works on is sent: the active cell, the selection or current region, or for whole-sheet commands the used range as text (truncated to a size limit, smaller for the on-device model). Formulas are included only for formula and whole-sheet commands. **AI & Privacy** summarizes this.
- API keys are never stored in settings, drafts, exported settings or workbooks. On Android they are encrypted with a key held in the Android Keystore; on desktop they are encrypted with the OS keychain (Electron `safeStorage`: Keychain, DPAPI or libsecret). In a plain browser they are kept in IndexedDB for this site only, which is not encrypted at rest; use the on-device model or a local server if that matters.
